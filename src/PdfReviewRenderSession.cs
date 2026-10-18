using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace ExcelMerger
{
    /// <summary>
    /// Один системный рендерер на окно сравнения. Панели «Бок о бок», общий холст и
    /// полноэкранный режим рендерят через него по очереди (lock): документ читается в память
    /// и разбирается WinRT один раз, а не на каждый показ страницы, и две панели больше не
    /// соревнуются за PdfMemoryBudget собственными копиями одного файла — LRU рендерера сам
    /// вытесняет старый документ под новый. После простоя рендерер освобождается вместе с
    /// документами и их резервом бюджета (MemoryReleased будит ждущих).
    ///
    /// Рендерер однопоточен, но потоки разные: воркеры панелей и таймер простоя (MTA), а
    /// Release перед сравнением и Dispose окна — UI-поток (STA). Освобождать документы с
    /// любого из них безопасно не из-за общей квартиры, а потому что PdfDocument и
    /// InMemoryRandomAccessStream помечены MarshalingBehavior(Agile) — их можно вызывать из
    /// любого потока без маршалинга (learn.microsoft.com/uwp/api/windows.data.pdf.pdfdocument,
    /// learn.microsoft.com/uwp/api/windows.storage.streams.inmemoryrandomaccessstream).
    /// Поэтому достаточно взаимного исключения.
    /// </summary>
    internal sealed class PdfReviewRenderSession : IDisposable
    {
        internal static readonly TimeSpan DefaultIdle = TimeSpan.FromSeconds(20);

        // Страница показа — A4: три буфера декодирования (как считает PdfThumbnailRenderer)
        // и копия под подсветку. Соседняя панель «Бок о бок» держит свой показанный растр
        // шириной до PanelWidth. Документ резервирует 2 × файл.
        private const int PanelWidth = 1200;
        private const double PageAspect = 297.0 / 210.0;

        private readonly object _renderGate = new object();
        private readonly object _stateGate = new object();
        private readonly Timer _idleTimer;
        private readonly int _idleMilliseconds;
        // Документы, которые читал текущий рендерер: держит ли он чужие для этой работы.
        private readonly HashSet<string> _documents =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private PdfThumbnailRenderer _renderer;
        private bool _busy;
        private bool _releaseRequested;
        private bool _disposed;
        private int _creations;

        internal PdfReviewRenderSession() : this(DefaultIdle) { }

        internal PdfReviewRenderSession(TimeSpan idle)
        {
            _idleMilliseconds = (int)Math.Max(1, Math.Min(int.MaxValue, idle.TotalMilliseconds));
            _idleTimer = new Timer(delegate { Release(); }, null, Timeout.Infinite,
                Timeout.Infinite);
            ShowBudgetBytes = PdfMemoryBudget.LimitBytes;
        }

        /// <summary>Сколько раз создавался рендерер (наблюдаемость повторного использования).</summary>
        internal int RendererCreations { get { lock (_stateGate) return _creations; } }

        internal bool HasRenderer { get { lock (_stateGate) return _renderer != null; } }

        /// <summary>Бюджет, от которого считается «слишком большой»; по умолчанию — бюджет процесса.</summary>
        internal long ShowBudgetBytes { get; set; }

        /// <summary>
        /// Самый большой файл, страницу которого вообще можно показать шириной renderWidth:
        /// один документ (2 × файл) — других сессия при нехватке не держит (см. Use), — растр
        /// этого показа и показанная страница соседней панели. На x86 (128 МБ) панель (1200 px)
        /// показывает файлы до ~43 МБ, полный экран (1600 px) — до ~31 МБ; на x64 (768 МБ)
        /// оба предела больше предела сравнения 200 МБ.
        /// </summary>
        internal static long MaxShowableFileBytes(long budgetLimitBytes, int renderWidth)
        {
            // Шире предела рендера растр не бывает; зажим заодно не даёт сумме переполниться.
            int width = Math.Max(1, Math.Min(RasterBudget.MaxRenderDimension, renderWidth));
            int height = RasterBudget.ExpectedHeight(width, PageAspect);
            int panelHeight = RasterBudget.ExpectedHeight(PanelWidth, PageAspect);
            long raster = RasterBudget.BitmapWorkingSetBytes(width, height, 3) +
                PdfMemoryBudget.EstimateBitmapBytes(width, height) +
                PdfMemoryBudget.EstimateBitmapBytes(PanelWidth, panelHeight);
            long documents = budgetLimitBytes - raster;
            return documents <= 0 ? 0 : documents / 2;
        }

        /// <summary>
        /// Файл заведомо не помещается в бюджет показа этой ширины: пользователь увидит
        /// «слишком большой», а не «недоступно». Нечитаемый размер — не повод отказывать:
        /// решит сам рендер.
        /// </summary>
        internal bool IsTooLargeToShow(string path, int renderWidth)
        {
            try
            {
                return new FileInfo(path).Length >
                    MaxShowableFileBytes(ShowBudgetBytes, renderWidth);
            }
            catch { return false; }
        }

        /// <summary>
        /// Выполнить работу с рендерером под взаимным исключением; documentPath — документ,
        /// который она читает. Если работа вернула null (обычно растру не хватило бюджета),
        /// а рендерер держал документы других панелей, они отпускаются и работа повторяется
        /// один раз со свежим рендерером: иначе вторая панель «Бок о бок» оставалась бы
        /// «недоступной» до простоя рендерера. null — сессия закрыта или рендерер не
        /// создаётся. Таймер простоя стоит, пока идёт работа.
        /// </summary>
        internal T Use<T>(string documentPath, Func<PdfThumbnailRenderer, T> work) where T : class
        {
            if (work == null)
                return null;
            lock (_renderGate)
            {
                PdfThumbnailRenderer renderer;
                bool heldOthers;
                lock (_stateGate)
                {
                    if (_disposed)
                        return null;
                    heldOthers = _documents.Count >
                        (documentPath != null && _documents.Contains(documentPath) ? 1 : 0);
                    renderer = AttachRendererLocked(documentPath);
                    if (renderer == null)
                        return null;
                    _busy = true;
                    _releaseRequested = false;
                    _idleTimer.Change(Timeout.Infinite, Timeout.Infinite);
                }
                try
                {
                    T result = work(renderer);
                    if (result != null || !heldOthers)
                        return result;
                    renderer = Restart(documentPath);
                    return renderer == null ? null : work(renderer);
                }
                finally
                {
                    PdfThumbnailRenderer release = null;
                    lock (_stateGate)
                    {
                        _busy = false;
                        if (_disposed || _releaseRequested)
                        {
                            release = DetachRendererLocked();
                            _releaseRequested = false;
                        }
                        else
                            _idleTimer.Change(_idleMilliseconds, Timeout.Infinite);
                    }
                    if (release != null)
                        release.Dispose();
                }
            }
        }

        /// <summary>
        /// Свежий рендерер посреди работы: прежний закрывается вместе с документами и их
        /// резервом. Показанные растры панелей своих резервов не теряют. null — сессию
        /// закрыли посреди работы (рендерер отпустит выход из Use).
        /// </summary>
        private PdfThumbnailRenderer Restart(string documentPath)
        {
            PdfThumbnailRenderer previous;
            lock (_stateGate)
            {
                if (_disposed)
                    return null;
                previous = DetachRendererLocked();
            }
            if (previous != null)
                previous.Dispose();
            lock (_stateGate)
                return _disposed ? null : AttachRendererLocked(documentPath);
        }

        private PdfThumbnailRenderer AttachRendererLocked(string documentPath)
        {
            if (_renderer == null)
            {
                try { _renderer = new PdfThumbnailRenderer(); }
                catch { return null; }
                _creations++;
            }
            if (documentPath != null)
                _documents.Add(documentPath);
            return _renderer;
        }

        private PdfThumbnailRenderer DetachRendererLocked()
        {
            PdfThumbnailRenderer renderer = _renderer;
            _renderer = null;
            _documents.Clear();
            return renderer;
        }

        /// <summary>
        /// Отпустить рендерер с документами сейчас (или сразу после текущей работы). Окно
        /// зовёт это перед сравнением, чтобы растровое уточнение получило бюджет документов.
        /// </summary>
        internal void Release()
        {
            PdfThumbnailRenderer release = null;
            lock (_stateGate)
            {
                if (_busy)
                    _releaseRequested = true;
                else
                    release = DetachRendererLocked();
            }
            if (release != null)
                release.Dispose();
        }

        /// <summary>Не ждёт текущий рендер: рендерер закроет сам воркер, выходя из Use.</summary>
        public void Dispose()
        {
            lock (_stateGate)
            {
                if (_disposed)
                    return;
                _disposed = true;
            }
            _idleTimer.Dispose();
            Release();
        }
    }
}
