using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Переиспользуемая сетка миниатюр страниц PDF: фоновый рендер (Windows.Data.Pdf),
    /// масштаб (Ctrl+колесо и SetTileWidth), мультивыбор, при AllowReorder —
    /// перетаскивание, буфер страниц (вырезать/копировать/вставить) с кареткой вставки
    /// и автопрокруткой у краёв, при AllowRotate — поворот выбранных страниц; дроп
    /// PDF-файлов прямо на сетку (с позицией вставки). Модель порядка страниц держит
    /// владелец (форма): сетка мутирует только Rotation ссылок (общих с моделью), а
    /// перестановки/вставки запрашивает событиями. Общая для «Объединения»,
    /// «Разделения» и «PDF → Word» (DRY).
    ///
    /// Память: отрендеренные страницы живут в LRU-кэше, ёмкость которого посчитана из
    /// байтового бюджета (на x86 — вдвое меньше): большой документ не накапливает
    /// сотни мегабайт, вытесненная страница перерендеривается при следующем показе.
    /// </summary>
    public class PdfPageGrid : UserControl
    {
        private const string PlaceholderKey = "__ph";
        private const int EnqueueBuffer = 16;   // докачивать миниатюры чуть за пределами видимого
        private const int DragEdgePx = 28;      // зона автопрокрутки у верх/низ края при перетаскивании
        private static readonly long PageCacheBudget =
            IntPtr.Size == 8 ? 192L << 20 : 48L << 20; // x86: адресное пространство ~2 ГБ

        /// <summary>ListView, извещающий о прокрутке — для ленивого рендера видимых страниц.</summary>
        private sealed class ScrollList : ListView
        {
            public event EventHandler Scrolled;
            protected override void WndProc(ref Message m)
            {
                base.WndProc(ref m);
                const int WM_VSCROLL = 0x115, WM_MOUSEWHEEL = 0x20A, WM_KEYUP = 0x101;
                if (m.Msg == WM_VSCROLL || m.Msg == WM_MOUSEWHEEL || m.Msg == WM_KEYUP)
                {
                    EventHandler h = Scrolled;
                    if (h != null) h(this, EventArgs.Empty);
                }
            }
        }

        /// <summary>Отрендеренная страница в LRU: ключ хранится в значении, чтобы вытеснение знало, чьи плитки снимать.</summary>
        private sealed class CachedPage
        {
            public string Key;
            public Bitmap Bmp;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        private const int LVM_SCROLL = 0x1014; // прокрутка LargeIcon на dy пикселей

        private readonly ScrollList _list = new ScrollList();
        private System.Windows.Forms.Timer _visibleTimer;
        private readonly LruCache<CachedPage> _pageCache;
        private readonly int _renderWidth; // физические пиксели с учётом DPI; постоянен на жизнь сетки
        private bool _shuttingDown;        // Dispose: вытеснение не трогает UI, только освобождает bitmap
        // Ключи страниц, показываемых сейчас (обновляется в SetPages, только UI-поток).
        // Поздний результат рендера уже снятой страницы отбрасывается по этому набору.
        private HashSet<string> _currentKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // Элементы списка по ключу СТРАНИЦЫ (без поворота): готовый рендер проставляется
        // адресно, без прохода по всем элементам (иначе рендер документа — O(n²)). Только UI-поток.
        private readonly Dictionary<string, List<ListViewItem>> _itemsByKey =
            new Dictionary<string, List<ListViewItem>>(StringComparer.OrdinalIgnoreCase);
        private ImageList _thumbs;
        private int _tileWidth = ThumbZoom.DefaultWidth;

        private readonly object _qLock = new object();
        private readonly Queue<PdfPageRef> _thumbQueue = new Queue<PdfPageRef>();
        private readonly HashSet<string> _thumbRequested = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly ManualResetEventSlim _thumbSignal = new ManualResetEventSlim(false);
        private Thread _thumbThread;
        private volatile bool _thumbStop;

        // Буфер страниц окна (не системный клипборд: ссылки имеют смысл только здесь).
        // Вырезанные до вставки остаются на месте и приглушаются (как в Проводнике).
        private readonly List<PdfPageRef> _clipboard = new List<PdfPageRef>();
        private bool _clipIsCut;
        private readonly HashSet<PdfPageRef> _cutRefs = new HashSet<PdfPageRef>();

        // Каретка вставки: клик в зазор между плитками ставит её (показывается штатным
        // InsertionMark). Сбрасывается при любой пересборке содержимого.
        private int _caretIndex = -1;
        private bool _caretAfter;

        // Автопрокрутка при перетаскивании у края сетки.
        private System.Windows.Forms.Timer _dragScrollTimer;
        private int _dragScrollDir; // -1 вверх, +1 вниз, 0 — нет
        private bool _dragHasFiles; // в текущей drag-сессии тянут PDF-файлы (решено в DragEnter)

        private ContextMenuStrip _menu;
        private ToolStripMenuItem _miCut, _miCopy, _miPaste, _miRotateRight, _miRotateLeft, _miDelete, _miGoTo;

        /// <summary>Разрешить перетаскивание/буфер страниц для смены порядка (для «Объединения» и «PDF → Word»).</summary>
        public bool AllowReorder { get; set; }

        /// <summary>Разрешить поворот выбранных страниц (для «Объединения» и «Разделения»; в «PDF → Word» поворота нет).</summary>
        public bool AllowRotate { get; set; }

        /// <summary>Идёт фоновая операция формы: правки (буфер, поворот, перетаскивание, дроп) заблокированы.</summary>
        public bool Locked { get; set; }

        public event EventHandler SelectionChanged;
        /// <summary>Перетащили элемент from на позицию вставки to (0..Count).</summary>
        public event Action<int, int> ReorderRequested;
        /// <summary>Масштаб изменён изнутри (Ctrl+колесо) — чтобы синхронизировать ползунок.</summary>
        public event Action<int> ZoomChanged;
        /// <summary>Вставка вырезанных: перенести страницы с индексами (по возрастанию) ПЕРЕД позицией.</summary>
        public event Action<int[], int> MoveRangeRequested;
        /// <summary>Вставка скопированных: вставить НОВЫЕ экземпляры страниц ПЕРЕД позицией.</summary>
        public event Action<PdfPageRef[], int> InsertPagesRequested;
        /// <summary>На сетку сбросили PDF-файлы; int — позиция вставки (Count — в конец).</summary>
        public event Action<string[], int> FilesDropped;
        /// <summary>«Удалить» из контекстного меню (клавиша Delete обрабатывается формой напрямую).</summary>
        public event EventHandler DeleteRequested;
        /// <summary>«Перейти к странице…» из контекстного меню (диалог показывает форма).</summary>
        public event EventHandler GoToRequested;

        public PdfPageGrid()
        {
            _renderWidth = ThumbZoom.RenderWidthFor(DeviceDpi);
            _pageCache = new LruCache<CachedPage>(
                ThumbZoom.PageCacheCapacity(PageCacheBudget, _renderWidth), OnPageEvicted);
            _thumbs = NewImageList(_tileWidth);
            _list.Dock = DockStyle.Fill;
            _list.View = View.LargeIcon;
            _list.LargeImageList = _thumbs;
            _list.MultiSelect = true;
            _list.HideSelection = false;
            _list.LabelWrap = true;
            _list.BorderStyle = BorderStyle.FixedSingle;
            _list.BackColor = Color.FromArgb(250, 250, 250);
            _list.SelectedIndexChanged += delegate { var h = SelectionChanged; if (h != null) h(this, EventArgs.Empty); };
            _list.AllowDrop = true;
            _list.ItemDrag += OnItemDrag;
            _list.DragEnter += OnListDragEnter;
            _list.DragOver += OnListDragOver;
            _list.DragDrop += OnListDragDrop;
            _list.DragLeave += delegate { _list.InsertionMark.Index = -1; StopDragScroll(); };
            _list.MouseWheel += OnListMouseWheel;
            _list.MouseUp += OnListMouseUp;
            _list.Scrolled += delegate { ScheduleVisibleUpdate(); };
            _list.Resize += delegate { ScheduleVisibleUpdate(); };
            _list.SelectedIndexChanged += delegate { ScheduleVisibleUpdate(); }; // навигация клавишами
            Ui.EnableDoubleBuffer(_list);
            Controls.Add(_list);

            BuildContextMenu();

            // Троттлинг: события прокрутки сливаются в одно обновление видимых миниатюр.
            _visibleTimer = new System.Windows.Forms.Timer();
            _visibleTimer.Interval = 100;
            _visibleTimer.Tick += delegate { _visibleTimer.Stop(); UpdateVisibleThumbs(); };

            // Автопрокрутка при перетаскивании: DragOver не приходит без движения мыши,
            // поэтому прокрутка живёт на таймере, а метка вставки обновляется в тике.
            _dragScrollTimer = new System.Windows.Forms.Timer();
            _dragScrollTimer.Interval = 60;
            _dragScrollTimer.Tick += OnDragScrollTick;

            StartThumbWorker();
        }

        // ---------- публичный API ----------

        /// <summary>
        /// Подписи под плитками: true — ПОЗИЦИЯ в итоговом наборе (объединение,
        /// PDF → Word), false — номер ИСХОДНОЙ страницы (разделение одного файла).
        /// </summary>
        public bool ShowPositionNumbers { get; set; }

        /// <summary>Заменить содержимое сетки списком страниц (в этом порядке).</summary>
        public void SetPages(IList<PdfPageRef> pages)
        {
            // Набор ключей нового содержимого. Кэш, плитки и очередь чистятся до тех,
            // что остались: смена документа («Разделение») или удаление страниц
            // («Объединение») освобождают память сразу; переупорядочивание — тот же
            // набор, поэтому ничего не вытесняется и не перерисовывается.
            _currentKeys = BuildKeySet(pages);
            lock (_qLock)
            {
                _thumbQueue.Clear();     // снятые заявки на рендер отсутствующих страниц
                _thumbRequested.Clear(); // дедуп сбрасывается; кэш-проверка в EnqueueThumb не даёт перерендер
            }

            // Вырезанные, которых больше нет в наборе, выпадают из буфера; пустой
            // буфер вырезания снимается целиком (копии от набора не зависят).
            if (_clipIsCut)
            {
                var present = new HashSet<PdfPageRef>();
                if (pages != null)
                    foreach (PdfPageRef page in pages)
                        present.Add(page);
                _cutRefs.IntersectWith(present);
                _clipboard.RemoveAll(delegate(PdfPageRef p) { return !_cutRefs.Contains(p); });
                if (_cutRefs.Count == 0)
                    ClearClipboard();
            }
            ClearCaret();

            _list.BeginUpdate();
            _list.Items.Clear();         // после очистки ни один элемент не ссылается на плитки
            _itemsByKey.Clear();
            PruneCache(_currentKeys);    // освободить bitmap и плитки страниц вне набора
            if (pages != null)
            {
                for (int i = 0; i < pages.Count; i++)
                {
                    PdfPageRef page = pages[i];
                    string key = ThumbKey(page);
                    var item = new ListViewItem(PageLabel(page, i, ShowPositionNumbers));
                    item.Tag = page;
                    item.ToolTipText = string.Format(Loc.T("grid.pageTip"), page.SourcePath, page.PageIndex + 1);
                    if (_cutRefs.Contains(page))
                        item.ForeColor = SystemColors.GrayText;
                    _list.Items.Add(item);
                    IndexItem(key, item);
                    EnsureTile(item); // плитка из кэша, если страница уже отрендерена; иначе заглушка
                }
            }
            _list.EndUpdate();
            ScheduleVisibleUpdate(); // рендерим только видимые страницы, а не все сразу
        }

        /// <summary>Подпись плитки: позиция в наборе или номер исходной страницы. Чистая — под тест.</summary>
        internal static string PageLabel(PdfPageRef page, int position, bool positionMode)
        {
            int number = positionMode ? position + 1 : page.PageIndex + 1;
            return number.ToString();
        }

        /// <summary>Множество ключей плиток для набора страниц (без дублей). Чистая — под тест.</summary>
        internal static HashSet<string> BuildKeySet(IList<PdfPageRef> pages)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (pages != null)
                foreach (PdfPageRef page in pages)
                    keys.Add(ThumbKey(page));
            return keys;
        }

        /// <summary>Ключи кэша, отсутствующие в наборе keep (их плитки пора вытеснить). Чистая — под тест.</summary>
        internal static List<string> StaleKeys(IEnumerable<string> cachedKeys, ICollection<string> keepKeys)
        {
            var stale = new List<string>();
            foreach (string key in cachedKeys)
                if (!keepKeys.Contains(key))
                    stale.Add(key);
            return stale;
        }

        /// <summary>Освобождает bitmap и плитки страниц, которых больше нет в наборе keep.</summary>
        private void PruneCache(ICollection<string> keepKeys)
        {
            foreach (string key in StaleKeys(_pageCache.KeysSnapshot(), keepKeys))
            {
                CachedPage page;
                if (_pageCache.Remove(key, out page))
                    page.Bmp.Dispose();
                RemoveTilesOf(key);
            }
        }

        /// <summary>
        /// Вытеснение страницы LRU-кэшем (переполнение бюджета): освободить bitmap,
        /// снять её плитки и вернуть элементам заглушку — при следующем показе страница
        /// перерендерится (заявки в _thumbRequested на неё уже нет). Только UI-поток.
        /// </summary>
        private void OnPageEvicted(CachedPage page)
        {
            page.Bmp.Dispose();
            if (_shuttingDown)
                return; // Dispose: ImageList и элементы освобождаются целиком
            RemoveTilesOf(page.Key);
            // Если вытесненная страница ещё видна (кэш меньше видимого окна — крайний
            // случай), заявка на перерендер уйдёт ближайшим тиком, а не ждёт скролла.
            ScheduleVisibleUpdate();
        }

        /// <summary>Снять из ImageList все плитки страницы (все её повороты) и вернуть её элементам заглушку.</summary>
        private void RemoveTilesOf(string pageKey)
        {
            string prefix = pageKey + "|r";
            var stale = new List<string>();
            foreach (string imageKey in _thumbs.Images.Keys)
                if (imageKey.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    stale.Add(imageKey);
            foreach (string imageKey in stale)
                _thumbs.Images.RemoveByKey(imageKey);

            List<ListViewItem> items;
            if (_itemsByKey.TryGetValue(pageKey, out items))
                foreach (ListViewItem item in items)
                    if (item.ImageKey != PlaceholderKey)
                        item.ImageKey = PlaceholderKey;
        }

        /// <summary>Регистрирует элемент под ключом страницы (один ключ — несколько элементов при повторах страницы).</summary>
        private void IndexItem(string key, ListViewItem item)
        {
            List<ListViewItem> items;
            if (!_itemsByKey.TryGetValue(key, out items))
            {
                items = new List<ListViewItem>();
                _itemsByKey[key] = items;
            }
            items.Add(item);
        }

        // ---------- ленивый рендер видимых ----------

        private void ScheduleVisibleUpdate()
        {
            if (_visibleTimer == null)
                return;
            _visibleTimer.Stop();
            _visibleTimer.Start();
        }

        /// <summary>Ставит в очередь рендера только видимые страницы (плюс небольшой буфер).</summary>
        private void UpdateVisibleThumbs()
        {
            int count = _list.Items.Count;
            if (count == 0)
                return;
            // Раскладка LargeIcon монотонна сверху вниз (Bounds.Top/Bottom не убывают
            // по индексу): видимый диапазон ищем бинарным поиском — O(log n) обращений
            // к Bounds вместо линейного скана от начала на каждый тик прокрутки.
            // (ListView.TopItem в LargeIcon бросает исключение, поэтому по Bounds.)
            int bottom = _list.ClientSize.Height;
            int first, last;
            VisibleRange(count,
                delegate(int i) { return _list.Items[i].Bounds.Top; },
                delegate(int i) { return _list.Items[i].Bounds.Bottom; },
                bottom, out first, out last);
            if (first > last)
            {
                first = 0;
                last = Math.Min(count - 1, EnqueueBuffer);
            }
            int lo, hi;
            ClampWindow(first, last, count, EnqueueBuffer, out lo, out hi);
            for (int i = lo; i <= hi; i++)
            {
                var page = _list.Items[i].Tag as PdfPageRef;
                if (page != null)
                    EnqueueThumb(page);
            }
        }

        /// <summary>Окно докачки [lo..hi] вокруг видимого диапазона. Чистая — под тест.</summary>
        internal static void ClampWindow(int first, int last, int count, int buffer, out int lo, out int hi)
        {
            lo = first - buffer;
            if (lo < 0)
                lo = 0;
            hi = last + buffer;
            if (hi > count - 1)
                hi = count - 1;
        }

        /// <summary>
        /// Наименьший индекс в [0,count), для которого pred истинно; count, если такого
        /// нет. pred монотонен (false…false, затем true…true). Чистая — под тест.
        /// </summary>
        internal static int LowerBound(int count, Predicate<int> pred)
        {
            int lo = 0, hi = count;
            while (lo < hi)
            {
                int mid = lo + (hi - lo) / 2;
                if (pred(mid)) hi = mid;
                else lo = mid + 1;
            }
            return lo;
        }

        /// <summary>
        /// Видимый диапазон [first,last] по монотонным Top/Bottom (координаты клиента):
        /// first — первый элемент с Bottom ≥ 0, last — последний с Top ≤ viewportBottom.
        /// first &gt; last — ничего целиком не видно. Чистая — под тест.
        /// </summary>
        internal static void VisibleRange(int count, Func<int, int> topOf, Func<int, int> bottomOf, int viewportBottom, out int first, out int last)
        {
            first = LowerBound(count, delegate(int i) { return bottomOf(i) >= 0; });
            last = LowerBound(count, delegate(int i) { return topOf(i) > viewportBottom; }) - 1;
        }

        public int Count { get { return _list.Items.Count; } }
        public int SelectedCount { get { return _list.SelectedIndices.Count; } }
        public bool ListFocused { get { return _list.Focused; } }

        public int[] GetSelectedIndices()
        {
            var arr = new int[_list.SelectedIndices.Count];
            _list.SelectedIndices.CopyTo(arr, 0);
            Array.Sort(arr);
            return arr;
        }

        public void SelectAll()
        {
            if (_list.Items.Count == 0)
                return;
            _list.BeginUpdate();
            foreach (ListViewItem item in _list.Items)
                item.Selected = true;
            _list.EndUpdate();
        }

        public void SelectIndex(int index)
        {
            if (index < 0 || index >= _list.Items.Count)
                return;
            _list.SelectedIndices.Clear();
            _list.Items[index].Selected = true;
            _list.Items[index].EnsureVisible();
            _list.Focus();
        }

        /// <summary>Выделить диапазон подряд (после вставки/переноса) и показать его начало.</summary>
        public void SelectRange(int start, int count)
        {
            if (start < 0 || count <= 0 || start >= _list.Items.Count)
                return;
            int end = Math.Min(start + count - 1, _list.Items.Count - 1);
            _list.BeginUpdate();
            _list.SelectedIndices.Clear();
            for (int i = start; i <= end; i++)
                _list.Items[i].Selected = true;
            _list.EndUpdate();
            _list.Items[start].EnsureVisible();
            _list.Focus();
        }

        public int TileWidth { get { return _tileWidth; } }

        /// <summary>Задать масштаб плиток и пересобрать их из кэша (без повторного WinRT).</summary>
        public void SetTileWidth(int width)
        {
            width = ThumbZoom.Clamp(width);
            if (width == _tileWidth)
                return;
            _tileWidth = width;
            RebuildTiles();
        }

        // ---------- буфер страниц (вырезать/копировать/вставить) ----------

        /// <summary>Есть что вставлять.</summary>
        public bool ClipboardAvailable { get { return _clipboard.Count > 0; } }

        /// <summary>Есть незавершённое вырезание (для Esc).</summary>
        public bool HasCutPending { get { return _clipIsCut && _cutRefs.Count > 0; } }

        /// <summary>Вырезать выбранные: остаются на месте приглушёнными до вставки (Esc — отмена).</summary>
        public void CutSelected()
        {
            if (Locked || !AllowReorder || _list.SelectedIndices.Count == 0)
                return;
            CancelCut(); // прежнее вырезание снимается (его приглушение возвращается)
            _clipboard.Clear();
            _clipIsCut = true;
            foreach (int index in GetSelectedIndices())
            {
                var page = _list.Items[index].Tag as PdfPageRef;
                if (page == null)
                    continue;
                _clipboard.Add(page);
                _cutRefs.Add(page);
                _list.Items[index].ForeColor = SystemColors.GrayText;
            }
        }

        /// <summary>Копировать выбранные: в буфере — снимки страниц (с текущим поворотом).</summary>
        public void CopySelected()
        {
            if (Locked || !AllowReorder || _list.SelectedIndices.Count == 0)
                return;
            CancelCut();
            _clipboard.Clear();
            _clipIsCut = false;
            foreach (int index in GetSelectedIndices())
            {
                var page = _list.Items[index].Tag as PdfPageRef;
                if (page != null)
                    _clipboard.Add(page.Clone());
            }
        }

        /// <summary>
        /// Вставить буфер: вырезанные переносятся (MoveRangeRequested), скопированные
        /// вставляются новыми экземплярами (InsertPagesRequested). Позиция — каретка;
        /// без каретки — после последнего выбранного; без выбора — в конец.
        /// </summary>
        public void PasteClipboard()
        {
            if (Locked || !AllowReorder || _clipboard.Count == 0)
                return;
            int target = PasteIndex(_caretIndex, _caretAfter, GetSelectedIndices(), _list.Items.Count);
            if (_clipIsCut)
            {
                // Текущие позиции вырезанных ссылок (часть могла быть удалена — берём живые).
                var indices = new List<int>();
                for (int i = 0; i < _list.Items.Count; i++)
                    if (_cutRefs.Contains(_list.Items[i].Tag as PdfPageRef))
                        indices.Add(i);
                CancelCut();      // цвет возвращаем сразу: даже если перенос не состоится, серых плиток не останется
                ClearClipboard();
                if (indices.Count == 0)
                    return;
                var h = MoveRangeRequested;
                if (h != null)
                    h(indices.ToArray(), target);
            }
            else
            {
                var clones = new PdfPageRef[_clipboard.Count];
                for (int i = 0; i < _clipboard.Count; i++)
                    clones[i] = _clipboard[i].Clone(); // повторная вставка независима от прежних
                var h = InsertPagesRequested;
                if (h != null)
                    h(clones, target);
            }
        }

        /// <summary>Позиция вставки буфера. Чистая — под тест.</summary>
        internal static int PasteIndex(int caretIndex, bool caretAfter, int[] selectedSorted, int count)
        {
            if (caretIndex >= 0)
            {
                int at = caretIndex + (caretAfter ? 1 : 0);
                return at > count ? count : at;
            }
            if (selectedSorted != null && selectedSorted.Length > 0)
                return selectedSorted[selectedSorted.Length - 1] + 1;
            return count;
        }

        /// <summary>Отменить вырезание (Esc): вернуть цвет, очистить буфер. true — было что отменять.</summary>
        public bool TryCancelCut()
        {
            if (!HasCutPending)
                return false;
            CancelCut();
            ClearClipboard();
            return true;
        }

        private void CancelCut()
        {
            if (_cutRefs.Count == 0)
                return;
            foreach (ListViewItem item in _list.Items)
                if (_cutRefs.Contains(item.Tag as PdfPageRef))
                    item.ForeColor = _list.ForeColor;
            _cutRefs.Clear();
        }

        private void ClearClipboard()
        {
            _clipboard.Clear();
            _cutRefs.Clear();
            _clipIsCut = false;
        }

        // ---------- каретка вставки ----------

        private void OnListMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;
            if (_list.HitTest(e.Location).Item != null)
            {
                ClearCaret(); // клик по плитке: якорь вставки — выделение
                return;
            }
            if (_list.Items.Count == 0 || !AllowReorder || Locked)
                return;
            int index = _list.InsertionMark.NearestIndex(e.Location);
            bool after;
            if (index < 0)
            {
                index = _list.Items.Count - 1; // клик ниже последней строки — вставка в конец
                after = true;
            }
            else
            {
                Rectangle bounds = _list.GetItemRect(index);
                after = e.Location.X > bounds.Left + bounds.Width / 2;
            }
            _caretIndex = index;
            _caretAfter = after;
            _list.InsertionMark.AppearsAfterItem = after;
            _list.InsertionMark.Index = index;
        }

        private void ClearCaret()
        {
            _caretIndex = -1;
            _caretAfter = false;
            _list.InsertionMark.Index = -1;
        }

        // ---------- поворот выбранных страниц ----------

        /// <summary>
        /// Повернуть выбранные страницы на delta градусов по часовой (±90). Поворот —
        /// свойство страницы В ИТОГОВОМ файле (модель делит эти же ссылки); исходный
        /// PDF не меняется. Плитки обновляются из кэша без повторного рендера.
        /// </summary>
        public void RotateSelected(int delta)
        {
            if (Locked || !AllowRotate || _list.SelectedIndices.Count == 0)
                return;
            var pageKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _list.BeginUpdate();
            foreach (int index in GetSelectedIndices())
            {
                ListViewItem item = _list.Items[index];
                var page = item.Tag as PdfPageRef;
                if (page == null)
                    continue;
                page.Rotation = PdfPageRef.ComposeRotation(page.Rotation, delta);
                pageKeys.Add(ThumbKey(page));
                EnsureTile(item);
            }
            foreach (string pageKey in pageKeys)
                PruneUnusedRotations(pageKey);
            _list.EndUpdate();
        }

        /// <summary>Снять плитки поворотов страницы, которые больше не использует ни один её элемент.</summary>
        private void PruneUnusedRotations(string pageKey)
        {
            List<ListViewItem> items;
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (_itemsByKey.TryGetValue(pageKey, out items))
                foreach (ListViewItem item in items)
                {
                    var page = item.Tag as PdfPageRef;
                    if (page != null)
                        used.Add(TileKey(page));
                }
            string prefix = pageKey + "|r";
            var stale = new List<string>();
            foreach (string imageKey in _thumbs.Images.Keys)
                if (imageKey.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && !used.Contains(imageKey))
                    stale.Add(imageKey);
            foreach (string imageKey in stale)
                _thumbs.Images.RemoveByKey(imageKey);
        }

        // ---------- перетаскивание и дроп файлов ----------

        private void OnItemDrag(object sender, ItemDragEventArgs e)
        {
            if (AllowReorder && !Locked)
                _list.DoDragDrop(e.Item, DragDropEffects.Move);
        }

        private void OnListDragEnter(object sender, DragEventArgs e)
        {
            // Наличие PDF среди файлов выясняем один раз на drag-сессию: DragOver
            // сыплется на каждое движение мыши, разбирать там HDROP было бы расточительно.
            _dragHasFiles = !Locked && e.Data.GetDataPresent(DataFormats.FileDrop) &&
                PdfDrop.ExtractPaths(e).Length > 0;
            // Effect обязан быть выставлен уже на входе: без этого курсор показывает
            // «нельзя» до первого движения, а дроп без движения мыши отвергается вовсе.
            e.Effect = _dragHasFiles
                ? DragDropEffects.Copy
                : (AllowReorder && !Locked && e.Data.GetDataPresent(typeof(ListViewItem))
                    ? DragDropEffects.Move
                    : DragDropEffects.None);
            // Drag-сессия забирает InsertionMark под себя — прежняя каретка вставки
            // теряет и смысл, и отображение; иначе после отменённого перетаскивания
            // Ctrl+V вставил бы в невидимую позицию.
            ClearCaret();
        }

        private void OnListDragOver(object sender, DragEventArgs e)
        {
            bool reorder = AllowReorder && !Locked && e.Data.GetDataPresent(typeof(ListViewItem));
            if (!reorder && !_dragHasFiles)
            {
                e.Effect = DragDropEffects.None;
                StopDragScroll();
                return;
            }
            e.Effect = reorder ? DragDropEffects.Move : DragDropEffects.Copy;
            Point pt = _list.PointToClient(new Point(e.X, e.Y));
            // Метка вставки: и для внутреннего переноса, и для дропа файлов в позицию.
            // В сетке без порядка («Разделение») позиция не имеет смысла — метку не показываем.
            if (AllowReorder)
                UpdateInsertionMark(pt);
            UpdateDragScroll(pt.Y);
        }

        private void UpdateInsertionMark(Point pt)
        {
            int index = _list.InsertionMark.NearestIndex(pt);
            if (index >= 0)
            {
                Rectangle bounds = _list.GetItemRect(index);
                _list.InsertionMark.AppearsAfterItem = pt.X > bounds.Left + bounds.Width / 2;
            }
            _list.InsertionMark.Index = index;
        }

        private void OnListDragDrop(object sender, DragEventArgs e)
        {
            int target = _list.InsertionMark.Index;
            bool after = _list.InsertionMark.AppearsAfterItem;
            _list.InsertionMark.Index = -1;
            StopDragScroll();

            if (_dragHasFiles)
            {
                _dragHasFiles = false;
                string[] paths = PdfDrop.ExtractPaths(e);
                if (Locked || paths.Length == 0)
                    return;
                int insertAt = DropInsertIndex(AllowReorder, target, after, _list.Items.Count);
                var hf = FilesDropped;
                if (hf != null)
                    hf(paths, insertAt);
                return;
            }

            var item = e.Data.GetData(typeof(ListViewItem)) as ListViewItem;
            if (!AllowReorder || Locked || item == null || target < 0)
                return;
            int from = item.Index;
            int to = after ? target + 1 : target;
            var h = ReorderRequested;
            if (h != null)
                h(from, to);
        }

        /// <summary>
        /// Позиция вставки при дропе файлов: в сетке с порядком — по метке вставки
        /// (после плитки markIndex, если markAfter), иначе в конец. Чистая — под тест.
        /// </summary>
        internal static int DropInsertIndex(bool allowReorder, int markIndex, bool markAfter, int count)
        {
            if (allowReorder && markIndex >= 0)
                return markAfter ? markIndex + 1 : markIndex;
            return count;
        }

        // Автопрокрутка: у краёв сетки прокручиваем на таймере (DragOver без движения
        // мыши не приходит) и обновляем метку вставки под курсором.
        private void UpdateDragScroll(int clientY)
        {
            int dir = 0;
            if (clientY < DragEdgePx)
                dir = -1;
            else if (clientY > _list.ClientSize.Height - DragEdgePx)
                dir = 1;
            _dragScrollDir = dir;
            if (dir != 0)
            {
                if (!_dragScrollTimer.Enabled)
                    _dragScrollTimer.Start();
            }
            else
                StopDragScroll();
        }

        private void StopDragScroll()
        {
            _dragScrollDir = 0;
            if (_dragScrollTimer != null && _dragScrollTimer.Enabled)
                _dragScrollTimer.Stop();
        }

        private void OnDragScrollTick(object sender, EventArgs e)
        {
            if (_dragScrollDir == 0 || !_list.IsHandleCreated)
            {
                StopDragScroll();
                return;
            }
            int step = Math.Max(24, _thumbs.ImageSize.Height / 2);
            SendMessage(_list.Handle, LVM_SCROLL, IntPtr.Zero, (IntPtr)(_dragScrollDir * step));
            if (AllowReorder)
                UpdateInsertionMark(_list.PointToClient(Control.MousePosition));
            ScheduleVisibleUpdate();
        }

        // ---------- контекстное меню ----------

        private void BuildContextMenu()
        {
            _menu = new ContextMenuStrip();
            _miCut = AddMenuItem(Loc.T("grid.menu.cut"), "Ctrl+X", delegate { CutSelected(); });
            _miCopy = AddMenuItem(Loc.T("grid.menu.copy"), "Ctrl+C", delegate { CopySelected(); });
            _miPaste = AddMenuItem(Loc.T("grid.menu.paste"), "Ctrl+V", delegate { PasteClipboard(); });
            _menu.Items.Add(new ToolStripSeparator());
            _miRotateRight = AddMenuItem(Loc.T("grid.menu.rotateRight"), "Ctrl+Shift+«+»", delegate { RotateSelected(90); });
            _miRotateLeft = AddMenuItem(Loc.T("grid.menu.rotateLeft"), "Ctrl+Shift+«−»", delegate { RotateSelected(-90); });
            _menu.Items.Add(new ToolStripSeparator());
            _miDelete = AddMenuItem(Loc.T("grid.menu.delete"), "Del", delegate { var h = DeleteRequested; if (h != null) h(this, EventArgs.Empty); });
            _miGoTo = AddMenuItem(Loc.T("grid.menu.goto"), "Ctrl+G", delegate { var h = GoToRequested; if (h != null) h(this, EventArgs.Empty); });
            _menu.Opening += OnMenuOpening;
            _list.ContextMenuStrip = _menu;
        }

        private ToolStripMenuItem AddMenuItem(string text, string shortcutHint, EventHandler onClick)
        {
            var mi = new ToolStripMenuItem(text);
            mi.ShortcutKeyDisplayString = shortcutHint; // клавиши обрабатывает ProcessCmdKey формы
            mi.Click += onClick;
            _menu.Items.Add(mi);
            return mi;
        }

        private void OnMenuOpening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            bool hasSelection = _list.SelectedIndices.Count > 0;
            _miCut.Visible = _miCopy.Visible = _miPaste.Visible = _miDelete.Visible = AllowReorder;
            _miCut.Enabled = _miCopy.Enabled = _miDelete.Enabled = !Locked && hasSelection;
            _miPaste.Enabled = !Locked && ClipboardAvailable;
            _miRotateRight.Visible = _miRotateLeft.Visible = AllowRotate;
            _miRotateRight.Enabled = _miRotateLeft.Enabled = !Locked && hasSelection;
            _miGoTo.Enabled = _list.Items.Count > 0;
            // Разделители между скрытыми группами не нужны.
            bool anyEdit = AllowReorder;
            bool anyRotate = AllowRotate;
            _menu.Items[3].Visible = anyEdit && anyRotate;       // разделитель буфер | поворот
            _menu.Items[6].Visible = anyEdit || anyRotate;       // разделитель | удалить/перейти
            e.Cancel = _menu.Items.Count == 0;
        }

        // ---------- масштаб ----------

        private void OnListMouseWheel(object sender, MouseEventArgs e)
        {
            if ((ModifierKeys & Keys.Control) == 0)
                return; // без Ctrl — обычная прокрутка
            int newWidth = ThumbZoom.StepFromWheel(_tileWidth, e.Delta);
            var handled = e as HandledMouseEventArgs;
            if (handled != null)
                handled.Handled = true;
            if (newWidth != _tileWidth)
            {
                SetTileWidth(newWidth);
                var h = ZoomChanged;
                if (h != null) h(_tileWidth);
            }
        }

        private void RebuildTiles()
        {
            ImageList old = _thumbs;
            var fresh = NewImageList(_tileWidth);

            _list.BeginUpdate();
            _thumbs = fresh;
            _list.LargeImageList = fresh;
            foreach (ListViewItem item in _list.Items)
            {
                var page = item.Tag as PdfPageRef;
                if (page == null)
                    continue;
                CachedPage cached;
                if (_pageCache.TryPeek(ThumbKey(page), out cached))
                {
                    string tileKey = TileKey(page);
                    if (!fresh.Images.ContainsKey(tileKey))
                        fresh.Images.Add(tileKey, ComposeTile(cached.Bmp, fresh.ImageSize, page.Rotation));
                    item.ImageKey = tileKey;
                }
                else
                    item.ImageKey = PlaceholderKey;
            }
            _list.EndUpdate();
            if (old != null)
                old.Dispose();
        }

        // ---------- фоновый рендер ----------

        internal static string ThumbKey(PdfPageRef page)
        {
            return page.SourcePath.ToLowerInvariant() + "|" + page.PageIndex;
        }

        /// <summary>Ключ плитки: страница + её поворот (одна страница может висеть в сетке с разными поворотами).</summary>
        internal static string TileKey(PdfPageRef page)
        {
            return ThumbKey(page) + "|r" + page.Rotation;
        }

        private void EnqueueThumb(PdfPageRef page)
        {
            string key = ThumbKey(page);
            CachedPage cached;
            if (_pageCache.TryGet(key, out cached)) // попадание освежает LRU: видимые не вытесняются
                return;
            lock (_qLock)
            {
                if (!_thumbRequested.Add(key))
                    return;
                _thumbQueue.Enqueue(page);
            }
            _thumbSignal.Set();
        }

        private void StartThumbWorker()
        {
            _thumbThread = new Thread(ThumbWorker);
            _thumbThread.IsBackground = true;
            _thumbThread.Name = "pdf-thumbs";
            _thumbThread.Start();
        }

        private void ThumbWorker()
        {
            PdfThumbnailRenderer renderer;
            try { renderer = new PdfThumbnailRenderer(); }
            catch { return; } // WinRT недоступен — останутся заглушки

            try
            {
                while (!_thumbStop)
                {
                    PdfPageRef req = null;
                    lock (_qLock)
                    {
                        if (_thumbQueue.Count > 0) req = _thumbQueue.Dequeue();
                        else _thumbSignal.Reset();
                    }
                    if (req == null)
                    {
                        _thumbSignal.Wait();
                        continue;
                    }
                    Bitmap page = renderer.Render(req.SourcePath, req.PageIndex, _renderWidth);
                    if (page == null)
                        continue;
                    PostPage(req, page);
                }
            }
            finally
            {
                renderer.Dispose();
            }
        }

        private void PostPage(PdfPageRef req, Bitmap page)
        {
            try
            {
                if (IsHandleCreated && !IsDisposed)
                    BeginInvoke((MethodInvoker)delegate { ApplyPage(req, page); });
                else
                    page.Dispose();
            }
            catch (InvalidOperationException)
            {
                page.Dispose();
            }
        }

        private void ApplyPage(PdfPageRef req, Bitmap page)
        {
            string key = ThumbKey(req);
            // Заявка исполнена: снять дедуп, чтобы вытесненная позже страница могла
            // перерендериться (иначе после вытеснения она навсегда осталась бы заглушкой).
            lock (_qLock)
                _thumbRequested.Remove(key);
            if (!_currentKeys.Contains(key))
            {
                page.Dispose(); // страница уже снята из набора — поздний результат рендера отбрасываем
                return;
            }
            CachedPage existing;
            if (_pageCache.TryPeek(key, out existing))
            {
                page.Dispose();
                return;
            }
            _pageCache.Add(key, new CachedPage { Key = key, Bmp = page });

            // Адресно по ключу, а не проходом по всем элементам (рендер всего
            // документа иначе O(n²)). Плитка — под фактический поворот каждого элемента.
            List<ListViewItem> items;
            if (_itemsByKey.TryGetValue(key, out items))
                foreach (ListViewItem item in items)
                    EnsureTile(item);
        }

        /// <summary>
        /// Обеспечить элементу плитку его страницы С ЕГО поворотом: если страница в
        /// кэше — составить/переиспользовать плитку, иначе — заглушка (рендер докачает).
        /// </summary>
        private void EnsureTile(ListViewItem item)
        {
            var page = item.Tag as PdfPageRef;
            if (page == null)
                return;
            CachedPage cached;
            if (!_pageCache.TryPeek(ThumbKey(page), out cached))
            {
                if (item.ImageKey != PlaceholderKey)
                    item.ImageKey = PlaceholderKey;
                return;
            }
            string tileKey = TileKey(page);
            if (!_thumbs.Images.ContainsKey(tileKey))
                _thumbs.Images.Add(tileKey, ComposeTile(cached.Bmp, _thumbs.ImageSize, page.Rotation));
            if (item.ImageKey != tileKey)
                item.ImageKey = tileKey;
        }

        private ImageList NewImageList(int tileWidth)
        {
            var list = new ImageList();
            list.ImageSize = ThumbZoom.TileSize(tileWidth);
            list.ColorDepth = ColorDepth.Depth32Bit;
            // Изображение НЕ освобождаем: ImageList удерживает ссылку на оригинал для
            // пересоздания нативного handle (смена DPI/темы/ColorDepth) — досрочный
            // Dispose дал бы «красный крест». Освобождение — задача самого ImageList.
            list.Images.Add(PlaceholderKey, MakePlaceholder(list.ImageSize));
            return list;
        }

        /// <summary>RotateFlip для поворота по часовой на 90/180/270. Чистая — под тест.</summary>
        internal static RotateFlipType FlipFor(int rotation)
        {
            switch (rotation)
            {
                case 90: return RotateFlipType.Rotate90FlipNone;
                case 180: return RotateFlipType.Rotate180FlipNone;
                case 270: return RotateFlipType.Rotate270FlipNone;
                default: return RotateFlipType.RotateNoneFlipNone;
            }
        }

        private static Bitmap ComposeTile(Bitmap page, Size tile, int rotation)
        {
            if (rotation == 0)
                return ComposeTileCore(page, tile);
            // RotateFlip мутирует изображение — поворачиваем КОПИЮ, кэш остаётся неповёрнутым.
            using (var rotated = (Bitmap)page.Clone())
            {
                rotated.RotateFlip(FlipFor(rotation));
                return ComposeTileCore(rotated, tile);
            }
        }

        private static Bitmap ComposeTileCore(Bitmap page, Size tile)
        {
            var bmp = new Bitmap(tile.Width, tile.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.FromArgb(250, 250, 250));
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                float scale = Math.Min((tile.Width - 12f) / page.Width, (tile.Height - 12f) / page.Height);
                int w = Math.Max(1, (int)(page.Width * scale));
                int h = Math.Max(1, (int)(page.Height * scale));
                int x = (tile.Width - w) / 2;
                int y = (tile.Height - h) / 2;
                g.DrawImage(page, x, y, w, h);
                using (var pen = new Pen(Color.FromArgb(200, 200, 200)))
                    g.DrawRectangle(pen, x, y, w - 1, h - 1);
            }
            return bmp;
        }

        private static Bitmap MakePlaceholder(Size tile)
        {
            var bmp = new Bitmap(tile.Width, tile.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.FromArgb(250, 250, 250));
                int w = (int)((tile.Height - 24) * 0.72f);
                int h = tile.Height - 24;
                int x = (tile.Width - w) / 2;
                int y = 12;
                using (var b = new SolidBrush(Color.White))
                    g.FillRectangle(b, x, y, w, h);
                using (var pen = new Pen(Color.FromArgb(205, 205, 205)))
                    g.DrawRectangle(pen, x, y, w, h);
            }
            return bmp;
        }

        /// <summary>Останавливает фоновый поток рендера — вызывать при закрытии окна.</summary>
        public void StopRendering()
        {
            _thumbStop = true;
            _thumbSignal.Set();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                StopRendering();
                // Дождаться выхода фонового рендера, прежде чем освобождать сигнал,
                // которого он касается (_thumbSignal.Wait/Reset): иначе поток упал бы
                // ObjectDisposedException. Таймаут — на случай долгого Render; сам
                // сигнал освобождаем только если поток гарантированно завершился,
                // поэтому медленный рендер «в полёте» не роняет процесс на диспоузе.
                bool stopped = _thumbThread == null || _thumbThread.Join(2000);
                if (_visibleTimer != null)
                    _visibleTimer.Dispose();
                if (_dragScrollTimer != null)
                    _dragScrollTimer.Dispose();
                if (_menu != null)
                    _menu.Dispose(); // ContextMenuStrip не дочерний контрол — сам не освободится
                _shuttingDown = true;    // вытеснение при Clear только освобождает bitmap
                _pageCache.Clear();
                if (_thumbs != null)
                    _thumbs.Dispose();
                if (stopped)
                    _thumbSignal.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
