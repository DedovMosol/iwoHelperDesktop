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
    /// и автопрокруткой у краёв, при AllowRotate — поворот выбранных или всех страниц
    /// с бейджем угла на плитке; дроп PDF-файлов прямо на сетку (с позицией вставки).
    /// Модель порядка страниц держит владелец (форма): сетка мутирует только Rotation
    /// ссылок (общих с моделью), а перестановки/вставки запрашивает событиями. Общая
    /// для «Объединения», «Разделения» и «PDF → Word» (DRY).
    ///
    /// Отрисовка — OWNER-DRAW: плитки, подписи-номера, выделение и приглушение
    /// вырезанных рисуются в DrawItem, шаг ячеек задаёт LVM_SETICONSPACING. Это
    /// снимает лимит ImageList 256×256 — масштаб до 400 px. Память ограничена:
    /// отрендеренные страницы живут в LRU-кэше с байтовым бюджетом (на x86 — вдвое
    /// меньше), плитки существуют только для закэшированных страниц; при крупном зуме
    /// видимые страницы дорендериваются в увеличенную ширину (прогрессивно).
    /// </summary>
    public class PdfPageGrid : UserControl
    {
        private const int EnqueueBuffer = 16;   // докачивать миниатюры чуть за пределами видимого
        private const int DragEdgePx = 28;      // зона автопрокрутки у верх/низ края при перетаскивании
        private const int DropFramePx = 3;      // толщина акцентной рамки-подсветки при дропе файла
        private static readonly long PageCacheBudget =
            IntPtr.Size == 8 ? 192L << 20 : 48L << 20; // x86: адресное пространство ~2 ГБ
        private static readonly Color HoverMarkColor = Color.FromArgb(185, 185, 185);
        private static readonly Color GridBack = Color.FromArgb(250, 250, 250);

        // Кисти/перья горячего paint-пути (DrawItemCore зовётся на каждую видимую плитку в
        // каждом кадре): цвета фиксированы — объекты живут весь процесс, а не создаются на
        // каждую отрисовку (тот же стандарт, что ChoiceCard.TitleFont). Только UI-поток.
        private static readonly Brush BackBrush = new SolidBrush(GridBack);
        private static readonly Brush DimBrush = new SolidBrush(Color.FromArgb(150, GridBack));   // приглушение вырезанных
        private static readonly Brush ChipFill = new SolidBrush(Color.FromArgb(200, 40, 40, 40)); // hover-кнопки поворота
        private static readonly Brush BadgeFill = new SolidBrush(Color.FromArgb(190, 40, 40, 40));// бейдж угла
        private static readonly Pen PlaceholderPen = new Pen(Color.FromArgb(205, 205, 205));
        private static readonly Pen TileBorderPen = new Pen(Color.FromArgb(200, 200, 200));
        // Заливка выделения зависит от системного цвета: кэш с ленивым пересозданием
        // (сбрасывается в OnSystemColorsChanged — смена темы не заморозит старый цвет).
        private Brush _selFill;

        /// <summary>
        /// ListView, извещающий о прокрутке (для ленивого рендера видимых страниц) и
        /// рисующий подсказку по центру, когда список ПУСТ (owner-draw не вызывает
        /// DrawItem без элементов, поэтому текст рисуем в WM_PAINT поверх фона).
        /// </summary>
        private sealed class ScrollList : ListView
        {
            private static readonly Color HintColor = Color.FromArgb(120, 120, 120);
            public event EventHandler Scrolled;
            public string HintText;   // текст-подсказка пустого списка; null — не рисовать
            public Font HintFont;     // общий кэшированный шрифт (Ui.Font) — не освобождать

            protected override void WndProc(ref Message m)
            {
                base.WndProc(ref m);
                const int WM_VSCROLL = 0x115, WM_MOUSEWHEEL = 0x20A, WM_KEYUP = 0x101, WM_PAINT = 0x0F;
                if (m.Msg == WM_VSCROLL || m.Msg == WM_MOUSEWHEEL || m.Msg == WM_KEYUP)
                {
                    EventHandler h = Scrolled;
                    if (h != null) h(this, EventArgs.Empty);
                }
                else if (m.Msg == WM_PAINT && Items.Count == 0 && !string.IsNullOrEmpty(HintText))
                {
                    // Пустой список: DrawItem не вызывается, база залила фон — рисуем подсказку.
                    using (var g = Graphics.FromHwnd(Handle))
                    {
                        Rectangle r = Rectangle.Inflate(ClientRectangle, -30, 0);
                        TextRenderer.DrawText(g, HintText, HintFont ?? Font, r, HintColor,
                            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
                            TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
                    }
                }
            }
        }

        /// <summary>Отрендеренная страница в LRU: ключ — для снятия плиток при вытеснении, Width — ширина рендера.</summary>
        private sealed class CachedPage
        {
            public string Key;
            public Bitmap Bmp;
            public int Width;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        private const int LVM_SCROLL = 0x1014;         // прокрутка LargeIcon на dy пикселей
        private const int LVM_SETICONSPACING = 0x1035; // шаг ячеек сетки (cx, cy) в пикселях

        private readonly ScrollList _list = new ScrollList();
        // ПУСТОЙ ImageList задаёт нативной ListView размер «иконной зоны» (якорь отрисовки
        // и хитбокс кликов): min(плитка, 256) — выше 256 WinForms не пускает. Плитки крупнее
        // рисуются поверх, а клики по их внешнему кольцу компенсируются в MouseUp
        // (нативный сброс выделения при клике «в пустоту» идёт ПОСЛЕ managed MouseDown).
        private readonly ImageList _metrics = new ImageList();
        private System.Windows.Forms.Timer _visibleTimer;
        private readonly LruCache<CachedPage> _pageCache;
        // Плитки текущего масштаба по ключу «страница|поворот»: составляются лениво при
        // отрисовке из кэша страниц, живут только для закэшированных страниц. UI-поток.
        private readonly Dictionary<string, Bitmap> _tiles =
            new Dictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);
        private readonly int _renderBase; // ширина рендера для обычного масштаба (физ. пиксели)
        private readonly int _renderHi;   // для крупного зума (> ThumbZoom.BaseWidth)
        private volatile int _renderTarget; // читает фоновый воркер в момент рендера
        private bool _shuttingDown;        // Dispose: вытеснение не трогает UI, только освобождает bitmap
        // Ключи страниц, показываемых сейчас (обновляется в SetPages, только UI-поток).
        // Поздний результат рендера уже снятой страницы отбрасывается по этому набору.
        private HashSet<string> _currentKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // Элементы списка по ключу СТРАНИЦЫ (без поворота): готовый рендер перерисовывает
        // адресно, без прохода по всем элементам (иначе рендер документа — O(n²)). Только UI-поток.
        private readonly Dictionary<string, List<ListViewItem>> _itemsByKey =
            new Dictionary<string, List<ListViewItem>>(StringComparer.OrdinalIgnoreCase);
        private int _tileWidth = ThumbZoom.DefaultWidth;
        private Font _badgeFont;
        private Font _hintFont;   // подсказка пустого списка (крупнее подписей)
        private bool _dropActive; // над сеткой тянут файл — рисуем акцентную рамку и меняем подсказку
        private string _emptyHint; // текст подсказки для пустого списка (задаёт форма)
        private string _dropHint;  // текст подсказки во время дропа файла (задаёт форма)

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
        // InsertionMark цветом выделения). До клика зазор ПОДСВЕЧИВАЕТСЯ при наведении
        // светлой меткой. Сбрасывается при любой пересборке содержимого и drag-сессией.
        private int _caretIndex = -1;
        private bool _caretAfter;
        private bool _hoverMark; // метка показана как подсветка наведения (не каретка)
        // Плитка под курсором: на ней рисуются hover-кнопки поворота ↺/↻ (как в Acrobat DC).
        private int _hotIndex = -1;
        private Font _chipFont;

        // Автопрокрутка при перетаскивании у края сетки.
        private System.Windows.Forms.Timer _dragScrollTimer;
        private int _dragScrollDir; // -1 вверх, +1 вниз, 0 — нет
        private bool _dragHasFiles; // в текущей drag-сессии тянут PDF-файлы (решено в DragEnter)

        private ContextMenuStrip _menu;
        private ToolStripMenuItem _miCut, _miCopy, _miPaste, _miMoveAfter, _miRotate, _miRotateRight, _miRotateLeft,
            _miRotateAllRight, _miRotateAllLeft, _miDelete, _miGoTo;

        /// <summary>Разрешить перетаскивание/буфер страниц для смены порядка (для «Объединения» и «PDF → Word»).</summary>
        public bool AllowReorder { get; set; }

        /// <summary>Разрешить поворот страниц (все три PDF-инструмента; в «PDF → Word» страница выправляется до анализа).</summary>
        public bool AllowRotate { get; set; }

        /// <summary>Идёт фоновая операция формы: правки (буфер, поворот, перетаскивание, дроп) заблокированы.</summary>
        public bool Locked { get; set; }

        public event EventHandler SelectionChanged;
        /// <summary>Перетащили элемент from на позицию вставки to (0..Count).</summary>
        public event Action<int, int> ReorderRequested;
        /// <summary>Масштаб изменён изнутри (Ctrl+колесо) — чтобы синхронизировать регулятор.</summary>
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
        /// <summary>«Переместить после страницы N…» из контекстного меню (диалог показывает форма).</summary>
        public event EventHandler MoveAfterRequested;
        /// <summary>
        /// Сейчас будут повёрнуты страницы (клавиши, меню или hover-кнопка): формы с моделью
        /// порядка делают снимок для Ctrl+Z. Поднимается ДО мутации поворотов.
        /// </summary>
        public event EventHandler BeforeRotate;

        public PdfPageGrid()
        {
            _renderBase = ThumbZoom.RenderWidthFor(DeviceDpi, ThumbZoom.BaseWidth);
            _renderHi = ThumbZoom.RenderWidthFor(DeviceDpi, ThumbZoom.MaxWidth);
            _renderTarget = _renderBase;
            _pageCache = new LruCache<CachedPage>(
                ThumbZoom.PageCacheCapacity(PageCacheBudget, _renderHi), OnPageEvicted);
            // Общие кэшированные шрифты (Ui.Font): грид ими не владеет и не освобождает.
            _badgeFont = Ui.Font(7.5f, FontStyle.Bold);
            _chipFont = Ui.Font(9.5f, FontStyle.Regular, "Segoe UI Symbol"); // глифы ↺/↻ hover-кнопок
            _hintFont = Ui.Font(11f); // подсказка пустого списка (крупнее подписей)
            // Тонкая внутренняя рамка (Padding) видна только когда закрашена акцентом при
            // дропе: список Dock.Fill занимает клиент за вычетом Padding, его частичные
            // перерисовки в неё не залезают — рамка держится без мерцания.
            BackColor = GridBack;
            Padding = new Padding(DropFramePx);
            _metrics.ColorDepth = ColorDepth.Depth32Bit;
            _list.HintFont = _hintFont;
            _list.Dock = DockStyle.Fill;
            _list.View = View.LargeIcon;
            _list.LargeImageList = _metrics;
            _list.MultiSelect = true;
            _list.HideSelection = false;
            _list.BorderStyle = BorderStyle.FixedSingle;
            _list.BackColor = GridBack; // единый источник с BackBrush/DimBrush пейнта
            _list.OwnerDraw = true;
            _list.DrawItem += OnDrawItem;
            _list.HandleCreated += delegate { ApplyIconSpacing(); };
            _list.SelectedIndexChanged += delegate
            {
                // Нативная инвалидация выделения покрывает лишь нативные границы (≤ 256):
                // при крупном зуме рамка выделения и подпись шире — перерисовать всё видимое.
                _list.Invalidate();
                var h = SelectionChanged;
                if (h != null) h(this, EventArgs.Empty);
                ScheduleVisibleUpdate(); // навигация клавишами меняет набор видимых миниатюр
            };
            _list.AllowDrop = true;
            _list.ItemDrag += OnItemDrag;
            _list.DragEnter += OnListDragEnter;
            _list.DragOver += OnListDragOver;
            _list.DragDrop += OnListDragDrop;
            _list.DragLeave += delegate { _list.InsertionMark.Index = -1; StopDragScroll(); SetDropActive(false); };
            _list.MouseWheel += OnListMouseWheel;
            _list.MouseUp += OnListMouseUp;
            _list.MouseDoubleClick += OnListDoubleClick;
            _list.MouseMove += OnListMouseMove;
            _list.MouseLeave += delegate { ClearHoverMark(); SetHotIndex(-1); };
            _list.Scrolled += delegate { ScheduleVisibleUpdate(); };
            _list.Resize += delegate { ScheduleVisibleUpdate(); };
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

        /// <summary>Подсказка по центру пустого списка (задаёт форма под свою кнопку добавления).</summary>
        public string EmptyHint
        {
            get { return _emptyHint; }
            set { _emptyHint = value; UpdateHint(); }
        }

        /// <summary>Подсказка по центру пустого списка во время дропа файла («Отпустите…»).</summary>
        public string DropHint
        {
            get { return _dropHint; }
            set { _dropHint = value; UpdateHint(); }
        }

        /// <summary>Текст подсказки пустого списка: во время дропа — «отпустите», иначе — «добавьте». Чистая — под тест.</summary>
        internal static string HintFor(int count, bool dropActive, string emptyHint, string dropHint)
        {
            if (count > 0)
                return null; // список не пуст — подсказка не рисуется
            return dropActive ? (dropHint ?? emptyHint) : emptyHint;
        }

        /// <summary>Пересчитать и применить подсказку пустого списка (пусто/дроп).</summary>
        private void UpdateHint()
        {
            string hint = HintFor(_list.Items.Count, _dropActive, _emptyHint, _dropHint);
            if (string.Equals(_list.HintText, hint, StringComparison.Ordinal))
                return;
            _list.HintText = hint;
            if (_list.IsHandleCreated)
                _list.Invalidate();
        }

        /// <summary>Включить/снять акцентную рамку-подсветку дропа файла (рамка — Padding, текст — список).</summary>
        private void SetDropActive(bool active)
        {
            if (_dropActive == active)
                return;
            _dropActive = active;
            Invalidate();  // перерисовать рамку (Padding)
            UpdateHint();  // возможная смена подсказки пусто→дроп
        }

        /// <summary>Рамка-подсветка при дропе рисуется закраской Padding (список поверх оставляет только рамку).</summary>
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            e.Graphics.Clear(_dropActive ? Theme.Accent : BackColor);
        }

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
            _hotIndex = -1; // индексы элементов сейчас пересоберутся

            _list.BeginUpdate();
            _list.Items.Clear();         // после очистки ни один элемент не ссылается на плитки
            _itemsByKey.Clear();
            PruneCache(_currentKeys);    // освободить bitmap и плитки страниц вне набора
            if (pages != null)
            {
                for (int i = 0; i < pages.Count; i++)
                {
                    PdfPageRef page = pages[i];
                    var item = new ListViewItem(PageLabel(page, i, ShowPositionNumbers));
                    item.Tag = page;
                    item.ToolTipText = string.Format(Loc.T("grid.pageTip"), page.SourcePath, page.PageIndex + 1);
                    _list.Items.Add(item);
                    IndexItem(ThumbKey(page), item);
                }
            }
            _list.EndUpdate();
            UpdateHint();            // пусто/не пусто → показать или снять подсказку
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
        /// снять её плитки — при следующем показе страница перерендерится (заявки в
        /// _thumbRequested на неё уже нет). Только UI-поток.
        /// </summary>
        private void OnPageEvicted(CachedPage page)
        {
            page.Bmp.Dispose();
            if (_shuttingDown)
                return; // Dispose: всё остальное освобождается целиком
            RemoveTilesOf(page.Key);
            // Если вытесненная страница ещё видна (кэш меньше видимого окна — крайний
            // случай), заявка на перерендер уйдёт ближайшим тиком, а не ждёт скролла.
            ScheduleVisibleUpdate();
        }

        /// <summary>Снять плитки страницы (все её повороты) и перерисовать её элементы (появится заглушка).</summary>
        private void RemoveTilesOf(string pageKey)
        {
            string prefix = pageKey + "|r";
            var stale = new List<string>();
            foreach (KeyValuePair<string, Bitmap> kv in _tiles)
                if (kv.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    stale.Add(kv.Key);
            foreach (string tileKey in stale)
            {
                _tiles[tileKey].Dispose();
                _tiles.Remove(tileKey);
            }
            RedrawItemsOf(pageKey);
        }

        /// <summary>Перерисовать элементы страницы (адресно, по индексной карте; полными ячейками).</summary>
        private void RedrawItemsOf(string pageKey)
        {
            if (!_list.IsHandleCreated)
                return;
            List<ListViewItem> items;
            if (_itemsByKey.TryGetValue(pageKey, out items))
                foreach (ListViewItem item in items)
                    InvalidateItem(item.Index);
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

        // ---------- отрисовка (owner-draw) ----------

        /// <summary>Логические пиксели → физические (сетка живёт в физических: и ячейки, и плитки).</summary>
        private int Px(int logical)
        {
            return (int)Math.Round(logical * (DeviceDpi / 96.0));
        }

        /// <summary>
        /// Настроить геометрию под текущий масштаб: иконная зона (хитбокс) —
        /// min(плитка, 256), шаг ячеек — LVM_SETICONSPACING; затем переразложить.
        /// </summary>
        private void ApplyIconSpacing()
        {
            if (!_list.IsHandleCreated)
                return;
            Size tile = ThumbZoom.TileSize(_tileWidth);
            _metrics.ImageSize = new Size(
                Math.Min(256, Px(tile.Width)),
                Math.Min(256, Px(tile.Height)));
            Size cell = ThumbZoom.CellSize(_tileWidth);
            int cx = Px(cell.Width), cy = Px(cell.Height);
            SendMessage(_list.Handle, LVM_SETICONSPACING, IntPtr.Zero, (IntPtr)((cy << 16) | (cx & 0xFFFF)));
            _list.ArrangeIcons();
            _list.Invalidate();
        }

        /// <summary>
        /// Рамка ПЛИТКИ по нативной иконной зоне элемента: горизонтально по её центру,
        /// вертикально от её верха (иконная зона может быть меньше плитки — хитбокс
        /// ограничен 256). Чистая — под тест.
        /// </summary>
        internal static Rectangle TileRectFromIcon(Rectangle iconRect, Size tilePx)
        {
            int centerX = iconRect.Left + iconRect.Width / 2;
            return new Rectangle(centerX - tilePx.Width / 2, iconRect.Top, tilePx.Width, tilePx.Height);
        }

        /// <summary>
        /// Рамка ЯЧЕЙКИ вокруг плитки — единая геометрия для отрисовки И инвалидации:
        /// при крупном зуме плитка больше нативных границ элемента (иконная зона ≤ 256),
        /// поэтому перерисовывать надо именно ячейку, иначе низ плитки с hover-кнопками
        /// не попадает в регион отрисовки и «съедается». Чистая — под тест.
        /// </summary>
        internal static Rectangle CellRectFromTile(Rectangle tileRect, Size cellPx, int topPad)
        {
            return new Rectangle(
                tileRect.Left - (cellPx.Width - tileRect.Width) / 2,
                tileRect.Top - topPad,
                cellPx.Width,
                cellPx.Height);
        }

        /// <summary>Рамки плитки и ячейки элемента (по нативной иконной зоне). Только с созданным handle.</summary>
        private void ItemGeometry(int index, out Rectangle tileRect, out Rectangle cellRect)
        {
            Size tileLogical = ThumbZoom.TileSize(_tileWidth);
            var tilePx = new Size(Px(tileLogical.Width), Px(tileLogical.Height));
            Size cellLogical = ThumbZoom.CellSize(_tileWidth);
            var cellPx = new Size(Px(cellLogical.Width), Px(cellLogical.Height));
            tileRect = TileRectFromIcon(_list.GetItemRect(index, ItemBoundsPortion.Icon), tilePx);
            cellRect = CellRectFromTile(tileRect, cellPx, Px(4));
        }

        /// <summary>Перерисовать ПОЛНУЮ ячейку элемента (RedrawItems инвалидирует лишь нативные границы ≤ 256).</summary>
        private void InvalidateItem(int index)
        {
            if (!_list.IsHandleCreated || index < 0 || index >= _list.Items.Count)
                return;
            Rectangle tile, cell;
            ItemGeometry(index, out tile, out cell);
            _list.Invalidate(cell);
        }

        private void OnDrawItem(object sender, DrawListViewItemEventArgs e)
        {
            // Ошибка в отрисовке не должна валить процесс паint-штормом — рисуем, что смогли.
            try
            {
                DrawItemCore(e);
            }
            catch { }
        }

        private void DrawItemCore(DrawListViewItemEventArgs e)
        {
            Graphics g = e.Graphics;
            var page = e.Item.Tag as PdfPageRef;

            // Геометрия — от нативной иконной зоны (якорь ячейки), а не от e.Bounds:
            // при плитке крупнее 256 нативные границы элемента меньше нарисованного.
            // Та же геометрия используется инвалидацией (InvalidateItem) — DRY.
            Rectangle tileRect, cell;
            ItemGeometry(e.ItemIndex, out tileRect, out cell);

            g.FillRectangle(BackBrush, cell);

            bool selected = e.Item.Selected;
            if (selected)
            {
                Rectangle sel = Rectangle.Inflate(tileRect, Px(4), Px(4));
                g.FillRectangle(SelectionFill(), sel);
                g.DrawRectangle(SystemPens.Highlight, sel.X, sel.Y, sel.Width - 1, sel.Height - 1);
            }

            Bitmap tile = page != null ? GetTile(page) : null;
            if (tile != null)
                g.DrawImage(tile, tileRect.Location); // плитка составлена ровно под tileRect
            else
                DrawPlaceholder(g, tileRect);

            if (page != null && _cutRefs.Contains(page))
                g.FillRectangle(DimBrush, tileRect); // вырезанные приглушены до вставки (Esc — отмена)

            if (page != null && page.Rotation != 0)
                DrawRotationBadge(g, tileRect, page.Rotation);

            // Hover-кнопки поворота на плитке под курсором (как в Acrobat DC).
            if (e.ItemIndex == _hotIndex && AllowRotate && !Locked)
                DrawHoverRotateButtons(g, tileRect);

            // Подпись-номер в полосе под плиткой.
            var labelRect = new Rectangle(cell.X, tileRect.Bottom + Px(2), cell.Width, cell.Bottom - tileRect.Bottom - Px(2));
            Color labelColor = page != null && _cutRefs.Contains(page) ? SystemColors.GrayText : ForeColor;
            TextRenderer.DrawText(g, e.Item.Text, Font, labelRect, labelColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.Top | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine);

            if (e.Item.Focused && _list.Focused)
                ControlPaint.DrawFocusRectangle(g, Rectangle.Inflate(tileRect, Px(6), Px(6)));
        }

        /// <summary>Индекс элемента, в НАРИСОВАННУЮ плитку которого попадает точка; -1 — мимо всех.</summary>
        private int FindTileHit(Point pt)
        {
            int count = _list.Items.Count;
            if (count == 0)
                return -1;
            Size tilePx = new Size(Px(ThumbZoom.TileSize(_tileWidth).Width), Px(ThumbZoom.TileSize(_tileWidth).Height));
            if (tilePx.Width <= _metrics.ImageSize.Width && tilePx.Height <= _metrics.ImageSize.Height)
                return -1; // плитка целиком в хитбоксе — компенсация не нужна
            // Кандидаты — только видимые (окно по бинарному поиску, как у ленивого рендера).
            int first, last;
            VisibleRange(count,
                delegate(int i) { return _list.Items[i].Bounds.Top; },
                delegate(int i) { return _list.Items[i].Bounds.Bottom; },
                _list.ClientSize.Height, out first, out last);
            if (first > last)
                return -1;
            for (int i = Math.Max(0, first - 2); i <= Math.Min(count - 1, last + 2); i++)
                if (TileRectFromIcon(_list.GetItemRect(i, ItemBoundsPortion.Icon), tilePx).Contains(pt))
                    return i;
            return -1;
        }

        /// <summary>
        /// Рамки двух hover-кнопок поворота (↺ влево, ↻ вправо) у нижней кромки плитки,
        /// по центру, не пересекаются. Чистая — под тест.
        /// </summary>
        internal static Rectangle[] HoverRotateButtons(Rectangle tileRect, int diameter, int gap)
        {
            int totalW = diameter * 2 + gap;
            int x = tileRect.Left + (tileRect.Width - totalW) / 2;
            int y = tileRect.Bottom - diameter - gap;
            return new[]
            {
                new Rectangle(x, y, diameter, diameter),
                new Rectangle(x + diameter + gap, y, diameter, diameter)
            };
        }

        private Rectangle[] HoverRotateButtonsFor(int index)
        {
            Rectangle tile, cell;
            ItemGeometry(index, out tile, out cell);
            return HoverRotateButtons(tile, Px(24), Px(6));
        }

        private void DrawHoverRotateButtons(Graphics g, Rectangle tileRect)
        {
            Rectangle[] buttons = HoverRotateButtons(tileRect, Px(24), Px(6));
            string[] glyphs = { "↺", "↻" }; // ↺ влево, ↻ вправо
            SmoothingMode prev = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            for (int i = 0; i < 2; i++)
            {
                g.FillEllipse(ChipFill, buttons[i]);
                TextRenderer.DrawText(g, glyphs[i], _chipFont, buttons[i], Color.White,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
            }
            g.SmoothingMode = prev;
        }

        /// <summary>Клик по hover-кнопке поворота: вращает страницу ПОД кнопкой. true — клик обработан.</summary>
        private bool TryHoverRotateClick(Point pt)
        {
            if (_hotIndex < 0 || _hotIndex >= _list.Items.Count || !AllowRotate || Locked)
                return false;
            Rectangle[] buttons = HoverRotateButtonsFor(_hotIndex);
            int delta = buttons[0].Contains(pt) ? -90 : (buttons[1].Contains(pt) ? 90 : 0);
            if (delta == 0)
                return false;
            RotateItems(new[] { _hotIndex }, delta);
            return true;
        }

        /// <summary>Сменить плитку под курсором и перерисовать прежнюю/новую (полными ячейками — hover-кнопки у низа плитки).</summary>
        private void SetHotIndex(int index)
        {
            if (index == _hotIndex)
                return;
            int old = _hotIndex;
            _hotIndex = index;
            InvalidateItem(old);
            InvalidateItem(_hotIndex);
        }

        /// <summary>Бейдж угла поворота («90°») в правом верхнем углу плитки.</summary>
        private void DrawRotationBadge(Graphics g, Rectangle tileRect, int rotation)
        {
            string text = rotation + "°";
            Size sz = TextRenderer.MeasureText(g, text, _badgeFont, Size.Empty, TextFormatFlags.NoPadding);
            int padX = Px(4), padY = Px(2);
            var badge = new Rectangle(tileRect.Right - sz.Width - 2 * padX - Px(3), tileRect.Y + Px(3),
                sz.Width + 2 * padX, sz.Height + 2 * padY);
            g.FillRectangle(BadgeFill, badge);
            TextRenderer.DrawText(g, text, _badgeFont, badge, Color.White,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        }

        private static void DrawPlaceholder(Graphics g, Rectangle tileRect)
        {
            int w = (int)((tileRect.Height - 24) * 0.72f);
            int h = tileRect.Height - 24;
            if (w < 8 || h < 8)
                return;
            int x = tileRect.X + (tileRect.Width - w) / 2;
            int y = tileRect.Y + 12;
            g.FillRectangle(Brushes.White, x, y, w, h);
            g.DrawRectangle(PlaceholderPen, x, y, w, h);
        }

        /// <summary>Заливка выделения (SystemColors.Highlight с прозрачностью) — лениво, до смены темы.</summary>
        private Brush SelectionFill()
        {
            if (_selFill == null)
                _selFill = new SolidBrush(Color.FromArgb(45, SystemColors.Highlight));
            return _selFill;
        }

        protected override void OnSystemColorsChanged(EventArgs e)
        {
            // Цвет выделения мог смениться вместе с темой — кэш пересоздастся при отрисовке.
            if (_selFill != null)
            {
                _selFill.Dispose();
                _selFill = null;
            }
            base.OnSystemColorsChanged(e);
        }

        /// <summary>
        /// Плитка страницы С ЕЁ поворотом: из словаря, либо составляется лениво из кэша
        /// отрендеренных страниц; страницы нет в кэше — null (рисуется заглушка, рендер докачает).
        /// </summary>
        private Bitmap GetTile(PdfPageRef page)
        {
            string tileKey = TileKey(page);
            Bitmap tile;
            if (_tiles.TryGetValue(tileKey, out tile))
                return tile;
            CachedPage cached;
            if (!_pageCache.TryPeek(ThumbKey(page), out cached))
                return null;
            Size logical = ThumbZoom.TileSize(_tileWidth);
            tile = ComposeTile(cached.Bmp, new Size(Px(logical.Width), Px(logical.Height)), page.Rotation);
            _tiles[tileKey] = tile;
            return tile;
        }

        /// <summary>Сбросить все плитки (смена масштаба): пересоставятся лениво при отрисовке.</summary>
        private void ClearTiles()
        {
            foreach (Bitmap tile in _tiles.Values)
                tile.Dispose();
            _tiles.Clear();
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
            Ui.SelectAllItems(_list);
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

        /// <summary>
        /// Задать масштаб плиток: плитки пересоставляются из кэша лениво при отрисовке;
        /// при крупном зуме видимые страницы дорендериваются в увеличенную ширину.
        /// </summary>
        public void SetTileWidth(int width)
        {
            width = ThumbZoom.Clamp(width);
            _wheelPending = 0; // цель колеса применена (или перебита регулятором)
            if (width == _tileWidth)
                return;
            _tileWidth = width;
            _renderTarget = width > ThumbZoom.BaseWidth ? _renderHi : _renderBase;
            ClearTiles();
            ApplyIconSpacing();
            ScheduleVisibleUpdate(); // дорендер видимых при переходе в крупный зум
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
            }
            _list.Invalidate(); // приглушение рисуется в DrawItem
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
                CancelCut();      // приглушение снимаем сразу: даже если перенос не состоится, серых плиток не останется
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

        /// <summary>
        /// Перенести ВЫБРАННЫЕ страницы ПЕРЕД позицией insertAt («после страницы N» = N):
        /// тот же конвейер, что у вырезать-вставить, но без буфера — форма получает
        /// MoveRangeRequested и правит модель.
        /// </summary>
        public void MoveSelectedTo(int insertAt)
        {
            if (Locked || !AllowReorder || _list.SelectedIndices.Count == 0)
                return;
            if (insertAt < 0) insertAt = 0;
            if (insertAt > _list.Items.Count) insertAt = _list.Items.Count;
            var h = MoveRangeRequested;
            if (h != null)
                h(GetSelectedIndices(), insertAt);
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

        /// <summary>Отменить вырезание (Esc): снять приглушение, очистить буфер. true — было что отменять.</summary>
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
            _cutRefs.Clear();
            _list.Invalidate(); // приглушение рисуется в DrawItem — просто перерисовать
        }

        private void ClearClipboard()
        {
            _clipboard.Clear();
            _cutRefs.Clear();
            _clipIsCut = false;
        }

        // ---------- каретка вставки и подсветка зазора ----------

        private void OnListMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;
            if (TryHoverRotateClick(e.Location))
                return; // клик пришёлся в hover-кнопку поворота
            if (_list.HitTest(e.Location).Item != null)
            {
                ClearCaret(); // клик по плитке: якорь вставки — выделение
                return;
            }
            int ringHit = FindTileHit(e.Location);
            if (ringHit >= 0)
            {
                // «Кольцо» крупной плитки вне нативного хитбокса (его предел 256): нативная
                // обработка уже сбросила выделение на MouseDown — выделяем страницу здесь,
                // как обычный клик (модификаторные клики по кольцу не эмулируем).
                ClearCaret();
                if (!Locked && ModifierKeys == Keys.None)
                    SelectIndex(ringHit);
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
            _hoverMark = false;
            ShowMark(index, after, SystemColors.Highlight);
        }

        /// <summary>Точка попадает в полосу номера под плиткой (не в саму плитку). Чистая — под тест.</summary>
        internal static bool IsOnLabel(Point pt, Rectangle tileRect, Rectangle cell)
        {
            return pt.Y > tileRect.Bottom && pt.Y <= cell.Bottom && pt.X >= cell.Left && pt.X <= cell.Right;
        }

        /// <summary>Что делает двойной клик по элементу сетки. Чистая — под тест.</summary>
        internal enum DoubleClickAction
        {
            Preview,    // по плитке — полноразмерный предпросмотр
            MoveAfter,  // по номеру под плиткой — «Переместить после…»
            RotateChip  // по hover-кнопке поворота — проглотить (крутит каждый MouseUp)
        }

        /// <summary>
        /// Классификация двойного клика: hover-кнопка поворота (если показана) старше
        /// номера, номер — старше плитки. hoverChips == null — кнопки не показаны.
        /// Чистая — под тест.
        /// </summary>
        internal static DoubleClickAction ClassifyDoubleClick(Point pt, Rectangle tileRect, Rectangle cell,
            Rectangle[] hoverChips, bool allowMoveAfter)
        {
            if (hoverChips != null)
                foreach (Rectangle chip in hoverChips)
                    if (chip.Contains(pt))
                        return DoubleClickAction.RotateChip;
            if (allowMoveAfter && IsOnLabel(pt, tileRect, cell))
                return DoubleClickAction.MoveAfter;
            return DoubleClickAction.Preview;
        }

        /// <summary>
        /// Двойной клик: по НОМЕРУ (в редактируемой сетке) — перенос страницы по номеру
        /// (как контекстное «Переместить после…»); по самой ПЛИТКЕ — полноразмерный
        /// предпросмотр страницы. Обе возможности доступны всем PDF-инструментам (DRY).
        /// Клик по hover-кнопке поворота предпросмотр НЕ открывает: поворот делает каждый
        /// MouseUp пары, поэтому быстрый двойной клик по ↻ честно даёт 180°.
        /// </summary>
        private void OnListDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left || _list.Items.Count == 0)
                return;
            ListViewHitTestInfo hit = _list.HitTest(e.Location);
            int index = hit.Item != null ? hit.Item.Index : FindTileHit(e.Location);
            if (index < 0)
                return;
            var page = _list.Items[index].Tag as PdfPageRef;
            if (page == null)
                return;

            Rectangle tileRect, cell;
            ItemGeometry(index, out tileRect, out cell);
            Rectangle[] chips = index == _hotIndex && AllowRotate && !Locked ? HoverRotateButtonsFor(index) : null;
            switch (ClassifyDoubleClick(e.Location, tileRect, cell, chips, AllowReorder && !Locked))
            {
                case DoubleClickAction.RotateChip:
                    return; // второй MouseUp пары сам довернёт страницу ещё на 90°
                case DoubleClickAction.MoveAfter:
                    // Двойной клик по номеру = быстрый перенос ЭТОЙ страницы (выделяем её и просим диалог).
                    SelectIndex(index);
                    var h = MoveAfterRequested;
                    if (h != null) h(this, EventArgs.Empty);
                    return;
            }
            // Иначе — предпросмотр страницы в полный размер. Поворот из просмотра идёт тем же
            // путём, что и из сетки (чекпойнт для Ctrl+Z, чистка лишних плиток, перерисовка).
            // Пока просмотр модален, состав сетки не меняется, поэтому индекс остаётся верным.
            int previewIndex = index;
            Action<int> rotate = null;
            if (AllowRotate && !Locked)
                rotate = delegate(int delta) { RotateItems(new[] { previewIndex }, delta); };
            PagePreviewForm.Show(FindForm(), page, string.Format(Loc.T("preview.title"), _list.Items[index].Text), rotate);
        }

        /// <summary>
        /// Наведение: плитка под курсором получает hover-кнопки поворота, а зазор —
        /// светлую метку там, куда лёг бы клик каретки.
        /// </summary>
        private void OnListMouseMove(object sender, MouseEventArgs e)
        {
            // Плитка под курсором (нативный хит или «кольцо» крупной плитки).
            ListViewHitTestInfo hit = _list.HitTest(e.Location);
            int hot = hit.Item != null ? hit.Item.Index : FindTileHit(e.Location);
            SetHotIndex(e.Button == MouseButtons.None && !Locked ? hot : -1);

            if (Locked || !AllowReorder || _caretIndex >= 0 || _list.Items.Count == 0 || e.Button != MouseButtons.None)
            {
                ClearHoverMark();
                return;
            }
            if (hot >= 0)
            {
                ClearHoverMark(); // курсор над плиткой — не над зазором
                return;
            }
            int index = _list.InsertionMark.NearestIndex(e.Location);
            bool after;
            if (index < 0)
            {
                index = _list.Items.Count - 1;
                after = true;
            }
            else
            {
                Rectangle bounds = _list.GetItemRect(index);
                after = e.Location.X > bounds.Left + bounds.Width / 2;
            }
            _hoverMark = true;
            ShowMark(index, after, HoverMarkColor);
        }

        private void ShowMark(int index, bool after, Color color)
        {
            _list.InsertionMark.Color = color;
            _list.InsertionMark.AppearsAfterItem = after;
            _list.InsertionMark.Index = index;
        }

        private void ClearHoverMark()
        {
            if (!_hoverMark)
                return;
            _hoverMark = false;
            if (_caretIndex < 0)
                _list.InsertionMark.Index = -1;
        }

        private void ClearCaret()
        {
            _caretIndex = -1;
            _caretAfter = false;
            _hoverMark = false;
            _list.InsertionMark.Index = -1;
        }

        // ---------- поворот страниц ----------

        /// <summary>
        /// Повернуть выбранные страницы на delta градусов по часовой (±90). Поворот —
        /// свойство страницы В ИТОГОВОМ файле (модель делит эти же ссылки); исходный
        /// PDF не меняется. Плитки обновляются из кэша без повторного рендера.
        /// </summary>
        public void RotateSelected(int delta)
        {
            if (_list.SelectedIndices.Count == 0)
                return;
            RotateItems(GetSelectedIndices(), delta);
        }

        /// <summary>Массовый поворот: ВСЕ страницы сетки на delta по часовой (пункт меню «все страницы»).</summary>
        public void RotateAll(int delta)
        {
            int count = _list.Items.Count;
            if (count == 0)
                return;
            var all = new int[count];
            for (int i = 0; i < count; i++)
                all[i] = i;
            RotateItems(all, delta);
        }

        private void RotateItems(int[] indices, int delta)
        {
            if (Locked || !AllowRotate || indices.Length == 0)
                return;
            var hb = BeforeRotate; // формы с моделью порядка снимают чекпойнт для Ctrl+Z
            if (hb != null)
                hb(this, EventArgs.Empty);
            var pageKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (int index in indices)
            {
                var page = _list.Items[index].Tag as PdfPageRef;
                if (page == null)
                    continue;
                page.Rotation = PdfPageRef.ComposeRotation(page.Rotation, delta);
                pageKeys.Add(ThumbKey(page));
            }
            foreach (string pageKey in pageKeys)
                PruneUnusedRotations(pageKey);
            _list.Invalidate(); // плитки нужных поворотов пересоставятся лениво при отрисовке
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
            foreach (KeyValuePair<string, Bitmap> kv in _tiles)
                if (kv.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && !used.Contains(kv.Key))
                    stale.Add(kv.Key);
            foreach (string tileKey in stale)
            {
                _tiles[tileKey].Dispose();
                _tiles.Remove(tileKey);
            }
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
            _list.InsertionMark.Color = SystemColors.Highlight; // метка драга — в полном цвете
            SetDropActive(_dragHasFiles); // акцентная рамка только для импорта файлов, не для перестановки
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
            SetDropActive(false); // снять рамку до обработки дропа

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
            Size cell = ThumbZoom.CellSize(_tileWidth);
            int step = Math.Max(24, Px(cell.Height) / 3);
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
            _miMoveAfter = AddMenuItem(Loc.T("grid.menu.moveAfter"), null, delegate { var h = MoveAfterRequested; if (h != null) h(this, EventArgs.Empty); });
            _menu.Items.Add(new ToolStripSeparator());
            // Поворот — подменю: выбранные и «все страницы» (массовый поворот документа).
            _miRotate = new ToolStripMenuItem(Loc.T("grid.menu.rotate"));
            _miRotateRight = AddSubItem(_miRotate, Loc.T("grid.menu.rotateRight"), Loc.T("grid.key.rotateRight"), delegate { RotateSelected(90); });
            _miRotateLeft = AddSubItem(_miRotate, Loc.T("grid.menu.rotateLeft"), Loc.T("grid.key.rotateLeft"), delegate { RotateSelected(-90); });
            _miRotate.DropDownItems.Add(new ToolStripSeparator());
            _miRotateAllRight = AddSubItem(_miRotate, Loc.T("grid.menu.rotateAllRight"), null, delegate { RotateAll(90); });
            _miRotateAllLeft = AddSubItem(_miRotate, Loc.T("grid.menu.rotateAllLeft"), null, delegate { RotateAll(-90); });
            _menu.Items.Add(_miRotate);
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

        private static ToolStripMenuItem AddSubItem(ToolStripMenuItem parent, string text, string shortcutHint, EventHandler onClick)
        {
            var mi = new ToolStripMenuItem(text);
            if (shortcutHint != null)
                mi.ShortcutKeyDisplayString = shortcutHint;
            mi.Click += onClick;
            parent.DropDownItems.Add(mi);
            return mi;
        }

        private void OnMenuOpening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            bool hasSelection = _list.SelectedIndices.Count > 0;
            bool hasPages = _list.Items.Count > 0;
            _miCut.Visible = _miCopy.Visible = _miPaste.Visible = _miMoveAfter.Visible = _miDelete.Visible = AllowReorder;
            _miCut.Enabled = _miCopy.Enabled = _miMoveAfter.Enabled = _miDelete.Enabled = !Locked && hasSelection;
            _miPaste.Enabled = !Locked && ClipboardAvailable;
            _miRotate.Visible = AllowRotate;
            _miRotate.Enabled = !Locked && hasPages;
            _miRotateRight.Enabled = _miRotateLeft.Enabled = !Locked && hasSelection;
            _miRotateAllRight.Enabled = _miRotateAllLeft.Enabled = !Locked && hasPages;
            _miGoTo.Enabled = hasPages;
            // Разделители между скрытыми группами не нужны. Порядок пунктов:
            // 0 cut, 1 copy, 2 paste, 3 moveAfter, 4 sep, 5 rotate, 6 sep, 7 delete, 8 goto.
            bool anyEdit = AllowReorder;
            bool anyRotate = AllowRotate;
            _menu.Items[4].Visible = anyEdit && anyRotate;       // разделитель буфер | поворот
            _menu.Items[6].Visible = anyEdit || anyRotate;       // разделитель | удалить/перейти
            e.Cancel = _menu.Items.Count == 0;
        }

        // ---------- масштаб ----------

        // Целевая ширина Ctrl+колеса, ещё не применённая троттлингом регулятора (0 — нет).
        // Пока цель не применена, следующие щелчки шагают от неё, а не от устаревшей
        // _tileWidth — иначе быстрое вращение теряло бы шаги.
        private int _wheelPending;

        /// <summary>База шага колеса: неприменённая цель прошлых щелчков или текущая ширина. Чистая — под тест.</summary>
        internal static int WheelBasis(int pendingWidth, int currentWidth)
        {
            return pendingWidth > 0 ? pendingWidth : currentWidth;
        }

        private void OnListMouseWheel(object sender, MouseEventArgs e)
        {
            if ((ModifierKeys & Keys.Control) == 0)
                return; // без Ctrl — обычная прокрутка
            var handled = e as HandledMouseEventArgs;
            if (handled != null)
                handled.Handled = true;
            int basis = WheelBasis(_wheelPending, _tileWidth);
            int newWidth = ThumbZoom.StepFromWheel(basis, e.Delta);
            if (newWidth == basis)
                return; // упёрлись в границу масштаба
            var h = ZoomChanged;
            if (h != null)
            {
                // Пересборка плиток (ClearTiles + Invalidate) — дорогая: не на каждый щелчок,
                // а через регулятор и его общий 60-мс троттлинг (как и перетаскивание регулятора).
                // Подпись «%» при этом обновляется мгновенно из значения регулятора.
                _wheelPending = newWidth;
                h(newWidth);
            }
            else
                SetTileWidth(newWidth); // сетка без регулятора — применяем сразу
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
            int target = _renderTarget;
            CachedPage cached;
            // Попадание освежает LRU (видимые не вытесняются); перерендер нужен, только
            // если закэшированная ширина меньше требуемой текущим масштабом.
            if (_pageCache.TryGet(key, out cached) && cached.Width >= target)
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
                        // Флаг перечитываем ПОСЛЕ Reset и ДО ожидания. StopRendering ставит
                        // _thumbStop и зовёт Set() НЕ под _qLock, поэтому его сигнал мог
                        // прийти между проверкой цикла и нашим Reset — и быть им стёрт.
                        // Без этой проверки поток уснул бы навсегда: Dispose провисел бы на
                        // Join(2000) и не освободил рендерер с копиями PDF в памяти.
                        if (_thumbStop)
                            break;
                        _thumbSignal.Wait();
                        continue;
                    }
                    int width = _renderTarget; // текущая целевая ширина (могла смениться зумом)
                    Bitmap page = renderer.Render(req.SourcePath, req.PageIndex, width);
                    if (page == null)
                    {
                        // Сбой рендера (файл занят антивирусом/сетью и т.п.): снять ключ дедупа,
                        // иначе страница осталась бы заглушкой НАВСЕГДА — EnqueueThumb больше не
                        // поставил бы её в очередь. Ближайший тик видимых повторит заявку.
                        lock (_qLock)
                            _thumbRequested.Remove(ThumbKey(req));
                        continue;
                    }
                    PostPage(req, page, width);
                }
            }
            finally
            {
                renderer.Dispose();
            }
        }

        private void PostPage(PdfPageRef req, Bitmap page, int width)
        {
            try
            {
                if (IsHandleCreated && !IsDisposed)
                    BeginInvoke((MethodInvoker)delegate { ApplyPage(req, page, width); });
                else
                    page.Dispose();
            }
            catch (InvalidOperationException)
            {
                page.Dispose();
            }
        }

        private void ApplyPage(PdfPageRef req, Bitmap page, int width)
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
                if (existing.Width >= width)
                {
                    page.Dispose(); // уже есть не хуже (гонка зум-аута) — свежий не нужен
                    return;
                }
                // Дорендер крупнее: прежний рендер и его плитки заменяются.
                CachedPage old;
                if (_pageCache.Remove(key, out old))
                    old.Bmp.Dispose();
                RemoveTilesOf(key);
            }
            _pageCache.Add(key, new CachedPage { Key = key, Bmp = page, Width = width });
            RedrawItemsOf(key); // плитка составится лениво при отрисовке
            // Пока страница рендерилась, масштаб мог вырасти: заявку на больший размер
            // дедуп подавил (ключ был занят ЭТОЙ заявкой), поэтому просим обновить видимые
            // ещё раз. Иначе плитка осталась бы растянутой из мелкого рендера до тех пор,
            // пока пользователь случайно не прокрутит сетку.
            if (width < _renderTarget)
                ScheduleVisibleUpdate();
        }

        private static Bitmap ComposeTile(Bitmap page, Size tile, int rotation)
        {
            if (rotation == 0)
                return ComposeTileCore(page, tile);
            // RotateFlip мутирует изображение — поворачиваем КОПИЮ, кэш остаётся неповёрнутым.
            using (var rotated = (Bitmap)page.Clone())
            {
                rotated.RotateFlip(PageRotation.FlipFor(rotation));
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
                g.DrawRectangle(TileBorderPen, x, y, w - 1, h - 1);
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
                // Шрифты — общие кэшированные (Ui.Font), грид их не освобождает.
                if (_selFill != null)
                    _selFill.Dispose();
                _metrics.Dispose(); // пустой ImageList метрик (не дочерний компонент)
                _shuttingDown = true;    // вытеснение при Clear только освобождает bitmap
                _pageCache.Clear();
                ClearTiles();
                if (stopped)
                    _thumbSignal.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
