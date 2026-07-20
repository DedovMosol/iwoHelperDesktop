using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Переиспользуемая сетка миниатюр страниц PDF: фоновый рендер (Windows.Data.Pdf),
    /// масштаб (Ctrl+колесо и SetTileWidth), мультивыбор, при AllowReorder —
    /// перетаскивание с событием <see cref="ReorderRequested"/>. Модель порядка
    /// страниц держит владелец (форма) и передаёт список через <see cref="SetPages"/>.
    /// Общая для «Объединения» и «Разделения» PDF (DRY).
    /// </summary>
    public class PdfPageGrid : UserControl
    {
        private const string PlaceholderKey = "__ph";
        private const int EnqueueBuffer = 16; // докачивать миниатюры чуть за пределами видимого

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

        private readonly ScrollList _list = new ScrollList();
        private System.Windows.Forms.Timer _visibleTimer;
        private readonly Dictionary<string, Bitmap> _pageCache =
            new Dictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);
        // Ключи страниц, показываемых сейчас (обновляется в SetPages, только UI-поток).
        // Поздний результат рендера уже снятой страницы отбрасывается по этому набору.
        private HashSet<string> _currentKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // Элементы списка по ключу плитки: готовый рендер проставляется адресно, без
        // прохода по всем элементам (иначе рендер всего документа — O(n²)). Только UI-поток.
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

        /// <summary>Разрешить перетаскивание для смены порядка (для «Объединения»).</summary>
        public bool AllowReorder { get; set; }

        public event EventHandler SelectionChanged;
        /// <summary>Перетащили элемент from на позицию вставки to (0..Count).</summary>
        public event Action<int, int> ReorderRequested;
        /// <summary>Масштаб изменён изнутри (Ctrl+колесо) — чтобы синхронизировать ползунок.</summary>
        public event Action<int> ZoomChanged;

        public PdfPageGrid()
        {
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
            _list.DragOver += OnListDragOver;
            _list.DragDrop += OnListDragDrop;
            _list.DragLeave += delegate { _list.InsertionMark.Index = -1; };
            _list.MouseWheel += OnListMouseWheel;
            _list.Scrolled += delegate { ScheduleVisibleUpdate(); };
            _list.Resize += delegate { ScheduleVisibleUpdate(); };
            _list.SelectedIndexChanged += delegate { ScheduleVisibleUpdate(); }; // навигация клавишами
            EnableDoubleBuffer(_list);
            Controls.Add(_list);

            // Троттлинг: события прокрутки сливаются в одно обновление видимых миниатюр.
            _visibleTimer = new System.Windows.Forms.Timer();
            _visibleTimer.Interval = 100;
            _visibleTimer.Tick += delegate { _visibleTimer.Stop(); UpdateVisibleThumbs(); };

            StartThumbWorker();
        }

        // ---------- публичный API ----------

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

            _list.BeginUpdate();
            _list.Items.Clear();         // после очистки ни один элемент не ссылается на плитки
            _itemsByKey.Clear();
            PruneCache(_currentKeys);    // освободить bitmap и плитку страниц вне набора
            if (pages != null)
            {
                foreach (PdfPageRef page in pages)
                {
                    string key = ThumbKey(page);
                    var item = new ListViewItem(MakeLabel(page));
                    item.Tag = page;
                    item.ToolTipText = page.SourcePath + " — стр. " + (page.PageIndex + 1);
                    item.ImageKey = _thumbs.Images.ContainsKey(key) ? key : PlaceholderKey;
                    _list.Items.Add(item);
                    IndexItem(key, item);
                }
            }
            _list.EndUpdate();
            ScheduleVisibleUpdate(); // рендерим только видимые страницы, а не все сразу
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

        /// <summary>Освобождает bitmap и плитку страниц, которых больше нет в наборе keep.</summary>
        private void PruneCache(ICollection<string> keepKeys)
        {
            foreach (string key in StaleKeys(_pageCache.Keys, keepKeys))
            {
                _pageCache[key].Dispose();
                _pageCache.Remove(key);
                if (_thumbs.Images.ContainsKey(key))
                    _thumbs.Images.RemoveByKey(key);
            }
        }

        /// <summary>Регистрирует элемент под ключом плитки (один ключ — несколько элементов при повторах страницы).</summary>
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

        // ---------- перетаскивание ----------

        private void OnItemDrag(object sender, ItemDragEventArgs e)
        {
            if (AllowReorder)
                _list.DoDragDrop(e.Item, DragDropEffects.Move);
        }

        private void OnListDragOver(object sender, DragEventArgs e)
        {
            if (!AllowReorder || !e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                e.Effect = DragDropEffects.None;
                return;
            }
            e.Effect = DragDropEffects.Move;
            Point pt = _list.PointToClient(new Point(e.X, e.Y));
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
            var item = e.Data.GetData(typeof(ListViewItem)) as ListViewItem;
            if (!AllowReorder || item == null || target < 0)
                return;
            int from = item.Index;
            int to = after ? target + 1 : target;
            var h = ReorderRequested;
            if (h != null)
                h(from, to);
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
            foreach (KeyValuePair<string, Bitmap> kv in _pageCache)
                fresh.Images.Add(kv.Key, ComposeTile(kv.Value, fresh.ImageSize));

            _list.BeginUpdate();
            _thumbs = fresh;
            _list.LargeImageList = fresh;
            foreach (ListViewItem item in _list.Items)
            {
                var p = item.Tag as PdfPageRef;
                string key = p != null ? ThumbKey(p) : null;
                item.ImageKey = key != null && fresh.Images.ContainsKey(key) ? key : PlaceholderKey;
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

        private void EnqueueThumb(PdfPageRef page)
        {
            string key = ThumbKey(page);
            if (_pageCache.ContainsKey(key))
                return; // уже отрендерено — не тревожим воркер и не переоткрываем документ
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
                    Bitmap page = renderer.Render(req.SourcePath, req.PageIndex, ThumbZoom.RenderWidth);
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
            if (!_currentKeys.Contains(key))
            {
                page.Dispose(); // страница уже снята из набора — поздний результат рендера отбрасываем
                return;
            }
            if (_pageCache.ContainsKey(key))
            {
                page.Dispose();
                return;
            }
            _pageCache[key] = page;

            if (!_thumbs.Images.ContainsKey(key))
                _thumbs.Images.Add(key, ComposeTile(page, _thumbs.ImageSize));

            // Адресно по ключу, а не проходом по всем элементам (рендер всего
            // документа иначе O(n²)).
            List<ListViewItem> items;
            if (_itemsByKey.TryGetValue(key, out items))
                foreach (ListViewItem item in items)
                    if (item.ImageKey != key)
                        item.ImageKey = key;
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

        private static string MakeLabel(PdfPageRef page)
        {
            string name = Path.GetFileNameWithoutExtension(page.FileName);
            if (name.Length > 18)
                name = name.Substring(0, 17) + "…";
            return name + "\nстр. " + (page.PageIndex + 1);
        }

        private static Bitmap ComposeTile(Bitmap page, Size tile)
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

        private static void EnableDoubleBuffer(ListView list)
        {
            var p = typeof(ListView).GetProperty("DoubleBuffered",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            if (p != null)
                p.SetValue(list, true, null);
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
                foreach (Bitmap page in _pageCache.Values)
                    page.Dispose();
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
