using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Инструмент «Объединение PDF»: сетка миниатюр страниц выбранных документов,
    /// масштаб (ползунок и Ctrl+колесо), перестановка кнопками и перетаскиванием,
    /// удаление, сохранение в один PDF. Страницы копируются без переконвертации
    /// (PDFsharp). Миниатюры рисует системный движок Windows.Data.Pdf в фоне;
    /// если он недоступен — показываются заглушки, инструмент работает как список.
    /// </summary>
    public class PdfMergeForm : Form
    {
        private const string Title = "Объединение PDF";
        private const string PlaceholderKey = "__ph";

        private readonly PdfPageOrder _order = new PdfPageOrder();
        // Отрендеренные страницы (в ширину RenderWidth), из них пересобираются
        // плитки при зуме. Только UI-поток.
        private readonly Dictionary<string, Bitmap> _pageCache =
            new Dictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);
        private ImageList _thumbs;
        private int _tileWidth = ThumbZoom.DefaultWidth;

        private ListView _list;
        private TrackBar _zoom;
        private System.Windows.Forms.Timer _zoomTimer;
        private Button _btnAdd;
        private Button _btnUp;
        private Button _btnDown;
        private Button _btnRemove;
        private Button _btnSave;
        private Label _lblStatus;
        private ToolTip _tips;
        private bool _busy; // идёт сохранение (только UI-поток)

        // Фоновый рендер миниатюр.
        private readonly object _qLock = new object();
        private readonly Queue<PdfPageRef> _thumbQueue = new Queue<PdfPageRef>();
        private readonly HashSet<string> _thumbRequested = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly ManualResetEventSlim _thumbSignal = new ManualResetEventSlim(false);
        private Thread _thumbThread;
        private volatile bool _thumbStop;

        public PdfMergeForm()
        {
            BuildUi();
            StartThumbWorker();
            UpdateButtons();
        }

        private void BuildUi()
        {
            Text = Title;
            Font = new Font("Segoe UI", 9.75f);
            BackColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(780, 620);
            MinimumSize = new Size(660, 500);
            ShowInTaskbar = false;
            AllowDrop = true;
            DragEnter += OnFileDragEnter;
            DragDrop += OnFileDragDrop;
            _tips = new ToolTip();

            MenuStrip menu = HelpMenu.Create(this, ShowHelp);
            MainMenuStrip = menu;
            Controls.Add(menu);

            int m = HelpMenu.Height; // содержимое ниже строки меню
            Ui.AccentBar(this, m);
            Ui.Label(this, "Объединение PDF", 20, m + 16,
                new Font("Segoe UI", 14f, FontStyle.Bold), Color.FromArgb(40, 40, 40));
            Ui.Label(this, "Перетаскивайте миниатюры, чтобы задать порядок; масштаб — ползунком или Ctrl+колесо.",
                22, m + 46, Font, Theme.TextMuted);

            int right = ClientSize.Width - 20;

            _thumbs = NewImageList(_tileWidth);

            _list = new ListView();
            _list.SetBounds(20, m + 80, right - 20 - 150, ClientSize.Height - (m + 80) - 112);
            _list.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _list.View = View.LargeIcon;
            _list.LargeImageList = _thumbs;
            _list.MultiSelect = true;
            _list.HideSelection = false;
            _list.LabelWrap = true;
            _list.BorderStyle = BorderStyle.FixedSingle;
            _list.BackColor = Color.FromArgb(250, 250, 250);
            _list.SelectedIndexChanged += delegate { UpdateButtons(); };
            _list.AllowDrop = true;
            _list.ItemDrag += OnItemDrag;
            _list.DragOver += OnListDragOver;
            _list.DragDrop += OnListDragDrop;
            _list.DragLeave += delegate { _list.InsertionMark.Index = -1; };
            _list.MouseWheel += OnListMouseWheel; // Ctrl+колесо = зум
            EnableDoubleBuffer(_list);
            Controls.Add(_list);

            int col = right - 130;
            _btnAdd = AddButton("Добавить PDF…", col, m + 80, 130, 32);
            _btnAdd.Click += OnAddClick;
            _tips.SetToolTip(_btnAdd, "Файлы также можно перетащить в окно");
            _btnUp = AddButton("◀ Раньше", col, m + 124, 130, 30);
            _btnUp.Click += delegate { MoveSelected(false); };
            _btnDown = AddButton("Позже ▶", col, m + 160, 130, 30);
            _btnDown.Click += delegate { MoveSelected(true); };
            _btnRemove = AddButton("Удалить", col, m + 204, 130, 30);
            _btnRemove.Click += OnRemoveClick;

            // Масштаб миниатюр
            Ui.Label(this, "Масштаб:", 20, ClientSize.Height - 104, Font, Theme.TextMuted)
                .Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            _zoom = new TrackBar();
            _zoom.SetBounds(85, ClientSize.Height - 108, 180, 30);
            _zoom.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            _zoom.Minimum = ThumbZoom.MinWidth;
            _zoom.Maximum = ThumbZoom.MaxWidth;
            _zoom.Value = _tileWidth;
            _zoom.TickFrequency = 32;
            _zoom.SmallChange = 16;
            _zoom.LargeChange = 32;
            _zoom.ValueChanged += delegate { ScheduleZoom(); };
            _tips.SetToolTip(_zoom, "Масштаб миниатюр (также Ctrl+колесо мыши)");
            Controls.Add(_zoom);

            // Троттлинг пересборки плиток при перетаскивании ползунка.
            _zoomTimer = new System.Windows.Forms.Timer();
            _zoomTimer.Interval = 60;
            _zoomTimer.Tick += OnZoomTick;

            var save = new RoundedButton(true);
            save.Text = "Сохранить PDF…";
            save.SetBounds(right - 190, ClientSize.Height - 58, 190, 38);
            save.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            save.Click += OnSaveClick;
            Controls.Add(save);
            _btnSave = save;
            AcceptButton = save;

            _lblStatus = Ui.Label(this, "Добавьте PDF-файлы — кнопкой или перетащив их в окно.",
                20, ClientSize.Height - 50, Font, Theme.TextMuted);
            _lblStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, Title, "Как пользоваться",
                "1. Добавьте PDF-файлы — кнопкой «Добавить PDF…» или перетащив их в окно.\n" +
                "2. Появится сетка миниатюр страниц. Масштаб — ползунком внизу или Ctrl+колесо мыши.\n" +
                "3. Задайте порядок: перетаскивайте миниатюры или используйте «◀ Раньше» / «Позже ▶».\n" +
                "   Лишние страницы удаляйте кнопкой «Удалить».\n" +
                "4. «Сохранить PDF…» соберёт один документ в выбранном порядке.\n\n" +
                "Страницы копируются как есть, без переконвертации — сканы, печати и подписи " +
                "не искажаются. Битые и защищённые паролем файлы пропускаются с причиной.");
        }

        private ImageList NewImageList(int tileWidth)
        {
            var list = new ImageList();
            list.ImageSize = ThumbZoom.TileSize(tileWidth);
            list.ColorDepth = ColorDepth.Depth32Bit;
            list.Images.Add(PlaceholderKey, MakePlaceholder(list.ImageSize));
            return list;
        }

        private Button AddButton(string text, int x, int y, int w, int h)
        {
            var b = new RoundedButton(false);
            b.Text = text;
            b.SetBounds(x, y, w, h);
            b.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(b);
            return b;
        }

        private static void EnableDoubleBuffer(ListView list)
        {
            var p = typeof(ListView).GetProperty("DoubleBuffered",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            if (p != null)
                p.SetValue(list, true, null);
        }

        // ---------- добавление файлов ----------

        private void OnAddClick(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Документы PDF (*.pdf)|*.pdf";
                dialog.Multiselect = true;
                dialog.Title = "Выберите PDF-файлы";
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    AddFiles(dialog.FileNames);
            }
        }

        private void OnFileDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = !_busy && ExtractPdfPaths(e).Length > 0
                ? DragDropEffects.Copy
                : DragDropEffects.None;
        }

        private void OnFileDragDrop(object sender, DragEventArgs e)
        {
            if (!_busy)
                AddFiles(ExtractPdfPaths(e));
        }

        private static string[] ExtractPdfPaths(DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return new string[0];
            var items = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (items == null)
                return new string[0];
            var pdfs = new List<string>();
            foreach (string item in items)
            {
                if (File.Exists(item) &&
                    string.Equals(Path.GetExtension(item), ".pdf", StringComparison.OrdinalIgnoreCase))
                    pdfs.Add(item);
            }
            return pdfs.ToArray();
        }

        private void AddFiles(string[] paths)
        {
            int added = 0;
            Cursor = Cursors.WaitCursor;
            try
            {
                foreach (string path in paths)
                {
                    try
                    {
                        int pages = PdfMergeService.LoadPages(path).Count;
                        _order.AddDocument(path, pages);
                        added += pages;
                    }
                    catch (MergeException ex)
                    {
                        Dialogs.Error(this, Title, "Файл не добавлен", ex.Message);
                    }
                }
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            if (added > 0)
            {
                RefreshList();
                SetStatus("Страниц в списке: " + _order.Count + ".", Theme.TextMuted);
            }
            UpdateButtons();
        }

        // ---------- перестановка и удаление ----------

        private void MoveSelected(bool later)
        {
            if (_list.SelectedIndices.Count != 1)
                return;
            int index = _list.SelectedIndices[0];
            int moved = later ? _order.MoveDown(index) : _order.MoveUp(index);
            if (moved == index)
                return;
            RefreshList();
            _list.Items[moved].Selected = true;
            _list.Items[moved].EnsureVisible();
            _list.Focus();
        }

        private void OnRemoveClick(object sender, EventArgs e)
        {
            if (_list.SelectedIndices.Count == 0)
                return;
            var indices = new int[_list.SelectedIndices.Count];
            _list.SelectedIndices.CopyTo(indices, 0);
            _order.RemoveAt(indices);
            RefreshList();
            SetStatus("Страниц в списке: " + _order.Count + ".", Theme.TextMuted);
            UpdateButtons();
        }

        private void OnItemDrag(object sender, ItemDragEventArgs e)
        {
            if (!_busy)
                _list.DoDragDrop(e.Item, DragDropEffects.Move);
        }

        private void OnListDragOver(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(ListViewItem)))
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
                // LargeIcon: плитки идут слева направо — вставка по горизонтали.
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
            if (item == null || target < 0)
                return;
            int from = item.Index;
            int to = after ? target + 1 : target;
            _order.Move(from, to);
            RefreshList();
            int landed = to > from ? to - 1 : to;
            if (landed >= 0 && landed < _list.Items.Count)
            {
                _list.Items[landed].Selected = true;
                _list.Items[landed].EnsureVisible();
            }
        }

        private void RefreshList()
        {
            _list.BeginUpdate();
            _list.Items.Clear();
            for (int i = 0; i < _order.Count; i++)
            {
                PdfPageRef page = _order[i];
                string key = ThumbKey(page);
                var item = new ListViewItem(MakeLabel(page));
                item.Tag = page;
                item.ToolTipText = page.SourcePath + " — стр. " + (page.PageIndex + 1);
                item.ImageKey = _thumbs.Images.ContainsKey(key) ? key : PlaceholderKey;
                _list.Items.Add(item);
                EnqueueThumb(page);
            }
            _list.EndUpdate();
        }

        private static string MakeLabel(PdfPageRef page)
        {
            string name = Path.GetFileNameWithoutExtension(page.FileName);
            if (name.Length > 18)
                name = name.Substring(0, 17) + "…";
            return name + "\nстр. " + (page.PageIndex + 1);
        }

        // ---------- масштаб ----------

        private void OnListMouseWheel(object sender, MouseEventArgs e)
        {
            if ((ModifierKeys & Keys.Control) == 0)
                return; // без Ctrl — обычная прокрутка списка
            int newWidth = ThumbZoom.StepFromWheel(_zoom.Value, e.Delta);
            var handled = e as HandledMouseEventArgs;
            if (handled != null)
                handled.Handled = true; // не прокручивать при зуме
            if (newWidth != _zoom.Value)
                _zoom.Value = newWidth; // -> ScheduleZoom
        }

        private void ScheduleZoom()
        {
            // Троттлинг: частые изменения (перетаскивание, серия щелчков колеса)
            // сливаются в одну пересборку плиток.
            _zoomTimer.Stop();
            _zoomTimer.Start();
        }

        private void OnZoomTick(object sender, EventArgs e)
        {
            _zoomTimer.Stop();
            if (_zoom.Value == _tileWidth)
                return;
            _tileWidth = _zoom.Value;
            RebuildTiles();
        }

        /// <summary>Пересобирает плитки из кэша страниц под текущий масштаб (без WinRT).</summary>
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

        // ---------- фоновый рендер миниатюр ----------

        private static string ThumbKey(PdfPageRef page)
        {
            return page.SourcePath.ToLowerInvariant() + "|" + page.PageIndex;
        }

        private void EnqueueThumb(PdfPageRef page)
        {
            string key = ThumbKey(page);
            lock (_qLock)
            {
                if (!_thumbRequested.Add(key)) // уже в очереди или готова
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
            catch { return; } // WinRT недоступен — заглушки останутся, инструмент работает

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
                page.Dispose(); // окно уже разрушено
            }
        }

        private void ApplyPage(PdfPageRef req, Bitmap page)
        {
            string key = ThumbKey(req);
            if (_pageCache.ContainsKey(key))
            {
                page.Dispose(); // уже есть (не должно, но защищаемся)
                return;
            }
            _pageCache[key] = page; // кэш владеет исходным изображением страницы

            if (!_thumbs.Images.ContainsKey(key))
                _thumbs.Images.Add(key, ComposeTile(page, _thumbs.ImageSize));

            foreach (ListViewItem item in _list.Items)
            {
                var p = item.Tag as PdfPageRef;
                if (p != null && item.ImageKey != key && ThumbKey(p) == key)
                    item.ImageKey = key;
            }
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

        // ---------- сохранение ----------

        private void OnSaveClick(object sender, EventArgs e)
        {
            if (_busy || _order.Count == 0)
                return;
            string outputPath;
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Документ PDF (*.pdf)|*.pdf";
                dialog.FileName = "Объединённый.pdf";
                dialog.InitialDirectory = Path.GetDirectoryName(_order[0].SourcePath);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                outputPath = dialog.FileName;
            }

            var pages = _order.ToList();
            _busy = true;
            UpdateButtons();
            SetStatus("Сохранение…", Theme.TextMuted);

            var thread = new Thread(delegate()
            {
                Exception error = null;
                try
                {
                    PdfMergeService.Merge(pages, outputPath);
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                try
                {
                    if (IsHandleCreated && !IsDisposed)
                        BeginInvoke((MethodInvoker)delegate { OnSaveFinished(error, outputPath, pages.Count); });
                }
                catch (InvalidOperationException) { }
            });
            thread.IsBackground = true;
            thread.Start();
        }

        private void OnSaveFinished(Exception error, string outputPath, int pageCount)
        {
            _busy = false;
            UpdateButtons();
            if (error != null)
            {
                SetStatus("PDF не сохранён.", Theme.ErrRed);
                Dialogs.Error(this, Title, "PDF не сохранён", error.Message);
                return;
            }
            SetStatus("✓ Сохранено страниц: " + pageCount + ".", Theme.OkGreen);
            try { Process.Start(outputPath); }
            catch { } // нет ассоциации PDF — файл всё равно сохранён
        }

        private void UpdateButtons()
        {
            bool hasSelection = _list.SelectedIndices.Count > 0;
            _btnAdd.Enabled = !_busy;
            _btnUp.Enabled = !_busy && _list.SelectedIndices.Count == 1;
            _btnDown.Enabled = !_busy && _list.SelectedIndices.Count == 1;
            _btnRemove.Enabled = !_busy && hasSelection;
            _btnSave.Enabled = !_busy && _order.Count > 0;
        }

        private void SetStatus(string text, Color color)
        {
            _lblStatus.Text = text;
            _lblStatus.ForeColor = color;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_busy)
            {
                SetStatus("Дождитесь завершения сохранения…", Theme.WarnOrange);
                e.Cancel = true;
                return;
            }
            _thumbStop = true;
            _thumbSignal.Set(); // разбудить воркер, чтобы он вышел и освободил рендерер
            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _thumbStop = true;
                _thumbSignal.Set();
                if (_zoomTimer != null)
                    _zoomTimer.Dispose();
                foreach (Bitmap page in _pageCache.Values)
                    page.Dispose();
                _pageCache.Clear();
                if (_thumbs != null)
                    _thumbs.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
