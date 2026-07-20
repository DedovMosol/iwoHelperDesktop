using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Инструмент «Разделение PDF»: открывает один документ, показывает его страницы
    /// сеткой миниатюр (<see cref="PdfPageGrid"/>) и в выбранном режиме извлекает
    /// выбранные страницы в один PDF либо разбивает документ на несколько (по
    /// диапазонам, каждые N страниц, по закладкам). Страницы копируются без
    /// переконвертации (PDFsharp); исходный файл не изменяется.
    /// </summary>
    public class PdfSplitForm : PdfToolFormBase
    {
        private const string Title = "Разделение PDF";
        private const int ModeExtract = 0, ModeRanges = 1, ModeEveryN = 2, ModeBookmarks = 3;

        // Сетка, зум, сжатие, статус, подсказки и флаг _busy — в базе PdfToolFormBase.
        private string _sourcePath;
        private int _pageCount;

        private Button _btnOpen;
        private ComboBox _cmbMode;
        private Label _lblRanges;
        private TextBox _txtRanges;
        private Label _lblN;
        private NumericUpDown _numN;
        private CheckBox _chkCombine;
        private Label _lblHint;
        private Button _btnDo;

        public PdfSplitForm() : this(null) { }

        public PdfSplitForm(Action showHub) : base(showHub)
        {
            BuildUi();
            UpdateModeInputs();
            UpdateControls();
        }

        private void BuildUi()
        {
            Text = Title;
            Icon icon = Ui.AppIcon();
            if (icon != null)
                Icon = icon;
            Font = new Font("Segoe UI", 9.75f);
            BackColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(800, 620);
            MinimumSize = new Size(700, 520);
            ShowInTaskbar = true;
            WindowChrome.Enable(this, Theme.PdfRed); // PDF-инструмент — красный заголовок
            AllowDrop = true;
            DragEnter += OnFileDragEnter;
            DragDrop += OnFileDragDrop;
            _tips = new ToolTip();

            MenuStrip menu = HelpMenu.Create(this, ShowHelp);
            MainMenuStrip = menu;
            Controls.Add(menu);

            int m = HelpMenu.Height;
            var header = new HeaderBand(Title,
                "Извлеките нужные страницы в один PDF или разбейте документ на несколько.",
                Theme.PdfRed, Theme.PdfRedDark);
            header.SetBounds(0, m, ClientSize.Width, 76);
            header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            header.TabIndex = 100;
            Controls.Add(header);
            if (_showHub != null)
            {
                Button home = Ui.HomeButton(_showHub);
                home.SetBounds(header.Width - 180, 22, 160, 30);
                home.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                _tips.SetToolTip(home, "Открыть экран выбора инструмента");
                header.Controls.Add(home);
            }

            int right = ClientSize.Width - 20;
            int panelW = 210;
            int gridBottom = ClientSize.Height - 112;

            _grid = new PdfPageGrid();
            _grid.AllowReorder = false; // разделение не меняет порядок исходника
            _grid.SetBounds(20, m + 84, right - 20 - panelW, gridBottom - (m + 84));
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _grid.SelectionChanged += delegate { UpdateControls(); };
            _grid.ZoomChanged += delegate(int w) { _zoom.Value = w; };
            Controls.Add(_grid);

            int px = right - panelW + 10; // левый край панели режима
            int pw = panelW - 10;
            _btnOpen = new RoundedButton(false);
            _btnOpen.Text = "Открыть PDF…";
            _btnOpen.SetBounds(px, m + 84, pw, 32);
            _btnOpen.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnOpen.Click += OnOpenClick;
            _tips.SetToolTip(_btnOpen, "Файл также можно перетащить в окно");
            Controls.Add(_btnOpen);

            Label lblMode = Ui.Label(this, "Режим:", px, m + 128, Font, Theme.TextPrimary);
            lblMode.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _cmbMode = new ComboBox();
            _cmbMode.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbMode.Items.AddRange(new object[] { "Извлечь выбранные", "По диапазонам", "Каждые N страниц", "По закладкам" });
            _cmbMode.SelectedIndex = ModeExtract;
            _cmbMode.SetBounds(px, m + 150, pw, 27);
            _cmbMode.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _cmbMode.SelectedIndexChanged += delegate { UpdateModeInputs(); UpdateControls(); };
            Controls.Add(_cmbMode);

            // Поля ввода режимов (в одном месте, показываются по режиму).
            _lblRanges = Ui.Label(this, "Диапазоны (напр. 1-3, 5, 8-):", px, m + 188, Font, Theme.TextMuted);
            _lblRanges.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _txtRanges = new TextBox();
            _txtRanges.SetBounds(px, m + 210, pw, 27);
            _txtRanges.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(_txtRanges);

            _lblN = Ui.Label(this, "Страниц в части:", px, m + 188, Font, Theme.TextMuted);
            _lblN.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _numN = new NumericUpDown();
            _numN.Minimum = 1;
            _numN.Maximum = 10000;
            _numN.Value = 1;
            _numN.SetBounds(px, m + 210, 70, 27);
            _numN.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(_numN);

            _lblHint = Ui.Label(this, "", px, m + 188, Font, Theme.TextMuted);
            _lblHint.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _lblHint.MaximumSize = new Size(pw, 0);
            _lblHint.AutoSize = true;

            // Только для «По диапазонам»: собрать все страницы в один файл.
            _chkCombine = new AccentCheckBox();
            _chkCombine.Text = "Объединить в один файл";
            _chkCombine.SetBounds(px, m + 244, pw, 22);
            _chkCombine.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _chkCombine.ForeColor = Theme.TextPrimary;
            _tips.SetToolTip(_chkCombine, "Все указанные страницы — в один PDF, а не по файлу на диапазон");
            _chkCombine.CheckedChanged += delegate { UpdateModeInputs(); };
            Controls.Add(_chkCombine);

            _btnDo = new RoundedButton(true);
            _btnDo.SetBounds(px, m + 284, pw, 38);
            _btnDo.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnDo.Click += OnDoClick;
            Controls.Add(_btnDo);

            // Масштаб миниатюр (как в «Объединении»).
            Ui.Label(this, "Масштаб:", 20, ClientSize.Height - 104, Font, Theme.TextMuted)
                .Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            _zoom = new TrackBar();
            _zoom.SetBounds(85, ClientSize.Height - 108, 180, 30);
            _zoom.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            _zoom.Minimum = ThumbZoom.MinWidth;
            _zoom.Maximum = ThumbZoom.MaxWidth;
            _zoom.Value = _grid.TileWidth;
            _zoom.TickFrequency = 32;
            _zoom.SmallChange = 16;
            _zoom.LargeChange = 32;
            _zoom.ValueChanged += delegate { ScheduleZoom(); };
            _tips.SetToolTip(_zoom, "Масштаб миниатюр (также Ctrl+колесо мыши)");
            Controls.Add(_zoom);

            _zoomTimer = new System.Windows.Forms.Timer();
            _zoomTimer.Interval = 60;
            _zoomTimer.Tick += delegate { _zoomTimer.Stop(); _grid.SetTileWidth(_zoom.Value); };

            _compress = new CompressionPicker();
            _compress.Location = new Point(right - _compress.Width, ClientSize.Height - 106);
            _compress.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Controls.Add(_compress);

            _lblStatus = Ui.Label(this, "Откройте PDF — кнопкой «Открыть PDF…» или перетащив его в окно.",
                20, ClientSize.Height - 50, Font, Theme.TextMuted);
            _lblStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, Title, "Как пользоваться",
                "1. Откройте PDF — кнопкой «Открыть PDF…» или перетащив его в окно; появится сетка страниц.\n" +
                "2. Выберите режим:\n" +
                "   • «Извлечь выбранные» — выделите страницы в сетке (Ctrl+A — все) → сохранит их в один PDF;\n" +
                "   • «По диапазонам» — «1-3, 5, 8-»: каждый диапазон → отдельный файл;\n" +
                "   • «Каждые N страниц» — равные части (1 — каждая страница отдельно);\n" +
                "   • «По закладкам» — по одному файлу на закладку верхнего уровня, имена из заголовков.\n" +
                "3. При необходимости выберите «Сжатие» (по умолчанию «Отлично» — без сжатия): " +
                "«Хорошо»/«Нормально» уменьшают размер за счёт понижения разрешения изображений " +
                "(как в Acrobat), текст сохраняется. Требуется Ghostscript.\n" +
                "4. Нажмите «Извлечь…»/«Разделить…» и укажите имя и папку для результата " +
                "(при разбиении к имени добавятся номера или метки).\n\n" +
                "Страницы копируются как есть, без переконвертации. Исходный файл не изменяется; " +
                "имена не перезаписываются (при совпадении добавляется номер).\n" +
                "Сжатие меняет содержимое файла, поэтому у подписанных PDF подпись станет " +
                "недействительной (как и при сжатии в Acrobat) — сжимайте до подписания.");
        }

        // ---------- открытие исходника ----------

        private void OnOpenClick(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Документы PDF (*.pdf)|*.pdf";
                dialog.Title = "Выберите PDF для разделения";
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    LoadSource(dialog.FileName);
            }
        }

        private void OnFileDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = !_busy && PdfDrop.ExtractPaths(e).Length > 0 ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void OnFileDragDrop(object sender, DragEventArgs e)
        {
            if (_busy)
                return;
            string[] paths = PdfDrop.ExtractPaths(e);
            if (paths.Length > 0)
                LoadSource(paths[0]); // разделение работает с одним документом
        }

        private void LoadSource(string path)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                _pageCount = PdfMergeService.LoadPages(path).Count;
            }
            catch (MergeException ex)
            {
                Cursor = Cursors.Default;
                Dialogs.Error(this, Title, "Файл не открыт", ex.Message);
                return;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            _sourcePath = path;
            var pages = new List<PdfPageRef>();
            for (int i = 0; i < _pageCount; i++)
                pages.Add(new PdfPageRef { SourcePath = path, PageIndex = i });
            _grid.SetPages(pages);
            SetStatus("Открыт «" + Path.GetFileName(path) + "»: страниц " + _pageCount + ".", Theme.TextMuted);
            UpdateControls();
        }

        // ---------- режимы ----------

        private void UpdateModeInputs()
        {
            int mode = _cmbMode.SelectedIndex;
            _lblRanges.Visible = _txtRanges.Visible = mode == ModeRanges;
            _lblN.Visible = _numN.Visible = mode == ModeEveryN;
            _chkCombine.Visible = mode == ModeRanges;
            _lblHint.Visible = mode == ModeExtract || mode == ModeBookmarks;
            if (mode == ModeExtract)
                _lblHint.Text = "Выделите нужные страницы в сетке (Ctrl+A — все).";
            else if (mode == ModeBookmarks)
                _lblHint.Text = "По одному файлу на закладку верхнего уровня.";
            // Извлечение (в т.ч. диапазоны+объединить) даёт один файл; иначе — несколько.
            bool oneFile = mode == ModeExtract || (mode == ModeRanges && _chkCombine.Checked);
            _btnDo.Text = oneFile ? "Извлечь…" : "Разделить…";
        }

        private void UpdateControls()
        {
            bool loaded = _sourcePath != null;
            _compress.Enabled = !_busy;
            _btnOpen.Enabled = !_busy;
            _cmbMode.Enabled = !_busy && loaded;
            _txtRanges.Enabled = !_busy && loaded;
            _numN.Enabled = !_busy && loaded;
            bool canDo = !_busy && loaded &&
                (_cmbMode.SelectedIndex != ModeExtract || _grid.SelectedCount > 0);
            _btnDo.Enabled = canDo;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (_grid != null && _grid.ListFocused && keyData == (Keys.Control | Keys.A))
            {
                _grid.SelectAll();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // ---------- выполнение ----------

        private void OnDoClick(object sender, EventArgs e)
        {
            if (_busy || _sourcePath == null)
                return;
            int mode = _cmbMode.SelectedIndex;
            bool combine = mode == ModeRanges && _chkCombine.Checked;
            string src = _sourcePath;
            CompressionLevel level = _compress.Level; // с UI-потока до старта воркера

            // Один файл: «Извлечь выбранные» ИЛИ «По диапазонам» + объединить.
            if (mode == ModeExtract || combine)
            {
                List<int> indices;
                if (mode == ModeExtract)
                {
                    int[] sel = _grid.GetSelectedIndices();
                    if (sel.Length == 0)
                    {
                        Dialogs.Error(this, Title, "Не выбраны страницы", "Выделите страницы в сетке (Ctrl+A — все).");
                        return;
                    }
                    indices = new List<int>(sel);
                }
                else
                {
                    try { indices = PageRanges.ToIndices(PageRanges.Parse(_txtRanges.Text, _pageCount)); }
                    catch (MergeException ex) { Dialogs.Error(this, Title, "Диапазоны заданы неверно", ex.Message); return; }
                }
                string outPath;
                using (var dialog = new SaveFileDialog())
                {
                    dialog.Filter = "Документ PDF (*.pdf)|*.pdf";
                    dialog.FileName = Path.GetFileNameWithoutExtension(src) +
                        (mode == ModeExtract ? "_выбранные.pdf" : "_объединённые.pdf");
                    dialog.InitialDirectory = Path.GetDirectoryName(src);
                    if (dialog.ShowDialog(this) != DialogResult.OK)
                        return;
                    outPath = dialog.FileName;
                }
                RunSplit(delegate { PdfSplitService.Extract(src, indices, outPath); return new List<string> { outPath }; },
                    level, outPath, false, UsageStats.RecordPdfExtract);
                return;
            }

            // Несколько файлов: диапазоны (без объединения) / каждые N / закладки.
            IList<PageRange> ranges = null;
            int everyN = 0;
            if (mode == ModeRanges)
            {
                try { ranges = PageRanges.Parse(_txtRanges.Text, _pageCount); }
                catch (MergeException ex) { Dialogs.Error(this, Title, "Диапазоны заданы неверно", ex.Message); return; }
            }
            else if (mode == ModeEveryN)
            {
                everyN = (int)_numN.Value;
            }

            // Даём выбрать и папку, и базовое имя: к нему добавятся номера/метки.
            string dir, baseName;
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Документ PDF (*.pdf)|*.pdf";
                dialog.Title = "Базовое имя и папка для частей (к имени добавятся номера)";
                dialog.FileName = Path.GetFileNameWithoutExtension(src) + ".pdf";
                dialog.InitialDirectory = Path.GetDirectoryName(src);
                dialog.OverwritePrompt = false; // создаются base_1.pdf и т.п., а не сам base.pdf
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                dir = Path.GetDirectoryName(dialog.FileName);
                baseName = Path.GetFileNameWithoutExtension(dialog.FileName);
                if (string.IsNullOrWhiteSpace(baseName))
                    baseName = Path.GetFileNameWithoutExtension(src);
            }

            Func<List<string>> work;
            Action record;
            switch (mode)
            {
                case ModeRanges:
                    work = delegate { return PdfSplitService.SplitByRanges(src, ranges, dir, baseName); };
                    record = UsageStats.RecordPdfSplitRanges;
                    break;
                case ModeEveryN:
                    work = delegate { return PdfSplitService.SplitEveryN(src, everyN, dir, baseName); };
                    record = UsageStats.RecordPdfSplitEveryN;
                    break;
                case ModeBookmarks:
                    work = delegate { return PdfSplitService.SplitByBookmarks(src, dir, baseName); };
                    record = UsageStats.RecordPdfSplitBookmarks;
                    break;
                default:
                    return;
            }
            RunSplit(work, level, dir, true, record);
        }

        /// <summary>
        /// Выполнить работу в фоне; сжать полученные файлы (на этом же воркере, до
        /// открытия результата); по завершении — статус, счётчик, открытие результата.
        /// </summary>
        private void RunSplit(Func<List<string>> work, CompressionLevel level, string openTarget, bool openAsFolder, Action record)
        {
            _busy = true;
            UpdateControls();
            SetStatus(openAsFolder ? "Разделение…" : "Извлечение…", Theme.TextMuted);
            var thread = new Thread(delegate()
            {
                Exception error = null;
                int count = 0, compressed = 0;
                try
                {
                    List<string> files = work();
                    count = files.Count;
                    foreach (string f in files)
                        if (PdfCompression.Compress(f, level))
                            compressed++;
                }
                catch (Exception ex) { error = ex; }
                int resultCount = count, resultCompressed = compressed;
                try
                {
                    if (IsHandleCreated && !IsDisposed)
                        BeginInvoke((MethodInvoker)delegate { OnSplitFinished(error, resultCount, resultCompressed, openTarget, openAsFolder, record); });
                }
                catch (InvalidOperationException) { }
            });
            thread.IsBackground = true;
            thread.Start();
        }

        private void OnSplitFinished(Exception error, int count, int compressed, string openTarget, bool openAsFolder, Action record)
        {
            _busy = false;
            UpdateControls();
            if (error != null)
            {
                SetStatus("Не выполнено.", Theme.ErrRed);
                Dialogs.Error(this, Title, openAsFolder ? "Разделение не выполнено" : "Извлечение не выполнено",
                    error.Message);
                return;
            }
            if (record != null)
                record(); // успех — учитываем в статистике
            if (compressed > 0)
                UsageStats.RecordPdfCompress(compressed);
            string suffix = compressed > 0 ? " · сжато: " + compressed : "";
            SetStatus(openAsFolder ? ("✓ Создано файлов: " + count + "." + suffix)
                : ("✓ Готово." + suffix), Theme.OkGreen);
            try
            {
                if (openAsFolder)
                    Process.Start("explorer.exe", "\"" + openTarget + "\"");
                else
                    Process.Start(openTarget);
            }
            catch { } // нет ассоциации/проводника — файлы всё равно созданы
        }

    }
}
