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
    public class PdfSplitForm : Form
    {
        private const string Title = "Разделение PDF";
        private const int ModeExtract = 0, ModeRanges = 1, ModeEveryN = 2, ModeBookmarks = 3;

        private readonly Action _showHub;
        private string _sourcePath;
        private int _pageCount;
        private bool _busy;

        private PdfPageGrid _grid;
        private TrackBar _zoom;
        private System.Windows.Forms.Timer _zoomTimer;
        private ToolTip _tips;
        private Button _btnOpen;
        private ComboBox _cmbMode;
        private Label _lblRanges;
        private TextBox _txtRanges;
        private Label _lblN;
        private NumericUpDown _numN;
        private Label _lblHint;
        private Button _btnDo;
        private Label _lblStatus;

        public PdfSplitForm() : this(null) { }

        public PdfSplitForm(Action showHub)
        {
            _showHub = showHub;
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

            _btnDo = new RoundedButton(true);
            _btnDo.SetBounds(px, m + 250, pw, 38);
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
                "3. Нажмите «Извлечь…»/«Разделить…» и укажите имя и папку для результата " +
                "(при разбиении к имени добавятся номера или метки).\n\n" +
                "Страницы копируются как есть, без переконвертации. Исходный файл не изменяется; " +
                "имена не перезаписываются (при совпадении добавляется номер).");
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
            _lblHint.Visible = mode == ModeExtract || mode == ModeBookmarks;
            if (mode == ModeExtract)
                _lblHint.Text = "Выделите нужные страницы в сетке (Ctrl+A — все).";
            else if (mode == ModeBookmarks)
                _lblHint.Text = "По одному файлу на закладку верхнего уровня.";
            _btnDo.Text = mode == ModeExtract ? "Извлечь…" : "Разделить…";
        }

        private void UpdateControls()
        {
            bool loaded = _sourcePath != null;
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

        private void ScheduleZoom()
        {
            _zoomTimer.Stop();
            _zoomTimer.Start();
        }

        // ---------- выполнение ----------

        private void OnDoClick(object sender, EventArgs e)
        {
            if (_busy || _sourcePath == null)
                return;
            int mode = _cmbMode.SelectedIndex;

            if (mode == ModeExtract)
            {
                int[] indices = _grid.GetSelectedIndices();
                if (indices.Length == 0)
                {
                    Dialogs.Error(this, Title, "Не выбраны страницы", "Выделите страницы в сетке (Ctrl+A — все).");
                    return;
                }
                string outPath;
                using (var dialog = new SaveFileDialog())
                {
                    dialog.Filter = "Документ PDF (*.pdf)|*.pdf";
                    dialog.FileName = Path.GetFileNameWithoutExtension(_sourcePath) + "_выбранные.pdf";
                    dialog.InitialDirectory = Path.GetDirectoryName(_sourcePath);
                    if (dialog.ShowDialog(this) != DialogResult.OK)
                        return;
                    outPath = dialog.FileName;
                }
                var idxCopy = new List<int>(indices);
                RunSplit(delegate { PdfSplitService.Extract(_sourcePath, idxCopy, outPath); return 1; }, outPath, false);
                return;
            }

            // Разбиение на несколько файлов — заранее проверяем ввод, затем папка.
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

            // Даём выбрать и папку, и базовое имя: к нему добавятся номера/метки
            // (base_1-3.pdf, base_часть_1.pdf, base_Глава.pdf).
            string dir, baseName;
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Документ PDF (*.pdf)|*.pdf";
                dialog.Title = "Базовое имя и папка для частей (к имени добавятся номера)";
                dialog.FileName = Path.GetFileNameWithoutExtension(_sourcePath) + ".pdf";
                dialog.InitialDirectory = Path.GetDirectoryName(_sourcePath);
                dialog.OverwritePrompt = false; // создаются base_1.pdf и т.п., а не сам base.pdf
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                dir = Path.GetDirectoryName(dialog.FileName);
                baseName = Path.GetFileNameWithoutExtension(dialog.FileName);
                if (string.IsNullOrWhiteSpace(baseName))
                    baseName = Path.GetFileNameWithoutExtension(_sourcePath);
            }
            string src = _sourcePath;
            RunSplit(delegate
            {
                switch (mode)
                {
                    case ModeRanges: return PdfSplitService.SplitByRanges(src, ranges, dir, baseName).Count;
                    case ModeEveryN: return PdfSplitService.SplitEveryN(src, everyN, dir, baseName).Count;
                    case ModeBookmarks: return PdfSplitService.SplitByBookmarks(src, dir, baseName).Count;
                    default: return 0;
                }
            }, dir, true);
        }

        /// <summary>Выполнить работу разбиения в фоне; по завершении — статус и открытие результата.</summary>
        private void RunSplit(Func<int> work, string openTarget, bool openAsFolder)
        {
            _busy = true;
            UpdateControls();
            SetStatus(openAsFolder ? "Разделение…" : "Извлечение…", Theme.TextMuted);
            var thread = new Thread(delegate()
            {
                Exception error = null;
                int count = 0;
                try { count = work(); }
                catch (Exception ex) { error = ex; }
                int result = count;
                try
                {
                    if (IsHandleCreated && !IsDisposed)
                        BeginInvoke((MethodInvoker)delegate { OnSplitFinished(error, result, openTarget, openAsFolder); });
                }
                catch (InvalidOperationException) { }
            });
            thread.IsBackground = true;
            thread.Start();
        }

        private void OnSplitFinished(Exception error, int count, string openTarget, bool openAsFolder)
        {
            _busy = false;
            UpdateControls();
            if (error != null)
            {
                SetStatus("Не выполнено.", Theme.ErrRed);
                Dialogs.Error(this, Title, openAsFolder ? "Разделение не выполнено" : "Извлечение не выполнено",
                    (error is MergeException) ? error.Message : error.Message);
                return;
            }
            SetStatus(openAsFolder ? ("✓ Создано файлов: " + count + ".") : "✓ Готово.", Theme.OkGreen);
            try
            {
                if (openAsFolder)
                    Process.Start("explorer.exe", "\"" + openTarget + "\"");
                else
                    Process.Start(openTarget);
            }
            catch { } // нет ассоциации/проводника — файлы всё равно созданы
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
                SetStatus("Дождитесь завершения…", Theme.WarnOrange);
                e.Cancel = true;
                return;
            }
            _grid.StopRendering();
            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _zoomTimer != null)
                _zoomTimer.Dispose();
            base.Dispose(disposing);
        }
    }
}
