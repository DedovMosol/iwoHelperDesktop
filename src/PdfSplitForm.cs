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
        private static string Title { get { return Loc.T("hub.split.name"); } }
        private const int ModeExtract = 0, ModeRanges = 1, ModeEveryN = 2, ModeBookmarks = 3;

        // Сетка, зум, сжатие, статус, подсказки и флаг _busy — в базе PdfToolFormBase.
        private string _sourcePath;
        private int _pageCount;
        // Страницы исходника, показанные сеткой: сетка мутирует их Rotation при повороте,
        // отсюда форма собирает карту поворотов для записи (все режимы разделения).
        private List<PdfPageRef> _pages = new List<PdfPageRef>();

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
            InitShell(Title, new Size(800, 660), new Size(700, 560), Theme.PdfRed);
            DragEnter += OnFileDragEnter;
            DragDrop += OnFileDragDrop;
            BuildHeaderWithHome(Title,
                Loc.T("split.header.subtitle"),
                Theme.PdfRed, Theme.PdfRedDark, ShowHelp);

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;
            int panelW = 210;
            int gridBottom = ClientSize.Height - 152;

            _grid = new PdfPageGrid();
            _grid.AllowReorder = false; // разделение не меняет порядок исходника
            _grid.AllowRotate = true;   // но поворот страниц в ИТОГОВЫХ файлах — можно
            // ShowPositionNumbers = false: под плиткой — номер исходной страницы.
            _grid.SetBounds(20, m + 84, right - 20 - panelW, gridBottom - (m + 84));
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _grid.SelectionChanged += delegate { UpdateControls(); };
            _grid.FilesDropped += delegate(string[] paths, int insertAt)
            {
                if (!_busy && paths.Length > 0)
                    LoadSource(paths[0]); // разделение работает с одним документом
            };
            WireGridMenu();
            Controls.Add(_grid);

            int px = right - panelW + 10; // левый край панели режима
            int pw = panelW - 10;
            _btnOpen = new RoundedButton(false);
            _btnOpen.Text = Loc.T("split.btn.open");
            _btnOpen.SetBounds(px, m + 84, pw, 32);
            _btnOpen.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnOpen.Click += OnOpenClick;
            _tips.SetToolTip(_btnOpen, Loc.T("split.tip.open"));
            Controls.Add(_btnOpen);

            Label lblMode = Ui.Label(this, Loc.T("split.lbl.mode"), px, m + 128, Font, Theme.TextPrimary);
            lblMode.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _cmbMode = new ComboBox();
            _cmbMode.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbMode.Items.AddRange(new object[] { Loc.T("split.mode.extract"), Loc.T("split.mode.ranges"), Loc.T("split.mode.everyN"), Loc.T("split.mode.bookmarks") });
            _cmbMode.SelectedIndex = ModeExtract;
            _cmbMode.SetBounds(px, m + 150, pw, 27);
            _cmbMode.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _cmbMode.SelectedIndexChanged += delegate { UpdateModeInputs(); UpdateControls(); };
            // Выбор режима из раскрытого списка переводит фокус в его поле ввода — можно
            // сразу печатать. Стрелки по закрытому списку фокус не трогают, иначе они
            // перестали бы листать режимы.
            int modeBeforeDrop = -1;
            _cmbMode.DropDown += delegate { modeBeforeDrop = _cmbMode.SelectedIndex; };
            _cmbMode.DropDownClosed += delegate
            {
                if (_cmbMode.SelectedIndex != modeBeforeDrop)
                    FocusModeInput();
            };
            Controls.Add(_cmbMode);

            // Поля ввода режимов (в одном месте, показываются по режиму).
            _lblRanges = Ui.Label(this, Loc.T("split.lbl.ranges"), px, m + 188, Font, Theme.TextMuted);
            _lblRanges.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _txtRanges = new TextBox();
            _txtRanges.SetBounds(px, m + 210, pw, 27);
            _txtRanges.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(_txtRanges);

            _lblN = Ui.Label(this, Loc.T("split.lbl.n"), px, m + 188, Font, Theme.TextMuted);
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
            _chkCombine.Text = Loc.T("split.chk.combine");
            _chkCombine.SetBounds(px, m + 244, pw, 22);
            _chkCombine.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _chkCombine.ForeColor = Theme.TextPrimary;
            _tips.SetToolTip(_chkCombine, Loc.T("split.tip.combine"));
            _chkCombine.CheckedChanged += delegate { UpdateModeInputs(); };
            Controls.Add(_chkCombine);

            // Масштаб, сжатие и статус — общий нижний строй (как в «Объединении»).
            BuildBottomStrip(right, Loc.T("split.status.openPdf"), 190);

            // Действие — в правом нижнем углу (как «Сохранить PDF» в «Объединении»).
            _btnDo = new RoundedButton(true);
            _btnDo.SetBounds(right - 190, ClientSize.Height - 58, 190, 38);
            _btnDo.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _btnDo.Click += OnDoClick;
            Controls.Add(_btnDo);
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, Title, Loc.T("menu.howTo"), Loc.T("split.help.body"));
        }

        // ---------- открытие исходника ----------

        private void OnOpenClick(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = Loc.T("common.pdfFilter");
                dialog.Title = Loc.T("split.pickPdf");
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
                Dialogs.Error(this, Title, Loc.T("split.err.fileNotOpened"), ex.Message);
                return;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            _sourcePath = path;
            _pages = new List<PdfPageRef>();
            for (int i = 0; i < _pageCount; i++)
                _pages.Add(new PdfPageRef { SourcePath = path, PageIndex = i });
            _grid.SetPages(_pages);
            SetStatus(string.Format(Loc.T("split.status.opened"), Path.GetFileName(path), _pageCount), Theme.TextMuted);
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
                _lblHint.Text = Loc.T("split.hint.extract");
            else if (mode == ModeBookmarks)
                _lblHint.Text = Loc.T("split.hint.bookmarks");
            // Извлечение (в т.ч. диапазоны+объединить) даёт один файл; иначе — несколько.
            bool oneFile = mode == ModeExtract || (mode == ModeRanges && _chkCombine.Checked);
            _btnDo.Text = oneFile ? Loc.T("split.btn.extract") : Loc.T("split.btn.split");
        }

        /// <summary>Фокус в поле ввода текущего режима (Focus() у недоступного поля — no-op).</summary>
        private void FocusModeInput()
        {
            if (_cmbMode.SelectedIndex == ModeRanges)
                _txtRanges.Focus();
            else if (_cmbMode.SelectedIndex == ModeEveryN)
                _numN.Focus();
        }

        private void UpdateControls()
        {
            bool loaded = _sourcePath != null;
            _grid.Locked = _busy; // поворот и дроп — только вне операции
            _compress.Enabled = !_busy;
            _btnOpen.Enabled = !_busy;
            _cmbMode.Enabled = !_busy && loaded;
            _txtRanges.Enabled = !_busy && loaded;
            _numN.Enabled = !_busy && loaded;
            bool canDo = !_busy && loaded &&
                (_cmbMode.SelectedIndex != ModeExtract || _grid.SelectedCount > 0);
            _btnDo.Enabled = canDo;
        }

        // Ctrl+A в сетке — в базе PdfToolFormBase (сетка без AllowReorder: только выделение).

        // ---------- выполнение ----------

        private void OnDoClick(object sender, EventArgs e)
        {
            if (_busy || _sourcePath == null)
                return;
            int mode = _cmbMode.SelectedIndex;
            bool combine = mode == ModeRanges && _chkCombine.Checked;
            string src = _sourcePath;
            CompressionLevel level = _compress.Level; // с UI-потока до старта воркера
            int[] rotations = CurrentRotations();     // снимок поворотов — тоже до старта воркера

            // Один файл: «Извлечь выбранные» ИЛИ «По диапазонам» + объединить.
            if (mode == ModeExtract || combine)
            {
                List<int> indices;
                if (mode == ModeExtract)
                {
                    int[] sel = _grid.GetSelectedIndices();
                    if (sel.Length == 0)
                    {
                        Dialogs.Error(this, Title, Loc.T("split.err.noPages.title"), Loc.T("split.err.noPages.body"));
                        return;
                    }
                    indices = new List<int>(sel);
                }
                else
                {
                    try { indices = PageRanges.ToIndices(PageRanges.Parse(_txtRanges.Text, _pageCount)); }
                    catch (MergeException ex) { Dialogs.Error(this, Title, Loc.T("split.err.badRanges"), ex.Message); return; }
                }
                string outPath;
                using (var dialog = new SaveFileDialog())
                {
                    dialog.Filter = Loc.T("common.pdfSaveFilter");
                    dialog.FileName = Path.GetFileNameWithoutExtension(src) +
                        (mode == ModeExtract ? Loc.T("split.suffix.selected") : Loc.T("split.suffix.combined"));
                    dialog.InitialDirectory = Path.GetDirectoryName(src);
                    if (dialog.ShowDialog(this) != DialogResult.OK)
                        return;
                    outPath = dialog.FileName;
                }
                RunSplit(delegate(Action<int, int> pr) { PdfSplitService.Extract(src, indices, outPath, pr, rotations); return new List<string> { outPath }; },
                    level, outPath, false, UsageStats.RecordPdfExtract);
                return;
            }

            // Несколько файлов: диапазоны (без объединения) / каждые N / закладки.
            IList<PageRange> ranges = null;
            int everyN = 0;
            if (mode == ModeRanges)
            {
                try { ranges = PageRanges.Parse(_txtRanges.Text, _pageCount); }
                catch (MergeException ex) { Dialogs.Error(this, Title, Loc.T("split.err.badRanges"), ex.Message); return; }
            }
            else if (mode == ModeEveryN)
            {
                everyN = (int)_numN.Value;
            }

            // Даём выбрать и папку, и базовое имя: к нему добавятся номера/метки.
            string dir, baseName;
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = Loc.T("common.pdfSaveFilter");
                dialog.Title = Loc.T("split.pickBase");
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

            Func<Action<int, int>, List<string>> work;
            Action record;
            switch (mode)
            {
                case ModeRanges:
                    work = delegate(Action<int, int> pr) { return PdfSplitService.SplitByRanges(src, ranges, dir, baseName, pr, rotations); };
                    record = UsageStats.RecordPdfSplitRanges;
                    break;
                case ModeEveryN:
                    work = delegate(Action<int, int> pr) { return PdfSplitService.SplitEveryN(src, everyN, dir, baseName, pr, rotations); };
                    record = UsageStats.RecordPdfSplitEveryN;
                    break;
                case ModeBookmarks:
                    work = delegate(Action<int, int> pr) { return PdfSplitService.SplitByBookmarks(src, dir, baseName, pr, rotations); };
                    record = UsageStats.RecordPdfSplitBookmarks;
                    break;
                default:
                    return;
            }
            RunSplit(work, level, dir, true, record);
        }

        /// <summary>
        /// Карта поворотов по индексу исходной страницы (для всех режимов записи);
        /// null — поворотов нет. Снимается на UI-потоке ДО старта воркера, а во время
        /// операции сетка заблокирована (Locked) — карта не меняется под ногами.
        /// </summary>
        private int[] CurrentRotations()
        {
            if (_pages == null || _pages.Count == 0)
                return null;
            bool any = false;
            var rotations = new int[_pageCount];
            foreach (PdfPageRef page in _pages)
                if (page.PageIndex >= 0 && page.PageIndex < rotations.Length && page.Rotation != 0)
                {
                    rotations[page.PageIndex] = page.Rotation;
                    any = true;
                }
            return any ? rotations : null;
        }

        /// <summary>
        /// Выполнить работу в фоне; сжать полученные файлы (на этом же воркере, до
        /// открытия результата); по завершении — статус, счётчик, открытие результата.
        /// </summary>
        private void RunSplit(Func<Action<int, int>, List<string>> work, CompressionLevel level, string openTarget, bool openAsFolder, Action record)
        {
            _busy = true;
            UpdateControls();
            SetStatus(openAsFolder ? Loc.T("split.status.splitting") : Loc.T("split.status.extracting"), Theme.TextMuted);
            BeginProgress();
            Action<int, int> onProgress = UiProgress();
            long sourceSize = SafeLength(_sourcePath); // для подсказки о сжатии (UI-поток, до старта воркера)
            var thread = new Thread(delegate()
            {
                Exception error = null;
                int count = 0, compressed = 0;
                long largest = 0;
                try
                {
                    // Две фазы в одну шкалу 2×частей: разбиение (0..P) и сжатие (P..2P).
                    List<string> files = work(delegate(int done, int total) { onProgress(done, 2 * total); });
                    count = files.Count;
                    for (int i = 0; i < files.Count; i++)
                    {
                        if (PdfCompression.Compress(files[i], level))
                            compressed++;
                        onProgress(files.Count + i + 1, 2 * files.Count);
                    }
                    foreach (string f in files) // итоговые размеры (уже после сжатия)
                    {
                        long len = SafeLength(f);
                        if (len > largest) largest = len;
                    }
                }
                catch (Exception ex) { error = ex; }
                int resultCount = count, resultCompressed = compressed;
                long resultLargest = largest;
                try
                {
                    if (IsHandleCreated && !IsDisposed)
                        BeginInvoke((MethodInvoker)delegate { OnSplitFinished(error, resultCount, resultCompressed, openTarget, openAsFolder, record, level, sourceSize, resultLargest); });
                }
                catch (InvalidOperationException) { }
            });
            thread.IsBackground = true;
            thread.Start();
        }

        private void OnSplitFinished(Exception error, int count, int compressed, string openTarget, bool openAsFolder, Action record,
            CompressionLevel level, long sourceSize, long largestOutput)
        {
            _busy = false;
            EndProgress();
            UpdateControls();
            if (error != null)
            {
                SetStatus(Loc.T("common.status.notDone"), Theme.ErrRed);
                Dialogs.Error(this, Title, Loc.T(openAsFolder ? "split.err.splitFailed" : "split.err.extractFailed"),
                    error.Message);
                return;
            }
            if (record != null)
                record(); // успех — учитываем в статистике
            if (compressed > 0)
                UsageStats.RecordPdfCompress(compressed);
            string suffix = compressed > 0 ? string.Format(Loc.T("split.suffix.compressed"), compressed) : "";
            string status = openAsFolder ? (string.Format(Loc.T("split.status.created"), count) + suffix)
                                          : (Loc.T("split.status.done") + suffix);
            // Если без сжатия результат вышел почти как исходник (общие ресурсы страниц
            // едут вместе с ними) — ненавязчиво подсказать про «Сжатие».
            if (ShouldSuggestCompression(level, sourceSize, largestOutput))
                status += Loc.T("split.status.largeHint");
            SetStatus(status, Theme.OkGreen);
            try
            {
                if (openAsFolder)
                    Process.Start("explorer.exe", "\"" + openTarget + "\"");
                else
                    Process.Start(openTarget);
            }
            catch { } // нет ассоциации/проводника — файлы всё равно созданы
        }

        /// <summary>Длина файла в байтах (0, если недоступен). Без исключений.</summary>
        private static long SafeLength(string path)
        {
            try { return !string.IsNullOrEmpty(path) && File.Exists(path) ? new FileInfo(path).Length : 0L; }
            catch { return 0L; }
        }

        /// <summary>
        /// Стоит ли ненавязчиво подсказать «включите Сжатие»: сжатие не выбрано, а
        /// результат вышел почти как исходник (≥ 90% и не мелочь). Общие ресурсы
        /// страниц копируются вместе с ними, поэтому подмножество может весить столько же.
        /// Чистая — под тест.
        /// </summary>
        internal static bool ShouldSuggestCompression(CompressionLevel level, long sourceSize, long largestOutputSize)
        {
            return level == CompressionLevel.None
                && sourceSize > 0                             // размер исходника известен
                && largestOutputSize >= 1024L * 1024          // не шумим на мелких файлах (< 1 МБ)
                && largestOutputSize * 10L >= sourceSize * 9L; // вышло ≥ 90% исходника
        }

    }
}
