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
            InitShell(Title, new Size(800, 660), new Size(700, 560), Theme.PdfRed);
            DragEnter += OnFileDragEnter;
            DragDrop += OnFileDragDrop;
            BuildHeaderWithHome(Title,
                "Извлечение страниц из документа формата *.pdf со сжатием.",
                Theme.PdfRed, Theme.PdfRedDark, ShowHelp);

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;
            int panelW = 210;
            int gridBottom = ClientSize.Height - 152;

            _grid = new PdfPageGrid();
            _grid.AllowReorder = false; // разделение не меняет порядок исходника
            _grid.SetBounds(20, m + 84, right - 20 - panelW, gridBottom - (m + 84));
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _grid.SelectionChanged += delegate { UpdateControls(); };
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

            // Масштаб, сжатие и статус — общий нижний строй (как в «Объединении»).
            BuildBottomStrip(right, "Откройте PDF — кнопкой «Открыть PDF…» или перетащив его в окно.", 190);

            // Действие — в правом нижнем углу (как «Сохранить PDF» в «Объединении»).
            _btnDo = new RoundedButton(true);
            _btnDo.SetBounds(right - 190, ClientSize.Height - 58, 190, 38);
            _btnDo.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _btnDo.Click += OnDoClick;
            Controls.Add(_btnDo);
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
                RunSplit(delegate(Action<int, int> pr) { PdfSplitService.Extract(src, indices, outPath, pr); return new List<string> { outPath }; },
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

            Func<Action<int, int>, List<string>> work;
            Action record;
            switch (mode)
            {
                case ModeRanges:
                    work = delegate(Action<int, int> pr) { return PdfSplitService.SplitByRanges(src, ranges, dir, baseName, pr); };
                    record = UsageStats.RecordPdfSplitRanges;
                    break;
                case ModeEveryN:
                    work = delegate(Action<int, int> pr) { return PdfSplitService.SplitEveryN(src, everyN, dir, baseName, pr); };
                    record = UsageStats.RecordPdfSplitEveryN;
                    break;
                case ModeBookmarks:
                    work = delegate(Action<int, int> pr) { return PdfSplitService.SplitByBookmarks(src, dir, baseName, pr); };
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
        private void RunSplit(Func<Action<int, int>, List<string>> work, CompressionLevel level, string openTarget, bool openAsFolder, Action record)
        {
            _busy = true;
            UpdateControls();
            SetStatus(openAsFolder ? "Разделение…" : "Извлечение…", Theme.TextMuted);
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
            string status = openAsFolder ? ("✓ Создано файлов: " + count + "." + suffix)
                                          : ("✓ Готово." + suffix);
            // Если без сжатия результат вышел почти как исходник (общие ресурсы страниц
            // едут вместе с ними) — ненавязчиво подсказать про «Сжатие».
            if (ShouldSuggestCompression(level, sourceSize, largestOutput))
                status += " · файл крупный — включите «Сжатие», чтобы уменьшить размер.";
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
