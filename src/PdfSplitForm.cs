using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
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
    public class PdfSplitForm : PdfSingleDocFormBase
    {
        private static string Title { get { return Loc.T("hub.split.name"); } }
        private const int ModeExtract = 0, ModeRanges = 1, ModeEveryN = 2, ModeBookmarks = 3;

        // Сетка, зум, сжатие, статус, подсказки и флаг _busy — в базе PdfToolFormBase.
        // Открытый документ (_sourcePath, _pageCount, _pages), его загрузка и выбор страниц —
        // в базе PdfSingleDocFormBase (общее с «Прочими операциями»).
        private Button _btnOpen;
        private ComboBox _cmbMode;
        private Label _lblRanges;
        private TextBox _txtRanges;
        private Label _lblN;
        private NumericUpDown _numN;
        private CheckBox _chkCombine;
        private TextBox _txtNameTemplate; // шаблон имени частей; пусто — прежние имена
        private Button _btnOps;
        private Button _btnPrint;
        private Label _lblHint;
        private Button _btnDo;

        public PdfSplitForm() : this(null) { }

        public PdfSplitForm(Action showHub) : this(showHub, null) { }

        public PdfSplitForm(Action showHub, Action<string> openOps) : base(showHub)
        {
            OpsBridge = openOps; // мост в «Прочие операции» с открытым документом (см. базу)
            BuildUi();
            UpdateModeInputs();
            SyncControls();
        }

        protected override string ToolTitle { get { return Title; } }

        /// <summary>Заголовок окна выбора файла — своя формулировка у каждого инструмента.</summary>
        protected override string PickFileTitle { get { return Loc.T("split.pickPdf"); } }

        private void BuildUi()
        {
            InitShell(Title, new Size(800, 660), new Size(700, 600), Theme.PdfRed);
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
            WireSingleDocGrid(); // подсказки, выделение, дроп на сетку и на окно (общая обвязка базы)
            WireGridMenu();
            Controls.Add(_grid);

            int px = right - panelW + 10; // левый край панели режима
            int pw = panelW - 10;
            _btnOpen = new RoundedButton(false);
            _btnOpen.Text = Loc.T("common.btn.openPdf");
            _btnOpen.SetBounds(px, m + 84, pw, 32);
            _btnOpen.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnOpen.Click += delegate { PickAndOpenFile(); };
            _tips.SetToolTip(_btnOpen, Loc.T("common.tip.openPdf"));
            Controls.Add(_btnOpen);

            Label lblMode = Ui.Label(this, Loc.T("split.lbl.mode"), px, m + 128, Font, Theme.TextPrimary);
            lblMode.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _cmbMode = new ComboBox();
            _cmbMode.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbMode.Items.AddRange(new object[] { Loc.T("split.mode.extract"), Loc.T("split.mode.ranges"), Loc.T("split.mode.everyN"), Loc.T("split.mode.bookmarks") });
            _cmbMode.SelectedIndex = ModeExtract;
            _cmbMode.SetBounds(px, m + 150, pw, 27);
            _cmbMode.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _cmbMode.SelectedIndexChanged += delegate { UpdateModeInputs(); SyncControls(); };
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
            _tips.SetToolTip(_txtRanges, Loc.T("split.tip.ranges"));
            Controls.Add(_txtRanges);

            _lblN = Ui.Label(this, Loc.T("split.lbl.n"), px, m + 188, Font, Theme.TextMuted);
            _lblN.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _numN = new NumericUpDown();
            _numN.Minimum = 1;
            _numN.Maximum = 10000;
            _numN.Value = 1;
            _numN.SetBounds(px, m + 210, 70, 27);
            _numN.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _tips.SetToolTip(_numN, Loc.T("split.tip.everyN"));
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

            Label lblTemplate = Ui.Label(this, Loc.T("split.lbl.template"), px, m + 276, Font, Theme.TextMuted);
            lblTemplate.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _txtNameTemplate = new TextBox();
            _txtNameTemplate.SetBounds(px, m + 298, pw, 27);
            _txtNameTemplate.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _tips.SetToolTip(_txtNameTemplate, Loc.T("split.tip.template"));
            // Подстановка по правому клику: набирать «[FILENUMBER###]» по памяти никто не станет,
            // а список доступных обозначений иначе негде увидеть.
            _txtNameTemplate.ContextMenuStrip = BuildTemplateMenu();
            Controls.Add(_txtNameTemplate);

            // Кнопка «Ещё с документом» рядом с остальным вводом: те же пункты, что в «☰».
            // Без неё картинки, текст, серое, восстановление и свойства оставались невидимыми —
            // искать их в меню окна никто не догадывался.
            _btnPrint = new RoundedButton(false);
            _btnPrint.Text = Loc.T("split.btn.print");
            _btnPrint.SetBounds(px, m + 336, 74, 32);
            _btnPrint.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _tips.SetToolTip(_btnPrint, Loc.T("split.tip.print"));
            _btnPrint.Click += delegate { PrintSelectedPages(); };
            Controls.Add(_btnPrint);

            // Шесть операций над одним документом переехали в своё окно (их не находили в меню).
            // Здесь остался переход туда с УЖЕ ОТКРЫТЫМ файлом — открывать его заново не нужно.
            _btnOps = new RoundedButton(false);
            _btnOps.Text = Loc.T("split.btn.ops");
            _btnOps.SetBounds(px + 80, m + 336, pw - 80, 32);
            _btnOps.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _tips.SetToolTip(_btnOps, Loc.T("split.tip.ops"));
            // Открытый документ уезжает туда сам, а если ничего не открыто — окно просто
            // откроется: нажатие обязано что-то делать, иначе кнопку считают сломанной.
            _btnOps.Click += delegate
            {
                if (OpsBridge != null)
                    OpsBridge(_sourcePath);
            };
            Controls.Add(_btnOps);

            // Масштаб, сжатие и статус — общий нижний строй (как в «Объединении»).
            BuildBottomStrip(right, Loc.T("common.status.openPdf"), 190);

            // Действие — в правом нижнем углу (как «Сохранить PDF» в «Объединении»).
            _btnDo = new RoundedButton(true);
            _btnDo.SetBounds(right - 190, ClientSize.Height - 58, 190, 38);
            _btnDo.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _btnDo.Click += OnDoClick;
            Controls.Add(_btnDo);
            RegisterActionButton(_btnDo); // база подменит её кнопкой «Отмена» во время операции
            AcceptButton = _btnDo; // Enter запускает действие — как в «Объединении» и «PDF → Word»
        }

        /// <summary>
        /// Меню подстановки обозначений в шаблон имени: вставляет выбранное в место каретки.
        /// Список берётся из <see cref="NameTemplate.Tokens"/> — единственного места, где он
        /// объявлен, поэтому новое обозначение появляется в меню само.
        /// </summary>
        private ContextMenuStrip BuildTemplateMenu()
        {
            var menu = new ContextMenuStrip();
            foreach (string token in NameTemplate.Tokens)
            {
                string t = token; // копия для замыкания
                menu.Items.Add("[" + t + "]", null, delegate { InsertToken("[" + t + "]"); });
            }
            _templateMenu = menu; // не дочерний контрол — освобождаем сами
            return menu;
        }

        private ContextMenuStrip _templateMenu;

        private void InsertToken(string token)
        {
            int at = _txtNameTemplate.SelectionStart;
            _txtNameTemplate.Text = _txtNameTemplate.Text.Remove(at, _txtNameTemplate.SelectionLength).Insert(at, token);
            _txtNameTemplate.SelectionStart = at + token.Length;
            _txtNameTemplate.Focus();
        }


        /// <summary>Печать выбранных (или всех) страниц — через общий механизм базы.</summary>
        private void PrintSelectedPages()
        {
            if (_sourcePath == null)
                return;
            PrintPages(SelectedPageRefs());
        }

        protected override void Dispose(bool disposing)
        {
            // Меню подстановки назначено свойством, а не добавлено в Controls: само не освободится.
            if (disposing && _templateMenu != null)
            {
                if (_txtNameTemplate != null)
                    _txtNameTemplate.ContextMenuStrip = null;
                _templateMenu.Dispose();
                _templateMenu = null;
            }
            base.Dispose(disposing);
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, Title, Loc.T("menu.howTo"), Loc.T("split.help.body"));
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

        protected override void SyncControls()
        {
            bool loaded = _sourcePath != null;
            _grid.Locked = Working; // поворот и дроп — только вне работы (операция/загрузка)
            _compress.Enabled = !Working;
            _btnOpen.Enabled = !Working;
            _cmbMode.Enabled = !Working && loaded;
            _txtRanges.Enabled = !Working && loaded;
            _numN.Enabled = !Working && loaded;
            // Галку «Объединить в один файл» гасим наравне с остальным вводом: иначе во время
            // разделения она оставалась живой и её переключение перестраивало режим на ходу.
            _chkCombine.Enabled = !Working && loaded;
            _txtNameTemplate.Enabled = !Working && loaded;
            _btnOps.Enabled = !Working; // работает и без открытого документа — просто откроет окно
            _btnPrint.Enabled = !Working && loaded;
            bool canDo = !Working && loaded &&
                (_cmbMode.SelectedIndex != ModeExtract || _grid.SelectedCount > 0);
            _btnDo.Enabled = canDo;
        }

        // Ctrl+A в сетке — в базе PdfToolFormBase (сетка без AllowReorder: только выделение).

        // ---------- выполнение ----------

        private void OnDoClick(object sender, EventArgs e)
        {
            if (Working || _sourcePath == null)
                return;
            int mode = _cmbMode.SelectedIndex;
            bool combine = mode == ModeRanges && _chkCombine.Checked;
            string src = _sourcePath;
            CompressionLevel level = _compress.Level; // с UI-потока до старта воркера
            int[] rotations = CurrentRotations();     // снимок поворотов — тоже до старта воркера
            Func<bool> cancel = CancelToken();        // кооперативная отмена длинной операции

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
                RunSplit(delegate(Action<int, int> pr) { PdfSplitService.Extract(src, indices, outPath, pr, rotations, cancel); return new List<string> { outPath }; },
                    level, outPath, false, UsageStats.RecordPdfExtract, indices.Count);
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

            string template = _txtNameTemplate.Text.Trim(); // с UI-потока до старта воркера
            Func<Action<int, int>, List<string>> work;
            Action record;
            switch (mode)
            {
                case ModeRanges:
                    work = delegate(Action<int, int> pr) { return PdfSplitService.SplitByRanges(src, ranges, dir, baseName, pr, rotations, cancel, template); };
                    record = UsageStats.RecordPdfSplitRanges;
                    break;
                case ModeEveryN:
                    work = delegate(Action<int, int> pr) { return PdfSplitService.SplitEveryN(src, everyN, dir, baseName, pr, rotations, cancel, template); };
                    record = UsageStats.RecordPdfSplitEveryN;
                    break;
                case ModeBookmarks:
                    work = delegate(Action<int, int> pr) { return PdfSplitService.SplitByBookmarks(src, dir, baseName, pr, rotations, cancel, template); };
                    record = UsageStats.RecordPdfSplitBookmarks;
                    break;
                default:
                    return;
            }
            RunSplit(work, level, dir, true, record, _pageCount);
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
        /// workUnits — фактический объём работы в страницах (у «Извлечь выбранные» —
        /// число выбранных, не страницы источника): от него порог кнопки «Отмена», а у
        /// извлечения ещё и число страниц в итоговом файле (Extract пишет ровно по странице
        /// на индекс), которое уходит в статус.
        /// </summary>
        private void RunSplit(Func<Action<int, int>, List<string>> work, CompressionLevel level, string openTarget, bool openAsFolder, Action record, int workUnits)
        {
            // Счётчик «страница N из M» для разделения не показываем — единица работы здесь «часть».
            BeginOperation(openAsFolder ? Loc.T("split.status.splitting") : Loc.T("split.status.extracting"), workUnits);
            Action<int, int> onProgress = UiProgress();
            long sourceSize = SafeLength(_sourcePath); // для подсказки о сжатии (UI-поток, до старта воркера)
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                int count = 0, compressed = 0;
                long largest = 0;
                try
                {
                    // Две фазы в одну шкалу 2×частей: разбиение (0..P) и сжатие (P..2P).
                    List<string> files = work(delegate(int done, int total) { onProgress(done, 2 * total); });
                    count = files.Count;
                    // Файлы записаны — точка невозврата: сжатие (Ghostscript) не прерываем, поэтому
                    // снимаем предложение отмены, чтобы кнопка не «зависала» на «Отмена…».
                    // Статус называет фазу: полоса на второй половине шкалы двигается по файлам,
                    // но по одной части она стоит, и без подписи это выглядело бы зависанием.
                    bool willCompress = level != CompressionLevel.None;
                    OnUi(delegate
                    {
                        StopOfferingCancel();
                        if (willCompress)
                            SetStatus(Loc.T("common.status.compressing"), Theme.TextMuted);
                    });
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
                OnUi(delegate { OnSplitFinished(error, resultCount, resultCompressed, openTarget, openAsFolder, record, level, sourceSize, resultLargest, workUnits); });
            });
        }

        private void OnSplitFinished(Exception error, int count, int compressed, string openTarget, bool openAsFolder, Action record,
            CompressionLevel level, long sourceSize, long largestOutput, int pageCount)
        {
            if (!FinishOperation(error, Loc.T("common.status.notDone"),
                    Loc.T(openAsFolder ? "split.err.splitFailed" : "split.err.extractFailed")))
                return; // отмена (частичные файлы удалены) или ошибка — база уже всё показала
            if (record != null)
                record(); // успех — учитываем в статистике
            if (compressed > 0)
                UsageStats.RecordPdfCompress(compressed);
            // Разбиение на части считаем файлами (страницы по частям расходятся неравномерно,
            // а при разбиении по диапазонам часть страниц может вообще не попасть в результат).
            // Извлечение даёт ОДИН файл, и число его страниц известно точно — показываем его.
            string status = DoneStatus(openAsFolder, count, pageCount, compressed, level);
            // Если без сжатия результат вышел почти как исходник (общие ресурсы страниц
            // едут вместе с ними) — ненавязчиво подсказать про «Сжатие».
            if (ShouldSuggestCompression(level, sourceSize, largestOutput))
                status += Loc.T("split.status.largeHint");
            SetStatus(status, Theme.OkGreen);
            Ui.OpenPath(openTarget, openAsFolder); // авто-открытие; молча — файлы всё равно созданы
        }

        /// <summary>
        /// Строка успешного завершения. Разбиение на части считаем ФАЙЛАМИ: страницы по частям
        /// расходятся неравномерно, а при разбиении по диапазонам часть страниц может вообще не
        /// попасть в результат, поэтому их число там было бы неправдой. Извлечение даёт ОДИН
        /// файл, и число его страниц известно точно (Extract пишет по странице на индекс).
        /// Отдельным методом, чтобы подстановка значений проверялась тестом, а не только глазами.
        /// Чистая — под тест.
        /// </summary>
        internal static string DoneStatus(bool openAsFolder, int count, int pageCount, int compressed, CompressionLevel level)
        {
            if (!openAsFolder)
                return SuccessStatus(string.Format(Loc.T("split.status.pagesExtracted"), pageCount),
                    CompressedPart(compressed > 0, level));
            string manyFiles = compressed > 0
                ? string.Format(Loc.T("split.suffix.compressed"), compressed, PdfCompression.ImageDpi(level))
                : null;
            return SuccessStatus(string.Format(Loc.T("split.status.filesCreated"), count), manyFiles);
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
