using System;
using System.Collections.Generic;
using System.Diagnostics;
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
    public class PdfSplitForm : PdfToolFormBase, IFileAcceptor
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
        private TextBox _txtNameTemplate; // шаблон имени частей; пусто — прежние имена
        private Label _lblHint;
        private Button _btnDo;

        public PdfSplitForm() : this(null) { }

        public PdfSplitForm(Action showHub) : base(showHub)
        {
            BuildUi();
            UpdateModeInputs();
            SyncControls();
        }

        private void BuildUi()
        {
            InitShell(Title, new Size(800, 660), new Size(700, 560), Theme.PdfRed);
            // Дроп PDF на окно — открыть первый файл (разделение работает с одним документом).
            WireFileDrop(delegate(string[] paths) { LoadSource(paths[0]); });
            BuildHeaderWithHome(Title,
                Loc.T("split.header.subtitle"),
                Theme.PdfRed, Theme.PdfRedDark, ShowHelp, BuildToolsMenu());

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;
            int panelW = 210;
            int gridBottom = ClientSize.Height - 152;

            _grid = new PdfPageGrid();
            _grid.AllowReorder = false; // разделение не меняет порядок исходника
            _grid.AllowRotate = true;   // но поворот страниц в ИТОГОВЫХ файлах — можно
            // ShowPositionNumbers = false: под плиткой — номер исходной страницы.
            _grid.EmptyHint = Loc.T("split.grid.empty");
            _grid.DropHint = Loc.T("split.grid.drop");
            _grid.SetBounds(20, m + 84, right - 20 - panelW, gridBottom - (m + 84));
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _grid.SelectionChanged += delegate { SyncControls(); RefreshRestingStatus(); };
            _grid.FilesDropped += delegate(string[] paths, int insertAt)
            {
                if (paths.Length > 0)
                    LoadSource(paths[0]); // разделение работает с одним документом (LoadSource гейтит Working)
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

            // Масштаб, сжатие и статус — общий нижний строй (как в «Объединении»).
            BuildBottomStrip(right, Loc.T("split.status.openPdf"), 190);

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

        // ---------- дополнительные преобразования одного документа ----------

        /// <summary>
        /// Пункт меню «Ещё»: сохранить страницы картинками, извлечь текст, перевести в серое,
        /// восстановить повреждённый файл. Все они работают с ОДНИМ документом, поэтому живут
        /// здесь, а не в объединении. Доступность пересчитывается при открытии меню: без
        /// загруженного документа остаётся только восстановление — оно как раз и нужно тогда,
        /// когда файл не открывается.
        /// </summary>
        private ToolStripMenuItem BuildToolsMenu()
        {
            var root = new ToolStripMenuItem(Loc.T("split.menu.more"));

            var images = new ToolStripMenuItem(Loc.T("split.menu.toImages"));
            foreach (int dpi in PdfExportService.DpiChoices)
            {
                int chosen = dpi; // копия для замыкания: иначе все пункты возьмут последнее значение
                images.DropDownItems.Add(string.Format(Loc.T("split.menu.dpi"), chosen), null,
                    delegate { ExportImages(chosen); });
            }
            root.DropDownItems.Add(images);
            var text = new ToolStripMenuItem(Loc.T("split.menu.toText"), null, delegate { ExportText(); });
            root.DropDownItems.Add(text);
            root.DropDownItems.Add(new ToolStripSeparator());
            var gray = new ToolStripMenuItem(Loc.T("split.menu.grayscale"), null,
                delegate { ConvertCopy(PdfConvertMode.Grayscale, _sourcePath); });
            root.DropDownItems.Add(gray);
            root.DropDownItems.Add(Loc.T("split.menu.repair"), null, delegate { RepairChosenFile(); });
            root.DropDownItems.Add(new ToolStripSeparator());
            var meta = new ToolStripMenuItem(Loc.T("split.menu.metadata"), null, delegate { EditMetadata(); });
            root.DropDownItems.Add(meta);

            root.DropDownOpening += delegate
            {
                bool ready = !Working && _sourcePath != null;
                images.Enabled = ready;
                text.Enabled = ready;
                gray.Enabled = ready && Ghostscript.Available;
                meta.Enabled = ready;
            };
            return root;
        }

        /// <summary>Сохранить выбранные (или все) страницы картинками в выбранную папку.</summary>
        private void ExportImages(int dpi)
        {
            if (Working || _sourcePath == null)
                return;
            List<int> pages = PagesForExport();
            string dir = FolderPicker.Show(this, Loc.T("split.pick.imagesDir"), Path.GetDirectoryName(_sourcePath));
            if (string.IsNullOrEmpty(dir))
                return;
            ImageExportFormat format = Dialogs.ConfirmWarning(this, Title, Loc.T("split.ask.jpeg.title"),
                Loc.T("split.ask.jpeg.body")) ? ImageExportFormat.Jpeg : ImageExportFormat.Png;

            string source = _sourcePath;
            BeginOperation(Loc.T("split.status.exporting"), pages.Count, Loc.T("split.status.exportingPage"));
            Action<int, int> onProgress = UiProgress();
            Func<bool> cancel = CancelToken();
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                List<string> files = null;
                try
                {
                    files = PdfExportService.ToImages(source, pages, dir, NameTemplate.Default, format, dpi,
                        onProgress, cancel);
                }
                catch (Exception ex) { error = ex; }
                int count = files == null ? 0 : files.Count;
                OnUi(delegate { OnExportFinished(error, count, dir, true); });
            });
        }

        /// <summary>Извлечь текстовый слой документа в .txt.</summary>
        private void ExportText()
        {
            if (Working || _sourcePath == null)
                return;
            string outPath;
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = Loc.T("split.txtFilter");
                dialog.FileName = Path.GetFileNameWithoutExtension(_sourcePath) + ".txt";
                dialog.InitialDirectory = Path.GetDirectoryName(_sourcePath);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                outPath = dialog.FileName;
            }
            string source = _sourcePath;
            BeginOperation(Loc.T("split.status.extractingText"), _pageCount);
            Action<int, int> onProgress = UiProgress();
            Func<bool> cancel = CancelToken();
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                try { PdfExportService.ToText(source, outPath, onProgress, cancel); }
                catch (Exception ex) { error = ex; }
                OnUi(delegate { OnExportFinished(error, 1, outPath, false); });
            });
        }

        /// <summary>
        /// Преобразовать документ, записав результат в НОВЫЙ файл: исходники приложение не
        /// меняет никогда. Движок правит файл на месте, поэтому сначала делаем копию — она и
        /// становится результатом. Не получилось — копию убираем, чтобы не оставлять огрызок.
        /// </summary>
        private void ConvertCopy(PdfConvertMode mode, string source)
        {
            if (Working || string.IsNullOrEmpty(source))
                return;
            string outPath;
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = Loc.T("common.pdfSaveFilter");
                dialog.FileName = Path.GetFileNameWithoutExtension(source) +
                    Loc.T(mode == PdfConvertMode.Grayscale ? "split.suffix.gray" : "split.suffix.repaired") + ".pdf";
                dialog.InitialDirectory = Path.GetDirectoryName(source);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                outPath = dialog.FileName;
            }
            BeginOperation(Loc.T("split.status.converting"), 0);
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                bool ok = false;
                try
                {
                    File.Copy(source, outPath, true);
                    ok = PdfConvert.Apply(outPath, mode);
                    if (!ok)
                        try { File.Delete(outPath); } catch { } // не вышло — огрызок не оставляем
                }
                catch (Exception ex) { error = ex; }
                bool applied = ok;
                OnUi(delegate
                {
                    if (!FinishOperation(error, Loc.T("common.status.notDone"), Loc.T("split.err.convertFailed")))
                        return;
                    if (!applied)
                    {
                        SetStatus(Loc.T("split.status.convertFailed"), Theme.ErrRed);
                        return;
                    }
                    SetStatus(SuccessStatus(Loc.T("split.status.converted")), Theme.OkGreen);
                    Ui.OpenPath(outPath);
                });
            });
        }

        /// <summary>
        /// Восстановление выбирает файл своим диалогом: повреждённый документ в сетку не
        /// открывается, а чинить нужно именно такой. Требовать сперва открыть его значило бы
        /// сделать функцию недоступной ровно тогда, когда она нужна.
        /// </summary>
        private void RepairChosenFile()
        {
            if (Working)
                return;
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = Loc.T("common.pdfFilter");
                dialog.Title = Loc.T("split.pick.repair");
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                ConvertCopy(PdfConvertMode.Repair, dialog.FileName);
            }
        }

        /// <summary>
        /// Правка свойств документа. Результат пишется в НОВЫЙ файл: исходники приложение не
        /// меняет. Пустое поле очищает свойство — так из файла убирают имя автора перед отправкой.
        /// </summary>
        private void EditMetadata()
        {
            if (Working || _sourcePath == null)
                return;
            PdfMetadata edited = MetadataForm.Edit(this, PdfMetadataService.Read(_sourcePath));
            if (edited == null)
                return; // пользователь отказался
            string outPath;
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = Loc.T("common.pdfSaveFilter");
                dialog.FileName = Path.GetFileNameWithoutExtension(_sourcePath) + Loc.T("split.suffix.meta") + ".pdf";
                dialog.InitialDirectory = Path.GetDirectoryName(_sourcePath);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                outPath = dialog.FileName;
            }
            string source = _sourcePath;
            BeginOperation(Loc.T("split.status.savingMeta"), 0);
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                try { PdfMetadataService.Write(source, outPath, edited); }
                catch (Exception ex) { error = ex; }
                OnUi(delegate
                {
                    if (!FinishOperation(error, Loc.T("common.status.notDone"), Loc.T("split.err.metaFailed")))
                        return;
                    SetStatus(SuccessStatus(Loc.T("split.status.metaSaved")), Theme.OkGreen);
                    Ui.OpenPath(outPath);
                });
            });
        }

        /// <summary>Страницы для экспорта: выбранные в сетке, а если ничего не выбрано — все.</summary>
        private List<int> PagesForExport()
        {
            var pages = new List<int>(_grid.GetSelectedIndices());
            if (pages.Count == 0)
                for (int i = 0; i < _pageCount; i++)
                    pages.Add(i);
            pages.Sort();
            return pages;
        }

        private void OnExportFinished(Exception error, int count, string openTarget, bool asFolder)
        {
            if (!FinishOperation(error, Loc.T("common.status.notDone"), Loc.T("split.err.exportFailed")))
                return;
            SetStatus(SuccessStatus(string.Format(Loc.T("split.status.exported"), count)), Theme.OkGreen);
            Ui.OpenPath(openTarget, asFolder);
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

        /// <summary>Idle-статус: не открыт — подсказка открытия, иначе — имя файла и число страниц.</summary>
        protected override string IdleStatusText()
        {
            return _sourcePath == null
                ? Loc.T("split.status.openPdf")
                : string.Format(Loc.T("split.status.opened"), Path.GetFileName(_sourcePath), _pageCount);
        }

        /// <summary>Дроп PDF на карточку хаба: открыть первый файл (разделение — один документ).</summary>
        public void AcceptFiles(string[] paths)
        {
            if (paths != null && paths.Length > 0)
                LoadSource(paths[0]); // LoadSource гейтит Working
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

        /// <summary>
        /// Открыть исходный документ: разбор PDF идёт в фоне (большой/сетевой файл не морозит
        /// окно), затем сетка страниц и статус. Битый/занятый файл — диалог, прежний документ цел.
        /// </summary>
        private void LoadSource(string path)
        {
            if (!BeginLoad(Loc.T("common.status.loading")))
                return; // уже идёт операция или загрузка
            Ui.RunWorker(delegate()
            {
                int count = 0;
                string error = null;
                // Ловим ШИРОКО: битый/занятый/аварийный файл (в т.ч. редкий OOM, который
                // LoadPages НЕ оборачивает) не должен ронять фоновый поток — только диалог.
                try { count = PdfMergeService.LoadPages(path).Count; }
                catch (MergeException ex) { error = ex.Message; }
                catch (Exception ex) { error = string.Format(Loc.T("err.pdf.cantOpen"), Path.GetFileName(path), ex.Message); }
                int pages = count;
                string err = error;
                OnUi(delegate { ApplyLoadedSource(path, pages, err); });
            });
        }

        /// <summary>Применить результат фонового разбора исходника (UI-поток).</summary>
        private void ApplyLoadedSource(string path, int pageCount, string error)
        {
            EndLoad();
            if (error != null)
            {
                RefreshRestingStatus(); // вернуть статус прежнего документа вместо залипшей «Загрузка…»
                Dialogs.Error(this, Title, Loc.T("split.err.fileNotOpened"), error); // прежний документ остаётся
                return;
            }
            _sourcePath = path;
            _pageCount = pageCount;
            _pages = new List<PdfPageRef>();
            for (int i = 0; i < _pageCount; i++)
                _pages.Add(new PdfPageRef { SourcePath = path, PageIndex = i });
            _grid.SetPages(_pages);
            SetStatus(string.Format(Loc.T("split.status.opened"), Path.GetFileName(path), _pageCount), Theme.TextMuted);
            SyncControls();
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
