using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger
{
    public class MainForm : Form
    {
        private const string AppTitle = "iwo Helper Desktop";
        private const int MenuHeight = HelpMenu.Height;

        private TextBox _txtInput;
        private Button _btnBrowseInput;
        private Label _lblFound;
        private TextBox _txtName;
        private TextBox _txtOutDir;
        private Button _btnBrowseOut;
        private CheckBox _chkToc;
        private CheckBox _chkValues;
        private ComboBox _cmbFormat;
        private ComboBox _cmbScope;
        private ToolTip _tips;
        private Button _btnMerge;
        private Button _btnCancel;
        private ProgressBar _progress;
        private Label _lblStatus;
        private ListView _list;
        private LinkLabel _lnkOpenFile;
        private LinkLabel _lnkOpenFolder;
        private LinkLabel _lnkOpenReport;
        private LinkLabel _lnkNote;
        private Button _btnRetry;

        private readonly Action _showHub;
        private MergeService _service;
        private Thread _worker;
        private UserSettings _settings;
        private readonly TaskbarProgress _taskbar = new TaskbarProgress();
        private string _lastOutputPath;
        private string _lastReportPath;
        private string _lastInputFolder;
        private MergeOptions _lastOptions;
        private DateTime _lastStartedAt;
        private MergeResult _lastResult;

        // Список файлов до слияния: порядок и включение.
        private readonly SourceFileList _files = new SourceFileList();
        private readonly Dictionary<string, ListViewItem> _rowByPath =
            new Dictionary<string, ListViewItem>(StringComparer.OrdinalIgnoreCase);
        private Button _btnUp;
        private Button _btnDown;
        private Button _btnSortName;
        private Button _btnCheckAll;
        private Button _btnUncheckAll;
        private Panel _dropLine;   // индикатор вставки при перетаскивании
        private System.Windows.Forms.Timer _inputDebounce; // отложенное сканирование папки
        private int _dragIndex = -1;
        private bool _populating;  // подавляет ItemChecked во время заполнения
        private bool _running;        // истина от нажатия «Объединить» до OnMergeFinished (только UI-поток)
        private bool _noteBusy;       // готовится записка Word (только UI-поток)
        private bool _closeRequested; // пользователь закрыл окно во время объединения
        private bool _isFreshRun;     // прогон — новый свод (не дослияние), для счётчика статистики

        public MainForm() : this(null) { }

        public MainForm(Action showHub)
        {
            _showHub = showHub;
            BuildUi();

            _settings = UserSettings.Load();
            if (!string.IsNullOrEmpty(_settings.LastInputFolder) && Directory.Exists(_settings.LastInputFolder))
                _txtInput.Text = _settings.LastInputFolder;
            if (!string.IsNullOrEmpty(_settings.LastOutputFolder) && Directory.Exists(_settings.LastOutputFolder))
                _txtOutDir.Text = _settings.LastOutputFolder;
            _txtName.Text = Loc.T("excel.defaultName") + DateTime.Now.ToString("yyyy-MM-dd");
            _chkToc.Checked = _settings.AddToc;
            // «Заменить формулы значениями» всегда стартует выключенным: режим
            // меняет содержимое свода и включается осознанно на каждый запуск.
            int formatIndex = Array.IndexOf(OutputFormats.Extensions, _settings.OutputExtension);
            _cmbFormat.SelectedIndex = formatIndex >= 0 ? formatIndex : 0;
            _cmbScope.SelectedIndex = _settings.AllSheets ? 1 : 0;
            if (_inputDebounce != null)
                _inputDebounce.Stop(); // стартовая загрузка — сразу, без отложенного повтора
            RefreshFileList();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_inputDebounce != null)
                    _inputDebounce.Dispose();
                // ToolTip — компонент, а не дочерний контрол: авто-освобождение не срабатывает.
                if (_tips != null)
                    _tips.Dispose();
            }
            base.Dispose(disposing);
        }

        // ---------- построение интерфейса ----------

        private void BuildUi()
        {
            Text = Loc.T("hub.excel.name"); // имя окна = имя инструмента (как у «Объединение PDF»), не как у хаба
            Icon appIcon = Ui.AppIcon();
            if (appIcon != null)
                Icon = appIcon;
            Font = new Font("Segoe UI", 9.75f);
            BackColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(780, 785);
            // Ширина минимума не даёт нижним ссылкам («Записка Word») перекрыться
            // с правой кнопкой «Повторить пропущенные».
            MinimumSize = new Size(760, 725);
            WindowChrome.Enable(this, Theme.Accent); // зелёный заголовок на Windows 11
            AllowDrop = true;
            DragEnter += OnDragEnter;
            DragDrop += OnDragDrop;
            _tips = new ToolTip();

            int right = ClientSize.Width - 20;
            var stretch = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            BuildMenu();
            var header = new HeaderBand(Loc.T("hub.excel.name"),
                Loc.T("excel.header.subtitle"),
                Theme.Accent, Theme.AccentPressed);
            header.SetBounds(0, MenuHeight, ClientSize.Width, 82);
            header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            header.TabIndex = 100; // «Главная» — в конце обхода Tab, а не в начале
            Controls.Add(header);
            AddHomeButton(header);

            // Шаг 1: исходная папка
            AddSectionLabel(Loc.T("excel.sec.inputFolder"), 115);
            _txtInput = AddTextBox(20, 135, right - 20 - 110);
            // Дебаунс: набор пути вручную не сканирует папку на каждый символ.
            _inputDebounce = new System.Windows.Forms.Timer();
            _inputDebounce.Interval = 300;
            _inputDebounce.Tick += delegate { _inputDebounce.Stop(); RefreshFileList(); };
            _txtInput.TextChanged += delegate { _inputDebounce.Stop(); _inputDebounce.Start(); };
            _btnBrowseInput = AddButton(Loc.T("common.browse"), false, right - 100, 134, 100, 29);
            _btnBrowseInput.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnBrowseInput.Click += OnBrowseInput;
            _lblFound = Ui.Label(this, "", 20, 167, Font, Theme.TextMuted);

            // Шаг 2: итоговый файл
            AddSectionLabel(Loc.T("excel.sec.output"), 199);
            Ui.Label(this, Loc.T("excel.lbl.name"), 20, 222, Font, Theme.TextPrimary);
            _txtName = AddTextBox(75, 219, 300);
            _txtName.Anchor = AnchorStyles.Top | AnchorStyles.Left; // фиксированная ширина: справа выбор формата
            _txtName.TextChanged += delegate { UpdateReadiness(); };
            _cmbFormat = new ComboBox();
            _cmbFormat.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbFormat.Items.AddRange(OutputFormats.Extensions);
            _cmbFormat.SetBounds(383, 219, 85, 27);
            Controls.Add(_cmbFormat);
            _tips.SetToolTip(_cmbFormat, Loc.T("excel.tip.format"));

            Ui.Label(this, Loc.T("excel.lbl.folder"), 20, 258, Font, Theme.TextPrimary);
            _txtOutDir = AddTextBox(75, 255, right - 75 - 110);
            _btnBrowseOut = AddButton(Loc.T("common.browse"), false, right - 100, 254, 100, 29);
            _btnBrowseOut.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnBrowseOut.Click += OnBrowseOutput;

            // Параметры
            AddSectionLabel(Loc.T("excel.sec.params"), 289);
            Ui.Label(this, Loc.T("excel.lbl.sheets"), 20, 312, Font, Theme.TextPrimary);
            _cmbScope = new ComboBox();
            _cmbScope.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbScope.Items.AddRange(new object[] { Loc.T("excel.scope.first"), Loc.T("excel.scope.all") });
            _cmbScope.SetBounds(75, 309, 200, 27);
            Controls.Add(_cmbScope);
            _tips.SetToolTip(_cmbScope, Loc.T("excel.tip.scope"));

            _chkToc = AddCheckBox(Loc.T("excel.chk.toc"), 20, 345, Loc.T("excel.tip.toc"));
            _chkValues = AddCheckBox(Loc.T("excel.chk.values"), 20, 371, Loc.T("excel.tip.values"));

            // Действия: основная слева, отмена — у правого края, подальше от основной
            _btnMerge = AddButton(Loc.T("excel.btn.merge"), true, 20, 407, 170, 40);
            _btnMerge.Click += OnMergeClick;
            AcceptButton = _btnMerge;
            _tips.SetToolTip(_btnMerge, Loc.T("excel.tip.merge"));

            _btnCancel = AddButton(Loc.T("excel.btn.cancel"), false, right - 130, 407, 130, 40);
            _btnCancel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnCancel.Enabled = false;
            _btnCancel.Click += OnCancelClick;
            CancelButton = _btnCancel;
            _tips.SetToolTip(_btnCancel, Loc.T("excel.tip.cancel"));

            _progress = new ProgressBar();
            _progress.SetBounds(20, 465, right - 20, 8);
            _progress.Anchor = stretch;
            _progress.Visible = false; // пустая полоса в простое только путает
            Controls.Add(_progress);

            _lblStatus = Ui.Label(this, Loc.T("excel.status.chooseFolder"), 20, 483, Font, Theme.TextMuted);

            // Файлы: порядок и состав до слияния; после — результат в тех же строках
            AddSectionLabel(Loc.T("excel.sec.files"), 508);
            _list = new ListView();
            _list.SetBounds(20, 528, right - 20 - 150, ClientSize.Height - 528 - 44);
            _list.Anchor = stretch | AnchorStyles.Bottom;
            _list.View = View.Details;
            _list.FullRowSelect = true;
            _list.CheckBoxes = true; // снятая галочка исключает файл из свода
            _list.HeaderStyle = ColumnHeaderStyle.Nonclickable; // порядок задаёт пользователь
            _list.BorderStyle = BorderStyle.FixedSingle;
            _list.Columns.Add(Loc.T("excel.col.file"), 300);
            _list.Columns.Add(Loc.T("excel.col.result"), 130);
            _list.Columns.Add(Loc.T("excel.col.note"), 180);
            EnableDoubleBuffer(_list);
            _list.ItemChecked += OnItemChecked;
            _list.ItemCheck += delegate(object s, ItemCheckEventArgs e)
            {
                if (_running) // во время прогона состав не меняем
                    e.NewValue = e.CurrentValue;
            };
            _list.SelectedIndexChanged += delegate { UpdateListButtons(); };
            _list.AllowDrop = true;
            _list.ItemDrag += OnFileItemDrag;
            _list.DragOver += OnFileDragOver;
            _list.DragDrop += OnFileDragDrop;
            _list.DragLeave += delegate { HideDropLine(); };
            var copyMenu = new ContextMenuStrip();
            var copyItem = new ToolStripMenuItem(Loc.T("common.copy"));
            copyItem.ShortcutKeyDisplayString = "Ctrl+C";
            copyItem.Click += delegate { CopySelectedRows(); };
            copyMenu.Items.Add(copyItem);
            copyMenu.Opening += delegate { copyItem.Enabled = _list.SelectedItems.Count > 0; };
            _list.ContextMenuStrip = copyMenu;
            // Все горячие клавиши списка — в ProcessCmdKey (см. ниже): Enter как
            // диалоговая клавиша AcceptButton перехватывается до KeyDown, поэтому
            // KeyDown для неё бесполезен; заодно единое место для остальных.
            Controls.Add(_list);

            // Кнопки управления списком (правая колонка на уровне списка)
            int fcol = right - 140;
            _btnUp = AddListButton(Loc.T("excel.btn.up"), fcol, 528);
            _btnUp.Click += delegate { MoveSelectedFile(false); };
            _tips.SetToolTip(_btnUp, Loc.T("excel.tip.up"));
            _btnDown = AddListButton(Loc.T("excel.btn.down"), fcol, 562);
            _btnDown.Click += delegate { MoveSelectedFile(true); };
            _tips.SetToolTip(_btnDown, Loc.T("excel.tip.down"));
            _btnSortName = AddListButton(Loc.T("excel.btn.sortName"), fcol, 604);
            _btnSortName.Click += delegate { SortFilesByName(); };
            _tips.SetToolTip(_btnSortName, Loc.T("excel.tip.sortName"));
            _btnCheckAll = AddListButton(Loc.T("excel.btn.checkAll"), fcol, 646);
            _btnCheckAll.Click += delegate { SetAllChecked(true); };
            _btnUncheckAll = AddListButton(Loc.T("excel.btn.uncheckAll"), fcol, 680);
            _btnUncheckAll.Click += delegate { SetAllChecked(false); };

            _dropLine = new Panel();
            _dropLine.Height = 2;
            _dropLine.BackColor = Theme.Accent;
            _dropLine.Visible = false;
            Controls.Add(_dropLine);
            _dropLine.BringToFront();

            _btnRetry = AddButton(Loc.T("excel.btn.retry"), false, right - 200, ClientSize.Height - 40, 200, 30);
            _btnRetry.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _btnRetry.Visible = false;
            _btnRetry.Click += OnRetryClick;
            _tips.SetToolTip(_btnRetry, Loc.T("excel.tip.retry"));

            _lnkOpenFile = AddBottomLink(Loc.T("excel.link.openFile"), 20);
            _lnkOpenFile.LinkClicked += delegate { OpenPath(_lastOutputPath, false); };
            _lnkOpenFolder = AddBottomLink(Loc.T("excel.link.openFolder"), 145);
            _lnkOpenFolder.LinkClicked += delegate { OpenPath(_lastOutputPath, true); };
            _lnkOpenReport = AddBottomLink(Loc.T("excel.link.openReport"), 280);
            _lnkOpenReport.LinkClicked += delegate { OpenPath(_lastReportPath, false); };
            _tips.SetToolTip(_lnkOpenReport, Loc.T("excel.tip.openReport"));
            _lnkNote = AddBottomLink(Loc.T("excel.link.note"), 415);
            _lnkNote.LinkClicked += delegate { OnNoteClick(); };
            _tips.SetToolTip(_lnkNote, Loc.T("excel.tip.note"));

            _tips.SetToolTip(_txtInput, Loc.T("excel.tip.input"));
            _tips.SetToolTip(_txtName, Loc.T("excel.tip.name"));
            _tips.SetToolTip(_txtOutDir, Loc.T("excel.tip.outDir"));

            Resize += delegate { AdjustNoteColumn(); };
            AdjustNoteColumn();
            UpdateReadiness();
        }

        private void AddHomeButton(HeaderBand header)
        {
            if (_showHub == null)
                return; // запущено вне хаба (напр. автотест)
            Button home = Ui.HomeButton(_showHub);
            home.SetBounds(header.Width - 180, 24, 160, 30);
            home.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _tips.SetToolTip(home, Loc.T("common.homeTip"));
            header.Controls.Add(home);
        }

        private void BuildMenu()
        {
            var reports = new ToolStripMenuItem(Loc.T("excel.menu.reports"));
            reports.Click += delegate { OpenReportsFolder(); };

            MenuStrip menu = HelpMenu.Create(this, ShowHelp, reports);
            MainMenuStrip = menu;
            Controls.Add(menu);
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, AppTitle, Loc.T("menu.howTo"), Loc.T("excel.help.body"));
        }

        private void OpenReportsFolder()
        {
            try
            {
                Directory.CreateDirectory(AppPaths.ReportsDir);
                Process.Start("explorer.exe", "\"" + AppPaths.ReportsDir + "\"");
            }
            catch (Exception ex)
            {
                Dialogs.Error(this, AppTitle, Loc.T("excel.err.openReports"), ex.Message);
            }
        }

        private void AddSectionLabel(string text, int y)
        {
            Ui.Label(this, text, 20, y, new Font("Segoe UI", 8f, FontStyle.Bold), Theme.TextMuted);
        }

        private CheckBox AddCheckBox(string text, int x, int y, string tooltip)
        {
            var c = new AccentCheckBox();
            c.Text = text;
            c.Location = new Point(x, y);
            c.AutoSize = true;
            c.ForeColor = Theme.TextPrimary;
            Controls.Add(c);
            _tips.SetToolTip(c, tooltip);
            return c;
        }

        private TextBox AddTextBox(int x, int y, int width)
        {
            var t = new TextBox();
            t.SetBounds(x, y, width, 27);
            t.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(t);
            return t;
        }

        private Button AddButton(string text, bool primary, int x, int y, int w, int h)
        {
            var b = new RoundedButton(primary);
            b.Text = text;
            b.SetBounds(x, y, w, h);
            Controls.Add(b);
            return b;
        }

        private LinkLabel AddBottomLink(string text, int x)
        {
            LinkLabel l = Ui.Link(this, text, x, ClientSize.Height - 32);
            l.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            l.Visible = false;
            return l;
        }

        /// <summary>Выбранные строки — в буфер: строки отчёта по листам, а до слияния — имена.</summary>
        private void CopySelectedRows()
        {
            if (_list.SelectedItems.Count == 0)
                return;
            var sb = new StringBuilder();
            foreach (ListViewItem item in _list.SelectedItems)
            {
                var results = item.Tag as List<FileResult>;
                if (results != null && results.Count > 0)
                    foreach (FileResult fr in results)
                        sb.AppendLine(ReportWriter.FormatFileLine(fr));
                else
                    sb.AppendLine(item.Text); // предпросмотр — имя файла
            }
            if (sb.Length == 0)
                return;
            try { Clipboard.SetText(sb.ToString()); }
            catch { } // буфер обмена занят другим приложением — не повод падать
        }

        private static void EnableDoubleBuffer(ListView list)
        {
            // Убирает мерцание при добавлении строк; свойство защищённое — только через reflection.
            PropertyInfo p = typeof(ListView).GetProperty("DoubleBuffered",
                BindingFlags.Instance | BindingFlags.NonPublic);
            if (p != null)
                p.SetValue(list, true, null);
        }

        // Доли колонок: Файл / Результат / Примечание.
        private static readonly float[] ColumnWeights = { 0.48f, 0.22f, 0.30f };

        /// <summary>Колонки списка делят ширину пропорционально.</summary>
        private void AdjustNoteColumn()
        {
            int width = _list.ClientSize.Width - 4;
            if (width < 280)
                return;
            for (int i = 0; i < _list.Columns.Count && i < ColumnWeights.Length; i++)
                _list.Columns[i].Width = (int)(width * ColumnWeights[i]);
        }

        private Button AddListButton(string text, int x, int y)
        {
            var b = new RoundedButton(false);
            b.Text = text;
            b.SetBounds(x, y, 130, 30);
            b.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(b);
            return b;
        }

        // ---------- выбор папок, drag&drop, живая валидация ----------

        private void OnBrowseInput(object sender, EventArgs e)
        {
            string path = FolderPicker.Show(this, Loc.T("excel.pick.input"), _txtInput.Text.Trim());
            if (path != null)
                SetInputFolder(path);
        }

        private void OnBrowseOutput(object sender, EventArgs e)
        {
            string initial = _txtOutDir.Text.Trim();
            if (initial.Length == 0)
                initial = _txtInput.Text.Trim();
            string path = FolderPicker.Show(this, Loc.T("excel.pick.output"), initial);
            if (path != null)
                _txtOutDir.Text = path;
        }

        private void SetInputFolder(string path)
        {
            _txtInput.Text = path;
            if (_txtOutDir.Text.Trim().Length == 0)
                _txtOutDir.Text = path;
        }

        private void OnDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = !_running && ExtractDroppedFolder(e) != null
                ? DragDropEffects.Copy
                : DragDropEffects.None;
        }

        private void OnDragDrop(object sender, DragEventArgs e)
        {
            if (_running)
                return;
            string folder = ExtractDroppedFolder(e);
            if (folder != null)
                SetInputFolder(folder);
        }

        private static string ExtractDroppedFolder(DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return null;
            var items = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (items == null || items.Length == 0)
                return null;
            if (Directory.Exists(items[0]))
                return items[0];
            if (File.Exists(items[0]))
                return Path.GetDirectoryName(items[0]); // брошен файл — берём его папку
            return null;
        }

        /// <summary>Перечитывает папку в список файлов (все включены, естественный порядок).</summary>
        private void RefreshFileList()
        {
            if (_running)
                return;
            string folder = _txtInput.Text.Trim();
            List<string> paths = new List<string>();
            if (folder.Length == 0)
                SetFoundLabel(Loc.T("excel.found.chooseFolder"), Theme.TextMuted);
            else if (!Directory.Exists(folder))
                SetFoundLabel(Loc.T("excel.found.notFound"), Theme.ErrRed);
            else
            {
                try { paths = MergeService.FindSourceFiles(folder, null); }
                catch (Exception ex) { SetFoundLabel(string.Format(Loc.T("excel.found.readError"), ex.Message), Theme.ErrRed); }
            }
            _files.SetFiles(paths);
            RebuildFileRows();
        }

        private void RebuildFileRows()
        {
            _populating = true;
            _list.BeginUpdate();
            _list.Items.Clear();
            _rowByPath.Clear();
            for (int i = 0; i < _files.Count; i++)
            {
                SourceFile f = _files[i];
                var item = new ListViewItem(f.FileName);
                item.Checked = f.Include;
                item.SubItems.Add("");   // Результат
                item.SubItems.Add("");   // Примечание
                item.Tag = new List<FileResult>(); // результаты по этому файлу
                item.ToolTipText = f.Path;
                _list.Items.Add(item);
                _rowByPath[Path.GetFullPath(f.Path)] = item;
            }
            _list.EndUpdate();
            _populating = false;
            UpdateFoundLabel();
            UpdateReadiness();
            UpdateListButtons();
        }

        private void UpdateFoundLabel()
        {
            if (_files.Count == 0)
            {
                if (_txtInput.Text.Trim().Length > 0 && Directory.Exists(_txtInput.Text.Trim()))
                    SetFoundLabel(Loc.T("excel.found.noExcel"), Theme.ErrRed);
                return;
            }
            int inc = _files.IncludedCount;
            SetFoundLabel(string.Format(Loc.T("excel.found.count"), _files.Count, inc), Theme.OkGreen);
        }

        private void SetFoundLabel(string text, Color color)
        {
            _lblFound.Text = text;
            _lblFound.ForeColor = color;
        }

        private void OnItemChecked(object sender, ItemCheckedEventArgs e)
        {
            if (_populating)
                return;
            int index = e.Item.Index;
            if (index >= 0 && index < _files.Count)
                _files[index].Include = e.Item.Checked;
            UpdateFoundLabel();
            UpdateReadiness();
        }

        // ---------- порядок и состав ----------

        /// <summary>Действие клавиатуры для сфокусированного файл-листа. Чистая — под тест.</summary>
        internal enum ListKeyAction { None, MoveUp, MoveDown, Copy, SelectAll, Exclude, Swallow }

        internal static ListKeyAction ClassifyListKey(Keys keyData)
        {
            if (keyData == (Keys.Alt | Keys.Up)) return ListKeyAction.MoveUp;
            if (keyData == (Keys.Alt | Keys.Down)) return ListKeyAction.MoveDown;
            if (keyData == (Keys.Control | Keys.C)) return ListKeyAction.Copy;
            if (keyData == (Keys.Control | Keys.A)) return ListKeyAction.SelectAll;
            if (keyData == Keys.Delete) return ListKeyAction.Exclude; // снять галочку у выбранных
            if (keyData == Keys.Enter) return ListKeyAction.Swallow;  // не запускать слияние из списка
            return ListKeyAction.None;
        }

        // ProcessCmdKey срабатывает раньше диалоговой обработки (AcceptButton),
        // поэтому и подавление Enter, и модификаторные сочетания здесь надёжны.
        // Copy/SelectAll допустимы всегда; перестановка и исключение — методы
        // сами не выполняются во время прогона.
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (_list != null && _list.Focused)
            {
                switch (ClassifyListKey(keyData))
                {
                    case ListKeyAction.Copy: CopySelectedRows(); return true;
                    case ListKeyAction.SelectAll: SelectAllRows(); return true;
                    case ListKeyAction.Swallow: return true;
                    case ListKeyAction.MoveUp: MoveSelectedFile(false); return true;
                    case ListKeyAction.MoveDown: MoveSelectedFile(true); return true;
                    case ListKeyAction.Exclude: ExcludeSelectedRows(); return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void SelectAllRows()
        {
            if (_list.Items.Count == 0)
                return;
            _list.BeginUpdate();
            foreach (ListViewItem item in _list.Items)
                item.Selected = true;
            _list.EndUpdate();
        }

        /// <summary>Delete в списке — снять галочку (исключить) у выбранных файлов.</summary>
        private void ExcludeSelectedRows()
        {
            if (_running || _list.SelectedItems.Count == 0)
                return;
            _populating = true;
            foreach (ListViewItem item in _list.SelectedItems)
            {
                item.Checked = false;
                int idx = item.Index;
                if (idx >= 0 && idx < _files.Count)
                    _files[idx].Include = false;
            }
            _populating = false;
            UpdateFoundLabel();
            UpdateReadiness();
        }

        private void MoveSelectedFile(bool down)
        {
            if (_running || _list.SelectedIndices.Count != 1)
                return;
            int index = _list.SelectedIndices[0];
            int moved = down ? _files.MoveDown(index) : _files.MoveUp(index);
            if (moved == index)
                return;
            RebuildFileRows();
            _list.Items[moved].Selected = true;
            _list.Items[moved].EnsureVisible();
            _list.Focus();
        }

        private void SortFilesByName()
        {
            if (_running)
                return;
            _files.SortByName();
            RebuildFileRows();
        }

        private void SetAllChecked(bool included)
        {
            if (_running)
                return;
            _files.SetAllIncluded(included);
            _populating = true;
            foreach (ListViewItem item in _list.Items)
                item.Checked = included;
            _populating = false;
            UpdateFoundLabel();
            UpdateReadiness();
        }

        private void UpdateListButtons()
        {
            bool one = !_running && _list.SelectedIndices.Count == 1;
            _btnUp.Enabled = one;
            _btnDown.Enabled = one;
            _btnSortName.Enabled = !_running && _files.Count > 1;
            _btnCheckAll.Enabled = !_running && _files.Count > 0;
            _btnUncheckAll.Enabled = !_running && _files.Count > 0;
        }

        // ---------- перетаскивание строк ----------

        private void OnFileItemDrag(object sender, ItemDragEventArgs e)
        {
            if (_running)
                return;
            var item = e.Item as ListViewItem;
            if (item != null)
            {
                _dragIndex = item.Index;
                _list.DoDragDrop(item, DragDropEffects.Move);
            }
        }

        private void OnFileDragOver(object sender, DragEventArgs e)
        {
            if (_running || _dragIndex < 0)
            {
                e.Effect = DragDropEffects.None;
                return;
            }
            e.Effect = DragDropEffects.Move;
            int target = DropTargetIndex(e);
            ShowDropLine(target);
        }

        private void OnFileDragDrop(object sender, DragEventArgs e)
        {
            HideDropLine();
            if (_running || _dragIndex < 0)
                return;
            int target = DropTargetIndex(e);
            int from = _dragIndex;
            _dragIndex = -1;
            _files.Move(from, target);
            RebuildFileRows();
            int landed = target > from ? target - 1 : target;
            if (landed >= 0 && landed < _list.Items.Count)
            {
                _list.Items[landed].Selected = true;
                _list.Items[landed].EnsureVisible();
            }
        }

        /// <summary>Индекс вставки по позиции курсора (0..Count).</summary>
        private int DropTargetIndex(DragEventArgs e)
        {
            Point pt = _list.PointToClient(new Point(e.X, e.Y));
            ListViewItem over = _list.GetItemAt(pt.X, pt.Y);
            if (over == null)
                return _list.Items.Count; // ниже последней строки
            Rectangle b = over.Bounds;
            return pt.Y > b.Top + b.Height / 2 ? over.Index + 1 : over.Index;
        }

        private void ShowDropLine(int target)
        {
            // y — в координатах клиента списка; переносим в координаты формы (+_list.Top).
            int y;
            if (_list.Items.Count == 0)
                y = 4;
            else if (target >= _list.Items.Count)
                y = _list.Items[_list.Items.Count - 1].Bounds.Bottom;
            else
                y = _list.Items[target].Bounds.Top;
            _dropLine.SetBounds(_list.Left + 2, _list.Top + y - 1, _list.Width - 4, 2);
            _dropLine.Visible = true;
            _dropLine.BringToFront();
        }

        private void HideDropLine()
        {
            if (_dropLine != null)
                _dropLine.Visible = false;
        }

        /// <summary>Кнопка «Объединить» активна при выбранных файлах и заданном имени.</summary>
        private void UpdateReadiness()
        {
            UpdateListButtons();
            if (_running)
                return;
            _btnMerge.Enabled = _files.IncludedCount > 0 && _txtName.Text.Trim().Length > 0;
        }

        // ---------- запуск объединения ----------

        private void OnMergeClick(object sender, EventArgs e)
        {
            string folder = _txtInput.Text.Trim();
            if (!Directory.Exists(folder))
            {
                Dialogs.Error(this, AppTitle, Loc.T("excel.err.folderNotFound.title"),
                    string.Format(Loc.T("excel.err.folderNotFound.body"), folder));
                return;
            }

            string name = OutputFormats.StripKnownExtension(_txtName.Text.Trim());
            if (name.Length == 0)
            {
                Dialogs.Error(this, AppTitle, Loc.T("excel.err.noName.title"),
                    Loc.T("excel.err.noName.body"));
                return;
            }
            if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                Dialogs.Error(this, AppTitle, Loc.T("excel.err.badName.title"),
                    Loc.T("excel.err.badName.body"));
                return;
            }

            string outDir = _txtOutDir.Text.Trim();
            if (outDir.Length == 0)
                outDir = folder;
            if (!Directory.Exists(outDir))
            {
                if (!Dialogs.ConfirmWarning(this, AppTitle, Loc.T("excel.confirm.createFolder.title"),
                        string.Format(Loc.T("excel.confirm.createFolder.body"), outDir)))
                    return;
                try
                {
                    Directory.CreateDirectory(outDir);
                }
                catch (Exception ex)
                {
                    Dialogs.Error(this, AppTitle, Loc.T("excel.err.createFolder"), ex.Message);
                    return;
                }
            }

            string outputPath = Path.Combine(outDir, name + (string)_cmbFormat.SelectedItem);

            List<string> files = _files.IncludedInOrder();
            if (files.Count == 0)
            {
                Dialogs.Error(this, AppTitle, Loc.T("excel.err.noFiles.title"),
                    Loc.T("excel.err.noFiles.body"));
                return;
            }

            if (File.Exists(outputPath) &&
                !Dialogs.ConfirmWarning(this, AppTitle, Loc.T("excel.confirm.overwrite.title"),
                    string.Format(Loc.T("excel.confirm.overwrite.body"), Path.GetFileName(outputPath))))
                return;

            // Занятый итоговый файл должен остановить работу сейчас,
            // а не после обработки всех исходных файлов.
            string lockError = MergeService.CheckOutputWritable(outputPath);
            if (lockError != null)
            {
                Dialogs.Error(this, AppTitle, Loc.T("excel.err.outputLocked"), lockError);
                return;
            }

            _settings.LastInputFolder = folder;
            _settings.LastOutputFolder = outDir;
            _settings.AddToc = _chkToc.Checked;
            _settings.AllSheets = _cmbScope.SelectedIndex == 1;
            _settings.OutputExtension = (string)_cmbFormat.SelectedItem;
            _settings.Save();

            var options = new MergeOptions();
            options.AddToc = _chkToc.Checked;
            options.ValuesOnly = _chkValues.Checked;
            options.AllSheets = _cmbScope.SelectedIndex == 1;
            StartMerge(files, folder, outputPath, options);
        }

        private void StartMerge(List<string> files, string folder, string outputPath, MergeOptions options)
        {
            PrepareRun(folder, outputPath, options, true); // свежий прогон — очистить результаты всех строк
            StartWorker(delegate { return _service.Merge(files, outputPath, options); });
        }

        private void OnRetryClick(object sender, EventArgs e)
        {
            MergeResult previous = _lastResult;
            if (_running || previous == null || previous.SkipCount == 0)
                return;
            string outputPath = previous.OutputPath;
            string lockError = MergeService.CheckOutputWritable(outputPath);
            if (lockError != null)
            {
                Dialogs.Error(this, AppTitle, Loc.T("excel.err.outputLocked"), lockError);
                return;
            }
            MergeOptions options = _lastOptions;
            PrepareRun(_lastInputFolder, outputPath, options, false); // повтор — не трогаем успешные строки
            // сбросить результат только у повторяемых (ранее пропущенных) файлов
            foreach (FileResult fr in previous.Files)
            {
                ListViewItem row;
                if (!fr.Ok && fr.FullPath != null && _rowByPath.TryGetValue(Path.GetFullPath(fr.FullPath), out row))
                    ResetRow(row);
            }
            StartWorker(delegate { return _service.RetrySkipped(outputPath, options, previous); });
        }

        private void ResetRow(ListViewItem row)
        {
            ((List<FileResult>)row.Tag).Clear();
            row.SubItems[1].Text = "";
            row.SubItems[2].Text = "";
            row.ForeColor = _list.ForeColor;
        }

        /// <summary>Общий сброс интерфейса перед обычным и повторным слиянием.</summary>
        private void PrepareRun(string folder, string outputPath, MergeOptions options, bool clearAllRows)
        {
            _running = true;
            _isFreshRun = clearAllRows; // clearAllRows=false — это дослияние пропущенных
            SetRunning(true);
            if (clearAllRows)
                foreach (ListViewItem row in _list.Items)
                    ResetRow(row);
            _lnkOpenFile.Visible = false;
            _lnkOpenFolder.Visible = false;
            _lnkOpenReport.Visible = false;
            _lnkNote.Visible = false;
            _btnRetry.Visible = false;
            _progress.Value = 0;
            _progress.Visible = true;
            SetStatus(Loc.T("excel.status.startingExcel"), Theme.TextMuted);
            _lastOutputPath = outputPath;
            _lastInputFolder = folder;
            _lastOptions = options;
            _lastStartedAt = DateTime.Now;
            _taskbar.SetState(Handle, TaskbarProgressState.Indeterminate);

            _service = new MergeService();
            _service.Progress += OnServiceProgress;
            _service.FileDone += OnServiceFileDone;
            _service.Restarting += OnServiceRestarting;
        }

        /// <summary>
        /// Excel перезапускается из-за зависшего файла: прошлый проход отменён и
        /// будет переотправлен, поэтому очищаем накопленные в строках результаты,
        /// чтобы они не задвоились.
        /// </summary>
        private void OnServiceRestarting()
        {
            OnUi(delegate
            {
                foreach (ListViewItem row in _list.Items)
                    ResetRow(row);
            });
        }

        /// <summary>Запуск фонового STA-потока: исключения доставляются в UI, поток не падает.</summary>
        private void StartWorker(Func<MergeResult> work)
        {
            _worker = new Thread(delegate()
            {
                MergeResult result = null;
                Exception error = null;
                try
                {
                    result = work();
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                OnUi(delegate { OnMergeFinished(result, error); });
            });
            _worker.SetApartmentState(ApartmentState.STA); // требование Excel COM
            _worker.IsBackground = true;
            _worker.Start();
        }

        /// <summary>Безопасная доставка действия в UI-поток из фонового.</summary>
        private void OnUi(Action action)
        {
            try
            {
                if (IsHandleCreated && !IsDisposed)
                    BeginInvoke(action);
            }
            catch (InvalidOperationException) { } // окно уже разрушено — доставлять некому
        }

        private void OnServiceProgress(int current, int total, string fileName)
        {
            OnUi(delegate
            {
                _progress.Maximum = total; // реальное число файлов известно только сервису
                if (current - 1 <= total)
                    _progress.Value = current - 1;
                SyncTaskbar();
                SetStatus(string.Format(Loc.T("excel.status.fileProgress"), current, total, fileName), Theme.TextMuted);
            });
        }

        private void OnServiceFileDone(FileResult fr)
        {
            OnUi(delegate
            {
                ListViewItem row;
                if (fr.FullPath == null || !_rowByPath.TryGetValue(Path.GetFullPath(fr.FullPath), out row))
                    return; // результат по файлу вне списка (не должно случаться)
                var results = (List<FileResult>)row.Tag;
                results.Add(fr); // в режиме «все листы» на файл несколько результатов
                UpdateRowDisplay(row, results);
                row.EnsureVisible();
                // Прогресс ведёт OnServiceProgress (по файлам).
            });
        }

        /// <summary>Свод результатов по файлу в его строку: статус, примечания, цвет.</summary>
        private void UpdateRowDisplay(ListViewItem row, List<FileResult> results)
        {
            int ok = 0, fail = 0;
            var notes = new List<string>();
            foreach (FileResult fr in results)
            {
                if (fr.Ok) ok++; else fail++;
                if (!string.IsNullOrEmpty(fr.Note) && !notes.Contains(fr.Note))
                    notes.Add(fr.Note);
            }
            string status;
            Color color;
            if (ok == 0)
            {
                status = Loc.T("excel.row.skipped");
                color = Theme.ErrRed;
            }
            else if (fail == 0)
            {
                status = ok == 1 ? Loc.T("excel.row.moved") : string.Format(Loc.T("excel.row.sheets"), ok);
                color = notes.Count > 0 ? Theme.WarnOrange : Theme.OkGreen;
            }
            else
            {
                status = string.Format(Loc.T("excel.row.sheetsPartial"), ok, ok + fail);
                color = Theme.WarnOrange;
            }
            row.SubItems[1].Text = status;
            row.SubItems[2].Text = string.Join("; ", notes.ToArray());
            row.ForeColor = color;
        }

        private void OnMergeFinished(MergeResult result, Exception error)
        {
            _running = false;
            _worker = null;
            SetRunning(false);
            UpdateReadiness();

            if (_closeRequested)
            {
                Close();
                return;
            }

            // Пользователь работает в другом окне — мигнуть кнопкой на панели задач.
            if (Form.ActiveForm == null)
                WindowFlasher.FlashUntilForeground(this);

            // Отменённый прогон не менял свод — прежний результат остаётся актуальным.
            if (result != null && !result.Cancelled)
                _lastResult = result;
            UpdateRetryButton();

            SaveReport(result);

            if (error != null)
            {
                _taskbar.SetState(Handle, TaskbarProgressState.Error);
                SetStatus(Loc.T("excel.status.failed"), Theme.ErrRed);
                Dialogs.Error(this, AppTitle, Loc.T("excel.err.mergeFailed.title"), error.Message);
                _taskbar.SetState(Handle, TaskbarProgressState.None);
                return;
            }

            if (result.Cancelled)
            {
                _taskbar.SetState(Handle, TaskbarProgressState.None);
                SetStatus(Loc.T("excel.status.cancelled"), Theme.WarnOrange);
                return;
            }

            if (_isFreshRun)
                UsageStats.RecordExcelDigest(); // успешный новый свод (дослияние не считаем)

            _progress.Value = _progress.Maximum;
            string text;
            if (result.SkipCount > 0)
                text = string.Format(Loc.T("excel.status.doneWithSkips"), result.OkCount, result.SkipCount);
            else
                text = string.Format(Loc.T("excel.status.doneClean"), result.OkCount);
            if (result.TocError != null)
                text += string.Format(Loc.T("excel.status.tocWarn"), result.TocError);
            bool clean = result.SkipCount == 0 && result.TocError == null;
            SetStatus(text, clean ? Theme.OkGreen : Theme.WarnOrange);
            _lnkOpenFile.Visible = true;
            _lnkOpenFolder.Visible = true;
            _lnkNote.Visible = true;
        }

        /// <summary>Сопроводительная записка .docx рядом со сводом — в фоновом STA-потоке.</summary>
        private void OnNoteClick()
        {
            if (_running || _noteBusy || _lastResult == null)
                return;
            MergeResult result = _lastResult;
            string folder = _lastInputFolder;
            MergeOptions options = _lastOptions;
            DateTime startedAt = _lastStartedAt;
            string notePath = Path.Combine(
                Path.GetDirectoryName(result.OutputPath),
                Path.GetFileNameWithoutExtension(result.OutputPath) + Loc.T("excel.noteFileSuffix"));

            _noteBusy = true;
            _lnkNote.Enabled = false;
            SetStatus(Loc.T("excel.status.noteBusy"), Theme.TextMuted);

            var thread = new Thread(delegate()
            {
                Exception error = null;
                try
                {
                    WordNoteWriter.Write(NoteText.Build(result, folder, options, startedAt), notePath);
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                OnUi(delegate
                {
                    _noteBusy = false;
                    _lnkNote.Enabled = true;
                    if (error != null)
                    {
                        if (!_running) // статус идущего слияния не перебиваем
                            SetStatus(Loc.T("excel.status.noteFailed"), Theme.ErrRed);
                        Dialogs.Error(this, AppTitle, Loc.T("excel.err.noteFailed.title"), error.Message);
                    }
                    else
                    {
                        if (!_running)
                            SetStatus(Loc.T("excel.status.noteSaved"), Theme.OkGreen);
                        OpenPath(notePath, false);
                    }
                });
            });
            thread.SetApartmentState(ApartmentState.STA); // требование Word COM
            thread.IsBackground = true;
            thread.Start();
        }

        /// <summary>«Повторить пропущенные» видима, когда есть что и куда дослить.</summary>
        private void UpdateRetryButton()
        {
            _btnRetry.Visible = !_running && _lastResult != null &&
                _lastResult.SkipCount > 0 && File.Exists(_lastResult.OutputPath);
        }

        /// <summary>Отчёт о завершённом слиянии — в историю (три последних, %APPDATA%).</summary>
        private void SaveReport(MergeResult result)
        {
            if (result == null)
                return; // ошибка до начала обработки — отчитываться не о чем
            try
            {
                string content = ReportWriter.BuildReport(result, _lastInputFolder, _lastOptions, _lastStartedAt);
                _lastReportPath = ReportWriter.SaveWithRotation(AppPaths.ReportsDir, content, DateTime.Now, 3);
                _lnkOpenReport.Visible = true;
            }
            catch { } // недоступный профиль не должен ломать результат слияния
        }

        /// <summary>Прогресс на кнопке панели задач зеркалит ProgressBar.</summary>
        private void SyncTaskbar()
        {
            _taskbar.SetState(Handle, TaskbarProgressState.Normal);
            _taskbar.SetValue(Handle, _progress.Value, _progress.Maximum);
        }

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);
            // Пользователь вернулся в окно и увидел результат — индикатор больше не нужен.
            if (!_running && _lastOutputPath != null)
                _taskbar.SetState(Handle, TaskbarProgressState.None);
        }

        private void OnCancelClick(object sender, EventArgs e)
        {
            if (!_running)
                return;
            if (_service != null)
                _service.Cancel();
            _btnCancel.Enabled = false;
            SetStatus(Loc.T("excel.status.cancelling"), Theme.WarnOrange);
        }

        private void SetStatus(string text, Color color)
        {
            _lblStatus.Text = text;
            _lblStatus.ForeColor = color;
        }

        private void SetRunning(bool running)
        {
            _txtInput.Enabled = !running;
            _btnBrowseInput.Enabled = !running;
            _txtName.Enabled = !running;
            _txtOutDir.Enabled = !running;
            _btnBrowseOut.Enabled = !running;
            _chkToc.Enabled = !running;
            _chkValues.Enabled = !running;
            _cmbFormat.Enabled = !running;
            _cmbScope.Enabled = !running;
            _btnMerge.Enabled = !running;
            _btnCancel.Enabled = running;
            UpdateListButtons(); // кнопки порядка/выбора блокируются во время прогона
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_noteBusy)
            {
                SetStatus(Loc.T("excel.status.waitNote"), Theme.WarnOrange);
                e.Cancel = true; // генерация занимает секунды; иначе остался бы зомби-WINWORD
                return;
            }
            if (_running)
            {
                if (!_closeRequested &&
                    Dialogs.ConfirmWarning(this, AppTitle, Loc.T("excel.confirm.closeBusy.title"),
                        Loc.T("excel.confirm.closeBusy.body")))
                {
                    _closeRequested = true;
                    if (_service != null)
                        _service.Cancel();
                    SetStatus(Loc.T("excel.status.finishing"), Theme.WarnOrange);
                }
                e.Cancel = true; // закроемся сами после корректного завершения Excel
                return;
            }
            SavePathsToSettings();
            base.OnFormClosing(e);
        }

        private void SavePathsToSettings()
        {
            string input = _txtInput.Text.Trim();
            string outDir = _txtOutDir.Text.Trim();
            if (Directory.Exists(input))
                _settings.LastInputFolder = input;
            if (Directory.Exists(outDir))
                _settings.LastOutputFolder = outDir;
            _settings.AddToc = _chkToc.Checked;
            _settings.AllSheets = _cmbScope.SelectedIndex == 1;
            _settings.OutputExtension = (string)_cmbFormat.SelectedItem;
            _settings.Save();
        }

        private void OpenPath(string filePath, bool selectInFolder)
        {
            try
            {
                if (selectInFolder)
                    Process.Start("explorer.exe", "/select,\"" + filePath + "\"");
                else
                    Process.Start(filePath);
            }
            catch (Exception ex)
            {
                Dialogs.Error(this, AppTitle, Loc.T("common.err.openFailed"), ex.Message);
            }
        }
    }
}
