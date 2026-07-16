using System;
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
        private const string AppTitle = "Свод листов Excel";
        private const int MenuHeight = 28;

        private TextBox _txtInput;
        private Button _btnBrowseInput;
        private Label _lblFound;
        private TextBox _txtName;
        private TextBox _txtOutDir;
        private Button _btnBrowseOut;
        private CheckBox _chkToc;
        private CheckBox _chkValues;
        private ComboBox _cmbFormat;
        private ToolTip _tips;
        private Button _btnMerge;
        private Button _btnCancel;
        private ProgressBar _progress;
        private Label _lblStatus;
        private ListView _list;
        private LinkLabel _lnkOpenFile;
        private LinkLabel _lnkOpenFolder;
        private LinkLabel _lnkOpenReport;
        private Button _btnRetry;

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
        private ListViewNaturalSorter _sorter;
        private Label _lblListHint;
        private int _foundCount;
        private bool _running;        // истина от нажатия «Объединить» до OnMergeFinished (только UI-поток)
        private bool _closeRequested; // пользователь закрыл окно во время объединения

        public MainForm()
        {
            BuildUi();

            _settings = UserSettings.Load();
            if (!string.IsNullOrEmpty(_settings.LastInputFolder) && Directory.Exists(_settings.LastInputFolder))
                _txtInput.Text = _settings.LastInputFolder;
            if (!string.IsNullOrEmpty(_settings.LastOutputFolder) && Directory.Exists(_settings.LastOutputFolder))
                _txtOutDir.Text = _settings.LastOutputFolder;
            _txtName.Text = "Свод_" + DateTime.Now.ToString("yyyy-MM-dd");
            _chkToc.Checked = _settings.AddToc;
            // «Заменить формулы значениями» всегда стартует выключенным: режим
            // меняет содержимое свода и включается осознанно на каждый запуск.
            int formatIndex = Array.IndexOf(OutputFormats.Extensions, _settings.OutputExtension);
            _cmbFormat.SelectedIndex = formatIndex >= 0 ? formatIndex : 0;
            RefreshFileCount();
        }

        // ---------- построение интерфейса ----------

        private void BuildUi()
        {
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            Text = AppTitle + " " + version.ToString(2);
            Icon icon = LoadAppIcon();
            if (icon != null)
                Icon = icon;
            Font = new Font("Segoe UI", 9.75f);
            BackColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(780, 755);
            MinimumSize = new Size(700, 695);
            AllowDrop = true;
            DragEnter += OnDragEnter;
            DragDrop += OnDragDrop;
            _tips = new ToolTip();

            int right = ClientSize.Width - 20;
            var stretch = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            BuildMenu();
            Ui.AccentBar(this, MenuHeight);

            // Шапка
            Ui.Label(this, "Свод листов Excel", 20, 47,
                new Font("Segoe UI", 15f, FontStyle.Bold), Color.FromArgb(40, 40, 40));
            Ui.Label(this, "Первый видимый лист каждого файла папки — в один итоговый файл .xlsx",
                22, 79, Font, Theme.TextMuted);

            // Шаг 1: исходная папка
            AddSectionLabel("ПАПКА С ИСХОДНЫМИ ФАЙЛАМИ", 115);
            _txtInput = AddTextBox(20, 135, right - 20 - 110);
            _txtInput.TextChanged += delegate { RefreshFileCount(); };
            _btnBrowseInput = AddButton("Обзор…", false, right - 100, 134, 100, 29);
            _btnBrowseInput.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnBrowseInput.Click += OnBrowseInput;
            _lblFound = Ui.Label(this, "", 20, 167, Font, Theme.TextMuted);

            // Шаг 2: итоговый файл
            AddSectionLabel("ИТОГОВЫЙ ФАЙЛ", 199);
            Ui.Label(this, "Имя:", 20, 222, Font, Theme.TextPrimary);
            _txtName = AddTextBox(75, 219, 300);
            _txtName.Anchor = AnchorStyles.Top | AnchorStyles.Left; // фиксированная ширина: справа выбор формата
            _txtName.TextChanged += delegate { UpdateReadiness(); };
            _cmbFormat = new ComboBox();
            _cmbFormat.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbFormat.Items.AddRange(OutputFormats.Extensions);
            _cmbFormat.SetBounds(383, 219, 85, 27);
            Controls.Add(_cmbFormat);
            _tips.SetToolTip(_cmbFormat, "Формат итогового файла; .xls — старый формат Excel 97–2003");

            Ui.Label(this, "Папка:", 20, 258, Font, Theme.TextPrimary);
            _txtOutDir = AddTextBox(75, 255, right - 75 - 110);
            _btnBrowseOut = AddButton("Обзор…", false, right - 100, 254, 100, 29);
            _btnBrowseOut.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnBrowseOut.Click += OnBrowseOutput;

            // Параметры
            AddSectionLabel("ПАРАМЕТРЫ", 289);
            _chkToc = AddCheckBox("Добавить лист «Содержание» с оглавлением и ссылками", 20, 309,
                "Первым листом свода будет оглавление: гиперссылки на листы и статусы всех файлов");
            _chkValues = AddCheckBox("Заменить формулы значениями", 20, 335,
                "Свод не будет зависеть от исходных файлов: вместо формул — вычисленные значения");

            // Действия: основная слева, отмена — у правого края, подальше от основной
            _btnMerge = AddButton("Объединить", true, 20, 371, 170, 40);
            _btnMerge.Click += OnMergeClick;
            AcceptButton = _btnMerge;
            _tips.SetToolTip(_btnMerge, "Собрать свод из файлов выбранной папки (Enter)");

            _btnCancel = AddButton("Отменить", false, right - 130, 371, 130, 40);
            _btnCancel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnCancel.Enabled = false;
            _btnCancel.Click += OnCancelClick;
            CancelButton = _btnCancel;
            _tips.SetToolTip(_btnCancel, "Остановить после текущего файла (Esc)");

            _progress = new ProgressBar();
            _progress.SetBounds(20, 429, right - 20, 8);
            _progress.Anchor = stretch;
            _progress.Visible = false; // пустая полоса в простое только путает
            Controls.Add(_progress);

            _lblStatus = Ui.Label(this, "Выберите папку с исходными файлами.", 20, 447, Font, Theme.TextMuted);

            // Журнал: построчный результат по каждому файлу текущего прогона
            AddSectionLabel("ЖУРНАЛ ОБРАБОТКИ", 472);
            _list = new ListView();
            _list.SetBounds(20, 492, right - 20, ClientSize.Height - 492 - 44);
            _list.Anchor = stretch | AnchorStyles.Bottom;
            _list.View = View.Details;
            _list.FullRowSelect = true;
            _list.HeaderStyle = ColumnHeaderStyle.Clickable; // клик по заголовку сортирует
            _list.BorderStyle = BorderStyle.FixedSingle;
            _sorter = new ListViewNaturalSorter(_list);
            _list.ColumnClick += delegate(object sender, ColumnClickEventArgs e)
            {
                if (!_running) // во время прогона строки идут в порядке обработки
                    _sorter.SortBy(e.Column);
            };
            _list.Columns.Add("Файл", 270);
            _list.Columns.Add("Лист", 170);
            _list.Columns.Add("Результат", 110);
            _list.Columns.Add("Примечание", 160);
            EnableDoubleBuffer(_list);
            // Подсказка пустого состояния — внутри списка, пока нет ни одной строки.
            _lblListHint = new Label();
            _lblListHint.Text = "Здесь появится результат по каждому файлу:\nимя листа в своде, статус и причина пропуска.";
            _lblListHint.AutoSize = true;
            _lblListHint.ForeColor = Theme.TextMuted;
            _lblListHint.BackColor = Color.White;
            _lblListHint.Location = new Point(14, 34);
            _list.Controls.Add(_lblListHint);
            var copyMenu = new ContextMenuStrip();
            var copyItem = new ToolStripMenuItem("Копировать");
            copyItem.ShortcutKeyDisplayString = "Ctrl+C";
            copyItem.Click += delegate { CopySelectedRows(); };
            copyMenu.Items.Add(copyItem);
            copyMenu.Opening += delegate { copyItem.Enabled = _list.SelectedItems.Count > 0; };
            _list.ContextMenuStrip = copyMenu;
            _list.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.Control && e.KeyCode == Keys.C)
                {
                    CopySelectedRows();
                    e.Handled = true;
                }
            };
            Controls.Add(_list);

            _btnRetry = AddButton("Повторить пропущенные", false, right - 200, ClientSize.Height - 40, 200, 30);
            _btnRetry.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _btnRetry.Visible = false;
            _btnRetry.Click += OnRetryClick;
            _tips.SetToolTip(_btnRetry, "Дослить исправленные файлы в существующий свод без полного пересбора");

            _lnkOpenFile = AddBottomLink("Открыть файл", 20);
            _lnkOpenFile.LinkClicked += delegate { OpenPath(_lastOutputPath, false); };
            _lnkOpenFolder = AddBottomLink("Открыть папку", 145);
            _lnkOpenFolder.LinkClicked += delegate { OpenPath(_lastOutputPath, true); };
            _lnkOpenReport = AddBottomLink("Открыть отчёт", 280);
            _lnkOpenReport.LinkClicked += delegate { OpenPath(_lastReportPath, false); };
            _tips.SetToolTip(_lnkOpenReport, "Отчёт о слиянии; в истории хранятся три последних");

            _tips.SetToolTip(_txtInput, "Папку можно перетащить мышью в окно программы");
            _tips.SetToolTip(_txtName, "Расширение .xlsx добавится автоматически");
            _tips.SetToolTip(_txtOutDir, "Пусто — итоговый файл сохранится в папку с исходными");

            Resize += delegate { AdjustNoteColumn(); };
            AdjustNoteColumn();
            UpdateReadiness();
        }

        private void BuildMenu()
        {
            var menu = new MenuStrip();
            menu.AutoSize = false;
            menu.Height = MenuHeight;
            menu.Dock = DockStyle.Top;
            menu.BackColor = Color.White;
            menu.Padding = new Padding(12, 4, 0, 0);

            var help = new ToolStripMenuItem("Справка");

            // По гайдлайнам Windows многоточие — только у команд, требующих
            // дополнительного ввода; просмотр справки и «О программе» — без него.
            var howTo = new ToolStripMenuItem("Как пользоваться");
            howTo.ShortcutKeys = Keys.F1;
            howTo.Click += delegate { ShowHelp(); };

            var reports = new ToolStripMenuItem("Папка отчётов");
            reports.Click += delegate { OpenReportsFolder(); };

            var about = new ToolStripMenuItem("О программе");
            about.Click += delegate
            {
                using (var form = new AboutForm())
                    form.ShowDialog(this);
            };

            help.DropDownItems.Add(howTo);
            help.DropDownItems.Add(new ToolStripSeparator());
            help.DropDownItems.Add(reports);
            help.DropDownItems.Add(new ToolStripSeparator());
            help.DropDownItems.Add(about);
            menu.Items.Add(help);

            MainMenuStrip = menu;
            Controls.Add(menu);
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, AppTitle, "Как пользоваться",
                "1. Укажите папку с исходными файлами — «Обзор…» или перетащите папку в окно.\n" +
                "2. Задайте имя свода; папку сохранения можно сменить (пустая — папка с исходными).\n" +
                "3. Нажмите «Объединить»: из каждого файла переносится первый видимый лист " +
                "со всем оформлением, формулами и диаграммами.\n\n" +
                "Параметры:\n" +
                "• лист «Содержание» — оглавление свода с гиперссылками и статусами файлов,\n" +
                "• «Заменить формулы значениями» — свод не зависит от исходных файлов.\n\n" +
                "Битые и запароленные файлы пропускаются, причина видна в списке и в отчёте.\n" +
                "Отчёты (три последних): Справка → «Папка отчётов».");
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
                Dialogs.Error(this, AppTitle, "Не удалось открыть папку отчётов", ex.Message);
            }
        }

        private static Icon LoadAppIcon()
        {
            try
            {
                return Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch
            {
                return null; // без иконки, со стандартной системной
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

        /// <summary>Выбранные строки журнала — в буфер обмена (формат как в отчёте).</summary>
        private void CopySelectedRows()
        {
            if (_list.SelectedItems.Count == 0)
                return;
            var sb = new StringBuilder();
            foreach (ListViewItem item in _list.SelectedItems)
            {
                var fr = item.Tag as FileResult;
                if (fr != null)
                    sb.AppendLine(ReportWriter.FormatFileLine(fr));
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

        // Доли колонок журнала: Файл / Лист / Результат / Примечание.
        private static readonly float[] ColumnWeights = { 0.36f, 0.22f, 0.14f, 0.28f };

        /// <summary>Колонки журнала делят ширину списка пропорционально.</summary>
        private void AdjustNoteColumn()
        {
            int width = _list.ClientSize.Width - 4;
            if (width < 300)
                return;
            for (int i = 0; i < _list.Columns.Count; i++)
                _list.Columns[i].Width = (int)(width * ColumnWeights[i]);
        }

        // ---------- выбор папок, drag&drop, живая валидация ----------

        private void OnBrowseInput(object sender, EventArgs e)
        {
            string path = FolderPicker.Show(this, "Папка с исходными файлами Excel", _txtInput.Text.Trim());
            if (path != null)
                SetInputFolder(path);
        }

        private void OnBrowseOutput(object sender, EventArgs e)
        {
            string initial = _txtOutDir.Text.Trim();
            if (initial.Length == 0)
                initial = _txtInput.Text.Trim();
            string path = FolderPicker.Show(this, "Папка для сохранения итогового файла", initial);
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

        private void RefreshFileCount()
        {
            string folder = _txtInput.Text.Trim();
            _foundCount = 0;
            if (folder.Length == 0)
            {
                SetFoundLabel("Укажите папку или перетащите её в окно.", Theme.TextMuted);
            }
            else if (!Directory.Exists(folder))
            {
                SetFoundLabel("Папка не найдена.", Theme.ErrRed);
            }
            else
            {
                try
                {
                    _foundCount = MergeService.FindSourceFiles(folder, null).Count;
                    if (_foundCount > 0)
                        SetFoundLabel("✓ Найдено файлов Excel: " + _foundCount, Theme.OkGreen);
                    else
                        SetFoundLabel("Файлы Excel (.xlsx, .xls, .xlsm, .xlsb) не найдены.", Theme.ErrRed);
                }
                catch (Exception ex)
                {
                    SetFoundLabel("Не удалось прочитать папку: " + ex.Message, Theme.ErrRed);
                }
            }
            UpdateReadiness();
        }

        private void SetFoundLabel(string text, Color color)
        {
            _lblFound.Text = text;
            _lblFound.ForeColor = color;
        }

        /// <summary>Кнопка «Объединить» активна только при валидных входных данных.</summary>
        private void UpdateReadiness()
        {
            if (_running)
                return;
            _btnMerge.Enabled = _foundCount > 0 && _txtName.Text.Trim().Length > 0;
        }

        // ---------- запуск объединения ----------

        private void OnMergeClick(object sender, EventArgs e)
        {
            string folder = _txtInput.Text.Trim();
            if (!Directory.Exists(folder))
            {
                Dialogs.Error(this, AppTitle, "Папка с исходными файлами не найдена",
                    "Проверьте путь: " + folder);
                return;
            }

            string name = OutputFormats.StripKnownExtension(_txtName.Text.Trim());
            if (name.Length == 0)
            {
                Dialogs.Error(this, AppTitle, "Укажите имя итогового файла",
                    "Поле «Имя» не заполнено.");
                return;
            }
            if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                Dialogs.Error(this, AppTitle, "Недопустимое имя файла",
                    "Имя не должно содержать символы  \\ / : * ? \" < > |");
                return;
            }

            string outDir = _txtOutDir.Text.Trim();
            if (outDir.Length == 0)
                outDir = folder;
            if (!Directory.Exists(outDir))
            {
                if (!Dialogs.ConfirmWarning(this, AppTitle, "Папка сохранения не существует",
                        "Создать папку?\n" + outDir))
                    return;
                try
                {
                    Directory.CreateDirectory(outDir);
                }
                catch (Exception ex)
                {
                    Dialogs.Error(this, AppTitle, "Не удалось создать папку", ex.Message);
                    return;
                }
            }

            string outputPath = Path.Combine(outDir, name + (string)_cmbFormat.SelectedItem);

            if (MergeService.FindSourceFiles(folder, outputPath).Count == 0)
            {
                Dialogs.Error(this, AppTitle, "В папке нет файлов Excel",
                    "Поддерживаются файлы .xlsx, .xls, .xlsm, .xlsb.");
                return;
            }

            if (File.Exists(outputPath) &&
                !Dialogs.ConfirmWarning(this, AppTitle, "Файл уже существует",
                    "Файл «" + Path.GetFileName(outputPath) + "» уже есть в папке сохранения.\nПерезаписать его?"))
                return;

            // Занятый итоговый файл должен остановить работу сейчас,
            // а не после обработки всех исходных файлов.
            string lockError = MergeService.CheckOutputWritable(outputPath);
            if (lockError != null)
            {
                Dialogs.Error(this, AppTitle, "Итоговый файл недоступен для записи", lockError);
                return;
            }

            _settings.LastInputFolder = folder;
            _settings.LastOutputFolder = outDir;
            _settings.AddToc = _chkToc.Checked;
            _settings.OutputExtension = (string)_cmbFormat.SelectedItem;
            _settings.Save();

            var options = new MergeOptions();
            options.AddToc = _chkToc.Checked;
            options.ValuesOnly = _chkValues.Checked;
            StartMerge(folder, outputPath, options);
        }

        private void StartMerge(string folder, string outputPath, MergeOptions options)
        {
            PrepareRun(folder, outputPath, options);
            StartWorker(delegate { return _service.Merge(folder, outputPath, options); });
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
                Dialogs.Error(this, AppTitle, "Итоговый файл недоступен для записи", lockError);
                return;
            }
            MergeOptions options = _lastOptions;
            PrepareRun(_lastInputFolder, outputPath, options);
            StartWorker(delegate { return _service.RetrySkipped(outputPath, options, previous); });
        }

        /// <summary>Общий сброс интерфейса перед обычным и повторным слиянием.</summary>
        private void PrepareRun(string folder, string outputPath, MergeOptions options)
        {
            _running = true;
            SetRunning(true);
            _sorter.Reset();
            _list.Items.Clear();
            _lblListHint.Visible = false;
            _lnkOpenFile.Visible = false;
            _lnkOpenFolder.Visible = false;
            _lnkOpenReport.Visible = false;
            _btnRetry.Visible = false;
            _progress.Value = 0;
            _progress.Visible = true;
            SetStatus("Запуск Excel…", Theme.TextMuted);
            _lastOutputPath = outputPath;
            _lastInputFolder = folder;
            _lastOptions = options;
            _lastStartedAt = DateTime.Now;
            _taskbar.SetState(Handle, TaskbarProgressState.Indeterminate);

            _service = new MergeService();
            _service.Progress += OnServiceProgress;
            _service.FileDone += OnServiceFileDone;
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
                SetStatus("Файл " + current + " из " + total + ": " + fileName, Theme.TextMuted);
            });
        }

        private void OnServiceFileDone(FileResult fr)
        {
            OnUi(delegate
            {
                var item = new ListViewItem(fr.FileName);
                item.Tag = fr; // для копирования строки в буфер
                item.SubItems.Add(fr.SheetName ?? "—");
                item.SubItems.Add(fr.Ok ? "✓ перенесён" : "✗ пропущен");
                item.SubItems.Add(fr.Note ?? "");
                item.ForeColor = !fr.Ok ? Theme.ErrRed : (fr.Note != null ? Theme.WarnOrange : Theme.OkGreen);
                _list.Items.Add(item);
                item.EnsureVisible();
                if (_progress.Value < _progress.Maximum)
                    _progress.Value++;
                SyncTaskbar();
            });
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
                SetStatus("Объединение не выполнено.", Theme.ErrRed);
                Dialogs.Error(this, AppTitle, "Объединение не выполнено", error.Message);
                _taskbar.SetState(Handle, TaskbarProgressState.None);
                return;
            }

            if (result.Cancelled)
            {
                _taskbar.SetState(Handle, TaskbarProgressState.None);
                SetStatus("Отменено — итоговый файл не создан.", Theme.WarnOrange);
                return;
            }

            _progress.Value = _progress.Maximum;
            string text;
            if (result.SkipCount > 0)
                text = "Готово: перенесено " + result.OkCount + ", пропущено " + result.SkipCount +
                    " — причины в списке.";
            else
                text = "✓ Готово: перенесено листов — " + result.OkCount + ".";
            if (result.TocError != null)
                text += " Внимание: " + result.TocError + ".";
            bool clean = result.SkipCount == 0 && result.TocError == null;
            SetStatus(text, clean ? Theme.OkGreen : Theme.WarnOrange);
            _lnkOpenFile.Visible = true;
            _lnkOpenFolder.Visible = true;
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
            SetStatus("Отмена после текущего файла…", Theme.WarnOrange);
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
            _btnMerge.Enabled = !running;
            _btnCancel.Enabled = running;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_running)
            {
                if (!_closeRequested &&
                    Dialogs.ConfirmWarning(this, AppTitle, "Идёт объединение",
                        "Прервать объединение и закрыть программу?"))
                {
                    _closeRequested = true;
                    if (_service != null)
                        _service.Cancel();
                    SetStatus("Завершение…", Theme.WarnOrange);
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
                Dialogs.Error(this, AppTitle, "Не удалось открыть", ex.Message);
            }
        }
    }
}
