using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger
{
    public class MainForm : Form
    {
        private const string AppTitle = "Свод листов Excel";

        private TextBox _txtInput;
        private Button _btnBrowseInput;
        private Label _lblFound;
        private TextBox _txtName;
        private TextBox _txtOutDir;
        private Button _btnBrowseOut;
        private Button _btnMerge;
        private Button _btnCancel;
        private ProgressBar _progress;
        private Label _lblStatus;
        private ListView _list;
        private LinkLabel _lnkOpenFile;
        private LinkLabel _lnkOpenFolder;

        private MergeService _service;
        private Thread _worker;
        private UserSettings _settings;
        private string _lastOutputPath;
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
            RefreshFileCount();
        }

        // ---------- построение интерфейса ----------

        private void BuildUi()
        {
            Text = AppTitle;
            Icon icon = LoadAppIcon();
            if (icon != null)
                Icon = icon;
            Font = new Font("Segoe UI", 9.75f);
            BackColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(780, 640);
            MinimumSize = new Size(700, 580);
            AllowDrop = true;
            DragEnter += OnDragEnter;
            DragDrop += OnDragDrop;

            int right = ClientSize.Width - 20;
            var stretch = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Шапка
            var topBar = new Panel();
            topBar.SetBounds(0, 0, ClientSize.Width, 4);
            topBar.Anchor = stretch;
            topBar.BackColor = Theme.Accent;
            Controls.Add(topBar);

            AddLabel("Свод листов Excel", 20, 16, new Font("Segoe UI", 15f, FontStyle.Bold), Color.FromArgb(40, 40, 40));
            AddLabel("Первый видимый лист каждого файла папки — в один итоговый файл .xlsx",
                22, 48, Font, Theme.TextMuted);

            // Шаг 1: исходная папка
            AddSectionLabel("ПАПКА С ИСХОДНЫМИ ФАЙЛАМИ", 84);
            _txtInput = AddTextBox(20, 104, right - 20 - 110);
            _txtInput.TextChanged += delegate { RefreshFileCount(); };
            _btnBrowseInput = AddButton("Обзор…", false, right - 100, 103, 100, 29);
            _btnBrowseInput.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnBrowseInput.Click += OnBrowseInput;
            _lblFound = AddLabel("", 20, 136, Font, Theme.TextMuted);

            // Шаг 2: итоговый файл
            AddSectionLabel("ИТОГОВЫЙ ФАЙЛ", 168);
            AddLabel("Имя:", 20, 191, Font, Theme.TextPrimary);
            _txtName = AddTextBox(75, 188, 300);
            _txtName.Anchor = AnchorStyles.Top | AnchorStyles.Left; // фиксированная ширина: справа подпись .xlsx
            _txtName.TextChanged += delegate { UpdateReadiness(); };
            AddLabel(".xlsx", 380, 191, Font, Theme.TextMuted);

            AddLabel("Папка:", 20, 227, Font, Theme.TextPrimary);
            _txtOutDir = AddTextBox(75, 224, right - 75 - 110);
            _btnBrowseOut = AddButton("Обзор…", false, right - 100, 223, 100, 29);
            _btnBrowseOut.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnBrowseOut.Click += OnBrowseOutput;

            // Действия
            _btnMerge = AddButton("Объединить", true, 20, 270, 170, 40);
            _btnMerge.Click += OnMergeClick;
            AcceptButton = _btnMerge;

            _btnCancel = AddButton("Отменить", false, 200, 270, 130, 40);
            _btnCancel.Enabled = false;
            _btnCancel.Click += OnCancelClick;
            CancelButton = _btnCancel;

            _progress = new ProgressBar();
            _progress.SetBounds(20, 328, right - 20, 8);
            _progress.Anchor = stretch;
            Controls.Add(_progress);

            _lblStatus = AddLabel("Выберите папку с исходными файлами.", 20, 346, Font, Theme.TextMuted);

            // Журнал
            _list = new ListView();
            _list.SetBounds(20, 376, right - 20, ClientSize.Height - 376 - 44);
            _list.Anchor = stretch | AnchorStyles.Bottom;
            _list.View = View.Details;
            _list.FullRowSelect = true;
            _list.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            _list.BorderStyle = BorderStyle.FixedSingle;
            _list.Columns.Add("Файл", 270);
            _list.Columns.Add("Лист", 170);
            _list.Columns.Add("Результат", 110);
            _list.Columns.Add("Примечание", 160);
            EnableDoubleBuffer(_list);
            Controls.Add(_list);

            _lnkOpenFile = AddLink("Открыть файл", 20, ClientSize.Height - 32);
            _lnkOpenFile.LinkClicked += delegate { OpenPath(_lastOutputPath, false); };
            _lnkOpenFolder = AddLink("Открыть папку", 145, ClientSize.Height - 32);
            _lnkOpenFolder.LinkClicked += delegate { OpenPath(_lastOutputPath, true); };

            var tips = new ToolTip();
            tips.SetToolTip(_txtInput, "Папку можно перетащить мышью в окно программы");
            tips.SetToolTip(_txtName, "Расширение .xlsx добавится автоматически");
            tips.SetToolTip(_txtOutDir, "Пусто — итоговый файл сохранится в папку с исходными");

            Resize += delegate { AdjustNoteColumn(); };
            AdjustNoteColumn();
            UpdateReadiness();
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

        private Label AddLabel(string text, int x, int y, Font font, Color color)
        {
            var l = new Label();
            l.Text = text;
            l.Location = new Point(x, y);
            l.AutoSize = true;
            l.Font = font;
            l.ForeColor = color;
            l.BackColor = Color.Transparent;
            Controls.Add(l);
            return l;
        }

        private void AddSectionLabel(string text, int y)
        {
            AddLabel(text, 20, y, new Font("Segoe UI", 8f, FontStyle.Bold), Theme.TextMuted);
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

        private LinkLabel AddLink(string text, int x, int y)
        {
            var l = new LinkLabel();
            l.Text = text;
            l.Location = new Point(x, y);
            l.AutoSize = true;
            l.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            l.LinkColor = Theme.Accent;
            l.ActiveLinkColor = Theme.AccentPressed;
            l.Visible = false;
            Controls.Add(l);
            return l;
        }

        private static void EnableDoubleBuffer(ListView list)
        {
            // Убирает мерцание при добавлении строк; свойство защищённое — только через reflection.
            PropertyInfo p = typeof(ListView).GetProperty("DoubleBuffered",
                BindingFlags.Instance | BindingFlags.NonPublic);
            if (p != null)
                p.SetValue(list, true, null);
        }

        private void AdjustNoteColumn()
        {
            int used = _list.Columns[0].Width + _list.Columns[1].Width + _list.Columns[2].Width;
            int rest = _list.ClientSize.Width - used - 4;
            if (rest > 100)
                _list.Columns[3].Width = rest;
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

            string name = _txtName.Text.Trim();
            if (name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                name = name.Substring(0, name.Length - 5).TrimEnd();
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

            string outputPath = Path.Combine(outDir, name + ".xlsx");

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

            _settings.LastInputFolder = folder;
            _settings.LastOutputFolder = outDir;
            _settings.Save();

            StartMerge(folder, outputPath);
        }

        private void StartMerge(string folder, string outputPath)
        {
            _running = true;
            SetRunning(true);
            _list.Items.Clear();
            _lnkOpenFile.Visible = false;
            _lnkOpenFolder.Visible = false;
            _progress.Value = 0;
            SetStatus("Запуск Excel…", Theme.TextMuted);
            _lastOutputPath = outputPath;

            _service = new MergeService();
            _service.Progress += OnServiceProgress;
            _service.FileDone += OnServiceFileDone;

            _worker = new Thread(delegate() { RunMergeWorker(folder, outputPath); });
            _worker.SetApartmentState(ApartmentState.STA); // требование Excel COM
            _worker.IsBackground = true;
            _worker.Start();
        }

        /// <summary>Тело фонового потока: любые исключения доставляются в UI, поток не падает.</summary>
        private void RunMergeWorker(string folder, string outputPath)
        {
            MergeResult result = null;
            Exception error = null;
            try
            {
                result = _service.Merge(folder, outputPath);
            }
            catch (Exception ex)
            {
                error = ex;
            }
            OnUi(delegate { OnMergeFinished(result, error); });
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
                SetStatus("Файл " + current + " из " + total + ": " + fileName, Theme.TextMuted);
            });
        }

        private void OnServiceFileDone(FileResult fr)
        {
            OnUi(delegate
            {
                var item = new ListViewItem(fr.FileName);
                item.SubItems.Add(fr.SheetName ?? "—");
                item.SubItems.Add(fr.Ok ? "✓ перенесён" : "✗ пропущен");
                item.SubItems.Add(fr.Note ?? "");
                item.ForeColor = !fr.Ok ? Theme.ErrRed : (fr.Note != null ? Theme.WarnOrange : Theme.OkGreen);
                _list.Items.Add(item);
                item.EnsureVisible();
                if (_progress.Value < _progress.Maximum)
                    _progress.Value++;
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

            if (error != null)
            {
                SetStatus("Объединение не выполнено.", Theme.ErrRed);
                Dialogs.Error(this, AppTitle, "Объединение не выполнено", error.Message);
                return;
            }

            if (result.Cancelled)
            {
                SetStatus("Отменено — итоговый файл не создан.", Theme.WarnOrange);
                return;
            }

            _progress.Value = _progress.Maximum;
            if (result.SkipCount > 0)
                SetStatus("Готово: перенесено " + result.OkCount + ", пропущено " + result.SkipCount +
                    " — причины в списке.", Theme.WarnOrange);
            else
                SetStatus("✓ Готово: перенесено листов — " + result.OkCount + ".", Theme.OkGreen);
            _lnkOpenFile.Visible = true;
            _lnkOpenFolder.Visible = true;
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
