using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Инструмент «Прочие операции»: шесть действий над ОДНИМ документом — сжать, перевести в
    /// оттенки серого, восстановить повреждённый, сохранить страницы картинками, извлечь текст,
    /// изменить свойства. До 1.17.9 пять из них прятались в меню «Доп. действия» внутри
    /// «Разделения PDF» (там их никто не находил), а сжатия одного файла не было вовсе — только
    /// как довесок к объединению и разделению.
    ///
    /// Каждая операция пишет результат в НОВЫЙ файл: исходники приложение не меняет никогда.
    /// Слой «открыть документ и показать его страницы» — в общей базе
    /// <see cref="PdfSingleDocFormBase"/> (та же, что у «Разделения»).
    /// </summary>
    public class PdfOpsForm : PdfSingleDocFormBase
    {
        private static string Title { get { return Loc.T("hub.ops.name"); } }

        // Ритм правой панели. Он плотнее, чем у соседних инструментов, потому что элементов
        // девять: иначе последняя кнопка не помещается над нижним строем на минимальном размере.
        private const int BtnH = 28;      // высота кнопки действия
        private const int BtnGap = 5;     // между кнопками одной группы
        private const int GroupGap = 12;  // между группами
        private const int HeaderH = 16;   // заголовок группы
        private const int HeaderGap = 2;  // от заголовка до первой кнопки группы

        private Button _btnOpen, _btnCompress, _btnGray, _btnRepair, _btnImages, _btnText, _btnMeta;
        private ContextMenuStrip _dpiMenu; // не дочерний контрол — освобождаем сами

        public PdfOpsForm() : this(null) { }

        public PdfOpsForm(Action showHub) : base(showHub)
        {
            BuildUi();
            SyncControls();
        }

        protected override string ToolTitle { get { return Title; } }

        protected override string PickFileTitle { get { return Loc.T("ops.pickPdf"); } }

        private void BuildUi()
        {
            // Минимум по высоте больше, чем у соседних инструментов: в правой панели девять
            // элементов (открытие, три заголовка групп и шесть действий), и при 600 последняя
            // кнопка наезжала на нижний строй — это поймал живой тест раскладки.
            InitShell(Title, new Size(820, 660), new Size(720, 620), Theme.PdfRed);
            BuildHeaderWithHome(Title, Loc.T("ops.header.subtitle"),
                Theme.PdfRed, Theme.PdfRedDark, ShowHelp);

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;
            int panelW = 230;
            int gridBottom = ClientSize.Height - 152;

            _grid = new PdfPageGrid();
            _grid.AllowReorder = false; // операции не меняют порядок исходника
            _grid.AllowRotate = false;  // и не поворачивают его: результат — копия документа
            _grid.SetBounds(20, m + 84, right - 20 - panelW, gridBottom - (m + 84));
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            WireSingleDocGrid();
            WireGridMenu();
            Controls.Add(_grid);

            int px = right - panelW + 10;
            int pw = panelW - 10;
            int y = m + 84;

            _btnOpen = AddButton(Loc.T("common.btn.openPdf"), px, y, pw, 32, Loc.T("common.tip.openPdf"));
            _btnOpen.Click += delegate { PickAndOpenFile(); };
            y += 32 + GroupGap;

            // Шесть действий тремя группами: подряд идущие шесть равнозначных кнопок читаются
            // как свалка, а по заголовкам сразу видно, что делает документ, а что достаёт из него.
            y = AddGroup(Loc.T("ops.group.convert"), px, y, pw);
            _btnCompress = AddButton(Loc.T("ops.btn.compress"), px, y, pw, BtnH, Loc.T("ops.tip.compress"));
            _btnCompress.Click += delegate { CompressCopy(); };
            y += BtnH + BtnGap;
            _btnGray = AddButton(Loc.T("ops.btn.grayscale"), px, y, pw, BtnH, Loc.T("ops.tip.grayscale"));
            _btnGray.Click += delegate { ConvertCopy(PdfConvertMode.Grayscale, _sourcePath); };
            y += BtnH + BtnGap;
            _btnRepair = AddButton(Loc.T("ops.btn.repair"), px, y, pw, BtnH, Loc.T("ops.tip.repair"));
            _btnRepair.Click += delegate { RepairChosenFile(); };
            y += BtnH + GroupGap;

            y = AddGroup(Loc.T("ops.group.extract"), px, y, pw);
            _btnImages = AddButton(Loc.T("ops.btn.images"), px, y, pw, BtnH, Loc.T("ops.tip.images"));
            _btnImages.Click += delegate { ShowDpiMenu(); };
            y += BtnH + BtnGap;
            _btnText = AddButton(Loc.T("ops.btn.text"), px, y, pw, BtnH, Loc.T("ops.tip.text"));
            _btnText.Click += delegate { ExportText(); };
            y += BtnH + GroupGap;

            y = AddGroup(Loc.T("ops.group.edit"), px, y, pw);
            _btnMeta = AddButton(Loc.T("ops.btn.metadata"), px, y, pw, BtnH, Loc.T("ops.tip.metadata"));
            _btnMeta.Click += delegate { EditMetadata(); };

            BuildBottomStrip(right, Loc.T("common.status.openPdf"), 190);
            // Кнопки действия у этого окна нет — каждая операция своя, поэтому «Отмену» база
            // показывает на её обычном месте: правый нижний угол, как у остальных инструментов.
            RegisterCancelArea(new Rectangle(right - 190, ClientSize.Height - 58, 190, 38),
                AnchorStyles.Bottom | AnchorStyles.Right);
        }

        /// <summary>Заголовок группы кнопок. Возвращает Y первой кнопки под ним.</summary>
        private int AddGroup(string text, int x, int y, int width)
        {
            Label label = Ui.Label(this, text, x, y, Ui.Font(8.25f, FontStyle.Bold), Theme.TextMuted);
            label.SetBounds(x, y, width, HeaderH);
            label.AutoSize = false;
            label.AutoEllipsis = true; // длинный перевод не должен уезжать за край панели
            label.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            return y + HeaderH + HeaderGap;
        }

        private Button AddButton(string text, int x, int y, int w, int h, string tip)
        {
            var b = new RoundedButton(false);
            b.Text = text;
            b.SetBounds(x, y, w, h);
            b.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _tips.SetToolTip(b, tip);
            Controls.Add(b);
            return b;
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, Title, Loc.T("menu.howTo"), Loc.T("ops.help.body"));
        }

        /// <summary>Доступность кнопок: без открытого документа доступны только открытие и починка.</summary>
        protected override void SyncControls()
        {
            bool loaded = _sourcePath != null;
            bool ready = !Working && loaded;
            _grid.Locked = Working;
            _compress.Enabled = !Working;
            _btnOpen.Enabled = !Working;
            _btnCompress.Enabled = ready && Ghostscript.Available;
            _btnGray.Enabled = ready && Ghostscript.Available;
            // Починка выбирает файл сама: повреждённый документ в сетку не открывается, и
            // требовать «сначала откройте» значило бы выключить её ровно тогда, когда она нужна.
            _btnRepair.Enabled = !Working && Ghostscript.Available;
            _btnImages.Enabled = ready;
            _btnText.Enabled = ready;
            _btnMeta.Enabled = ready;
        }

        // ---------- преобразования документа ----------

        /// <summary>
        /// Сжать документ в новый файл. Уровень берётся из общего списка «Сжатие» внизу окна —
        /// того же, которым пользуются объединение и разделение (DRY). «Без сжатия» здесь не
        /// действие, а отсутствие действия, поэтому честно об этом говорим вместо тихого отказа.
        /// </summary>
        private void CompressCopy()
        {
            if (Working || _sourcePath == null)
                return;
            CompressionLevel level = _compress.Level;
            if (level == CompressionLevel.None)
            {
                Dialogs.Info(this, Title, Loc.T("ops.compress.pickLevel.title"),
                    Loc.T("ops.compress.pickLevel.body"));
                return;
            }
            string outPath = AskOutputPath(_sourcePath, Loc.T("ops.suffix.compressed"));
            if (outPath == null)
                return;
            string source = _sourcePath;
            BeginOperation(Loc.T("ops.status.compressing"), 0);
            BeginIndeterminate(Loc.T("ops.status.compressing")); // ход сжатия движок не сообщает
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                bool shrank = false;
                try
                {
                    File.Copy(source, outPath, true);
                    shrank = PdfCompression.Compress(outPath, level);
                }
                catch (Exception ex) { error = ex; }
                bool smaller = shrank;
                OnUi(delegate
                {
                    if (!FinishOperation(error, Loc.T("common.status.notDone"), Loc.T("ops.err.convertFailed")))
                        return;
                    // Копия создана в любом случае — она и есть то, о чём просили. Но если файл
                    // уже оптимизирован и меньше не стал, молчать нельзя: человек ждал сжатия.
                    SetStatus(smaller
                        ? SuccessStatus(string.Format(Loc.T("ops.status.compressed"),
                            PdfCompression.ImageDpi(level)))
                        : SuccessStatus(Loc.T("ops.status.notCompressed")), Theme.OkGreen);
                    Ui.OpenPath(outPath);
                });
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
            string outPath = AskOutputPath(source,
                Loc.T(mode == PdfConvertMode.Grayscale ? "ops.suffix.gray" : "ops.suffix.repaired"));
            if (outPath == null)
                return;
            BeginOperation(Loc.T("ops.status.converting"), 0);
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
                    if (!FinishOperation(error, Loc.T("common.status.notDone"), Loc.T("ops.err.convertFailed")))
                        return;
                    if (!applied)
                    {
                        SetStatus(Loc.T("ops.status.convertFailed"), Theme.ErrRed);
                        return;
                    }
                    SetStatus(SuccessStatus(Loc.T("ops.status.converted")), Theme.OkGreen);
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
                dialog.Title = Loc.T("ops.pick.repair");
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                ConvertCopy(PdfConvertMode.Repair, dialog.FileName);
            }
        }

        /// <summary>
        /// Спросить имя результата рядом с исходником и не дать записать поверх него. null —
        /// пользователь отказался или выбрал сам исходник (общая часть всех операций-копий).
        /// </summary>
        private string AskOutputPath(string source, string suffix)
        {
            string outPath;
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = Loc.T("common.pdfSaveFilter");
                dialog.FileName = Path.GetFileNameWithoutExtension(source) + suffix + ".pdf";
                dialog.InitialDirectory = Path.GetDirectoryName(source);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return null;
                outPath = dialog.FileName;
            }
            if (OutputFile.IsSameFile(outPath, source))
            {
                Dialogs.Error(this, Title, Loc.T("ops.err.sameFile"), Loc.T("ops.err.sameFile.body"));
                return null;
            }
            return outPath;
        }

        // ---------- извлечение ----------

        /// <summary>Меню разрешений под кнопкой картинок: сохранять ли в 96, 150, 300 или 600 dpi.</summary>
        private void ShowDpiMenu()
        {
            if (Working || _sourcePath == null)
                return;
            if (_dpiMenu == null)
            {
                _dpiMenu = new ContextMenuStrip();
                foreach (int dpi in PdfExportService.DpiChoices)
                {
                    int chosen = dpi; // копия для замыкания: иначе все пункты возьмут последнее значение
                    _dpiMenu.Items.Add(string.Format(Loc.T("ops.menu.dpi"), chosen), null,
                        delegate { ExportImages(chosen); });
                }
            }
            // Направление задаём явно: без него список раскрывается вверх и уезжает за край окна.
            _dpiMenu.Show(_btnImages, new Point(0, _btnImages.Height), ToolStripDropDownDirection.BelowRight);
        }

        /// <summary>Сохранить выбранные (или все) страницы картинками в выбранную папку.</summary>
        private void ExportImages(int dpi)
        {
            if (Working || _sourcePath == null)
                return;
            List<int> pages = PagesForExport();
            string dir = FolderPicker.Show(this, Loc.T("ops.pick.imagesDir"), Path.GetDirectoryName(_sourcePath));
            if (string.IsNullOrEmpty(dir))
                return;
            ImageExportFormat format = Dialogs.ConfirmWarning(this, Title, Loc.T("ops.ask.jpeg.title"),
                Loc.T("ops.ask.jpeg.body")) ? ImageExportFormat.Jpeg : ImageExportFormat.Png;

            string source = _sourcePath;
            BeginOperation(Loc.T("ops.status.exporting"), pages.Count, Loc.T("ops.status.exportingPage"));
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
                dialog.Filter = Loc.T("ops.txtFilter");
                dialog.FileName = Path.GetFileNameWithoutExtension(_sourcePath) + ".txt";
                dialog.InitialDirectory = Path.GetDirectoryName(_sourcePath);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                outPath = dialog.FileName;
            }
            string source = _sourcePath;
            BeginOperation(Loc.T("ops.status.extractingText"), _pageCount);
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

        private void OnExportFinished(Exception error, int count, string openTarget, bool asFolder)
        {
            if (!FinishOperation(error, Loc.T("common.status.notDone"), Loc.T("ops.err.exportFailed")))
                return;
            SetStatus(SuccessStatus(string.Format(Loc.T("ops.status.exported"), count)), Theme.OkGreen);
            Ui.OpenPath(openTarget, asFolder);
        }

        // ---------- свойства документа ----------

        /// <summary>
        /// Правка свойств документа. Результат пишется в НОВЫЙ файл: исходники приложение не
        /// меняет. Пустое поле очищает свойство — так из файла убирают имя автора перед отправкой.
        /// </summary>
        private void EditMetadata()
        {
            if (Working || _sourcePath == null)
                return;
            PdfMetadata edited = MetadataForm.Edit(this, PdfMetadataService.Read(_sourcePath),
                Path.GetFileName(_sourcePath));
            if (edited == null)
                return; // пользователь отказался
            string outPath = AskOutputPath(_sourcePath, Loc.T("ops.suffix.meta"));
            if (outPath == null)
                return;
            string source = _sourcePath;
            BeginOperation(Loc.T("ops.status.savingMeta"), 0);
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                try { PdfMetadataService.Write(source, outPath, edited); }
                catch (Exception ex) { error = ex; }
                OnUi(delegate
                {
                    if (!FinishOperation(error, Loc.T("common.status.notDone"), Loc.T("ops.err.metaFailed")))
                        return;
                    SetStatus(SuccessStatus(Loc.T("ops.status.metaSaved")), Theme.OkGreen);
                    Ui.OpenPath(outPath);
                });
            });
        }

        protected override void Dispose(bool disposing)
        {
            // Меню разрешений назначено не дочерним контролом, а показывается вручную:
            // само оно не освободится.
            if (disposing && _dpiMenu != null)
            {
                _dpiMenu.Dispose();
                _dpiMenu = null;
            }
            base.Dispose(disposing);
        }
    }
}
