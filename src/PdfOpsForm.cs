using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Инструмент «Прочие операции»: семь действий над ОДНИМ документом — сжать, перевести в
    /// оттенки серого, восстановить повреждённый, сохранить страницы картинками, извлечь текст,
    /// напечатать, изменить свойства. До 1.17.9 пять из них прятались в меню «Доп. действия» внутри
    /// «Разделения PDF» (там их никто не находил), а сжатия одного файла не было вовсе — только
    /// как довесок к объединению и разделению.
    ///
    /// Каждая операция пишет результат в НОВЫЙ файл: исходники приложение не меняет никогда.
    /// Слой «открыть документ и показать его страницы» — в общей базе
    /// <see cref="PdfSingleDocFormBase"/> (та же, что у «Разделения»).
    ///
    /// С 1.18.1 сетка здесь полноправная: страницы можно переставить, повернуть и убрать, и
    /// ЛЮБОЕ действие применяется к документу таким, каким он собран в сетке, — иначе окно
    /// показывало бы одно, а делало другое. Пока страницы не трогали, исходник просто
    /// копируется; стоит тронуть — документ сначала собирается заново (<see cref="AssembledDoc"/>).
    /// </summary>
    public class PdfOpsForm : PdfSingleDocFormBase
    {
        private static string Title { get { return Loc.T("hub.ops.name"); } }

        // Ритм правой панели. Он плотнее, чем у соседних инструментов, потому что элементов
        // десять: иначе последняя кнопка не помещается над нижним строем на минимальном размере.
        private const int BtnH = 28;      // высота кнопки действия
        private const int BtnGap = 5;     // между кнопками одной группы
        private const int GroupGap = 12;  // между группами
        private const int HeaderH = 16;   // заголовок группы
        private const int HeaderGap = 2;  // от заголовка до первой кнопки группы

        private Button _btnOpen, _btnAddImages, _btnSave, _btnCompress, _btnGray, _btnRepair,
            _btnImages, _btnText, _btnPrint, _btnMeta;
        private ContextMenuStrip _dpiMenu; // не дочерний контрол — освобождаем сами

        // Картинки живут в наборе одностраничными PDF-обёртками: их показывает та же сетка,
        // что и страницы документа, и они попадают во все действия окна без единой особой
        // ветки — включая поворот, порядок, печать и сжатие.
        private string _imageTempDir;
        private int _imageCounter; // номер обёртки: имена файлов не должны совпадать
        // Обёртка → настоящее имя картинки: от него предлагается имя результата (иначе человеку
        // предложили бы имя временного файла, которого он никогда не видел).
        private readonly Dictionary<string, string> _imageOrigins =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // Картинки, брошенные ВМЕСТЕ с PDF: пока идёт разбор документа, окно занято и добавить
        // их нельзя — ждут его конца, иначе пропали бы молча.
        private string[] _pendingImages;

        public PdfOpsForm() : this(null) { }

        public PdfOpsForm(Action showHub) : base(showHub)
        {
            BuildUi();
            SyncControls();
        }

        protected override string ToolTitle { get { return Title; } }

        protected override string PickFileTitle { get { return Loc.T("ops.pickPdf"); } }

        /// <summary>
        /// Idle-статус. Набор из одних картинок описывать «открытым документом» нечем — говорим
        /// о том, что есть: сколько страниц собрано и что с ними делать дальше.
        /// </summary>
        protected override string IdleStatusText()
        {
            if (_sourcePath == null && _order.Count > 0)
                return string.Format(Loc.T("ops.status.onlyImages"), _order.Count);
            return base.IdleStatusText();
        }

        private void BuildUi()
        {
            // Своя высота больше, чем у соседних инструментов: в правой панели одиннадцать
            // элементов (два способа набрать страницы, три заголовка групп и восемь действий).
            // Минимум по высоте дальше считается ПО ФАКТУ собранной панели — см. конец метода.
            InitShell(Title, new Size(820, 720), new Size(720, 620), Theme.PdfRed);
            BuildHeaderWithHome(Title, Loc.T("ops.header.subtitle"),
                Theme.PdfRed, Theme.PdfRedDark, ShowHelp);

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;
            int panelW = 230;
            int gridBottom = ClientSize.Height - BottomStripHeight;

            _grid = new PdfPageGrid();
            // Сетка правится как в остальных инструментах: порядок, повороты, удаление, буфер
            // и Ctrl+Z. Исходный файл при этом цел — правки описывают РЕЗУЛЬТАТ операции.
            _grid.AllowReorder = true;
            _grid.AllowRotate = true;
            _grid.ShowPositionNumbers = true; // под плиткой — место страницы в собранном документе
            _grid.SetBounds(20, m + 84, right - 20 - panelW, gridBottom - (m + 84));
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            WireSingleDocGrid(); // подсказки, правка порядка, контекстное меню, дроп (общая обвязка баз)
            // Своя подсказка пустой сетки: сюда кладут не только PDF, и общая про «Открыть PDF…»
            // умалчивала бы о половине возможностей окна.
            _grid.EmptyHint = Loc.T("ops.grid.empty");
            Controls.Add(_grid);

            int px = right - panelW + 10;
            int pw = panelW - 10;
            int y = m + 84;

            _btnOpen = AddButton(Loc.T("common.btn.openPdf"), px, y, pw, 32, Loc.T("common.tip.openPdf"));
            _btnOpen.Click += delegate { PickAndOpenFile(); };
            y += 32 + BtnGap;
            // Картинки — второй способ набрать страницы: снимки и сканы становятся документом
            // здесь же, а не в отдельном окне, потому что дальше с ними делают ровно то же —
            // поворачивают, переставляют, сжимают, печатают.
            _btnAddImages = AddButton(Loc.T("ops.btn.addImages"), px, y, pw, BtnH, Loc.T("ops.tip.addImages"));
            _btnAddImages.Click += delegate { PickAndAddImages(); };
            y += BtnH + GroupGap;

            // Действия тремя группами: подряд идущие равнозначные кнопки читаются как свалка,
            // а по заголовкам сразу видно, что делает документ, а что достаёт из него.
            y = AddGroup(Loc.T("ops.group.convert"), px, y, pw);
            _btnSave = AddButton(Loc.T("ops.btn.savePdf"), px, y, pw, BtnH, Loc.T("ops.tip.savePdf"));
            _btnSave.Click += delegate { SavePdf(); };
            y += BtnH + BtnGap;
            _btnCompress = AddButton(Loc.T("ops.btn.compress"), px, y, pw, BtnH, Loc.T("ops.tip.compress"));
            _btnCompress.Click += delegate { CompressCopy(); };
            y += BtnH + BtnGap;
            _btnGray = AddButton(Loc.T("ops.btn.grayscale"), px, y, pw, BtnH, Loc.T("ops.tip.grayscale"));
            _btnGray.Click += delegate { ConvertCopy(PdfConvertMode.Grayscale, NameSource(), true); };
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
            y += BtnH + BtnGap;
            // Печать — тоже «достать из документа», только не в файл, а на бумагу, и печатает
            // она собранное в сетке: выделенные страницы или все, с их поворотами.
            _btnPrint = AddButton(Loc.T("common.btn.print"), px, y, pw, BtnH, Loc.T("common.tip.print"));
            _btnPrint.Click += delegate { PrintPages(SelectedOrAllPages()); };
            y += BtnH + GroupGap;

            y = AddGroup(Loc.T("ops.group.edit"), px, y, pw);
            _btnMeta = AddButton(Loc.T("ops.btn.metadata"), px, y, pw, BtnH, Loc.T("ops.tip.metadata"));
            _btnMeta.Click += delegate { EditMetadata(); };

            BuildBottomStrip(right, Loc.T("common.status.openPdf"), 190);
            // Кнопки действия у этого окна нет — каждая операция своя, поэтому «Отмену» база
            // показывает на её обычном месте: правый нижний угол, как у остальных инструментов.
            RegisterCancelArea(new Rectangle(right - 190, ClientSize.Height - 58, 190, 38),
                AnchorStyles.Bottom | AnchorStyles.Right);

            // Минимум по высоте — ПО ФАКТУ собранной панели, а не числом: элементов в ней
            // одиннадцать, и каждое новое действие иначе тихо наезжало бы на нижний строй, причём
            // на машине с другим масштабом экрана — по-своему (числом это уже ловил живой тест).
            int neededClient = _btnMeta.Bottom + BottomStripHeight + BtnGap;
            int frame = Height - ClientSize.Height;
            MinimumSize = new Size(MinimumSize.Width, Math.Max(MinimumSize.Height, frame + neededClient));
        }

        /// <summary>Высота нижнего строя (масштаб, сжатие, статус) — от неё считается низ сетки.</summary>
        private const int BottomStripHeight = 152;

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

        /// <summary>
        /// Доступность кнопок. Действия смотрят на СОБРАННЫЕ страницы, а не на открытый файл:
        /// набор можно набрать одними картинками, и требовать при этом «сначала откройте PDF»
        /// значило бы выключить окно ровно там, где оно нужно.
        /// </summary>
        protected override void SyncControls()
        {
            bool ready = !Working && _order.Count > 0;
            _grid.Locked = Working;
            _compress.Enabled = !Working;
            _btnOpen.Enabled = !Working;
            _btnAddImages.Enabled = !Working;
            _btnSave.Enabled = ready;
            _btnCompress.Enabled = ready && Ghostscript.Available;
            _btnGray.Enabled = ready && Ghostscript.Available;
            // Починка выбирает файл сама: повреждённый документ в сетку не открывается, и
            // требовать «сначала откройте» значило бы выключить её ровно тогда, когда она нужна.
            _btnRepair.Enabled = !Working && Ghostscript.Available;
            _btnImages.Enabled = ready;
            _btnText.Enabled = ready;
            _btnPrint.Enabled = ready;
            // Свойства читаются из ОТКРЫТОГО документа: у набора из картинок их взять негде, а
            // предлагать правку пустоты — обещать то, чего не будет.
            _btnMeta.Enabled = !Working && _sourcePath != null;
        }

        // ---------- картинки страницами ----------

        /// <summary>Принимаем и PDF, и картинки: их кладут в набор ровно так же, как страницы.</summary>
        protected override string[] DropExtensions { get { return PdfDrop.PdfAndImages; } }

        /// <summary>
        /// Брошенное разбираем по сути: картинки ДОБАВЛЯЮТСЯ страницами, а PDF открывается —
        /// он заменяет набор целиком, как и по кнопке. Смешали в одном перетаскивании — сделаем
        /// и то, и другое: сначала откроем документ, потом добавим к нему картинки.
        /// </summary>
        protected override void AcceptDroppedPaths(string[] paths)
        {
            if (paths == null || paths.Length == 0)
                return;
            var images = new List<string>();
            string pdf = null;
            foreach (string path in paths)
            {
                if (ImageToPdfService.IsImage(path))
                    images.Add(path);
                else if (pdf == null)
                    pdf = path; // документ окно держит один — остальные не его дело
            }
            if (pdf == null)
            {
                AddImages(images.ToArray());
                return;
            }
            // Документ разбирается в фоне, и добавление в этот момент окно отклонит: картинки
            // ждут конца разбора (см. OnLoadAttemptFinished), а не теряются.
            _pendingImages = images.Count > 0 ? images.ToArray() : null;
            LoadSource(pdf); // гейтит занятость и спрашивает про собранное сам
            // Разбор не начался (заняты или человек отказался заменять набор) — очередь никто
            // не разберёт, поэтому добавляем картинки сами: они от судьбы документа не зависят.
            if (!Working && _pendingImages != null)
                OnLoadAttemptFinished(false);
        }

        /// <summary>
        /// Документ разобран (или не открылся) — добавляем картинки, брошенные вместе с ним.
        /// Не открылся — всё равно добавляем: человек бросил их не за компанию, а по делу, и
        /// сбой чужого файла не причина молча выбросить его картинки.
        /// </summary>
        protected override void OnLoadAttemptFinished(bool loaded)
        {
            string[] pending = _pendingImages;
            _pendingImages = null;
            if (pending != null)
                AddImages(pending);
        }

        /// <summary>
        /// Открытие документа заменяет набор целиком. Если в нём есть добавленные картинки, их
        /// правка исчезнет — спрашиваем: молча выбросить чужую работу одним нажатием нельзя.
        /// </summary>
        protected override bool ConfirmReplacingPages()
        {
            if (!HasWrappedImages())
                return true;
            return Dialogs.ConfirmWarning(this, Title, Loc.T("ops.ask.replacePages.title"),
                Loc.T("ops.ask.replacePages.body"));
        }

        private bool HasWrappedImages()
        {
            for (int i = 0; i < _order.Count; i++)
                if (_imageOrigins.ContainsKey(_order[i].SourcePath))
                    return true;
            return false;
        }

        /// <summary>Спросить картинки и добавить их страницами в конец набора.</summary>
        private void PickAndAddImages()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = ImageToPdfService.DialogFilter();
                dialog.Multiselect = true;
                dialog.Title = Loc.T("ops.pick.images");
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    AddImages(dialog.FileNames);
            }
        }

        /// <summary>
        /// Обернуть картинки в одностраничные PDF и добавить их страницами в конец набора.
        /// Работа идёт в фоне: снимок на 12 Мп раскладывается заметно дольше, чем окно вправе
        /// не отвечать. Сбой одной картинки не отменяет остальных — как и при добавлении PDF.
        /// </summary>
        private void AddImages(string[] paths)
        {
            if (paths == null || paths.Length == 0)
                return;
            string dir;
            try { dir = ImageTempDir(); }
            catch (Exception ex)
            {
                // Некуда положить обёртки (нет места, запрет на запись) — это отказ операции,
                // а не повод падать: говорим прямо и оставляем набор как был.
                Dialogs.Error(this, Title, Loc.T("common.fileNotAdded"), ex.Message);
                return;
            }
            if (!BeginLoad(Loc.T("ops.status.addingImages")))
                return; // уже идёт операция или загрузка
            int firstNumber = _imageCounter;
            _imageCounter += paths.Length;
            string[] toLoad = (string[])paths.Clone(); // воркер работает со снимком
            Ui.RunWorker(delegate()
            {
                var loaded = new List<LoadedDoc>();
                var origins = new List<string>();
                var errors = new List<string>();
                for (int i = 0; i < toLoad.Length; i++)
                {
                    string wrapper = Path.Combine(dir, "img" + (firstNumber + i).ToString("D4") + ".pdf");
                    // Ловим ШИРОКО (как остальные воркеры): битая, чужая или огромная картинка
                    // (в том числе OOM, который сервис НЕ маскирует) не должна ронять поток —
                    // остальные картинки пакета всё равно добавляются.
                    try
                    {
                        int pages = ImageToPdfService.WritePages(toLoad[i], wrapper);
                        loaded.Add(new LoadedDoc { Path = wrapper, PageCount = pages });
                        origins.Add(toLoad[i]);
                    }
                    catch (MergeException ex) { errors.Add(ex.Message); }
                    catch (Exception ex)
                    {
                        errors.Add(string.Format(Loc.T("err.img.cantRead"),
                            Path.GetFileName(toLoad[i]), ex.Message));
                    }
                }
                OnUi(delegate { ApplyAddedImages(loaded, origins, errors); });
            });
        }

        /// <summary>Применить результат обёртывания: запомнить настоящие имена и вставить страницы. UI-поток.</summary>
        private void ApplyAddedImages(List<LoadedDoc> loaded, List<string> origins, List<string> errors)
        {
            EndLoad(); // снять загрузку до вставки: Working=false, статус освобождается
            for (int i = 0; i < loaded.Count; i++)
                _imageOrigins[loaded[i].Path] = origins[i];
            InsertLoaded(loaded, errors, _order.Count, true); // вставка — общая с «Объединением»
        }

        /// <summary>
        /// Папка обёрток этого окна (создаётся при первой картинке). Заодно уносим мусор
        /// прошлых сеансов: аварийное завершение оставляет папку на диске, а сама она не уйдёт.
        /// </summary>
        private string ImageTempDir()
        {
            if (_imageTempDir == null)
            {
                SweepOldWrapperDirs();
                _imageTempDir = Path.Combine(Path.GetTempPath(), WrapperPrefix + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(_imageTempDir);
            }
            return _imageTempDir;
        }

        private const string WrapperPrefix = "iwo_ops_img_";

        /// <summary>
        /// Убрать папки обёрток, брошенные прошлыми сеансами. Сутки — чтобы не задеть чужой
        /// работающий экземпляр: приложение и так одно на систему, но полагаться на это здесь
        /// не стоит, а лишние сутки на диске ничего не стоят.
        /// </summary>
        private static void SweepOldWrapperDirs()
        {
            try
            {
                foreach (string dir in Directory.GetDirectories(Path.GetTempPath(), WrapperPrefix + "*"))
                    if (Directory.GetLastWriteTimeUtc(dir) < DateTime.UtcNow.AddDays(-1))
                        try { Directory.Delete(dir, true); } catch { } // занята — уйдёт в следующий раз
            }
            catch { } // нет доступа к папке временных файлов — не повод отказывать в добавлении
        }

        // ---------- документ, собранный в сетке ----------

        /// <summary>
        /// Записать в outPath документ, над которым будет работать движок: пока страницы не
        /// трогали (или файл вообще посторонний — починка), это копия исходника, быстрая и без
        /// потерь; стоит убрать, переставить или повернуть страницу — документ собирается
        /// заново, иначе действие пошло бы не по тому, что показано в сетке. Зовётся ИЗ ВОРКЕРА:
        /// сборка документа пишет на диск. pages = null — посторонний файл, собирать нечего.
        /// </summary>
        private static void WriteWorkingCopy(string source, IList<PdfPageRef> pages, int pageCount, string outPath)
        {
            if (pages == null || IsPristine(pages, source, pageCount))
                File.Copy(source, outPath, true);
            else
                PdfMergeService.Merge(pages, outPath);
        }

        /// <summary>
        /// Файл, из которого действие ЧИТАЕТ: сам исходник, пока страницы не трогали, иначе
        /// временный PDF, собранный из показанных. Временный удаляется в Dispose — операция не
        /// оставляет следов. Создаётся ИЗ ВОРКЕРА: сборка документа пишет на диск.
        /// </summary>
        private sealed class AssembledDoc : IDisposable
        {
            public readonly string FilePath;
            /// <summary>Страницы FilePath для постраничного действия (с нуля); null — читается весь файл.</summary>
            public readonly List<int> Pages;
            private readonly string _temp;

            private AssembledDoc(string path, List<int> pages, string temp)
            {
                FilePath = path;
                Pages = pages;
                _temp = temp;
            }

            /// <summary>
            /// Для постраничных действий (картинки). Пока страницы идут из исходника по
            /// возрастанию и без поворотов, читаем прямо из него: в именах картинок остаются
            /// НАСТОЯЩИЕ номера страниц документа. Иначе собираем временный PDF, и страницы в
            /// нём идут по порядку сетки — так, как человек её собрал.
            /// </summary>
            public static AssembledDoc PageWise(string source, IList<PdfPageRef> pages, Func<bool> cancelled)
            {
                var indexes = new List<int>(pages.Count);
                if (IsPlainSubset(pages, source))
                {
                    foreach (PdfPageRef page in pages)
                        indexes.Add(page.PageIndex);
                    return new AssembledDoc(source, indexes, null);
                }
                string temp = Materialize(pages, cancelled);
                for (int i = 0; i < pages.Count; i++)
                    indexes.Add(i);
                return new AssembledDoc(temp, indexes, temp);
            }

            /// <summary>
            /// Для действий, читающих документ ЦЕЛИКОМ (текст, свойства): здесь мало порядка
            /// страниц — содержимое файла должно совпадать с собранным, иначе в текст попали бы
            /// и убранные страницы.
            /// </summary>
            public static AssembledDoc Whole(string source, IList<PdfPageRef> pages, int pageCount,
                Func<bool> cancelled)
            {
                if (IsPristine(pages, source, pageCount))
                    return new AssembledDoc(source, null, null);
                string temp = Materialize(pages, cancelled);
                return new AssembledDoc(temp, null, temp);
            }

            private static string Materialize(IList<PdfPageRef> pages, Func<bool> cancelled)
            {
                string temp = Path.Combine(Path.GetTempPath(), "iwo_ops_" + Guid.NewGuid().ToString("N") + ".pdf");
                PdfMergeService.Merge(pages, temp, null, cancelled);
                return temp;
            }

            public void Dispose()
            {
                if (_temp != null)
                    try { File.Delete(_temp); } catch { } // не удалился — это папка временных файлов
            }
        }

        /// <summary>
        /// Страницы идут из ЭТОГО файла, по возрастанию номеров и без поворотов — постранично
        /// читать можно прямо из него, ничего не собирая. Чистая — под тест.
        /// </summary>
        internal static bool IsPlainSubset(IList<PdfPageRef> pages, string sourcePath)
        {
            if (pages == null || pages.Count == 0)
                return false;
            int previous = -1;
            foreach (PdfPageRef page in pages)
            {
                if (page == null || page.Rotation != 0 || page.PageIndex <= previous)
                    return false;
                if (!string.Equals(page.SourcePath, sourcePath, StringComparison.OrdinalIgnoreCase))
                    return false;
                previous = page.PageIndex;
            }
            return true;
        }

        // ---------- преобразования документа ----------

        /// <summary>
        /// Сохранить собранное в новый PDF. Так картинки становятся документом, а
        /// переставленные, повёрнутые и прореженные страницы — файлом: до 1.18.1 из этого окна
        /// нельзя было забрать собранное, можно было только применить к нему движок.
        /// Уровень сжатия берётся из общего списка внизу окна, «Без сжатия» — просто без него.
        /// </summary>
        private void SavePdf()
        {
            if (Working || _order.Count == 0)
                return;
            // Суффикс нужен там, где рядом лежит исходный PDF с тем же именем: у картинок его
            // нет, и «снимок_собранный.pdf» выглядело бы придумкой программы.
            string outPath = AskOutputPath(NameSource(), _sourcePath != null ? Loc.T("ops.suffix.saved") : "");
            if (outPath == null)
                return;
            string source = _sourcePath;
            List<PdfPageRef> pages = _order.ToList();
            int pageCount = _pageCount;
            CompressionLevel level = _compress.Level;
            BeginOperation(Loc.T("ops.status.savingPdf"), pages.Count, Loc.T("ops.status.savingPage"));
            Action<int, int> onProgress = UiProgress();
            Func<bool> cancel = CancelToken();
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                try
                {
                    // Отмена не оставляет половины документа: файл появляется в самом конце
                    // сборки, а брошенный на сжатии — удаляется (общий инвариант приложения).
                    Cancellation.NoPartialOutput(delegate(List<string> created)
                    {
                        if (IsPristine(pages, source, pageCount))
                            File.Copy(source, outPath, true); // ничего не меняли — копия без потерь
                        else
                            PdfMergeService.Merge(pages, outPath, onProgress, cancel);
                        created.Add(outPath);
                        Cancellation.ThrowIf(cancel);
                        if (level != CompressionLevel.None)
                            PdfCompression.Compress(outPath, level);
                        Cancellation.ThrowIf(cancel);
                    });
                }
                catch (Exception ex) { error = ex; }
                OnUi(delegate
                {
                    if (!FinishOperation(error, Loc.T("common.status.notDone"), Loc.T("ops.err.saveFailed")))
                        return;
                    SetStatus(SuccessStatus(string.Format(Loc.T("ops.status.savedPdf"), pages.Count)),
                        Theme.OkGreen);
                    Ui.OpenPath(outPath);
                });
            });
        }

        /// <summary>
        /// Файл, от которого предлагается имя и папка результата: открытый документ, а если
        /// открыт не он, а картинки — первая из них. Без этого человеку предлагалось бы имя
        /// временной обёртки, которого он никогда не видел.
        /// </summary>
        private string NameSource()
        {
            if (_sourcePath != null)
                return _sourcePath;
            if (_order.Count == 0)
                return null;
            string first = _order[0].SourcePath;
            string origin;
            return _imageOrigins.TryGetValue(first, out origin) ? origin : first;
        }



        /// <summary>
        /// Сжать документ в новый файл. Уровень берётся из общего списка «Сжатие» внизу окна —
        /// того же, которым пользуются объединение и разделение (DRY). «Без сжатия» здесь не
        /// действие, а отсутствие действия, поэтому честно об этом говорим вместо тихого отказа.
        /// </summary>
        private void CompressCopy()
        {
            if (Working || _order.Count == 0)
                return;
            CompressionLevel level = _compress.Level;
            if (level == CompressionLevel.None)
            {
                Dialogs.Info(this, Title, Loc.T("ops.compress.pickLevel.title"),
                    Loc.T("ops.compress.pickLevel.body"));
                return;
            }
            string outPath = AskOutputPath(NameSource(), Loc.T("ops.suffix.compressed"));
            if (outPath == null)
                return;
            string source = _sourcePath;
            List<PdfPageRef> pages = _order.ToList(); // снимок сетки: во время операции она заблокирована
            int pageCount = _pageCount;
            BeginUnmeasuredOperation(Loc.T("ops.status.compressing")); // ход сжатия движок не сообщает
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                bool shrank = false;
                try
                {
                    WriteWorkingCopy(source, pages, pageCount, outPath);
                    shrank = PdfCompression.Compress(outPath, level);
                }
                catch (Exception ex) { error = ex; }
                bool smaller = shrank;
                OnUi(delegate
                {
                    if (!FinishOperation(error, Loc.T("common.status.notDone"), Loc.T("ops.err.compressFailed")))
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
        /// fromGrid — брать документ собранным в сетке; у починки файл посторонний (его и в
        /// сетке нет), поэтому она идёт по нему как есть.
        /// </summary>
        private void ConvertCopy(PdfConvertMode mode, string source, bool fromGrid)
        {
            if (Working || string.IsNullOrEmpty(source))
                return;
            string outPath = AskOutputPath(source,
                Loc.T(mode == PdfConvertMode.Grayscale ? "ops.suffix.gray" : "ops.suffix.repaired"));
            if (outPath == null)
                return;
            List<PdfPageRef> pages = fromGrid ? _order.ToList() : null;
            int pageCount = _pageCount;
            BeginUnmeasuredOperation(Loc.T("ops.status.converting")); // движок о ходе не сообщает
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                bool ok = false;
                try
                {
                    WriteWorkingCopy(source, pages, pageCount, outPath);
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
                ConvertCopy(PdfConvertMode.Repair, dialog.FileName, false);
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
            if (Working || _order.Count == 0)
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

        /// <summary>
        /// Сохранить выбранные (или все) страницы картинками в выбранную папку — в том виде,
        /// в каком они собраны в сетке: порядок и повороты уезжают в картинки вместе с ними.
        /// </summary>
        private void ExportImages(int dpi)
        {
            if (Working || _order.Count == 0)
                return;
            List<PdfPageRef> pages = SelectedOrAllPages();
            if (pages.Count == 0)
                return;
            string dir = FolderPicker.Show(this, Loc.T("ops.pick.imagesDir"),
                Path.GetDirectoryName(NameSource()));
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
                    using (AssembledDoc doc = AssembledDoc.PageWise(source, pages, cancel))
                        files = PdfExportService.ToImages(doc.FilePath, doc.Pages, dir, NameTemplate.Default,
                            format, dpi, onProgress, cancel);
                }
                catch (Exception ex) { error = ex; }
                int count = files == null ? 0 : files.Count;
                OnUi(delegate { OnExportFinished(error, count, dir, true); });
            });
        }

        /// <summary>Извлечь текстовый слой документа в .txt.</summary>
        private void ExportText()
        {
            if (Working || _order.Count == 0)
                return;
            string nameSource = NameSource();
            string outPath;
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = Loc.T("ops.txtFilter");
                dialog.FileName = Path.GetFileNameWithoutExtension(nameSource) + ".txt";
                dialog.InitialDirectory = Path.GetDirectoryName(nameSource);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                outPath = dialog.FileName;
            }
            string source = _sourcePath;
            List<PdfPageRef> pages = _order.ToList(); // текст берётся из собранного документа, а не из файла
            int pageCount = _pageCount;
            BeginOperation(Loc.T("ops.status.extractingText"), pages.Count);
            Action<int, int> onProgress = UiProgress();
            Func<bool> cancel = CancelToken();
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                try
                {
                    using (AssembledDoc doc = AssembledDoc.Whole(source, pages, pageCount, cancel))
                        PdfExportService.ToText(doc.FilePath, outPath, onProgress, cancel);
                }
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
            List<PdfPageRef> pages = _order.ToList(); // свойства пишутся тому документу, что собран
            int pageCount = _pageCount;
            BeginUnmeasuredOperation(Loc.T("ops.status.savingMeta")); // одна запись, шагов нет
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                try
                {
                    using (AssembledDoc doc = AssembledDoc.Whole(source, pages, pageCount, null))
                        PdfMetadataService.Write(doc.FilePath, outPath, edited);
                }
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
            // Обёртки картинок удаляем ПОСЛЕ базы: до этого сетка ещё жива и держит их
            // миниатюры. Окно закрылось — временные файлы не должны его переживать.
            if (disposing && _imageTempDir != null)
            {
                try { Directory.Delete(_imageTempDir, true); } catch { } // занят — уберёт следующий сеанс
                _imageTempDir = null;
            }
        }
    }
}
