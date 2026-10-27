using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// База PDF-инструментов, у которых сетка страниц ПРАВИТСЯ: перестановка перетаскиванием,
    /// буфер страниц, удаление, поворот и откат Ctrl+Z/Y. Владеет
    /// <see cref="PdfPageOrder"/> и держит весь слой «порядок ↔ сетка» единожды.
    ///
    /// Выделена из <see cref="PdfOrderedToolFormBase"/> в 1.18.1, когда правка страниц
    /// понадобилась и окну «Прочие операции»: оно работает с ОДНИМ документом и потому на
    /// той базе жить не может, а копия слоя означала бы две правки на каждый будущий баг.
    /// Откуда берутся страницы, решает наследник: несколько файлов подряд
    /// (<see cref="PdfOrderedToolFormBase"/>) или один открытый документ
    /// (<see cref="PdfSingleDocFormBase"/>).
    /// </summary>
    public abstract class PdfPageOrderFormBase : PdfToolFormBase, IUnsavedStateAware
    {
        /// <summary>Порядок страниц, показанный сеткой (общие ссылки с ней).</summary>
        protected readonly PdfPageOrder _order = new PdfPageOrder();

        /// <summary>
        /// Любой набранный список страниц — рабочее состояние, которое нельзя терять при
        /// пересборке окна из-за языка. Чистое пустое окно можно пересоздать.
        /// </summary>
        public bool HasUncommittedState { get { return _order.Count > 0; } }

        protected PdfPageOrderFormBase(Action showHub) : base(showHub) { }

        /// <summary>
        /// Подключить события сетки к общим обработчикам модели порядка (вызвать в BuildUi
        /// наследника сразу после создания <c>_grid</c> и настройки его свойств). Включает и
        /// контекстное меню (<see cref="PdfToolFormBase.WireGridMenu"/>). Приём брошенных
        /// файлов наследники вешают сами: одни добавляют их к набору, другие — открывают
        /// вместо текущего документа.
        /// </summary>
        protected virtual void WireOrderGrid()
        {
            _grid.SelectionChanged += delegate { SyncControls(); RefreshRestingStatus(); };
            _grid.ReorderRequested += OnReorder;
            _grid.MoveRangeRequested += OnMoveRange;
            _grid.InsertPagesRequested += OnInsertPages;
            _grid.BeforeRotate += delegate { _order.Checkpoint(); }; // повороты — в историю Ctrl+Z
            // Пустой лист — такая же правка набора, как перестановка или удаление, поэтому он
            // доступен везде, где набор правится: и там, где документ собирают из нескольких
            // файлов, и там, где работают с одним. Прежняя привязка к «Прочим операциям» была
            // ограничением реализации (там лежала папка обёрток), а не смыслом.
            _grid.AllowInsertBlank = true;
            _grid.InsertBlankRequested += delegate { AddBlankPage(); };
            WireGridMenu();
        }

        // ---------- пустой лист и папка обёрток ----------

        private string _wrapDir;
        private int _wrapCounter;                                  // имена обёрток не должны совпадать
        private const string WrapperPrefix = "iwo_wrap_";
        // Имена файлов-обёрток: пустой лист и картинка (префикс + номер, см. NextWrapperPath) и
        // картинки «Прочих» прежних версий (img + 32 hex из Guid). Уборка трогает ТОЛЬКО их:
        // папка лежит во временных, но диалог сохранения мог в неё заглянуть, и чужой файл в
        // ней — это чей-то результат, а не наш мусор. Поэтому имя целиком, а не начало: маска
        // «img*» забрала бы и «IMG_2024 скан.pdf» человека.
        private static readonly Regex WrapperFileName = new Regex(
            @"^(blank|img)[0-9]{4,}\.pdf$|^img[0-9a-f]{32}\.pdf$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        // Обёртки, которые создало ЭТО окно: только их оно и вправе удалить при закрытии.
        private readonly HashSet<string> _ownWrappers =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Папка обёрток этого окна (создаётся при первой надобности). Мусор прошлых сеансов
        /// убирает уборка при запуске (<see cref="SweepOldWrapperDirs"/> из Program), а не окно:
        /// уборка из окна смотрела только на возраст папки и стирала живые обёртки СОСЕДНЕГО окна
        /// того же процесса, открытого дольше суток, — у него переставали сохраняться страницы.
        /// </summary>
        private string WrapperDir()
        {
            if (_wrapDir == null)
            {
                _wrapDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                    WrapperPrefix + Guid.NewGuid().ToString("N"));
                System.IO.Directory.CreateDirectory(_wrapDir);
            }
            return _wrapDir;
        }

        /// <summary>
        /// Путь следующей обёртки в папке окна (prefix0001.pdf, prefix0002.pdf, …). Путь сразу
        /// записывается за окном — при закрытии уйдёт именно он, а не всё содержимое папки.
        /// prefix — «blank» или «img»: уборка прошлых сеансов узнаёт обёртки по этим именам.
        /// Только UI-поток.
        /// </summary>
        protected string NextWrapperPath(string prefix)
        {
            string path = System.IO.Path.Combine(WrapperDir(), prefix + (_wrapCounter++).ToString("D4") + ".pdf");
            _ownWrappers.Add(path);
            return path;
        }

        /// <summary>
        /// Убрать обёртки, брошенные прошлыми сеансами в root. Зовётся один раз при запуске —
        /// до того, как у этого процесса появятся свои обёртки. Сутки — чтобы не задеть
        /// работающий экземпляр в другом сеансе того же пользователя (одиночный экземпляр
        /// держится на сеанс, а %TEMP% у сеансов общий); лишние сутки на диске ничего не стоят.
        ///
        /// Удаляем только файлы с именами обёрток и затем лишь ПУСТУЮ папку: рекурсивное
        /// удаление стёрло бы и документ, который человек сохранил туда сам (диалог сохранения
        /// предлагал папку первой страницы, а у пустого листа это папка обёрток).
        /// </summary>
        internal static void SweepOldWrapperDirs(string root, DateTime nowUtc)
        {
            try
            {
                foreach (string dir in System.IO.Directory.GetDirectories(root, WrapperPrefix + "*"))
                    if (System.IO.Directory.GetLastWriteTimeUtc(dir) < nowUtc.AddDays(-1))
                    {
                        foreach (string file in System.IO.Directory.GetFiles(dir, "*.pdf"))
                            if (WrapperFileName.IsMatch(System.IO.Path.GetFileName(file)))
                                try { System.IO.File.Delete(file); } catch { } // занят — уйдёт в следующий раз
                        TryDeleteEmptyDir(dir);
                    }
            }
            catch { } // нет доступа к папке временных файлов — не повод отказывать во вставке
        }

        /// <summary>Удалить папку, только если она пуста (без рекурсии). Занята или не пуста — оставить.</summary>
        private static void TryDeleteEmptyDir(string dir)
        {
            try { System.IO.Directory.Delete(dir, false); } catch { }
        }

        /// <summary>
        /// Лежит ли файл в папке обёрток — этой или любой другой (iwo_wrap_* прямо в root,
        /// папке временных файлов). Такой путь нельзя ни предлагать местом сохранения, ни
        /// считать «папкой документа». Чистая — под тест.
        /// </summary>
        internal static bool IsWrapperPath(string path, string root)
        {
            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(root))
                return false;
            try
            {
                string dir = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(path));
                string parent = dir == null ? null : System.IO.Path.GetDirectoryName(dir);
                return parent != null
                    && System.IO.Path.GetFileName(dir).StartsWith(WrapperPrefix, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(parent.TrimEnd('\\', '/'),
                        System.IO.Path.GetFullPath(root).TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; } // кривой путь (подпись вместо пути) — точно не наша обёртка
        }

        /// <summary>
        /// Папка, которую диалог сохранения предлагает первой: папка первого из paths, который
        /// — настоящий файл человека (путь с корнем и НЕ в папке обёрток). Пустой лист или
        /// картинка в начале набора иначе подставляли папку временных файлов окна — а её
        /// содержимое удаляется при закрытии, вместе с сохранённым туда результатом. Подпись
        /// вместо пути («пустая страница») пропускается. null — подходящего нет, решает диалог.
        /// Чистая — под тест.
        /// </summary>
        internal static string FirstUserFolder(IEnumerable<string> paths, string root)
        {
            if (paths == null)
                return null;
            foreach (string path in paths)
            {
                if (string.IsNullOrEmpty(path) || IsWrapperPath(path, root))
                    continue;
                try
                {
                    if (System.IO.Path.IsPathRooted(path))
                        return System.IO.Path.GetDirectoryName(path);
                }
                catch (ArgumentException) { } // подпись с недопустимыми символами — не путь
            }
            return null;
        }

        /// <summary>
        /// Папка для диалога сохранения: папка preferred (файла, от которого берётся имя
        /// результата), если это настоящий файл, иначе — первой настоящей страницы набора.
        /// </summary>
        protected string OutputFolderHint(string preferred = null)
        {
            var paths = new List<string>(_order.Count + 1) { preferred };
            for (int i = 0; i < _order.Count; i++)
                paths.Add(UserPathOf(_order[i].SourcePath));
            return FirstUserFolder(paths, System.IO.Path.GetTempPath());
        }

        /// <summary>
        /// Обёртка во временной папке окна здесь — вставленный пустой лист (картинки-обёртки
        /// бывают только в «Прочих операциях», и там свой UserPathOf): по имени файла его никто
        /// не знает, для человека это «пустая страница».
        /// </summary>
        protected override string UserPathOf(string sourcePath)
        {
            return IsWrapperPath(sourcePath, System.IO.Path.GetTempPath())
                ? Loc.T("ops.blankPage.name") : sourcePath;
        }

        /// <summary>
        /// Единый диалог «Сохранить как…» PDF-инструментов (DRY: шесть копий различались лишь
        /// фильтром, именем и папкой). initialDir — из <see cref="OutputFolderHint"/> или папка
        /// открытого документа; null — папку выбирает сам диалог. null в ответ — человек отказался.
        /// overwritePrompt = false — для «базового имени», к которому допишутся номера частей.
        /// </summary>
        protected string AskSavePath(string filter, string fileName, string initialDir,
            string title = null, bool overwritePrompt = true)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = filter;
                dialog.FileName = fileName;
                if (!string.IsNullOrEmpty(initialDir))
                    dialog.InitialDirectory = initialDir;
                if (title != null)
                    dialog.Title = title;
                dialog.OverwritePrompt = overwritePrompt;
                return dialog.ShowDialog(this) == DialogResult.OK ? dialog.FileName : null;
            }
        }

        /// <summary>
        /// Наследнику дают знать, что в набор попала страница-обёртка, которой нет на диске у
        /// пользователя: «Прочие операции» подставляют по ней осмысленное имя результата
        /// вместо «blank0001». Остальным это безразлично.
        /// </summary>
        protected virtual void OnWrapperInserted(string path, string displayName) { }

        /// <summary>
        /// Вставить пустой лист ПОСЛЕ выделенной страницы (без выделения — в конец). Формат
        /// берётся у соседа, чтобы лист не выпадал из документа.
        ///
        /// В фон уходит не запись листа (она мгновенная), а ЧТЕНИЕ размеров соседа: размеры
        /// страниц в модели не хранятся, и разбор документа ради них на потоке интерфейса
        /// подвешивал бы окно — 900-страничный файл читается 278 мс, а сетевой дольше.
        /// </summary>
        protected void AddBlankPage()
        {
            if (Working)
                return;
            string wrapper;
            try { wrapper = NextWrapperPath("blank"); }
            catch (Exception ex)
            {
                // Некуда положить обёртку (нет места, запрет на запись) — отказ операции,
                // а не повод падать: говорим прямо и оставляем набор как был.
                Dialogs.Error(this, Text, Loc.T("common.fileNotAdded"), ex.Message);
                return;
            }
            // После последней выделенной страницы: так же ведёт себя вставка из буфера.
            int at = _order.Count;
            int[] selected = _grid.GetSelectedIndices();
            if (selected.Length > 0)
                at = selected[selected.Length - 1] + 1;
            PdfPageRef neighbour = NeighbourOf(at);

            if (!BeginLoad(Loc.T("common.status.loading")))
                return;
            Func<PdfPageRef, PdfPageInfo> sizeOf = SizeOfForTests ?? SizeOf;
            Ui.RunWorker(delegate()
            {
                string error = null;
                bool cancelled = false;
                try
                {
                    if (LoadCancellationRequested)
                        cancelled = true;
                    else
                    {
                        double width, height;
                        // Чтение соседнего документа — только здесь, в фоне: на 900-страничном
                        // файле разбор ради размеров занимал 278 мс на потоке интерфейса.
                        BlankPages.SheetSize(sizeOf(neighbour), out width, out height);
                        if (LoadCancellationRequested)
                            cancelled = true;
                        else
                            BlankPages.WriteSheet(wrapper, width, height);
                    }
                }
                catch (Exception ex) { error = ex.Message; }
                string err = error;
                bool wasCancelled = cancelled || LoadCancellationRequested;
                OnUi(delegate
                {
                    if (wasCancelled || LoadCancellationRequested)
                    {
                        try { System.IO.File.Delete(wrapper); } catch { }
                        FinishCanceledLoad();
                        return;
                    }
                    EndLoad();
                    if (err != null)
                    {
                        Dialogs.Error(this, Text, Loc.T("common.fileNotAdded"), err);
                        return;
                    }
                    OnWrapperInserted(wrapper, Loc.T("ops.blankPage.name"));
                    var loaded = new List<LoadedDoc> { new LoadedDoc { Path = wrapper, PageCount = 1 } };
                    InsertLoaded(loaded, new List<string>(), at, true);
                });
            });
        }

        /// <summary>Страница, рядом с которой встаёт пустой лист (та, что на позиции at−1), или null.</summary>
        private PdfPageRef NeighbourOf(int at)
        {
            int index = at - 1;
            if (index < 0 || index >= _order.Count)
                index = _order.Count - 1;
            return index < 0 ? null : _order[index];
        }

        /// <summary>Шов для теста: подмена чтения размеров соседней страницы (null — настоящее).</summary>
        internal static Func<PdfPageRef, PdfPageInfo> SizeOfForTests { get; set; }

        /// <summary>
        /// Размеры страницы. Только с фонового потока: разбор документа не мгновенный.
        /// Не прочитали — null, и лист станет A4: это не повод отказывать во вставке.
        /// </summary>
        private static PdfPageInfo SizeOf(PdfPageRef page)
        {
            if (page == null)
                return null;
            try
            {
                List<PdfPageInfo> pages = PdfMergeService.LoadPages(page.SourcePath);
                return page.PageIndex >= 0 && page.PageIndex < pages.Count ? pages[page.PageIndex] : null;
            }
            catch { return null; }
        }

        /// <summary>
        /// Обёртки живут ровно столько, сколько окно. Удаляем только СВОИ файлы и затем папку,
        /// если она опустела: прежнее рекурсивное удаление стирало и результат, который человек
        /// сохранил в эту папку (диалог предлагал её, когда первой в наборе стоял пустой лист).
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && _wrapDir != null)
            {
                foreach (string file in _ownWrappers)
                    try { System.IO.File.Delete(file); } catch { } // занят — уберёт следующий сеанс
                _ownWrappers.Clear();
                TryDeleteEmptyDir(_wrapDir);
                _wrapDir = null;
            }
            base.Dispose(disposing);
        }

        /// <summary>Перестроить сетку из текущего порядка.</summary>
        protected void RefreshGrid()
        {
            _grid.SetPages(_order.ToList());
        }

        /// <summary>Один разобранный источник, готовый лечь страницами в набор.</summary>
        protected sealed class LoadedDoc
        {
            public string Path;
            public int PageCount;
        }

        /// <summary>
        /// Вставить разобранные источники в набор ПЕРЕД позицией at (за пределами набора —
        /// в конец), выделить вставленное, вернуть статус и показать ошибки. Один снимок
        /// Ctrl+Z на весь пакет: добавление — это один жест человека, а не десять.
        ///
        /// Общая половина добавления: разбирают источники по-разному (PDF читается сразу,
        /// картинка сперва оборачивается в страницу), а в набор ложатся одинаково.
        /// </summary>
        protected void InsertLoaded(List<LoadedDoc> loaded, List<string> errors, int at, bool select)
        {
            int added = 0, firstAdded = -1;
            if (loaded != null && loaded.Count > 0)
            {
                if (at < 0 || at > _order.Count)
                    at = _order.Count;
                _order.Checkpoint();
                foreach (LoadedDoc doc in loaded)
                {
                    int landed = _order.InsertDocument(at, doc.Path, doc.PageCount);
                    if (firstAdded < 0)
                        firstAdded = landed;
                    at = landed + doc.PageCount; // следующий источник — сразу за вставленным
                    added += doc.PageCount;
                }
            }
            if (added > 0)
            {
                RefreshGrid();
                if (select)
                    _grid.SelectRange(firstAdded, added); // показать, что именно добавилось
            }
            // Вернуть idle-статус вместо «Загрузка…» на ЛЮБОМ исходе (в том числе когда не
            // добавился ни один источник) — иначе статус залипал бы на загрузке.
            RefreshRestingStatus();
            SyncControls();
            ShowLoadErrors(errors);
        }

        // ---------- перестановка, вставка, удаление ----------

        private void ShowLoadErrors(IList<string> errors)
        {
            if (errors == null || errors.Count == 0)
                return;
            const int MaxShown = 30;
            var text = new StringBuilder();
            int shown = Math.Min(MaxShown, errors.Count);
            for (int i = 0; i < shown; i++)
            {
                if (i > 0)
                    text.AppendLine();
                text.Append("• ").Append(errors[i]);
            }
            if (errors.Count > shown)
                text.AppendLine().AppendLine().Append(string.Format(
                    Loc.T("common.fileNotAddedMore"), errors.Count - shown));
            Dialogs.Error(this, ToolTitle, Loc.T("common.fileNotAdded"), text.ToString());
        }

        private void OnReorder(int from, int to)
        {
            // Бросили страницу на её же место (или пришёл чужой индекс) — правки нет, и
            // снимок Ctrl+Z был бы пустым шагом отмены, стирающим ветку Ctrl+Y.
            if (Working || IsNoOpMove(_order.Count, from, to))
                return;
            _order.Checkpoint();
            _order.Move(from, to);
            RefreshGrid();
            _grid.SelectIndex(to > from ? to - 1 : to); // выделить страницу на новом месте
        }

        /// <summary>Вставка вырезанных страниц (Ctrl+X → Ctrl+V) — перенос набора внутри порядка.</summary>
        private void OnMoveRange(int[] indices, int insertAt)
        {
            if (Working || IsNoOpRangeMove(_order.Count, indices, insertAt))
                return; // пустой набор или вставка «на себя» — ни правки, ни шага отмены
            _order.Checkpoint();
            int landed = _order.MoveRange(indices, insertAt);
            RefreshGrid();
            _grid.SelectRange(landed, ListReorder.NormalizeIndices(indices, _order.Count).Count);
        }

        /// <summary>
        /// Перенос элемента from ПЕРЕД позицией to ничего не меняет: индекс вне набора (элемент
        /// чужого списка), to == from или to == from + 1 (бросили сразу за собой). Та же
        /// семантика, что у <see cref="ListReorder.Move{T}"/>. Чистая — под тест.
        /// </summary>
        internal static bool IsNoOpMove(int count, int from, int to)
        {
            return from < 0 || from >= count || to < 0 || to > count || to == from || to == from + 1;
        }

        /// <summary>
        /// Перенос набора (<see cref="ListReorder.MoveRange{T}"/>) ничего не меняет: живых индексов
        /// нет, или они идут подряд и встают ровно туда, где стоят. Чистая — под тест.
        /// </summary>
        internal static bool IsNoOpRangeMove(int count, IList<int> indices, int insertAt)
        {
            List<int> sorted = ListReorder.NormalizeIndices(indices, count);
            if (sorted.Count == 0)
                return true;
            if (insertAt < 0) insertAt = 0;
            if (insertAt > count) insertAt = count;
            int landed = ListReorder.AdjustedInsertIndex(sorted, insertAt);
            for (int i = 0; i < sorted.Count; i++)
                if (sorted[i] != landed + i)
                    return false;
            return true;
        }

        /// <summary>Вставка скопированных страниц (Ctrl+C → Ctrl+V) — новые экземпляры в позиции.</summary>
        private void OnInsertPages(PdfPageRef[] pages, int insertAt)
        {
            if (Working || pages == null || pages.Length == 0)
                return;
            _order.Checkpoint();
            int landed = _order.InsertAt(insertAt, pages);
            RefreshGrid();
            _grid.SelectRange(landed, pages.Length);
            RefreshRestingStatus();
            SyncControls();
        }

        /// <summary>Сдвинуть единственную выбранную страницу раньше/позже (Alt+←/→, кнопки).</summary>
        protected void MoveSelected(bool later)
        {
            if (Working || _grid.SelectedCount != 1)
                return;
            int index = _grid.GetSelectedIndices()[0];
            bool willMove = later ? index < _order.Count - 1 : index > 0;
            if (!willMove)
                return; // уже с краю — снимок для Ctrl+Z не нужен
            _order.Checkpoint();
            int moved = later ? _order.MoveDown(index) : _order.MoveUp(index);
            RefreshGrid();
            _grid.SelectIndex(moved);
        }

        /// <summary>Удалить выбранные страницы (кнопка, Delete).</summary>
        protected void RemoveSelected()
        {
            if (Working || _grid.SelectedCount == 0)
                return;
            _order.Checkpoint();
            _order.RemoveAt(_grid.GetSelectedIndices());
            RefreshGrid();
            RefreshRestingStatus();
            SyncControls();
        }

        /// <summary>
        /// Страницы для действия над частью набора (печать, картинки): выделенные в сетке, а
        /// если ничего не выделено — весь собранный порядок. Одно правило на все инструменты,
        /// чтобы «выделил — сделал с выделенным» работало везде одинаково.
        /// </summary>
        protected List<PdfPageRef> SelectedOrAllPages()
        {
            var all = _order.ToList();
            int[] selected = _grid.GetSelectedIndices();
            if (selected.Length == 0)
                return all;
            var picked = new List<PdfPageRef>(selected.Length);
            foreach (int i in selected)
                if (i >= 0 && i < all.Count)
                    picked.Add(all[i]);
            return picked;
        }

        // Горячие клавиши сетки (Delete, Alt+←/→) — в базе PdfToolFormBase.
        protected override void RemoveSelectedPages() { RemoveSelected(); }
        protected override void MoveSelectedPage(bool later) { MoveSelected(later); }

        /// <summary>Ctrl+Z: откат последнего жеста (перенос, удаление, вставка, добавление, поворот).</summary>
        protected override void UndoOrder()
        {
            if (Working || !_order.Undo())
                return;
            RefreshGrid();
            RefreshRestingStatus();
            SyncControls();
        }

        /// <summary>Ctrl+Y / Ctrl+Shift+Z: возврат откаченного жеста.</summary>
        protected override void RedoOrder()
        {
            if (Working || !_order.Redo())
                return;
            RefreshGrid();
            RefreshRestingStatus();
            SyncControls();
        }
    }
}
