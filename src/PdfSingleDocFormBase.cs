using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// База инструментов, работающих с ОДНИМ открытым документом — «Разделение PDF» и
    /// «Прочие операции». Держит весь слой «открыть файл и показать его страницы»: выбор
    /// файла, фоновый разбор, заполнение сетки, статус и приём файла дропом на окно, на сетку
    /// и с карточки стартового экрана.
    ///
    /// Выделена в 1.17.9, когда шесть действий над одним документом переехали из «Разделения»
    /// в собственное окно: копировать этот слой во второе окно значило бы держать две правки
    /// на каждый будущий баг. Правка страниц (порядок, поворот, удаление, Ctrl+Z) — общая с
    /// многофайловыми инструментами и живёт в <see cref="PdfPageOrderFormBase"/>; сколько из
    /// неё разрешено, каждое окно решает флагами сетки.
    ///
    /// Наследник строит свою раскладку сам и зовёт <see cref="WireSingleDocGrid"/> после
    /// создания сетки. Только UI-поток, кроме тела фонового разбора.
    /// </summary>
    public abstract class PdfSingleDocFormBase : PdfPageOrderFormBase, IFileAcceptor
    {
        /// <summary>Путь открытого документа или null, если ещё ничего не открыли.</summary>
        protected string _sourcePath;
        /// <summary>Число страниц В ФАЙЛЕ (не в сетке: из неё страницы могли убрать).</summary>
        protected int _pageCount;

        protected PdfSingleDocFormBase(Action showHub) : base(showHub) { }

        /// <summary>Заголовок окна выбора файла — у каждого инструмента свой («…для разделения»).</summary>
        protected abstract string PickFileTitle { get; }

        /// <summary>
        /// Idle-статус: не открыт — подсказка открытия, иначе имя файла и число страниц В ФАЙЛЕ.
        /// Если в сетке собрано другое (страницы убрали, переставили или повернули), говорим и
        /// это: иначе окно молчит о том, что дальше операции пойдут не по исходнику.
        /// </summary>
        protected override string IdleStatusText()
        {
            if (_sourcePath == null)
                return Loc.T("common.status.openPdf");
            string status = string.Format(Loc.T("common.status.opened"), Path.GetFileName(_sourcePath), _pageCount);
            if (!IsPristine(_order.ToList(), _sourcePath, _pageCount))
                status += string.Format(Loc.T("common.status.assembled"), _order.Count);
            return status;
        }

        /// <summary>
        /// Общая обвязка сетки одного документа: правка страниц как у всех (база), а брошенные
        /// файлы уходят в <see cref="AcceptDroppedPaths"/> — что с ними делать, решает
        /// инструмент. Набор принимаемых расширений тоже его: «Прочие операции» берут и картинки.
        /// </summary>
        protected void WireSingleDocGrid()
        {
            _grid.EmptyHint = Loc.T("common.grid.emptyOpen");
            _grid.DropHint = Loc.T("common.grid.dropOpen");
            _grid.DropExtensions = DropExtensions;
            WireOrderGrid();
            _grid.FilesDropped += delegate(string[] paths, int insertAt) { AcceptDroppedPaths(paths); };
            WireFileDrop(AcceptDroppedPaths, DropExtensions);
        }

        /// <summary>Что инструмент принимает перетаскиванием. По умолчанию только PDF.</summary>
        protected virtual string[] DropExtensions { get { return PdfDrop.PdfOnly; } }

        /// <summary>
        /// Брошенные файлы (на окно, на сетку, на карточку хаба — одна точка на все три):
        /// инструмент одного документа открывает ПЕРВЫЙ, остальные не его дело. Наследник может
        /// решить иначе — «Прочие операции» ещё и добавляют картинки страницами.
        /// </summary>
        protected virtual void AcceptDroppedPaths(string[] paths)
        {
            if (paths != null && paths.Length > 0)
                LoadSource(paths[0]); // LoadSource гейтит занятость сам
        }

        /// <summary>Дроп файлов на карточку стартового экрана — тем же путём, что и на окно.</summary>
        public void AcceptFiles(string[] paths)
        {
            AcceptDroppedPaths(paths);
        }

        /// <summary>Спросить файл и открыть его.</summary>
        protected void PickAndOpenFile()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = Loc.T("common.pdfFilter");
                dialog.Title = PickFileTitle;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    LoadSource(dialog.FileName);
            }
        }

        /// <summary>
        /// Открыть документ: разбор PDF идёт в фоне (большой или сетевой файл не морозит окно),
        /// затем сетка страниц и статус. Битый или занятый файл — диалог, прежний документ цел.
        /// </summary>
        protected void LoadSource(string path)
        {
            if (Working)
                return; // идёт операция или загрузка: спрашивать не о чем, ответ ничего не изменит
            if (!ConfirmReplacingPages())
                return; // человек не согласился потерять собранное — прежний набор цел
            if (!BeginLoad(Loc.T("common.status.loading")))
                return; // защёлка на случай гонки: решение всё равно за BeginLoad
            LoadSourcePass(path, false);
        }

        /// <summary>
        /// Один заход открытия документа. Защищённый файл не ошибка, а вопрос: спрашиваем
        /// пароль и заходим ещё раз (retry — прошлый пароль не подошёл, скажем об этом).
        /// Отказ от ввода завершает круг обычным сообщением «файл не открыт».
        /// </summary>
        private void LoadSourcePass(string path, bool retry)
        {
            Ui.RunWorker(delegate()
            {
                int count = 0;
                string error = null;
                bool locked = false;
                // Ловим ШИРОКО: битый, занятый или аварийный файл (в том числе редкий OOM,
                // который LoadPages НЕ оборачивает) не должен ронять фоновый поток — только диалог.
                try { count = PdfMergeService.LoadPages(path).Count; }
                catch (Exception ex)
                {
                    if (PdfPasswords.LooksPasswordProtected(ex))
                        locked = true;
                    else if (ex is MergeException)
                        error = ex.Message;
                    else
                        error = string.Format(Loc.T("err.pdf.cantOpen"), Path.GetFileName(path), ex.Message);
                }
                int pages = count;
                string err = error;
                bool needPassword = locked;
                OnUi(delegate
                {
                    if (needPassword)
                    {
                        var errors = new List<string>();
                        // Сбой при опросе пароля не должен оставить окно заблокированным:
                        // тогда инструмент пришлось бы закрывать и открывать заново.
                        try
                        {
                            List<string> again = AskPasswords(new[] { path },
                                retry ? new[] { path } : new string[0], errors);
                            if (again.Count > 0)
                            {
                                LoadSourcePass(path, true); // пароль дали — пробуем ещё раз
                                return;
                            }
                        }
                        catch (Exception ex)
                        {
                            errors.Add(ex.Message);
                        }
                        ApplyLoadedSource(path, 0, errors.Count > 0 ? errors[0] : null);
                        return;
                    }
                    ApplyLoadedSource(path, pages, err);
                });
            });
        }

        /// <summary>Применить результат фонового разбора документа (UI-поток).</summary>
        private void ApplyLoadedSource(string path, int pageCount, string error)
        {
            EndLoad();
            if (error != null)
            {
                RefreshRestingStatus(); // вернуть статус прежнего документа вместо залипшей «Загрузка…»
                Dialogs.Error(this, ToolTitle, Loc.T("common.err.fileNotOpened"), error); // прежний документ остаётся
                OnLoadAttemptFinished(false);
                return;
            }
            _sourcePath = path;
            _pageCount = pageCount;
            _order.Clear(); // другой документ — и правки прежнего, и их история ни к чему
            _order.AddDocument(path, pageCount);
            RefreshGrid();
            SetStatus(string.Format(Loc.T("common.status.opened"), Path.GetFileName(path), _pageCount),
                Theme.TextMuted);
            OnSourceLoaded();
            SyncControls();
            OnLoadAttemptFinished(true);
        }

        /// <summary>
        /// Попытка открыть документ завершилась — удачно или нет. Нужен тем, у кого на очереди
        /// ждёт своя работа: «Прочие операции» так добавляют картинки, брошенные вместе с PDF
        /// (сразу их добавить нельзя — окно занято разбором, и они пропали бы молча).
        /// </summary>
        protected virtual void OnLoadAttemptFinished(bool loaded) { }

        /// <summary>Документ открыт — наследник может подстроить свои поля. По умолчанию ничего.</summary>
        protected virtual void OnSourceLoaded() { }

        /// <summary>
        /// Спросить разрешение заменить собранный набор новым документом. По умолчанию нечего и
        /// спрашивать: набор — это и есть открытый файл. Смысл появляется там, где в сетку
        /// добавляют своё («Прочие операции» — картинки): молча выбросить их значило бы стереть
        /// работу человека одним нажатием «Открыть PDF…».
        /// </summary>
        protected virtual bool ConfirmReplacingPages() { return true; }

        /// <summary>
        /// Набор страниц — это открытый файл КАК ЕСТЬ: все его страницы, в исходном порядке и
        /// без поворотов. Тогда операция идёт прямо по исходнику, а иначе документ приходится
        /// сначала собрать заново — тем, что человек видит в сетке. Чистая — под тест.
        /// </summary>
        internal static bool IsPristine(IList<PdfPageRef> pages, string sourcePath, int pageCount)
        {
            if (pages == null || pages.Count != pageCount)
                return false;
            for (int i = 0; i < pages.Count; i++)
            {
                PdfPageRef page = pages[i];
                if (page == null || page.PageIndex != i || page.Rotation != 0)
                    return false;
                if (!string.Equals(page.SourcePath, sourcePath, StringComparison.OrdinalIgnoreCase))
                    return false;
            }
            return true;
        }
    }
}
