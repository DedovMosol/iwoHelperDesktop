using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// База инструментов, работающих с ОДНИМ открытым документом — «Разделение PDF» и
    /// «Прочие операции». Держит весь слой «открыть файл и показать его страницы»: выбор
    /// файла, фоновый разбор, заполнение сетки, статус, приём файла дропом на окно, на сетку
    /// и с карточки стартового экрана, а также выбор страниц для операций.
    ///
    /// Выделена в 1.17.9, когда шесть действий над одним документом переехали из «Разделения»
    /// в собственное окно: копировать этот слой во второе окно значило бы держать две правки
    /// на каждый будущий баг. Инструменты, собирающие результат из НЕСКОЛЬКИХ файлов, живут
    /// на другой базе — <see cref="PdfOrderedToolFormBase"/>.
    ///
    /// Наследник строит свою раскладку сам и зовёт <see cref="WireSingleDocGrid"/> после
    /// создания сетки. Только UI-поток, кроме тела фонового разбора.
    /// </summary>
    public abstract class PdfSingleDocFormBase : PdfToolFormBase, IFileAcceptor
    {
        /// <summary>Путь открытого документа или null, если ещё ничего не открыли.</summary>
        protected string _sourcePath;
        /// <summary>Число страниц открытого документа.</summary>
        protected int _pageCount;
        /// <summary>
        /// Страницы открытого документа, показанные сеткой. Сетка мутирует их
        /// <see cref="PdfPageRef.Rotation"/> при повороте, отсюда форма собирает карту
        /// поворотов для записи.
        /// </summary>
        protected List<PdfPageRef> _pages = new List<PdfPageRef>();

        protected PdfSingleDocFormBase(Action showHub) : base(showHub) { }

        /// <summary>Заголовок окна выбора файла — у каждого инструмента свой («…для разделения»).</summary>
        protected abstract string PickFileTitle { get; }

        /// <summary>Idle-статус: не открыт — подсказка открытия, иначе имя файла и число страниц.</summary>
        protected override string IdleStatusText()
        {
            return _sourcePath == null
                ? Loc.T("common.status.openPdf")
                : string.Format(Loc.T("common.status.opened"), Path.GetFileName(_sourcePath), _pageCount);
        }

        /// <summary>
        /// Общая обвязка сетки одного документа: выделение отражается на кнопках и статусе,
        /// дроп на сетку и на окно открывает первый брошенный файл (инструмент работает с
        /// одним документом, поэтому остальные игнорируются).
        /// </summary>
        protected void WireSingleDocGrid()
        {
            _grid.EmptyHint = Loc.T("common.grid.emptyOpen");
            _grid.DropHint = Loc.T("common.grid.dropOpen");
            _grid.SelectionChanged += delegate { SyncControls(); RefreshRestingStatus(); };
            _grid.FilesDropped += delegate(string[] paths, int insertAt)
            {
                if (paths.Length > 0)
                    LoadSource(paths[0]); // LoadSource гейтит занятость сам
            };
            WireFileDrop(delegate(string[] paths) { LoadSource(paths[0]); });
        }

        /// <summary>Дроп PDF на карточку стартового экрана: открыть первый файл.</summary>
        public void AcceptFiles(string[] paths)
        {
            if (paths != null && paths.Length > 0)
                LoadSource(paths[0]); // LoadSource гейтит занятость сам
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
            if (!BeginLoad(Loc.T("common.status.loading")))
                return; // уже идёт операция или загрузка
            Ui.RunWorker(delegate()
            {
                int count = 0;
                string error = null;
                // Ловим ШИРОКО: битый, занятый или аварийный файл (в том числе редкий OOM,
                // который LoadPages НЕ оборачивает) не должен ронять фоновый поток — только диалог.
                try { count = PdfMergeService.LoadPages(path).Count; }
                catch (MergeException ex) { error = ex.Message; }
                catch (Exception ex) { error = string.Format(Loc.T("err.pdf.cantOpen"), Path.GetFileName(path), ex.Message); }
                int pages = count;
                string err = error;
                OnUi(delegate { ApplyLoadedSource(path, pages, err); });
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
                return;
            }
            _sourcePath = path;
            _pageCount = pageCount;
            _pages = new List<PdfPageRef>();
            for (int i = 0; i < _pageCount; i++)
                _pages.Add(new PdfPageRef { SourcePath = path, PageIndex = i });
            _grid.SetPages(_pages);
            SetStatus(string.Format(Loc.T("common.status.opened"), Path.GetFileName(path), _pageCount),
                Theme.TextMuted);
            OnSourceLoaded();
            SyncControls();
        }

        /// <summary>Документ открыт — наследник может подстроить свои поля. По умолчанию ничего.</summary>
        protected virtual void OnSourceLoaded() { }

        /// <summary>Страницы для операции: выделенные в сетке, а если ничего не выделено — все.</summary>
        protected List<int> PagesForExport()
        {
            var pages = new List<int>(_grid.GetSelectedIndices());
            if (pages.Count == 0)
                for (int i = 0; i < _pageCount; i++)
                    pages.Add(i);
            pages.Sort();
            return pages;
        }

        /// <summary>
        /// Ссылки на выбранные страницы (для печати): берутся из показанных сеткой, чтобы
        /// вместе со страницей уехал и назначенный ей поворот.
        /// </summary>
        protected List<PdfPageRef> SelectedPageRefs()
        {
            var refs = new List<PdfPageRef>();
            foreach (int i in PagesForExport())
                refs.Add(i < _pages.Count ? _pages[i] : new PdfPageRef { SourcePath = _sourcePath, PageIndex = i });
            return refs;
        }
    }
}
