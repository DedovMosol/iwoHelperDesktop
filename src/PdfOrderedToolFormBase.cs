using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// База PDF-инструментов с МОДЕЛЬЮ ПОРЯДКА страниц — «Объединение» и «PDF → Word»:
    /// обе собирают один результат из страниц нескольких файлов в выбранном порядке.
    /// Владеет <see cref="PdfPageOrder"/> и держит весь слой «порядок ↔ сетка»
    /// (добавление файлов, перестановка, вставка/перенос из буфера, удаление, Ctrl+Z/Y)
    /// единожды (DRY): раньше он был посимвольной копией в двух формах. Наследник строит
    /// свой интерфейс (кнопки, действие), создаёт <c>_grid</c> и вызывает
    /// <see cref="WireOrderGrid"/> и <see cref="WireFileDropAppend"/>, а различия сводит в
    /// <see cref="PdfToolFormBase.SyncControls"/> и <see cref="ToolTitle"/>.
    ///
    /// Разбор PDF при добавлении идёт в фоне (<see cref="PdfToolFormBase.BeginLoad"/>) —
    /// большой или сетевой файл не морозит окно.
    /// </summary>
    public abstract class PdfOrderedToolFormBase : PdfToolFormBase, IFileAcceptor
    {
        /// <summary>Модель порядка страниц будущего результата (общие ссылки с сеткой).</summary>
        protected readonly PdfPageOrder _order = new PdfPageOrder();

        protected PdfOrderedToolFormBase(Action showHub) : base(showHub) { }

        /// <summary>Один разобранный файл, готовый к вставке в порядок (результат фонового разбора).</summary>
        private sealed class LoadedDoc { public string Path; public int PageCount; }

        /// <summary>
        /// Подключить события сетки к общим обработчикам модели порядка (вызвать в BuildUi
        /// наследника сразу после создания <c>_grid</c> и настройки его свойств). Включает и
        /// контекстное меню (<see cref="PdfToolFormBase.WireGridMenu"/>).
        /// </summary>
        protected void WireOrderGrid()
        {
            _grid.SelectionChanged += delegate { SyncControls(); RefreshRestingStatus(); };
            _grid.ReorderRequested += OnReorder;
            _grid.MoveRangeRequested += OnMoveRange;
            _grid.InsertPagesRequested += OnInsertPages;
            _grid.FilesDropped += delegate(string[] paths, int insertAt) { AddFiles(paths, insertAt); };
            _grid.BeforeRotate += delegate { _order.Checkpoint(); }; // повороты — в историю Ctrl+Z
            WireGridMenu();
        }

        /// <summary>Приём PDF, брошенных на ОКНО: добавить в конец (общая обвязка базы + модель порядка).</summary>
        protected void WireFileDropAppend()
        {
            WireFileDrop(delegate(string[] paths) { AddFiles(paths); });
        }

        /// <summary>Дроп PDF на карточку хаба (<see cref="IFileAcceptor"/>): добавить файлы в конец.</summary>
        public void AcceptFiles(string[] paths)
        {
            AddFiles(paths);
        }

        // ---------- добавление файлов (фоновый разбор) ----------

        /// <summary>
        /// Добавить файлы в конец (кнопка, дроп на окно) или ПЕРЕД позицией insertAt (дроп на
        /// сетку). Разбор PDF идёт в фоне; ошибка файла — диалог, остальные добавляются.
        /// Повторный вызов во время операции/загрузки игнорируется.
        /// </summary>
        protected void AddFiles(string[] paths, int insertAt = -1)
        {
            if (paths == null || paths.Length == 0)
                return;
            if (!BeginLoad(Loc.T("common.status.loading")))
                return; // уже идёт операция или загрузка
            int at = insertAt < 0 || insertAt > _order.Count ? _order.Count : insertAt;
            bool insertMode = insertAt >= 0;
            string[] toLoad = (string[])paths.Clone(); // воркер работает со снимком
            Ui.RunWorker(delegate()
            {
                var loaded = new List<LoadedDoc>();
                var errors = new List<string>();
                foreach (string path in toLoad)
                {
                    // Ловим ШИРОКО (как операционные воркеры): один битый/занятый/аварийный файл
                    // (в т.ч. редкий OOM, который LoadPages НЕ оборачивает) не должен ронять
                    // фоновый поток — остальные файлы пакета всё равно добавляются.
                    try { loaded.Add(new LoadedDoc { Path = path, PageCount = PdfMergeService.LoadPages(path).Count }); }
                    catch (MergeException ex) { errors.Add(ex.Message); } // понятное локализованное сообщение
                    catch (Exception ex) { errors.Add(string.Format(Loc.T("err.pdf.cantOpen"), Path.GetFileName(path), ex.Message)); }
                }
                OnUi(delegate { ApplyAdded(loaded, errors, at, insertMode); });
            });
        }

        /// <summary>Применить результат фонового разбора: вставить страницы, показать ошибки. UI-поток.</summary>
        private void ApplyAdded(List<LoadedDoc> loaded, List<string> errors, int at, bool insertMode)
        {
            EndLoad(); // снять загрузку до вставки: Working=false, статус освобождается
            int added = 0, firstAdded = -1;
            if (loaded.Count > 0)
            {
                _order.Checkpoint(); // один снимок Ctrl+Z на весь добавленный пакет
                foreach (LoadedDoc doc in loaded)
                {
                    int landed = _order.InsertDocument(at, doc.Path, doc.PageCount);
                    if (firstAdded < 0)
                        firstAdded = landed;
                    at = landed + doc.PageCount; // следующий файл — сразу за вставленным
                    added += doc.PageCount;
                }
            }
            if (added > 0)
            {
                RefreshGrid();
                if (insertMode)
                    _grid.SelectRange(firstAdded, added); // показать место вставки дропа
            }
            // Вернуть idle/счётчик-статус вместо «Загрузка…» на ЛЮБОМ исходе (в т.ч. когда
            // ни один файл не добавился) — иначе статус залипал бы на «Загрузка…».
            RefreshRestingStatus();
            SyncControls();
            foreach (string err in errors)
                Dialogs.Error(this, ToolTitle, Loc.T("common.fileNotAdded"), err);
        }

        /// <summary>Перестроить сетку из текущего порядка.</summary>
        protected void RefreshGrid()
        {
            _grid.SetPages(_order.ToList());
        }

        // ---------- перестановка, вставка, удаление ----------

        private void OnReorder(int from, int to)
        {
            if (Working)
                return;
            _order.Checkpoint();
            _order.Move(from, to);
            RefreshGrid();
            _grid.SelectIndex(to > from ? to - 1 : to); // выделить страницу на новом месте
        }

        /// <summary>Вставка вырезанных страниц (Ctrl+X → Ctrl+V) — перенос набора внутри порядка.</summary>
        private void OnMoveRange(int[] indices, int insertAt)
        {
            if (Working)
                return;
            _order.Checkpoint();
            int landed = _order.MoveRange(indices, insertAt);
            if (landed < 0)
                return;
            RefreshGrid();
            _grid.SelectRange(landed, indices.Length);
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
        /// Диалог выбора PDF-файлов и добавление выбранного в порядок страниц. Оба
        /// инструмента, собирающих результат из нескольких файлов («Объединение» и
        /// «PDF → Word»), держали его дословно одинаковой копией.
        /// </summary>
        protected void PickAndAddFiles()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = Loc.T("common.pdfFilter");
                dialog.Multiselect = true;
                dialog.Title = Loc.T("common.pickPdf");
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    AddFiles(dialog.FileNames);
            }
        }

        /// <summary>
        /// Что печатать: выделенные в сетке страницы, а если ничего не выделено — весь
        /// собранный порядок. Одно правило на оба инструмента, чтобы кнопка «Печать» вела
        /// себя одинаково и там, и там.
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
