using System;
using System.Collections.Generic;
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
    public abstract class PdfPageOrderFormBase : PdfToolFormBase
    {
        /// <summary>Порядок страниц, показанный сеткой (общие ссылки с ней).</summary>
        protected readonly PdfPageOrder _order = new PdfPageOrder();

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

        /// <summary>
        /// Папка обёрток этого окна (создаётся при первой надобности). Заодно уносим мусор
        /// прошлых сеансов: аварийное завершение оставляет папку на диске, сама она не уйдёт.
        /// </summary>
        protected string WrapperDir()
        {
            if (_wrapDir == null)
            {
                SweepOldWrapperDirs();
                _wrapDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                    WrapperPrefix + Guid.NewGuid().ToString("N"));
                System.IO.Directory.CreateDirectory(_wrapDir);
            }
            return _wrapDir;
        }

        /// <summary>Путь следующей обёртки в папке окна (prefix0001.pdf, prefix0002.pdf, …).</summary>
        protected string NextWrapperPath(string prefix)
        {
            return System.IO.Path.Combine(WrapperDir(), prefix + (_wrapCounter++).ToString("D4") + ".pdf");
        }

        /// <summary>
        /// Убрать папки обёрток, брошенные прошлыми сеансами. Сутки — чтобы не задеть чужой
        /// работающий экземпляр: приложение и так одно на систему, но полагаться на это здесь
        /// не стоит, а лишние сутки на диске ничего не стоят.
        /// </summary>
        private static void SweepOldWrapperDirs()
        {
            try
            {
                foreach (string dir in System.IO.Directory.GetDirectories(System.IO.Path.GetTempPath(),
                             WrapperPrefix + "*"))
                    if (System.IO.Directory.GetLastWriteTimeUtc(dir) < DateTime.UtcNow.AddDays(-1))
                        try { System.IO.Directory.Delete(dir, true); } catch { } // занята — уйдёт в следующий раз
            }
            catch { } // нет доступа к папке временных файлов — не повод отказывать во вставке
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
            Ui.RunWorker(delegate()
            {
                string error = null;
                try
                {
                    double width, height;
                    BlankPages.SheetSize(SizeOf(neighbour), out width, out height);
                    BlankPages.WriteSheet(wrapper, width, height);
                }
                catch (Exception ex) { error = ex.Message; }
                string err = error;
                OnUi(delegate
                {
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

        /// <summary>Папка обёрток живёт ровно столько, сколько окно: её содержимое больше никому не нужно.</summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && _wrapDir != null)
            {
                try { System.IO.Directory.Delete(_wrapDir, true); } catch { } // занята — уберёт следующий сеанс
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
            if (errors != null)
                foreach (string err in errors)
                    Dialogs.Error(this, ToolTitle, Loc.T("common.fileNotAdded"), err);
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
