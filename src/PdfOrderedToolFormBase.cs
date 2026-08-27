using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// База PDF-инструментов, собирающих результат из страниц НЕСКОЛЬКИХ файлов —
    /// «Объединение» и «PDF → Word». Добавляет к правке порядка
    /// (<see cref="PdfPageOrderFormBase"/>) свой слой: выбор файлов, фоновый разбор каждого
    /// и вставку его страниц в набор — раньше он был посимвольной копией в двух формах.
    /// Наследник строит свой интерфейс (кнопки, действие), создаёт <c>_grid</c> и вызывает
    /// <see cref="WireOrderGrid"/> и <see cref="WireFileDropAppend"/>, а различия сводит в
    /// <see cref="PdfToolFormBase.SyncControls"/> и <see cref="ToolTitle"/>.
    ///
    /// Разбор PDF при добавлении идёт в фоне (<see cref="PdfToolFormBase.BeginLoad"/>) —
    /// большой или сетевой файл не морозит окно.
    /// </summary>
    public abstract class PdfOrderedToolFormBase : PdfPageOrderFormBase, IFileAcceptor
    {
        protected PdfOrderedToolFormBase(Action showHub) : base(showHub) { }

        /// <summary>Кроме общей правки порядка — приём файлов, брошенных на сетку: вставить их страницы.</summary>
        protected override void WireOrderGrid()
        {
            base.WireOrderGrid();
            _grid.FilesDropped += delegate(string[] paths, int insertAt) { AddFiles(paths, insertAt); };
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
            LoadPass(toLoad, new List<LoadedDoc>(), new List<string>(), new List<string>(), at, insertMode);
        }

        /// <summary>
        /// Один заход разбора: прочитать пачку файлов в фоне, а защищённые отложить и спросить
        /// по ним пароль. Введённые пароли отправляют файл на следующий заход — так неверный
        /// пароль даёт ещё одну попытку, а отказ прекращает круг по этому файлу.
        /// Круг конечен: каждый заход либо добавляет файл, либо получает отказ по нему.
        /// </summary>
        private void LoadPass(string[] toLoad, List<LoadedDoc> loaded, List<string> errors,
            List<string> tried, int at, bool insertMode)
        {
            Ui.RunWorker(delegate()
            {
                var locked = new List<string>();
                bool cancelled = false;
                foreach (string path in toLoad)
                {
                    if (LoadCancellationRequested)
                    {
                        cancelled = true;
                        break;
                    }
                    // Ловим ШИРОКО (как операционные воркеры): один битый/занятый/аварийный файл
                    // (в т.ч. редкий OOM, который LoadPages НЕ оборачивает) не должен ронять
                    // фоновый поток — остальные файлы пакета всё равно добавляются.
                    try { loaded.Add(new LoadedDoc { Path = path, PageCount = PdfMergeService.LoadPages(path).Count }); }
                    catch (Exception ex)
                    {
                        // Защищённый файл — не ошибка, а вопрос: его отложим и спросим пароль.
                        if (PdfPasswords.LooksPasswordProtected(ex))
                            locked.Add(path);
                        else if (ex is MergeException)
                            errors.Add(ex.Message); // понятное локализованное сообщение
                        else
                            errors.Add(string.Format(Loc.T("err.pdf.cantOpen"), Path.GetFileName(path), ex.Message));
                    }
                }
                bool wasCancelled = cancelled || LoadCancellationRequested;
                OnUi(delegate
                {
                    if (wasCancelled || LoadCancellationRequested)
                    {
                        FinishCanceledLoad();
                        return; // пакет — один жест: частично разобранные файлы не публикуются
                    }
                    // Что бы ни случилось при опросе паролей, окно обязано выйти из состояния
                    // загрузки: иначе кнопки и сетка останутся заблокированными навсегда, и
                    // человеку придётся закрывать инструмент.
                    try
                    {
                        if (locked.Count > 0)
                        {
                            List<string> retry = AskPasswords(locked, tried, errors);
                            tried.AddRange(locked);
                            if (retry.Count > 0)
                            {
                                LoadPass(retry.ToArray(), loaded, errors, tried, at, insertMode);
                                return; // итог подведёт последний заход
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        errors.Add(ex.Message);
                    }
                    ApplyAdded(loaded, errors, at, insertMode);
                });
            });
        }

        /// <summary>Применить результат фонового разбора: вставить страницы, показать ошибки. UI-поток.</summary>
        private void ApplyAdded(List<LoadedDoc> loaded, List<string> errors, int at, bool insertMode)
        {
            EndLoad(); // снять загрузку до вставки: Working=false, статус освобождается
            InsertLoaded(loaded, errors, at, insertMode); // вставка — общая с «Прочими операциями»
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
    }
}
