using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Порядок страниц будущего PDF: добавление документов, перестановка
    /// и удаление. Чистая модель без UI и PDFsharp — покрыта юнит-тестами.
    /// Перестановка делегируется общему <see cref="ListReorder"/>.
    /// </summary>
    public class PdfPageOrder
    {
        private const int UndoLimit = 50; // хватает на сессию правок и не копит память

        private readonly List<PdfPageRef> _items = new List<PdfPageRef>();
        // Стек снимков ПОРЯДКА для Ctrl+Z: копии списка ссылок. Ссылки общие с сеткой,
        // поэтому откатывается только порядок/состав — повороты страниц не трогаются.
        private readonly List<List<PdfPageRef>> _undo = new List<List<PdfPageRef>>();

        public int Count
        {
            get { return _items.Count; }
        }

        public PdfPageRef this[int index]
        {
            get { return _items[index]; }
        }

        /// <summary>Есть что откатывать (для доступности Ctrl+Z).</summary>
        public bool CanUndo
        {
            get { return _undo.Count > 0; }
        }

        /// <summary>
        /// Снимок порядка ПЕРЕД пользовательским жестом (один жест — один снимок:
        /// форма зовёт перед добавлением, переносом, удалением, вставкой). Стек ограничен.
        /// </summary>
        public void Checkpoint()
        {
            _undo.Add(new List<PdfPageRef>(_items));
            if (_undo.Count > UndoLimit)
                _undo.RemoveAt(0);
        }

        /// <summary>Откатить последний жест (Ctrl+Z). false — откатывать нечего.</summary>
        public bool Undo()
        {
            if (_undo.Count == 0)
                return false;
            List<PdfPageRef> snapshot = _undo[_undo.Count - 1];
            _undo.RemoveAt(_undo.Count - 1);
            _items.Clear();
            _items.AddRange(snapshot);
            return true;
        }

        /// <summary>Очистить список (например, при открытии другого документа). История отката тоже очищается.</summary>
        public void Clear()
        {
            _items.Clear();
            _undo.Clear();
        }

        /// <summary>Все страницы документа добавляются в конец списка по порядку.</summary>
        public void AddDocument(string path, int pageCount)
        {
            InsertDocument(Count, path, pageCount);
        }

        /// <summary>
        /// Страницы документа вставляются подряд ПЕРЕД позицией insertAt (за пределами
        /// списка — прижимается к краю). Возвращает позицию первой вставленной страницы.
        /// </summary>
        public int InsertDocument(int insertAt, string path, int pageCount)
        {
            if (insertAt < 0) insertAt = 0;
            if (insertAt > _items.Count) insertAt = _items.Count;
            for (int i = 0; i < pageCount; i++)
            {
                var page = new PdfPageRef();
                page.SourcePath = path;
                page.PageIndex = i;
                _items.Insert(insertAt + i, page);
            }
            return insertAt;
        }

        /// <summary>
        /// Вставка готовых страниц (вставка из буфера) ПЕРЕД позицией insertAt.
        /// Возвращает позицию первой вставленной страницы.
        /// </summary>
        public int InsertAt(int insertAt, IList<PdfPageRef> pages)
        {
            if (insertAt < 0) insertAt = 0;
            if (insertAt > _items.Count) insertAt = _items.Count;
            if (pages != null)
                for (int i = 0; i < pages.Count; i++)
                    _items.Insert(insertAt + i, pages[i]);
            return insertAt;
        }

        /// <summary>
        /// Перенос набора страниц (вырезать → вставить) ПЕРЕД позицией insertAt в исходной
        /// нумерации. Возвращает позицию первой перенесённой страницы (-1 — пустой набор).
        /// </summary>
        public int MoveRange(IList<int> indices, int insertAt)
        {
            return ListReorder.MoveRange(_items, indices, insertAt);
        }

        public int MoveUp(int index)
        {
            return ListReorder.MoveUp(_items, index);
        }

        public int MoveDown(int index)
        {
            return ListReorder.MoveDown(_items, index);
        }

        public void Move(int from, int to)
        {
            ListReorder.Move(_items, from, to);
        }

        public void RemoveAt(IList<int> indices)
        {
            ListReorder.RemoveAt(_items, indices);
        }

        public List<PdfPageRef> ToList()
        {
            return new List<PdfPageRef>(_items);
        }
    }
}
