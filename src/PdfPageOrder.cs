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
        private readonly List<PdfPageRef> _items = new List<PdfPageRef>();

        public int Count
        {
            get { return _items.Count; }
        }

        public PdfPageRef this[int index]
        {
            get { return _items[index]; }
        }

        /// <summary>Очистить список (например, при открытии другого документа).</summary>
        public void Clear()
        {
            _items.Clear();
        }

        /// <summary>Все страницы документа добавляются в конец списка по порядку.</summary>
        public void AddDocument(string path, int pageCount)
        {
            for (int i = 0; i < pageCount; i++)
            {
                var page = new PdfPageRef();
                page.SourcePath = path;
                page.PageIndex = i;
                _items.Add(page);
            }
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
