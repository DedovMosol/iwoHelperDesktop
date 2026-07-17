using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Порядок страниц будущего PDF: добавление документов, перестановка
    /// и удаление. Чистая модель без UI и PDFsharp — покрыта юнит-тестами.
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

        /// <summary>Обмен с соседом сверху; возвращает новый индекс строки.</summary>
        public int MoveUp(int index)
        {
            if (index <= 0 || index >= _items.Count)
                return index;
            Swap(index, index - 1);
            return index - 1;
        }

        /// <summary>Обмен с соседом снизу; возвращает новый индекс строки.</summary>
        public int MoveDown(int index)
        {
            if (index < 0 || index >= _items.Count - 1)
                return index;
            Swap(index, index + 1);
            return index + 1;
        }

        /// <summary>Перенос строки (drag&amp;drop): вставить from ПЕРЕД текущей позицией to.</summary>
        public void Move(int from, int to)
        {
            if (from < 0 || from >= _items.Count || to < 0 || to > _items.Count || from == to)
                return;
            PdfPageRef item = _items[from];
            _items.RemoveAt(from);
            if (to > from)
                to--; // после изъятия строки цель сместилась
            _items.Insert(to, item);
        }

        /// <summary>Удаление набора строк по индексам (в любом порядке).</summary>
        public void RemoveAt(IList<int> indices)
        {
            var sorted = new List<int>(indices);
            sorted.Sort();
            for (int i = sorted.Count - 1; i >= 0; i--)
            {
                int index = sorted[i];
                if (index >= 0 && index < _items.Count)
                    _items.RemoveAt(index);
            }
        }

        public List<PdfPageRef> ToList()
        {
            return new List<PdfPageRef>(_items);
        }

        private void Swap(int a, int b)
        {
            PdfPageRef tmp = _items[a];
            _items[a] = _items[b];
            _items[b] = tmp;
        }
    }
}
