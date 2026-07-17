using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Операции перестановки/удаления над списком по индексам — общая логика
    /// для порядка страниц PDF и списка файлов свода. Чистая, покрыта тестами.
    /// </summary>
    public static class ListReorder
    {
        /// <summary>Обмен с соседом сверху; возвращает новый индекс элемента.</summary>
        public static int MoveUp<T>(IList<T> items, int index)
        {
            if (index <= 0 || index >= items.Count)
                return index;
            Swap(items, index, index - 1);
            return index - 1;
        }

        /// <summary>Обмен с соседом снизу; возвращает новый индекс элемента.</summary>
        public static int MoveDown<T>(IList<T> items, int index)
        {
            if (index < 0 || index >= items.Count - 1)
                return index;
            Swap(items, index, index + 1);
            return index + 1;
        }

        /// <summary>Перенос (drag&amp;drop): вставить элемент from ПЕРЕД позицией to.</summary>
        public static void Move<T>(IList<T> items, int from, int to)
        {
            if (from < 0 || from >= items.Count || to < 0 || to > items.Count || from == to)
                return;
            T item = items[from];
            items.RemoveAt(from);
            if (to > from)
                to--; // после изъятия элемента цель сместилась
            items.Insert(to, item);
        }

        /// <summary>Удаление набора элементов по индексам (в любом порядке).</summary>
        public static void RemoveAt<T>(IList<T> items, IList<int> indices)
        {
            var sorted = new List<int>(indices);
            sorted.Sort();
            for (int i = sorted.Count - 1; i >= 0; i--)
            {
                int index = sorted[i];
                if (index >= 0 && index < items.Count)
                    items.RemoveAt(index);
            }
        }

        private static void Swap<T>(IList<T> items, int a, int b)
        {
            T tmp = items[a];
            items[a] = items[b];
            items[b] = tmp;
        }
    }
}
