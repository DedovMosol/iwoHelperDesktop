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

        /// <summary>
        /// Перенос НАБОРА элементов (вырезать → вставить): элементы с индексами indices
        /// (в любом порядке, дубли игнорируются) изымаются и вставляются подряд ПЕРЕД
        /// позицией insertAt, посчитанной в исходной нумерации. Возвращает индекс первого
        /// перенесённого элемента после операции (для выделения); при пустом наборе — -1.
        /// </summary>
        public static int MoveRange<T>(IList<T> items, IList<int> indices, int insertAt)
        {
            List<int> sorted = NormalizeIndices(indices, items.Count);
            if (sorted.Count == 0)
                return -1;
            if (insertAt < 0) insertAt = 0;
            if (insertAt > items.Count) insertAt = items.Count;

            int adjusted = AdjustedInsertIndex(sorted, insertAt);
            var taken = new List<T>(sorted.Count);
            foreach (int index in sorted)
                taken.Add(items[index]);
            for (int i = sorted.Count - 1; i >= 0; i--)
                items.RemoveAt(sorted[i]);
            for (int i = 0; i < taken.Count; i++)
                items.Insert(adjusted + i, taken[i]);
            return adjusted;
        }

        /// <summary>
        /// Позиция вставки после изъятия элементов: каждый изъятый ЛЕВЕЕ цели сдвигает
        /// её на один влево. sortedIndices — по возрастанию. Чистая — под тест.
        /// </summary>
        internal static int AdjustedInsertIndex(IList<int> sortedIndices, int insertAt)
        {
            int before = 0;
            foreach (int index in sortedIndices)
                if (index < insertAt)
                    before++;
                else
                    break;
            return insertAt - before;
        }

        /// <summary>Индексы по возрастанию, без дублей и выходов за [0, count). Чистая — под тест.</summary>
        internal static List<int> NormalizeIndices(IList<int> indices, int count)
        {
            var result = new List<int>();
            if (indices == null)
                return result;
            var seen = new HashSet<int>();
            foreach (int index in indices)
                if (index >= 0 && index < count && seen.Add(index))
                    result.Add(index);
            result.Sort();
            return result;
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
