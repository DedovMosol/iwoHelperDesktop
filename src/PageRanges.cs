using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>Диапазон страниц (индексы с нуля, включительно).</summary>
    public struct PageRange
    {
        public int Start;
        public int End;

        public PageRange(int start, int end) { Start = start; End = end; }

        public int Count { get { return End - Start + 1; } }

        /// <summary>Человекочитаемая метка (1-базная): «5» или «1-3» — для имён файлов.</summary>
        public string Label
        {
            get { return Start == End ? (Start + 1).ToString() : (Start + 1) + "-" + (End + 1); }
        }
    }

    /// <summary>
    /// Разбор пользовательских диапазонов страниц («1-3, 5, 8-») и нарезка на
    /// равные части. Страницы 1-базные на входе, 0-базные внутри. Чистые функции —
    /// покрыты юнит-тестами.
    /// </summary>
    public static class PageRanges
    {
        /// <summary>
        /// «1-3, 5, 8-» при pageCount=10 → [0-2], [4-4], [7-9]. Каждый элемент —
        /// отдельный будущий файл, порядок сохраняется. Ошибка ввода — MergeException.
        /// </summary>
        public static List<PageRange> Parse(string spec, int pageCount)
        {
            var result = new List<PageRange>();
            if (spec == null || spec.Trim().Length == 0)
                throw new MergeException("Укажите диапазоны страниц, например: 1-3, 5, 8-");

            foreach (string raw in spec.Split(','))
            {
                string token = raw.Trim();
                if (token.Length == 0)
                    continue;

                int start, end;
                int dash = token.IndexOf('-');
                if (dash < 0)
                {
                    start = end = ParseNumber(token, token);
                }
                else
                {
                    string left = token.Substring(0, dash).Trim();
                    string right = token.Substring(dash + 1).Trim();
                    start = left.Length == 0 ? 1 : ParseNumber(left, token);
                    end = right.Length == 0 ? pageCount : ParseNumber(right, token);
                }

                if (start < 1 || end > pageCount || start > end)
                    throw new MergeException("Диапазон «" + token + "» вне 1–" + pageCount + ".");
                result.Add(new PageRange(start - 1, end - 1));
            }

            if (result.Count == 0)
                throw new MergeException("Не задано ни одного диапазона.");
            return result;
        }

        /// <summary>Нарезка pageCount страниц на части по n: 10 при n=3 → [0-2],[3-5],[6-8],[9-9].</summary>
        public static List<PageRange> EveryN(int pageCount, int n)
        {
            if (n < 1)
                throw new MergeException("Число страниц в части должно быть не меньше 1.");
            var result = new List<PageRange>();
            for (int start = 0; start < pageCount; start += n)
            {
                int end = start + n - 1;
                if (end >= pageCount)
                    end = pageCount - 1;
                result.Add(new PageRange(start, end));
            }
            return result;
        }

        private static int ParseNumber(string s, string token)
        {
            int value;
            if (!int.TryParse(s, out value))
                throw new MergeException("Не понял номер страницы в «" + token + "».");
            return value;
        }
    }
}
