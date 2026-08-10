using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Генераторы page range строк для типовых сценариев: нечётные, чётные, каждая N-я.
    /// Чистые функции — покрыты юнит-тестами (v1.18.4).
    /// </summary>
    public static class PageRangePresets
    {
        /// <summary>
        /// Нечётные страницы: "1,3,5,7,..." до pageCount.
        /// </summary>
        public static string Odd(int pageCount)
        {
            if (pageCount < 1) return "";
            var parts = new List<string>();
            for (int i = 1; i <= pageCount; i += 2)
                parts.Add(i.ToString());
            return string.Join(",", parts.ToArray());
        }

        /// <summary>
        /// Чётные страницы: "2,4,6,8,..." до pageCount. Пустая строка, если чётных страниц нет.
        /// </summary>
        public static string Even(int pageCount)
        {
            if (pageCount < 2) return "";
            var parts = new List<string>();
            for (int i = 2; i <= pageCount; i += 2)
                parts.Add(i.ToString());
            return string.Join(",", parts.ToArray());
        }

        /// <summary>
        /// Каждая N-я страница начиная с первой: step=2 → "1,3,5,...", step=3 → "1,4,7,...".
        /// </summary>
        public static string EveryNth(int pageCount, int step)
        {
            if (pageCount < 1 || step < 1) return "";
            var parts = new List<string>();
            for (int i = 1; i <= pageCount; i += step)
                parts.Add(i.ToString());
            return string.Join(",", parts.ToArray());
        }

        /// <summary>
        /// Все страницы: "1-{pageCount}".
        /// </summary>
        public static string All(int pageCount)
        {
            if (pageCount < 1) return "";
            return pageCount == 1 ? "1" : "1-" + pageCount;
        }

        /// <summary>Длина подписи, после которой перечень страниц сокращается многоточием.</summary>
        internal const int MaxLabelLength = 40;

        /// <summary>
        /// Подпись для выпадающего списка: длинный перечень («1,3,5,…» на 500 страниц — это
        /// 1250 символов) сокращается до начала и конца. Полная строка живёт отдельно и в
        /// разбор идёт именно она. Чистая функция — под тест.
        /// </summary>
        public static string Shorten(string rangeSpec)
        {
            if (rangeSpec == null || rangeSpec.Length <= MaxLabelLength)
                return rangeSpec ?? "";
            // Режем по границам номеров, а не посреди числа: иначе подпись показывает «49» от «499».
            const int headLen = 15, tailLen = 12;
            int head = rangeSpec.LastIndexOf(',', headLen);
            int tail = rangeSpec.IndexOf(',', rangeSpec.Length - tailLen);
            if (head <= 0 || tail < 0 || tail <= head)
                return rangeSpec.Substring(0, MaxLabelLength - 1) + "…";
            return rangeSpec.Substring(0, head) + ",…" + rangeSpec.Substring(tail);
        }
    }
}
