using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Одна часть результата разделения: страницы берутся из рабочего порядка,
    /// а Label/BookmarkPageIndex сохраняют исходную семантику имени и оглавления.
    /// Чистая модель без PDFsharp — покрыта тестами.
    /// </summary>
    internal sealed class PdfSplitPart
    {
        public readonly List<PdfPageRef> Pages;
        public readonly string Label;
        public readonly string BookmarkTitle;
        public readonly int BookmarkPageIndex;
        public readonly bool NumberedPart;

        public PdfSplitPart(IList<PdfPageRef> pages, string label)
            : this(pages, label, null, -1, false) { }

        public PdfSplitPart(IList<PdfPageRef> pages, string label, bool numberedPart)
            : this(pages, label, null, -1, numberedPart) { }

        public PdfSplitPart(IList<PdfPageRef> pages, string label,
            string bookmarkTitle, int bookmarkPageIndex, bool numberedPart = false)
        {
            Pages = new List<PdfPageRef>();
            if (pages != null)
                foreach (PdfPageRef page in pages)
                    if (page != null)
                        Pages.Add(page.Clone());
            Label = label ?? "";
            BookmarkTitle = bookmarkTitle;
            BookmarkPageIndex = bookmarkPageIndex;
            NumberedPart = numberedPart;
        }
    }

    /// <summary>Плоская граница верхнеуровневой закладки, без типов PDF-библиотеки.</summary>
    internal sealed class PdfSplitBookmark
    {
        public int PageIndex;
        public string Title;
    }

    /// <summary>
    /// Чистое планирование Split поверх рабочего списка PdfPageRef.
    /// Grid-позиции здесь никогда не считаются номерами страниц источника.
    /// </summary>
    internal static class PdfSplitPlan
    {
        public static List<PdfPageRef> ClonePages(IList<PdfPageRef> pages)
        {
            var result = new List<PdfPageRef>();
            if (pages == null)
                return result;
            foreach (PdfPageRef page in pages)
                if (page != null)
                    result.Add(page.Clone());
            return result;
        }

        /// <summary>Выбранные позиции рабочего списка → страницы источника в текущем порядке.</summary>
        public static List<PdfPageRef> Selected(IList<PdfPageRef> pages, IList<int> positions)
        {
            var result = new List<PdfPageRef>();
            if (pages == null || positions == null)
                return result;
            var seen = new HashSet<int>();
            foreach (int position in positions)
            {
                if (position < 0 || position >= pages.Count || !seen.Add(position))
                    continue;
                if (pages[position] != null)
                    result.Add(pages[position].Clone());
            }
            return result;
        }

        /// <summary>
        /// Каждый диапазон выбирает исходные номера, но результат сохраняет текущий порядок
        /// рабочего списка. Дубликаты страницы в рабочем списке намеренно сохраняются.
        /// Источник разделения выводится из набора (<see cref="InferSource"/>).
        /// </summary>
        public static List<PdfSplitPart> ByRanges(IList<PdfPageRef> pages, IList<PageRange> ranges)
        {
            return ByRanges(pages, ranges, InferSource(pages));
        }

        /// <summary>
        /// Диапазоны по страницам источника sourcePath. Страница ДРУГОГО файла (вставленный
        /// пустой лист: его PageIndex — 0 в своей обёртке) номера источника не имеет, поэтому
        /// идёт вслед за предыдущей страницей источника, а стоящая в самом начале — в первую
        /// часть. Прежде такой лист по своему нулю уезжал в часть со страницей 1.
        /// </summary>
        public static List<PdfSplitPart> ByRanges(IList<PdfPageRef> pages, IList<PageRange> ranges,
            string sourcePath)
        {
            var bounds = new List<KeyValuePair<int, int>>();
            if (ranges != null)
                foreach (PageRange range in ranges)
                    bounds.Add(new KeyValuePair<int, int>(range.Start, range.End));
            List<List<PdfPageRef>> parts = AssignByIntervals(pages, bounds, sourcePath);
            var result = new List<PdfSplitPart>();
            for (int i = 0; i < parts.Count; i++)
                if (parts[i].Count > 0)
                    result.Add(new PdfSplitPart(parts[i], ranges[i].Label));
            return result;
        }

        /// <summary>
        /// Разложить рабочий порядок по интервалам исходных номеров [Key..Value]. Страница
        /// источника попадает во все интервалы, куда входит её номер (диапазоны могут
        /// перекрываться). Чужая страница наследует интервалы ближайшей предыдущей страницы
        /// источника — лист, вставленный «после стр. 5», живёт там же, где стр. 5; ведущие
        /// чужие страницы встают в начало первой непустой части. Порядок внутри части — рабочий.
        /// </summary>
        private static List<List<PdfPageRef>> AssignByIntervals(IList<PdfPageRef> pages,
            IList<KeyValuePair<int, int>> intervals, string sourcePath)
        {
            var parts = new List<List<PdfPageRef>>();
            foreach (KeyValuePair<int, int> unused in intervals)
                parts.Add(new List<PdfPageRef>());
            if (pages == null || intervals.Count == 0)
                return parts;
            var leading = new List<PdfPageRef>();
            List<int> anchor = null; // интервалы последней страницы источника; null — её ещё не было
            foreach (PdfPageRef page in pages)
            {
                if (page == null)
                    continue;
                if (IsFromSource(page, sourcePath))
                {
                    anchor = new List<int>();
                    for (int i = 0; i < intervals.Count; i++)
                        if (page.PageIndex >= intervals[i].Key && page.PageIndex <= intervals[i].Value)
                            anchor.Add(i);
                }
                else if (anchor == null)
                {
                    leading.Add(page);
                    continue;
                }
                foreach (int i in anchor)
                    parts[i].Add(page);
            }
            if (leading.Count > 0)
                foreach (List<PdfPageRef> part in parts)
                    if (part.Count > 0)
                    {
                        part.InsertRange(0, leading);
                        break;
                    }
            return parts;
        }

        /// <summary>Страница источника разделения? null-источник — каждая страница своя (прежнее поведение).</summary>
        private static bool IsFromSource(PdfPageRef page, string sourcePath)
        {
            return sourcePath == null ||
                string.Equals(page.SourcePath, sourcePath, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Источник разделения, когда вызывающий его не назвал: файл с наибольшим числом
        /// РАЗНЫХ номеров страниц (при равенстве — НЕ обёртка, затем — с большим числом
        /// вхождений, затем — первый встреченный). Вставленный лист — обёртка из одной страницы
        /// и не перевесит документ, даже если его скопировали несколько раз; а когда от
        /// документа осталась одна страница, ничья решается в пользу документа — иначе его
        /// страница уезжала бы в чужую главу. Набор из одного файла — он и есть источник.
        /// Чистая — под тест.
        /// </summary>
        internal static string InferSource(IList<PdfPageRef> pages)
        {
            if (pages == null)
                return null;
            var order = new List<string>();
            var distinct = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
            var hits = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (PdfPageRef page in pages)
            {
                if (page == null || page.SourcePath == null)
                    continue;
                HashSet<int> indexes;
                if (!distinct.TryGetValue(page.SourcePath, out indexes))
                {
                    indexes = new HashSet<int>();
                    distinct.Add(page.SourcePath, indexes);
                    hits.Add(page.SourcePath, 0);
                    order.Add(page.SourcePath);
                }
                indexes.Add(page.PageIndex);
                hits[page.SourcePath]++;
            }
            string best = null;
            string root = System.IO.Path.GetTempPath(); // папки обёрток лежат прямо в ней
            foreach (string path in order)
            {
                if (best == null)
                {
                    best = path;
                    continue;
                }
                int cmp = distinct[path].Count.CompareTo(distinct[best].Count);
                if (cmp == 0) // поровну страниц: документ важнее обёртки (false < true)
                    cmp = PdfPageOrderFormBase.IsWrapperPath(best, root)
                        .CompareTo(PdfPageOrderFormBase.IsWrapperPath(path, root));
                if (cmp == 0)
                    cmp = hits[path].CompareTo(hits[best]);
                if (cmp > 0)
                    best = path; // при полной ничьей остаётся первый встреченный
            }
            return best;
        }

        /// <summary>Нарезать текущий рабочий порядок на пачки по n оставшихся страниц.</summary>
        public static List<PdfSplitPart> EveryN(IList<PdfPageRef> pages, int n)
        {
            if (n < 1)
                throw new MergeException(Loc.T("err.split.badN"));
            var result = new List<PdfSplitPart>();
            if (pages == null)
                return result;
            int part = 1;
            for (int start = 0; start < pages.Count; start += n)
            {
                var chunk = new List<PdfPageRef>();
                int end = Math.Min(pages.Count, start + n);
                for (int i = start; i < end; i++)
                    if (pages[i] != null)
                        chunk.Add(pages[i]);
                if (chunk.Count > 0)
                    result.Add(new PdfSplitPart(chunk, part.ToString(), true));
                part++;
            }
            return result;
        }

        /// <summary>
        /// Верхнеуровневые закладки задают интервалы исходных страниц; внутри интервала
        /// сохраняется пользовательский рабочий порядок. Пустые интервалы пропускаются.
        /// </summary>
        public static List<PdfSplitPart> ByBookmarks(IList<PdfPageRef> pages,
            IList<PdfSplitBookmark> marks, int sourcePageCount)
        {
            // Источник не передан (так зовёт сервис разделения) — выводим его из набора.
            return ByBookmarks(pages, marks, sourcePageCount, InferSource(pages));
        }

        /// <summary>
        /// Главы по страницам источника sourcePath; чужие страницы (вставленный пустой лист)
        /// идут за предыдущей страницей источника, ведущие — в первую главу (см. ByRanges).
        /// </summary>
        public static List<PdfSplitPart> ByBookmarks(IList<PdfPageRef> pages,
            IList<PdfSplitBookmark> marks, int sourcePageCount, string sourcePath)
        {
            var result = new List<PdfSplitPart>();
            if (pages == null || marks == null || marks.Count == 0 || sourcePageCount <= 0)
                return result;
            var ordered = new List<PdfSplitBookmark>(marks);
            ordered.Sort(delegate(PdfSplitBookmark a, PdfSplitBookmark b)
            {
                return a.PageIndex.CompareTo(b.PageIndex);
            });
            var intervals = new List<KeyValuePair<int, int>>();
            for (int i = 0; i < ordered.Count; i++)
            {
                int start = i == 0 ? 0 : ordered[i].PageIndex;
                int end = i + 1 < ordered.Count ? ordered[i + 1].PageIndex - 1 : sourcePageCount - 1;
                if (start < 0) start = 0;
                if (end >= sourcePageCount) end = sourcePageCount - 1;
                intervals.Add(new KeyValuePair<int, int>(start, end));
            }
            List<List<PdfPageRef>> parts = AssignByIntervals(pages, intervals, sourcePath);
            for (int i = 0; i < parts.Count; i++)
                if (parts[i].Count > 0)
                {
                    string title = ordered[i].Title ?? "";
                    result.Add(new PdfSplitPart(parts[i], title, title, ordered[i].PageIndex));
                }
            return result;
        }
    }
}
