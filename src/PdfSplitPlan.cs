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
        /// </summary>
        public static List<PdfSplitPart> ByRanges(IList<PdfPageRef> pages, IList<PageRange> ranges)
        {
            var result = new List<PdfSplitPart>();
            if (pages == null || ranges == null)
                return result;
            foreach (PageRange range in ranges)
            {
                var part = new List<PdfPageRef>();
                foreach (PdfPageRef page in pages)
                    if (page != null && page.PageIndex >= range.Start && page.PageIndex <= range.End)
                        part.Add(page);
                if (part.Count > 0)
                    result.Add(new PdfSplitPart(part, range.Label));
            }
            return result;
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
            var result = new List<PdfSplitPart>();
            if (pages == null || marks == null || marks.Count == 0 || sourcePageCount <= 0)
                return result;
            var ordered = new List<PdfSplitBookmark>(marks);
            ordered.Sort(delegate(PdfSplitBookmark a, PdfSplitBookmark b)
            {
                return a.PageIndex.CompareTo(b.PageIndex);
            });
            for (int i = 0; i < ordered.Count; i++)
            {
                PdfSplitBookmark mark = ordered[i];
                int start = i == 0 ? 0 : mark.PageIndex;
                int end = i + 1 < ordered.Count ? ordered[i + 1].PageIndex - 1 : sourcePageCount - 1;
                if (start < 0) start = 0;
                if (end >= sourcePageCount) end = sourcePageCount - 1;
                var part = new List<PdfPageRef>();
                foreach (PdfPageRef page in pages)
                    if (page != null && page.PageIndex >= start && page.PageIndex <= end)
                        part.Add(page);
                if (part.Count > 0)
                {
                    string title = mark.Title ?? "";
                    result.Add(new PdfSplitPart(part, title, title, mark.PageIndex));
                }
            }
            return result;
        }
    }
}
