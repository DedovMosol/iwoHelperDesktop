using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Разделение одного PDF: извлечение выбранных страниц в один файл и разбиение
    /// на несколько (по диапазонам, каждые N страниц, по закладкам). Страницы
    /// копируются как есть, без переконвертации (PDFsharp, MIT). Рабочий порядок
    /// приходит как PdfPageRef — поэтому удаление, перестановка и повороты не
    /// путаются с исходными номерами страниц.
    /// </summary>
    public static class PdfSplitService
    {
        // Во всех методах rotations — дополнительные повороты страниц ПО ИНДЕКСУ ИСХОДНОЙ
        // страницы (0/90/180/270 по часовой; null или короче документа — без поворота).
        // Единая конвенция остаётся у совместимых source-based overloads.

        /// <summary>Извлечь выбранные исходные страницы в один новый PDF.</summary>
        public static void Extract(string sourcePath, IList<int> pageIndices, string outputPath,
            Action<int, int> progress = null, IList<int> rotations = null, Func<bool> cancelled = null)
        {
            if (pageIndices == null || pageIndices.Count == 0)
                throw new MergeException(Loc.T("err.split.noPages"));
            var pages = new List<PdfPageRef>();
            foreach (int idx in pageIndices)
            {
                if (idx < 0)
                    throw new MergeException(Loc.T("err.split.noPages"));
                pages.Add(new PdfPageRef
                {
                    SourcePath = sourcePath,
                    PageIndex = idx,
                    Rotation = RotationAt(rotations, idx)
                });
            }
            ExtractPlanned(sourcePath, pages, outputPath, progress, cancelled);
        }

        /// <summary>
        /// Внутренний путь для формы: pages уже являются снимком рабочего порядка.
        /// </summary>
        internal static void ExtractPlanned(string sourcePath, IList<PdfPageRef> pages,
            string outputPath, Action<int, int> progress = null, Func<bool> cancelled = null)
        {
            if (pages == null || pages.Count == 0)
                throw new MergeException(Loc.T("err.split.noPages"));
            if (OutputFile.IsSameFile(sourcePath, outputPath))
                throw new MergeException(Loc.T("err.output.sameSource"));
            EmbeddedAssemblies.Ensure();
            ExtractPlannedCore(sourcePath, pages, outputPath, progress, cancelled);
        }

        /// <summary>Разбить по диапазонам («1-3, 5, 8-») — каждый retained диапазон в свой файл.</summary>
        public static List<string> SplitByRanges(string sourcePath, IList<PageRange> ranges,
            string outDir, string baseName, Action<int, int> progress = null,
            IList<int> rotations = null, Func<bool> cancelled = null, string template = null)
        {
            if (ranges == null || ranges.Count == 0)
                throw new MergeException(Loc.T("err.split.noRanges"));
            List<PdfPageRef> pages = PristinePages(sourcePath, rotations);
            return SplitPlanned(sourcePath, PdfSplitPlan.ByRanges(pages, ranges),
                outDir, baseName, progress, cancelled, template);
        }

        /// <summary>Разбить текущий порядок на пачки по n страниц.</summary>
        public static List<string> SplitEveryN(string sourcePath, int n, string outDir,
            string baseName, Action<int, int> progress = null, IList<int> rotations = null,
            Func<bool> cancelled = null, string template = null)
        {
            if (n < 1)
                throw new MergeException(Loc.T("err.split.badN"));
            List<PdfPageRef> pages = PristinePages(sourcePath, rotations);
            return SplitPlanned(sourcePath, PdfSplitPlan.EveryN(pages, n),
                outDir, baseName, progress, cancelled, template);
        }

        /// <summary>Разбить по закладкам верхнего уровня.</summary>
        public static List<string> SplitByBookmarks(string sourcePath, string outDir,
            string baseName, Action<int, int> progress = null, IList<int> rotations = null,
            Func<bool> cancelled = null, string template = null)
        {
            EmbeddedAssemblies.Ensure();
            return SplitBookmarksCore(sourcePath, outDir, baseName, progress, rotations, cancelled, template);
        }

        /// <summary>
        /// Внутренний путь формы: закладки определяют принадлежность страниц, а рабочий
        /// порядок определяет порядок страниц внутри частей.
        /// </summary>
        internal static List<string> SplitByBookmarksPlanned(string sourcePath,
            IList<PdfPageRef> working, string outDir, string baseName,
            Action<int, int> progress = null, Func<bool> cancelled = null, string template = null)
        {
            if (working == null || working.Count == 0)
                throw new MergeException(Loc.T("err.split.noPages"));
            EmbeddedAssemblies.Ensure();
            return SplitBookmarksPlannedCore(sourcePath, working, outDir, baseName,
                progress, cancelled, template);
        }

        /// <summary>
        /// Внутренний общий путь для range/every-N планов. Пустой plan — неуспех,
        /// а не зелёный результат с нулём файлов.
        /// </summary>
        internal static List<string> SplitPlanned(string sourcePath, IList<PdfSplitPart> parts,
            string outDir, string baseName, Action<int, int> progress = null,
            Func<bool> cancelled = null, string template = null)
        {
            if (parts == null || parts.Count == 0)
                throw new MergeException(Loc.T("err.split.noPages"));
            EmbeddedAssemblies.Ensure();
            return SplitPlannedCore(sourcePath, parts, outDir, baseName,
                progress, cancelled, template);
        }

        /// <summary>Поворот страницы pageIndex из старой карты.</summary>
        internal static int RotationAt(IList<int> rotations, int pageIndex)
        {
            return PageRotation.At(rotations, pageIndex);
        }

        private static List<PdfPageRef> PristinePages(string sourcePath, IList<int> rotations)
        {
            List<PdfPageInfo> info = PdfMergeService.LoadPages(sourcePath);
            var pages = new List<PdfPageRef>(info.Count);
            foreach (PdfPageInfo page in info)
                pages.Add(new PdfPageRef
                {
                    SourcePath = sourcePath,
                    PageIndex = page.PageIndex,
                    Rotation = RotationAt(rotations, page.PageIndex)
                });
            return pages;
        }

        private static List<PdfSplitPart> BookmarkParts(IList<PdfPageRef> pages,
            IList<PdfSplitBookmark> marks, int pageCount)
        {
            List<PdfSplitPart> result = PdfSplitPlan.ByBookmarks(pages, marks, pageCount);
            if (result.Count == 0)
                throw new MergeException(Loc.T("err.split.noPages"));
            return result;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ExtractPlannedCore(string sourcePath, IList<PdfPageRef> pages,
            string outputPath, Action<int, int> progress, Func<bool> cancelled)
        {
            PdfMergeService.Merge(pages, outputPath, null, cancelled);
            if (progress != null)
                progress(1, 1);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<string> SplitPlannedCore(string sourcePath,
            IList<PdfSplitPart> parts, string outDir, string baseName,
            Action<int, int> progress, Func<bool> cancelled, string template)
        {
            return WriteParts(sourcePath, parts, outDir, baseName,
                progress, cancelled, template);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<string> SplitBookmarksCore(string sourcePath, string outDir,
            string baseName, Action<int, int> progress, IList<int> rotations,
            Func<bool> cancelled, string template)
        {
            using (PdfDocument source = OpenSource(sourcePath))
            {
                List<PdfSplitBookmark> marks = ReadTopLevelBookmarks(source);
                if (marks.Count == 0)
                    throw new MergeException(Loc.T("err.split.noBookmarks"));
                var pages = new List<PdfPageRef>(source.PageCount);
                for (int i = 0; i < source.PageCount; i++)
                    pages.Add(new PdfPageRef
                    {
                        SourcePath = sourcePath,
                        PageIndex = i,
                        Rotation = RotationAt(rotations, i)
                    });
                return WriteParts(sourcePath,
                    BookmarkParts(pages, marks, source.PageCount), outDir, baseName,
                    progress, cancelled, template);
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<string> SplitBookmarksPlannedCore(string sourcePath,
            IList<PdfPageRef> working, string outDir, string baseName,
            Action<int, int> progress, Func<bool> cancelled, string template)
        {
            using (PdfDocument source = OpenSource(sourcePath))
            {
                List<PdfSplitBookmark> marks = ReadTopLevelBookmarks(source);
                if (marks.Count == 0)
                    throw new MergeException(Loc.T("err.split.noBookmarks"));
                return WriteParts(sourcePath,
                    BookmarkParts(working, marks, source.PageCount), outDir, baseName,
                    progress, cancelled, template);
            }
        }

        private static List<string> WriteParts(string sourcePath,
            IList<PdfSplitPart> parts, string outDir, string baseName,
            Action<int, int> progress, Func<bool> cancelled, string template)
        {
            if (parts == null || parts.Count == 0)
                throw new MergeException(Loc.T("err.split.noPages"));
            Directory.CreateDirectory(outDir);
            DateTime startedAt = DateTime.Now;
            return Cancellation.NoPartialOutput(delegate(List<string> created)
            {
                for (int i = 0; i < parts.Count; i++)
                {
                    Cancellation.ThrowIf(cancelled);
                    PdfSplitPart part = parts[i];
                    string legacy = LegacyPartName(baseName, part, i + 1);
                    string safeName = Sanitize(PartName(template, legacy,
                        ValuesFor(part, baseName, i + 1, parts.Count, startedAt)));
                    string path = WritePart(sourcePath, part, outDir, safeName, cancelled);
                    created.Add(path);
                    if (progress != null)
                        progress(created.Count, parts.Count);
                }
                if (created.Count == 0)
                    throw new MergeException(Loc.T("err.split.noPages"));
            });
        }

        private static string WritePart(string sourcePath, PdfSplitPart part,
            string directory, string safeName, Func<bool> cancelled)
        {
            if (part == null || part.Pages.Count == 0)
                throw new MergeException(Loc.T("err.split.noPages"));
            using (var output = OutputFile.CreateUnique(directory, safeName, ".pdf"))
            {
                PdfMergeService.WriteUnpublished(part.Pages, output.TempPath, null, cancelled);
                EnsurePartBookmark(output.TempPath, sourcePath, part);
                Cancellation.ThrowIf(cancelled);
                return output.Commit();
            }
        }

        /// <summary>
        /// Если страница-цель верхнеуровневой закладки была удалена, сам Merge по правилам
        /// оглавления её отбросит. Для части возвращаем заголовок на следующую уцелевшую
        /// исходную страницу (или на первую страницу части, когда следующей нет).
        /// </summary>
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void EnsurePartBookmark(string path, string sourcePath, PdfSplitPart part)
        {
            if (part == null || string.IsNullOrEmpty(part.BookmarkTitle) || part.Pages.Count == 0)
                return;
            string sourceFull = Path.GetFullPath(sourcePath);
            foreach (PdfPageRef page in part.Pages)
                if (page != null && page.PageIndex == part.BookmarkPageIndex &&
                    string.Equals(Path.GetFullPath(page.SourcePath), sourceFull,
                        StringComparison.OrdinalIgnoreCase))
                    return; // исходная закладка уже перенесена PdfMergeService

            int bestSource = int.MaxValue, target = -1;
            for (int i = 0; i < part.Pages.Count; i++)
            {
                PdfPageRef page = part.Pages[i];
                if (page == null || !string.Equals(Path.GetFullPath(page.SourcePath),
                    sourceFull, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (page.PageIndex >= part.BookmarkPageIndex && page.PageIndex < bestSource)
                {
                    bestSource = page.PageIndex;
                    target = i;
                }
            }
            if (target < 0) target = 0;
            using (PdfDocument doc = PdfReader.Open(path, PdfDocumentOpenMode.Modify))
            {
                doc.Outlines.Add(part.BookmarkTitle, doc.Pages[target], true);
                doc.Save(path);
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<PdfSplitBookmark> ReadTopLevelBookmarks(PdfDocument source)
        {
            var result = new List<PdfSplitBookmark>();
            var all = PdfBookmarks.Read(source);
            foreach (PdfBookmark mark in all)
                if (mark != null && mark.Level == 0)
                    result.Add(new PdfSplitBookmark { PageIndex = mark.PageIndex, Title = mark.Title });
            result.Sort(delegate(PdfSplitBookmark a, PdfSplitBookmark b)
            {
                return a.PageIndex.CompareTo(b.PageIndex);
            });
            return result;
        }

        private static string LegacyPartName(string baseName, PdfSplitPart part, int number)
        {
            if (!string.IsNullOrEmpty(part.BookmarkTitle))
                return baseName + "_" + Sanitize(part.BookmarkTitle);
            if (part.NumberedPart)
                return baseName + Loc.T("split.partInfix") + number;
            if (!string.IsNullOrEmpty(part.Label))
                return baseName + "_" + Sanitize(part.Label);
            return baseName + Loc.T("split.partInfix") + number;
        }

        private static NameValues ValuesFor(PdfSplitPart part, string baseName,
            int number, int total, DateTime startedAt)
        {
            int current = 1;
            if (part != null && part.Pages.Count > 0)
            {
                current = part.Pages[0].PageIndex + 1;
                foreach (PdfPageRef page in part.Pages)
                    if (page.PageIndex + 1 < current)
                        current = page.PageIndex + 1;
            }
            return new NameValues
            {
                BaseName = baseName,
                FileNumber = number,
                TotalFiles = total,
                CurrentPage = current,
                BookmarkName = part == null ? null : part.BookmarkTitle,
                Timestamp = startedAt
            };
        }

        internal static string PartName(string template, string legacyName, NameValues values)
        {
            return string.IsNullOrEmpty(template) ? legacyName : NameTemplate.Apply(template, values);
        }

        internal static string Sanitize(string name)
        {
            string s = NameTemplate.Sanitize(name);
            return s.Length == 0 ? Loc.T("split.unnamed") : s;
        }

        private static PdfDocument OpenSource(string path)
        {
            try
            {
                return PdfReader.Open(path, PdfPasswords.For(path), PdfDocumentOpenMode.Import);
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                throw new MergeException(string.Format(Loc.T("err.pdf.cantOpen"),
                    Path.GetFileName(path), ex.Message));
            }
        }
    }
}
