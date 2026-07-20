using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Разделение одного PDF: извлечение выбранных страниц в один файл и разбиение
    /// на несколько (по диапазонам, каждые N страниц, по закладкам). Страницы
    /// копируются как есть, без переконвертации (PDFsharp, MIT). Как и в
    /// <see cref="PdfMergeService"/>, публичные методы не содержат типов PdfSharp
    /// в телах: сначала EmbeddedAssemblies.Ensure(), затем [NoInlining]-ядро.
    /// </summary>
    public static class PdfSplitService
    {
        /// <summary>Извлечь выбранные страницы (индексы с нуля) в один новый PDF.</summary>
        public static void Extract(string sourcePath, IList<int> pageIndices, string outputPath)
        {
            if (pageIndices == null || pageIndices.Count == 0)
                throw new MergeException("Не выбрано ни одной страницы.");
            // Переиспользуем протестированное ядро объединения: набор страниц → один файл.
            var order = new List<PdfPageRef>();
            foreach (int idx in pageIndices)
                order.Add(new PdfPageRef { SourcePath = sourcePath, PageIndex = idx });
            PdfMergeService.Merge(order, outputPath);
        }

        /// <summary>Разбить по диапазонам («1-3, 5, 8-») — каждый диапазон в свой файл.</summary>
        public static List<string> SplitByRanges(string sourcePath, IList<PageRange> ranges, string outDir, string baseName)
        {
            if (ranges == null || ranges.Count == 0)
                throw new MergeException("Не задано ни одного диапазона.");
            EmbeddedAssemblies.Ensure();
            return SplitRangesCore(sourcePath, ranges, outDir, baseName);
        }

        /// <summary>Разбить на части по n страниц (n=1 — каждая страница отдельным файлом).</summary>
        public static List<string> SplitEveryN(string sourcePath, int n, string outDir, string baseName)
        {
            if (n < 1)
                throw new MergeException("Число страниц в части должно быть не меньше 1.");
            EmbeddedAssemblies.Ensure();
            return SplitEveryNCore(sourcePath, n, outDir, baseName);
        }

        /// <summary>Разбить по закладкам верхнего уровня — файлы именуются заголовками.</summary>
        public static List<string> SplitByBookmarks(string sourcePath, string outDir, string baseName)
        {
            EmbeddedAssemblies.Ensure();
            return SplitBookmarksCore(sourcePath, outDir, baseName);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<string> SplitRangesCore(string sourcePath, IList<PageRange> ranges, string outDir, string baseName)
        {
            var created = new List<string>();
            using (PdfDocument source = OpenSource(sourcePath))
            {
                foreach (PageRange r in ranges)
                {
                    if (r.Start < 0 || r.End >= source.PageCount)
                        throw new MergeException("Диапазон " + r.Label + " вне файла (страниц: " + source.PageCount + ").");
                    string path = UniquePath(outDir, baseName + "_" + r.Label);
                    WriteRange(source, r, path);
                    created.Add(path);
                }
            }
            return created;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<string> SplitEveryNCore(string sourcePath, int n, string outDir, string baseName)
        {
            var created = new List<string>();
            using (PdfDocument source = OpenSource(sourcePath))
            {
                List<PageRange> chunks = PageRanges.EveryN(source.PageCount, n);
                int part = 1;
                foreach (PageRange r in chunks)
                {
                    string path = UniquePath(outDir, baseName + "_часть_" + part);
                    WriteRange(source, r, path);
                    created.Add(path);
                    part++;
                }
            }
            return created;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<string> SplitBookmarksCore(string sourcePath, string outDir, string baseName)
        {
            var created = new List<string>();
            using (PdfDocument source = OpenSource(sourcePath))
            {
                // Индекс страницы по её объекту (ObjectID надёжнее ссылочного равенства обёрток).
                var indexByObject = new Dictionary<PdfObjectID, int>();
                for (int i = 0; i < source.PageCount; i++)
                    indexByObject[source.Pages[i].Reference.ObjectID] = i;

                var marks = new List<KeyValuePair<int, string>>();
                foreach (PdfOutline outline in source.Outlines)
                {
                    PdfPage dest = outline.DestinationPage;
                    int idx;
                    if (dest != null && dest.Reference != null &&
                        indexByObject.TryGetValue(dest.Reference.ObjectID, out idx))
                        marks.Add(new KeyValuePair<int, string>(idx, outline.Title));
                }
                if (marks.Count == 0)
                    throw new MergeException("В файле нет закладок верхнего уровня — этот режим не применим.");
                marks.Sort(delegate(KeyValuePair<int, string> a, KeyValuePair<int, string> b) { return a.Key.CompareTo(b.Key); });

                for (int m = 0; m < marks.Count; m++)
                {
                    int start = m == 0 ? 0 : marks[m].Key; // ведущие страницы — в первый файл
                    int end = (m + 1 < marks.Count ? marks[m + 1].Key : source.PageCount) - 1;
                    if (start > end)
                        continue; // две закладки на одной странице — пустой раздел пропускаем
                    string path = UniquePath(outDir, baseName + "_" + Sanitize(marks[m].Value));
                    WriteRange(source, new PageRange(start, end), path);
                    created.Add(path);
                }
            }
            return created;
        }

        private static void WriteRange(PdfDocument source, PageRange r, string path)
        {
            using (PdfDocument outDoc = new PdfDocument())
            {
                for (int i = r.Start; i <= r.End; i++)
                    outDoc.AddPage(source.Pages[i]);
                try
                {
                    outDoc.Save(path);
                }
                catch (Exception ex)
                {
                    throw new MergeException("Не удалось сохранить «" + Path.GetFileName(path) + "»: " + ex.Message);
                }
            }
        }

        private static PdfDocument OpenSource(string path)
        {
            try
            {
                return PdfReader.Open(path, PdfDocumentOpenMode.Import);
            }
            catch (Exception ex)
            {
                throw new MergeException("Не удалось открыть «" + Path.GetFileName(path) +
                    "»: файл повреждён, защищён паролем или использует неподдерживаемые возможности PDF. (" + ex.Message + ")");
            }
        }

        /// <summary>Путь в папке с уникальным именем (без перезаписи): name.pdf, name_2.pdf, …</summary>
        private static string UniquePath(string dir, string baseName)
        {
            string safe = Sanitize(baseName);
            string path = Path.Combine(dir, safe + ".pdf");
            int n = 2;
            while (File.Exists(path))
            {
                path = Path.Combine(dir, safe + "_" + n + ".pdf");
                n++;
            }
            return path;
        }

        /// <summary>Замена недопустимых для имени файла символов на «_». Чистая — под тест.</summary>
        internal static string Sanitize(string name)
        {
            if (string.IsNullOrEmpty(name))
                return "без_имени";
            char[] invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);
            foreach (char c in name)
                sb.Append(Array.IndexOf(invalid, c) >= 0 ? '_' : c);
            string s = sb.ToString().Trim();
            return s.Length == 0 ? "без_имени" : s;
        }
    }
}
