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
        // Во всех методах rotations — дополнительные повороты страниц ПО ИНДЕКСУ ИСХОДНОЙ
        // страницы (0/90/180/270 по часовой; null или короче документа — без поворота).
        // Единая конвенция на все режимы, чтобы форма собирала одну карту.

        /// <summary>Извлечь выбранные страницы (индексы с нуля) в один новый PDF. progress — «сделано/всего» частей (здесь одна).</summary>
        public static void Extract(string sourcePath, IList<int> pageIndices, string outputPath, Action<int, int> progress = null, IList<int> rotations = null)
        {
            if (pageIndices == null || pageIndices.Count == 0)
                throw new MergeException(Loc.T("err.split.noPages"));
            // Переиспользуем протестированное ядро объединения: набор страниц → один файл.
            var order = new List<PdfPageRef>();
            foreach (int idx in pageIndices)
                order.Add(new PdfPageRef { SourcePath = sourcePath, PageIndex = idx, Rotation = RotationAt(rotations, idx) });
            PdfMergeService.Merge(order, outputPath);
            if (progress != null)
                progress(1, 1); // одна часть (один итоговый файл)
        }

        /// <summary>Разбить по диапазонам («1-3, 5, 8-») — каждый диапазон в свой файл. progress — «сделано/всего» частей.</summary>
        public static List<string> SplitByRanges(string sourcePath, IList<PageRange> ranges, string outDir, string baseName, Action<int, int> progress = null, IList<int> rotations = null)
        {
            if (ranges == null || ranges.Count == 0)
                throw new MergeException(Loc.T("err.split.noRanges"));
            EmbeddedAssemblies.Ensure();
            return SplitRangesCore(sourcePath, ranges, outDir, baseName, progress, rotations);
        }

        /// <summary>Разбить на части по n страниц (n=1 — каждая страница отдельным файлом). progress — «сделано/всего» частей.</summary>
        public static List<string> SplitEveryN(string sourcePath, int n, string outDir, string baseName, Action<int, int> progress = null, IList<int> rotations = null)
        {
            if (n < 1)
                throw new MergeException(Loc.T("err.split.badN"));
            EmbeddedAssemblies.Ensure();
            return SplitEveryNCore(sourcePath, n, outDir, baseName, progress, rotations);
        }

        /// <summary>Разбить по закладкам верхнего уровня — файлы именуются заголовками. progress — «сделано/всего» частей.</summary>
        public static List<string> SplitByBookmarks(string sourcePath, string outDir, string baseName, Action<int, int> progress = null, IList<int> rotations = null)
        {
            EmbeddedAssemblies.Ensure();
            return SplitBookmarksCore(sourcePath, outDir, baseName, progress, rotations);
        }

        /// <summary>Поворот страницы pageIndex из карты (вне карты — 0). Чистая — под тест.</summary>
        internal static int RotationAt(IList<int> rotations, int pageIndex)
        {
            return rotations != null && pageIndex >= 0 && pageIndex < rotations.Count ? rotations[pageIndex] : 0;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<string> SplitRangesCore(string sourcePath, IList<PageRange> ranges, string outDir, string baseName, Action<int, int> progress, IList<int> rotations)
        {
            var created = new List<string>();
            using (PdfDocument source = OpenSource(sourcePath))
            {
                foreach (PageRange r in ranges)
                {
                    if (r.Start < 0 || r.End >= source.PageCount)
                        throw new MergeException(string.Format(Loc.T("err.split.rangeOutside"), r.Label, source.PageCount));
                    string path = UniquePath(outDir, baseName + "_" + r.Label);
                    WriteRange(source, r, path, rotations);
                    created.Add(path);
                    if (progress != null)
                        progress(created.Count, ranges.Count);
                }
            }
            return created;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<string> SplitEveryNCore(string sourcePath, int n, string outDir, string baseName, Action<int, int> progress, IList<int> rotations)
        {
            var created = new List<string>();
            using (PdfDocument source = OpenSource(sourcePath))
            {
                List<PageRange> chunks = PageRanges.EveryN(source.PageCount, n);
                int part = 1;
                foreach (PageRange r in chunks)
                {
                    string path = UniquePath(outDir, baseName + Loc.T("split.partInfix") + part);
                    WriteRange(source, r, path, rotations);
                    created.Add(path);
                    part++;
                    if (progress != null)
                        progress(created.Count, chunks.Count);
                }
            }
            return created;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<string> SplitBookmarksCore(string sourcePath, string outDir, string baseName, Action<int, int> progress, IList<int> rotations)
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
                    throw new MergeException(Loc.T("err.split.noBookmarks"));
                marks.Sort(delegate(KeyValuePair<int, string> a, KeyValuePair<int, string> b) { return a.Key.CompareTo(b.Key); });

                for (int m = 0; m < marks.Count; m++)
                {
                    int start = m == 0 ? 0 : marks[m].Key; // ведущие страницы — в первый файл
                    int end = (m + 1 < marks.Count ? marks[m + 1].Key : source.PageCount) - 1;
                    if (start > end)
                        continue; // две закладки на одной странице — пустой раздел пропускаем
                    string path = UniquePath(outDir, baseName + "_" + Sanitize(marks[m].Value));
                    WriteRange(source, new PageRange(start, end), path, rotations);
                    created.Add(path);
                    if (progress != null)
                        progress(created.Count, marks.Count);
                }
            }
            return created;
        }

        private static void WriteRange(PdfDocument source, PageRange r, string path, IList<int> rotations)
        {
            using (PdfDocument outDoc = new PdfDocument())
            {
                for (int i = r.Start; i <= r.End; i++)
                {
                    PdfPage copied = outDoc.AddPage(source.Pages[i]);
                    int extra = RotationAt(rotations, i);
                    if (extra != 0) // поворот пользователя поверх собственного /Rotate страницы
                        copied.Rotate = PdfPageRef.ComposeRotation(copied.Rotate, extra);
                }
                try
                {
                    outDoc.Save(path);
                }
                catch (Exception ex) when (MergeException.ShouldWrap(ex))
                {
                    throw new MergeException(string.Format(Loc.T("err.split.saveFailed"), Path.GetFileName(path), ex.Message));
                }
            }
        }

        private static PdfDocument OpenSource(string path)
        {
            try
            {
                return PdfReader.Open(path, PdfDocumentOpenMode.Import);
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                throw new MergeException(string.Format(Loc.T("err.pdf.cantOpen"), Path.GetFileName(path), ex.Message));
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
                return Loc.T("split.unnamed");
            char[] invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);
            foreach (char c in name)
                sb.Append(Array.IndexOf(invalid, c) >= 0 ? '_' : c);
            string s = sb.ToString().Trim();
            return s.Length == 0 ? Loc.T("split.unnamed") : s;
        }
    }
}
