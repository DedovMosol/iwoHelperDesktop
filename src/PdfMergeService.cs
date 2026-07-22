using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ExcelMerger
{
    /// <summary>Одна страница одного исходного PDF (индекс с нуля).</summary>
    public class PdfPageRef
    {
        public string SourcePath;
        public int PageIndex;

        public string FileName
        {
            get { return Path.GetFileName(SourcePath); }
        }
    }

    /// <summary>Размеры страницы в пунктах — для проверок и подписей.</summary>
    public class PdfPageInfo
    {
        public int PageIndex;
        public double WidthPt;
        public double HeightPt;
    }

    /// <summary>
    /// Слияние PDF-документов с произвольным порядком страниц (PDFsharp, MIT).
    /// Страницы копируются как есть, без переконвертации — сканы, печати
    /// и подписи не искажаются. Публичные методы не содержат типов PdfSharp
    /// в телах: сначала EmbeddedAssemblies.Ensure(), затем [NoInlining]-ядро.
    /// </summary>
    public static class PdfMergeService
    {
        /// <summary>Страницы документа с размерами. Битый/зашифрованный файл — MergeException.</summary>
        public static List<PdfPageInfo> LoadPages(string path)
        {
            EmbeddedAssemblies.Ensure();
            return LoadPagesCore(path);
        }

        /// <summary>Собирает страницы в порядке order в новый PDF. Пустой порядок — ошибка.</summary>
        public static void Merge(IList<PdfPageRef> order, string outputPath, Action<int, int> progress = null)
        {
            if (order == null || order.Count == 0)
                throw new MergeException("Нет страниц для объединения.");
            string lockError = MergeService.CheckOutputWritable(outputPath);
            if (lockError != null)
                throw new MergeException(lockError.Replace("Итоговый файл", "Файл PDF"));
            EmbeddedAssemblies.Ensure();
            MergeCore(order, outputPath, progress);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<PdfPageInfo> LoadPagesCore(string path)
        {
            try
            {
                using (PdfDocument doc = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                {
                    var pages = new List<PdfPageInfo>();
                    for (int i = 0; i < doc.PageCount; i++)
                    {
                        PdfPage page = doc.Pages[i];
                        var info = new PdfPageInfo();
                        info.PageIndex = i;
                        info.WidthPt = page.Width.Point;
                        info.HeightPt = page.Height.Point;
                        pages.Add(info);
                    }
                    if (pages.Count == 0)
                        throw new MergeException("В файле «" + Path.GetFileName(path) + "» нет страниц.");
                    return pages;
                }
            }
            catch (MergeException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new MergeException("Не удалось открыть «" + Path.GetFileName(path) +
                    "»: файл повреждён, защищён паролем или использует неподдерживаемые возможности PDF. (" +
                    ex.Message + ")");
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void MergeCore(IList<PdfPageRef> order, string outputPath, Action<int, int> progress)
        {
            // Каждый источник открывается один раз, сколько бы страниц из него ни брали.
            var sources = new Dictionary<string, PdfDocument>(StringComparer.OrdinalIgnoreCase);
            PdfDocument output = null;
            try
            {
                output = new PdfDocument();
                int added = 0;
                foreach (PdfPageRef page in order)
                {
                    PdfDocument source;
                    string key = Path.GetFullPath(page.SourcePath);
                    if (!sources.TryGetValue(key, out source))
                    {
                        try
                        {
                            source = PdfReader.Open(key, PdfDocumentOpenMode.Import);
                        }
                        catch (Exception ex)
                        {
                            throw new MergeException("Не удалось открыть «" + page.FileName + "»: " + ex.Message);
                        }
                        sources.Add(key, source);
                    }
                    if (page.PageIndex < 0 || page.PageIndex >= source.PageCount)
                        throw new MergeException("В «" + page.FileName + "» нет страницы " + (page.PageIndex + 1) +
                            " — файл изменился после добавления в список.");
                    output.AddPage(source.Pages[page.PageIndex]);
                    added++;
                    if (progress != null)
                        progress(added, order.Count);
                }

                try
                {
                    output.Save(outputPath);
                }
                catch (Exception ex)
                {
                    throw new MergeException("Не удалось сохранить PDF: " + ex.Message);
                }
            }
            finally
            {
                if (output != null)
                    try { output.Dispose(); } catch { }
                foreach (PdfDocument doc in sources.Values)
                    try { doc.Dispose(); } catch { }
            }
        }
    }
}
