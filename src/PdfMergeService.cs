using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Одна страница одного исходного PDF (индекс с нуля). Rotation — дополнительный
    /// поворот по часовой (0/90/180/270), назначенный пользователем в сетке миниатюр;
    /// при записи складывается с собственным /Rotate исходной страницы.
    /// Каждая позиция списка страниц — ОТДЕЛЬНЫЙ экземпляр (инвариант моделей):
    /// копирование страниц идёт через <see cref="Clone"/>, а не общие ссылки.
    /// </summary>
    public class PdfPageRef
    {
        public string SourcePath;
        public int PageIndex;
        public int Rotation;

        public string FileName
        {
            get { return Path.GetFileName(SourcePath); }
        }

        public PdfPageRef Clone()
        {
            return new PdfPageRef { SourcePath = SourcePath, PageIndex = PageIndex, Rotation = Rotation };
        }

        /// <summary>Сумма поворотов, нормализованная в {0, 90, 180, 270} (отрицательные допустимы). Чистая — под тест.</summary>
        internal static int ComposeRotation(int a, int b)
        {
            int sum = ((a + b) % 360 + 360) % 360;
            return sum - sum % 90; // не кратный 90 мусор из чужого PDF срезаем вниз до кратного
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
                throw new MergeException(Loc.T("err.pdf.noPages"));
            string lockError = MergeService.CheckOutputWritable(outputPath);
            if (lockError != null)
                throw new MergeException(Loc.T("err.pdf.fileBusy"));
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
                        throw new MergeException(string.Format(Loc.T("err.pdf.noPagesIn"), Path.GetFileName(path)));
                    return pages;
                }
            }
            catch (MergeException)
            {
                throw;
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                throw new MergeException(string.Format(Loc.T("err.pdf.cantOpen"), Path.GetFileName(path), ex.Message));
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
                        catch (Exception ex) when (MergeException.ShouldWrap(ex))
                        {
                            throw new MergeException(string.Format(Loc.T("err.pdf.cantOpenShort"), page.FileName, ex.Message));
                        }
                        sources.Add(key, source);
                    }
                    if (page.PageIndex < 0 || page.PageIndex >= source.PageCount)
                        throw new MergeException(string.Format(Loc.T("err.pdf.pageGone"), page.FileName, page.PageIndex + 1));
                    PdfPage copied = output.AddPage(source.Pages[page.PageIndex]);
                    if (page.Rotation != 0) // поворот пользователя поверх собственного /Rotate страницы
                        copied.Rotate = PdfPageRef.ComposeRotation(copied.Rotate, page.Rotation);
                    added++;
                    if (progress != null)
                        progress(added, order.Count);
                }

                try
                {
                    output.Save(outputPath);
                }
                catch (Exception ex) when (MergeException.ShouldWrap(ex))
                {
                    throw new MergeException(string.Format(Loc.T("err.pdf.saveFailed"), ex.Message));
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
