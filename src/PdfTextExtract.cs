using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

namespace ExcelMerger
{
    /// <summary>Текст одной страницы born-digital PDF (в порядке чтения).</summary>
    public class PdfPageText
    {
        public int PageIndex;      // с нуля
        public string Text;
        public double WidthPt;
        public double HeightPt;
    }

    /// <summary>
    /// Извлечение текстового слоя born-digital PDF (PdfPig, Apache 2.0) — без OCR.
    /// Публичные методы не содержат типов PdfPig в телах: сначала
    /// <see cref="EmbeddedAssemblies.Ensure"/>, затем [NoInlining]-ядро (как в
    /// <see cref="PdfMergeService"/>), иначе JIT падает на резолве вшитой сборки.
    /// </summary>
    public static class PdfTextExtract
    {
        /// <summary>
        /// Текст всех страниц. Битый/зашифрованный файл или запрет извлечения —
        /// <see cref="MergeException"/> с понятным сообщением.
        /// </summary>
        public static List<PdfPageText> Extract(string path)
        {
            EmbeddedAssemblies.Ensure();
            return ExtractCore(path);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<PdfPageText> ExtractCore(string path)
        {
            try
            {
                using (UglyToad.PdfPig.PdfDocument doc = UglyToad.PdfPig.PdfDocument.Open(path))
                {
                    var pages = new List<PdfPageText>();
                    foreach (UglyToad.PdfPig.Content.Page page in doc.GetPages())
                    {
                        pages.Add(new PdfPageText
                        {
                            PageIndex = page.Number - 1, // PdfPig нумерует страницы с 1
                            Text = page.Text ?? string.Empty,
                            WidthPt = page.Width,
                            HeightPt = page.Height
                        });
                    }
                    return pages;
                }
            }
            catch (Exception ex)
            {
                throw new MergeException("Не удалось извлечь текст из «" + Path.GetFileName(path) +
                    "»: файл повреждён, зашифрован или без прав на извлечение. (" + ex.Message + ")");
            }
        }
    }
}
