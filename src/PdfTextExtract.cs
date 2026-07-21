using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

namespace ExcelMerger
{
    /// <summary>Текст одной страницы born-digital PDF (абзацы в порядке чтения).</summary>
    public class PdfPageText
    {
        public int PageIndex;                              // с нуля
        public List<OcrParagraph> Paragraphs = new List<OcrParagraph>();
        public double FirstLineIndentPt;                   // отступ красной строки (pt); 0 — без отступов
        public double WidthPt;
        public double HeightPt;
        // Поля страницы из рамок текста (pt); 0 — на странице не было текста.
        public double LeftMarginPt;
        public double RightMarginPt;
        public double TopMarginPt;
        public double BottomMarginPt;

        /// <summary>Весь текст страницы: абзацы через пустую строку.</summary>
        public string Text
        {
            get
            {
                var parts = new List<string>(Paragraphs.Count);
                foreach (OcrParagraph p in Paragraphs)
                    parts.Add(p.Text);
                return string.Join("\n\n", parts);
            }
        }
    }

    /// <summary>
    /// Извлечение текстового слоя born-digital PDF (PdfPig, Apache 2.0) — без OCR.
    /// Слова с рамками собираются в порядок чтения (<see cref="OcrLayout"/>). Публичные
    /// методы не содержат типов PdfPig в телах: сначала
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
                        var words = new List<PdfWord>();
                        foreach (UglyToad.PdfPig.Content.Word w in page.GetWords())
                        {
                            UglyToad.PdfPig.Core.PdfRectangle bb = w.BoundingBox;
                            double size = 0;
                            bool bold = false, italic = false;
                            int color = 0;
                            if (w.Letters != null && w.Letters.Count > 0)
                            {
                                UglyToad.PdfPig.Content.Letter first = w.Letters[0];
                                size = first.PointSize;
                                string fn = first.FontName ?? "";
                                bold = fn.IndexOf("Bold", StringComparison.OrdinalIgnoreCase) >= 0;
                                italic = fn.IndexOf("Italic", StringComparison.OrdinalIgnoreCase) >= 0
                                      || fn.IndexOf("Oblique", StringComparison.OrdinalIgnoreCase) >= 0;
                                color = ColorArgb(first.Color);
                            }
                            words.Add(new PdfWord
                            {
                                Text = w.Text,
                                Left = bb.Left,
                                Right = bb.Right,
                                Bottom = bb.Bottom,
                                Top = bb.Top,
                                FontSizePt = size,
                                Bold = bold,
                                Italic = italic,
                                ColorArgb = color
                            });
                        }
                        OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
                        var pt = new PdfPageText
                        {
                            PageIndex = page.Number - 1, // PdfPig нумерует страницы с 1
                            Paragraphs = layout.Paragraphs,
                            FirstLineIndentPt = layout.FirstLineIndentPt,
                            WidthPt = page.Width,
                            HeightPt = page.Height
                        };
                        SetMargins(pt, words, page.Width, page.Height);
                        pages.Add(pt);
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

        /// <summary>Поля страницы из рамок текста (pt, ось Y вверх). Пустая страница — поля 0.</summary>
        private static void SetMargins(PdfPageText pt, List<PdfWord> words, double pageW, double pageH)
        {
            if (words.Count == 0)
                return;
            double minL = double.MaxValue, maxR = double.MinValue, minB = double.MaxValue, maxT = double.MinValue;
            foreach (PdfWord w in words)
            {
                if (w.Left < minL) minL = w.Left;
                if (w.Right > maxR) maxR = w.Right;
                if (w.Bottom < minB) minB = w.Bottom;
                if (w.Top > maxT) maxT = w.Top;
            }
            pt.LeftMarginPt = minL;
            pt.RightMarginPt = pageW - maxR;
            pt.TopMarginPt = pageH - maxT;
            pt.BottomMarginPt = minB;
        }

        /// <summary>Цвет буквы PdfPig → 0xRRGGBB; null/чёрный → 0.</summary>
        private static int ColorArgb(UglyToad.PdfPig.Graphics.Colors.IColor color)
        {
            if (color == null)
                return 0;
            var rgb = color.ToRGBValues();
            int r = Clamp255(rgb.r), g = Clamp255(rgb.g), b = Clamp255(rgb.b);
            return (r << 16) | (g << 8) | b;
        }

        private static int Clamp255(double v)
        {
            int n = (int)Math.Round(v * 255.0);
            return n < 0 ? 0 : (n > 255 ? 255 : n);
        }
    }
}
