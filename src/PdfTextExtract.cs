using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

namespace ExcelMerger
{
    /// <summary>Растровое изображение страницы (PNG) с рамкой в координатах PDF (ось Y вверх).</summary>
    public class OcrImage
    {
        public byte[] Png;
        public double TopPt;    // верх изображения — для порядка чтения вместе с абзацами
        public double LeftPt;
        public double WidthPt;
        public double HeightPt;
    }

    /// <summary>Текст одной страницы born-digital PDF (абзацы в порядке чтения).</summary>
    public class PdfPageText
    {
        public int PageIndex;                              // с нуля
        public List<OcrParagraph> Paragraphs = new List<OcrParagraph>();
        public List<OcrImage> Images = new List<OcrImage>();
        public List<OcrTable> Tables = new List<OcrTable>(); // восстановленные таблицы (порядок чтения — по TopPt)
        internal List<PdfLine> Lines = new List<PdfLine>(); // линовка страницы (границы таблиц, подчёркивания)
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
        public static List<PdfPageText> Extract(string path, Action<int, int> progress = null)
        {
            EmbeddedAssemblies.Ensure();
            return ExtractCore(path, progress);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<PdfPageText> ExtractCore(string path, Action<int, int> progress)
        {
            try
            {
                using (UglyToad.PdfPig.PdfDocument doc = UglyToad.PdfPig.PdfDocument.Open(path))
                {
                    var pages = new List<PdfPageText>();
                    int pageCount = doc.NumberOfPages;
                    foreach (UglyToad.PdfPig.Content.Page page in doc.GetPages())
                    {
                        var words = new List<PdfWord>();
                        foreach (UglyToad.PdfPig.Content.Word w in page.GetWords())
                        {
                            UglyToad.PdfPig.Core.PdfRectangle bb = w.BoundingBox;
                            double size = 0;
                            bool bold = false, italic = false;
                            int color = 0;
                            string family = null;
                            if (w.Letters != null && w.Letters.Count > 0)
                            {
                                UglyToad.PdfPig.Content.Letter first = w.Letters[0];
                                size = first.PointSize;
                                string fn = first.FontName ?? "";
                                bold = fn.IndexOf("Bold", StringComparison.OrdinalIgnoreCase) >= 0;
                                italic = fn.IndexOf("Italic", StringComparison.OrdinalIgnoreCase) >= 0
                                      || fn.IndexOf("Oblique", StringComparison.OrdinalIgnoreCase) >= 0;
                                color = ColorArgb(first.Color);
                                family = FontNames.Clean(fn);
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
                                ColorArgb = color,
                                FontName = family
                            });
                        }
                        AssignHyperlinks(words, page);
                        List<PdfLine> lines = ExtractLines(page);
                        // Слова таблиц уходят в ячейки; абзацы строятся из ОСТАВШИХСЯ (внетабличных) слов.
                        TableDetectResult det = TableDetector.Detect(lines, words, page.Width, page.Height);
                        UnderlineDetector.Mark(det.RemainingWords, lines); // подчёркивания по линовке под словами
                        OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(det.RemainingWords);
                        var pt = new PdfPageText
                        {
                            PageIndex = page.Number - 1, // PdfPig нумерует страницы с 1
                            Paragraphs = layout.Paragraphs,
                            FirstLineIndentPt = layout.FirstLineIndentPt,
                            WidthPt = page.Width,
                            HeightPt = page.Height,
                            Lines = lines,
                            Tables = det.Tables
                        };
                        SetMargins(pt, words, page.Width, page.Height);
                        pt.Images = ExtractImages(page);
                        pages.Add(pt);
                        if (progress != null)
                            progress(pages.Count, pageCount);
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

        /// <summary>
        /// Пометить слова гиперссылками: если центр слова попадает в рамку ссылки страницы —
        /// проставить её URI. Сбой получения ссылок не срывает извлечение. Вызывать из ядра.
        /// </summary>
        private static void AssignHyperlinks(List<PdfWord> words, UglyToad.PdfPig.Content.Page page)
        {
            System.Collections.Generic.IReadOnlyList<UglyToad.PdfPig.Content.Hyperlink> links;
            try { links = page.GetHyperlinks(); }
            catch { return; }
            if (links == null || links.Count == 0)
                return;
            foreach (PdfWord w in words)
            {
                double mx = (w.Left + w.Right) / 2, my = (w.Top + w.Bottom) / 2;
                foreach (UglyToad.PdfPig.Content.Hyperlink h in links)
                {
                    if (string.IsNullOrEmpty(h.Uri))
                        continue;
                    UglyToad.PdfPig.Core.PdfRectangle b = h.Bounds;
                    if (mx >= b.Left && mx <= b.Right && my >= b.Bottom && my <= b.Top)
                    {
                        w.Uri = h.Uri;
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Растровые изображения страницы в PNG с рамками. Недекодируемые форматы (JBIG2,
        /// CCITT, CMYK-JPEG и т.п.) пропускаются; сбой одной картинки не срывает остальные и
        /// текст. Вызывать из ядра (после Ensure) — тело ссылается на типы PdfPig.
        /// </summary>
        private static List<OcrImage> ExtractImages(UglyToad.PdfPig.Content.Page page)
        {
            var result = new List<OcrImage>();
            IEnumerable<UglyToad.PdfPig.Content.IPdfImage> images;
            try { images = page.GetImages(); }
            catch { return result; }
            foreach (UglyToad.PdfPig.Content.IPdfImage img in images)
            {
                try
                {
                    byte[] png = EncodePng(img);
                    if (png == null || png.Length == 0)
                        continue;
                    if (IsSolidColor(png))
                        continue; // одноцветный растр — обычно сбой декодера (напр. штрих-код чёрным ящиком); пропускаем
                    UglyToad.PdfPig.Core.PdfRectangle b = img.Bounds;
                    if (b.Width <= 0 || b.Height <= 0)
                        continue;
                    result.Add(new OcrImage
                    {
                        Png = png,
                        LeftPt = b.Left,
                        TopPt = b.Top,
                        WidthPt = b.Width,
                        HeightPt = b.Height
                    });
                }
                catch { } // одна битая картинка не ломает остальные
            }
            return result;
        }

        /// <summary>
        /// PNG для изображения PDF: сначала штатный TryGetPng; если он не смог (частый случай —
        /// DCTDecode/JPEG, для которого PdfPig не строит PNG), декодируем сырые байты потока через
        /// GDI (System.Drawing понимает JPEG/PNG/GIF/BMP) и пересохраняем в PNG. Возвращает null,
        /// если ни один путь не дал картинку (напр. сырьё — недекодируемый формат).
        /// </summary>
        private static byte[] EncodePng(UglyToad.PdfPig.Content.IPdfImage img)
        {
            try
            {
                byte[] direct;
                if (img.TryGetPng(out direct) && direct != null && direct.Length > 0)
                    return direct;
            }
            catch { } // TryGetPng может кинуть на экзотическом формате — уходим в фолбэк

            try
            {
                byte[] raw = img.RawBytes.ToArray(); // ReadOnlyMemory<byte> -> массив (напр. JPEG)
                if (raw.Length == 0)
                    return null;
                using (var ms = new MemoryStream(raw))
                using (System.Drawing.Image bmp = System.Drawing.Image.FromStream(ms))
                using (var outMs = new MemoryStream())
                {
                    bmp.Save(outMs, System.Drawing.Imaging.ImageFormat.Png);
                    return outMs.ToArray();
                }
            }
            catch { return null; } // сырьё не декодируется в растр — пропускаем картинку
        }

        /// <summary>
        /// Одноцветный ли растр (все опрошенные пиксели одного цвета). Признак сбоя декодера
        /// PDF-картинки (напр. штрих-код выходит сплошным чёрным прямоугольником) — такое
        /// изображение бессмысленно, его лучше не переносить. Сбой чтения — не считаем одноцветным.
        /// </summary>
        private static bool IsSolidColor(byte[] png)
        {
            try
            {
                using (var ms = new MemoryStream(png))
                using (var bmp = new System.Drawing.Bitmap(ms))
                {
                    int w = bmp.Width, h = bmp.Height;
                    if (w < 2 || h < 2)
                        return false;
                    int stepX = Math.Max(1, w / 16), stepY = Math.Max(1, h / 16);
                    int first = bmp.GetPixel(0, 0).ToArgb();
                    for (int y = 0; y < h; y += stepY)
                        for (int x = 0; x < w; x += stepX)
                            if (bmp.GetPixel(x, y).ToArgb() != first)
                                return false; // нашёлся другой цвет — не одноцветный
                    return true;
                }
            }
            catch { return false; } // не декодировали — не считаем одноцветным (пусть проходит)
        }

        // Классификация отрезка как линовки: отклонение от оси (pt) и минимальная длина.
        private const double LineAxisTol = 1.0; // |dy| для горизонтали / |dx| для вертикали
        private const double MinLineLen = 2.0;  // короче — графический шум, не линовка
        private const double ThinFillMax = 2.5; // залитая полоска тоньше этого — это линия, не фон

        /// <summary>
        /// Линовка страницы (границы таблиц, подчёркивания) из векторной графики. Берём только
        /// строго горизонтальные/вертикальные штрихи и рёбра прямоугольников; диагонали, кривые
        /// и крупные заливки (фон ячеек) пропускаем. Сбой одной фигуры не срывает остальные и
        /// текст. Вызывать из ядра (после Ensure) — тело ссылается на типы PdfPig.
        /// </summary>
        private static List<PdfLine> ExtractLines(UglyToad.PdfPig.Content.Page page)
        {
            var result = new List<PdfLine>();
            System.Collections.Generic.IReadOnlyList<UglyToad.PdfPig.Graphics.PdfPath> paths;
            try { paths = page.ExperimentalAccess.Paths; }
            catch { return result; }
            if (paths == null)
                return result;
            foreach (UglyToad.PdfPig.Graphics.PdfPath path in paths)
            {
                try
                {
                    if (path == null || path.IsClipping)
                        continue; // невидимая обтравка — не ink
                    bool stroked = path.IsStroked;
                    bool filled = path.IsFilled;
                    if (!stroked && !filled)
                        continue;
                    double thickness = path.LineWidth;
                    for (int i = 0; i < path.Count; i++)
                    {
                        UglyToad.PdfPig.Core.PdfSubpath sub = path[i];
                        if (sub == null)
                            continue;
                        if (stroked && sub.IsDrawnAsRectangle)
                            AddRectangleEdges(result, sub, thickness);
                        else if (stroked)
                            AddLineCommands(result, sub, thickness);
                        else if (sub.IsDrawnAsRectangle) // filled, не stroked
                            AddThinFilledRect(result, sub);
                    }
                }
                catch { } // одна битая фигура не ломает остальные
            }
            return result;
        }

        /// <summary>Рёбра нарисованного прямоугольника → 4 линии (границы ячейки/рамки).</summary>
        private static void AddRectangleEdges(List<PdfLine> result, UglyToad.PdfPig.Core.PdfSubpath sub, double thickness)
        {
            UglyToad.PdfPig.Core.PdfRectangle? r = sub.GetDrawnRectangle();
            if (r == null)
                return;
            UglyToad.PdfPig.Core.PdfRectangle rect = r.Value;
            AddSegment(result, rect.Left, rect.Top, rect.Right, rect.Top, thickness);
            AddSegment(result, rect.Left, rect.Bottom, rect.Right, rect.Bottom, thickness);
            AddSegment(result, rect.Left, rect.Bottom, rect.Left, rect.Top, thickness);
            AddSegment(result, rect.Right, rect.Bottom, rect.Right, rect.Top, thickness);
        }

        /// <summary>Прямые команды Line подпути → линовка (диагонали отсеются в AddSegment).</summary>
        private static void AddLineCommands(List<PdfLine> result, UglyToad.PdfPig.Core.PdfSubpath sub, double thickness)
        {
            System.Collections.Generic.IReadOnlyList<UglyToad.PdfPig.Core.PdfSubpath.IPathCommand> cmds = sub.Commands;
            if (cmds == null)
                return;
            foreach (UglyToad.PdfPig.Core.PdfSubpath.IPathCommand c in cmds)
            {
                var line = c as UglyToad.PdfPig.Core.PdfSubpath.Line;
                if (line != null)
                    AddSegment(result, line.From.X, line.From.Y, line.To.X, line.To.Y, thickness);
            }
        }

        /// <summary>Тонкая залитая полоска (линия, нарисованная заливкой) → одна линия по длинной оси.</summary>
        private static void AddThinFilledRect(List<PdfLine> result, UglyToad.PdfPig.Core.PdfSubpath sub)
        {
            UglyToad.PdfPig.Core.PdfRectangle? r = sub.GetDrawnRectangle();
            if (r == null)
                return;
            UglyToad.PdfPig.Core.PdfRectangle rect = r.Value;
            double w = rect.Width, h = rect.Height;
            if (Math.Min(w, h) > ThinFillMax)
                return; // крупная заливка — фон ячейки, а не линия
            double midX = (rect.Left + rect.Right) / 2, midY = (rect.Bottom + rect.Top) / 2;
            if (w >= h)
                AddSegment(result, rect.Left, midY, rect.Right, midY, Math.Min(w, h));
            else
                AddSegment(result, midX, rect.Bottom, midX, rect.Top, Math.Min(w, h));
        }

        /// <summary>Классифицировать отрезок и добавить, если это строго H/V-линовка достаточной длины.</summary>
        private static void AddSegment(List<PdfLine> result, double x1, double y1, double x2, double y2, double thickness)
        {
            double dx = Math.Abs(x2 - x1), dy = Math.Abs(y2 - y1);
            if (dy <= LineAxisTol && dx > dy && dx >= MinLineLen)
                result.Add(new PdfLine { Orientation = LineOrientation.Horizontal, X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Thickness = thickness });
            else if (dx <= LineAxisTol && dy > dx && dy >= MinLineLen)
                result.Add(new PdfLine { Orientation = LineOrientation.Vertical, X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Thickness = thickness });
            // иначе диагональ/кривая — не линовка
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
