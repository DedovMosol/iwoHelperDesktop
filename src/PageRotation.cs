using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Поворот геометрии страницы PDF (pt, ось Y вверх) на 0/90/180/270 по часовой —
    /// как страницу видит пользователь в сетке миниатюр. Применяется в «PDF → Word»
    /// ДО анализа макета: боковой текст становится горизонтальным, и весь конвейер
    /// (таблицы, абзацы, поля) работает в выправленном пространстве. Чистая математика
    /// без PdfPig — покрыта юнит-тестами.
    ///
    /// Маппинг точки (x, y) при повороте страницы W×H по часовой:
    ///   90°: (x, y) → (y, W − x), страница становится H×W
    ///  180°: (x, y) → (W − x, H − y)
    ///  270°: (x, y) → (H − y, x), страница становится H×W
    /// (проверка углами: нижний-левый угол при 90° уходит в верхний-левый).
    /// </summary>
    internal static class PageRotation
    {
        /// <summary>Поворот страницы pageIndex из карты (null/вне карты — 0). Чистая — под тест.</summary>
        public static int At(IList<int> rotations, int pageIndex)
        {
            return rotations != null && pageIndex >= 0 && pageIndex < rotations.Count ? rotations[pageIndex] : 0;
        }

        /// <summary>Меняет ли поворот местами ширину и высоту страницы.</summary>
        public static bool SwapsDimensions(int rotation)
        {
            return rotation == 90 || rotation == 270;
        }

        /// <summary>Обратный поворот (для обратного маппинга областей в исходное пространство).</summary>
        public static int Inverse(int rotation)
        {
            return (360 - PdfPageRef.ComposeRotation(rotation, 0)) % 360;
        }

        /// <summary>RotateFlip для поворота растра по часовой на 90/180/270. Чистая — под тест.</summary>
        public static RotateFlipType FlipFor(int rotation)
        {
            switch (rotation)
            {
                case 90: return RotateFlipType.Rotate90FlipNone;
                case 180: return RotateFlipType.Rotate180FlipNone;
                case 270: return RotateFlipType.Rotate270FlipNone;
                default: return RotateFlipType.RotateNoneFlipNone;
            }
        }

        /// <summary>Точка (x, y) исходной страницы W×H → точка повёрнутой страницы.</summary>
        public static void MapPoint(double x, double y, int rotation, double pageW, double pageH,
            out double nx, out double ny)
        {
            switch (rotation)
            {
                case 90: nx = y; ny = pageW - x; break;
                case 180: nx = pageW - x; ny = pageH - y; break;
                case 270: nx = pageH - y; ny = x; break;
                default: nx = x; ny = y; break;
            }
        }

        /// <summary>
        /// Прямоугольник (left/bottom/right/top) исходной страницы W×H → рамка на повёрнутой
        /// странице (поворот двух углов и нормализация min/max).
        /// </summary>
        public static void MapBox(double left, double bottom, double right, double top,
            int rotation, double pageW, double pageH,
            out double nLeft, out double nBottom, out double nRight, out double nTop)
        {
            double x1, y1, x2, y2;
            MapPoint(left, bottom, rotation, pageW, pageH, out x1, out y1);
            MapPoint(right, top, rotation, pageW, pageH, out x2, out y2);
            nLeft = Math.Min(x1, x2);
            nRight = Math.Max(x1, x2);
            nBottom = Math.Min(y1, y2);
            nTop = Math.Max(y1, y2);
        }

        /// <summary>PNG, повёрнутый по часовой на rotation. Сбой декодирования — исходные байты (картинку не теряем).</summary>
        public static byte[] RotatePng(byte[] png, int rotation)
        {
            if (png == null || png.Length == 0 || rotation == 0)
                return png;
            try
            {
                using (var ms = new MemoryStream(png))
                using (var bmp = new Bitmap(ms))
                {
                    bmp.RotateFlip(FlipFor(rotation)); // собственная копия из потока — мутировать безопасно
                    using (var outMs = new MemoryStream())
                    {
                        bmp.Save(outMs, System.Drawing.Imaging.ImageFormat.Png);
                        return outMs.ToArray();
                    }
                }
            }
            catch
            {
                return png;
            }
        }

        /// <summary>
        /// Повернуть содержимое страницы НА МЕСТЕ: слова, линовку (с переклассификацией
        /// H↔V), изображения (рамки и пиксели) и размеры страницы. rotation == 0 — no-op.
        /// Вызывать ДО детекторов таблиц и разбиения на абзацы.
        /// </summary>
        public static void RotatePage(List<PdfWord> words, List<PdfLine> lines, List<OcrImage> images,
            int rotation, ref double pageW, ref double pageH)
        {
            if (rotation == 0)
                return;
            double w = pageW, h = pageH;

            if (words != null)
                foreach (PdfWord word in words)
                {
                    double l, b, r, t;
                    MapBox(word.Left, word.Bottom, word.Right, word.Top, rotation, w, h, out l, out b, out r, out t);
                    word.Left = l;
                    word.Bottom = b;
                    word.Right = r;
                    word.Top = t;
                    // Базовая линия — ТОЧКА, а не сторона рамки: у боковой строки она вертикальна,
                    // и «нижний край рамки» после разворота указал бы не туда. Поворачиваем как
                    // точку — в выправленном пространстве её Y и есть базовая линия строки.
                    if (word.BaselineXPt != 0 || word.BaselineYPt != 0)
                    {
                        double bx, by;
                        MapPoint(word.BaselineXPt, word.BaselineYPt, rotation, w, h, out bx, out by);
                        word.BaselineXPt = bx;
                        word.BaselineYPt = by;
                    }
                }

            if (lines != null)
                foreach (PdfLine line in lines)
                {
                    double x1, y1, x2, y2;
                    MapPoint(line.X1, line.Y1, rotation, w, h, out x1, out y1);
                    MapPoint(line.X2, line.Y2, rotation, w, h, out x2, out y2);
                    line.X1 = x1;
                    line.Y1 = y1;
                    line.X2 = x2;
                    line.Y2 = y2;
                    if (SwapsDimensions(rotation))
                        line.Orientation = line.Orientation == LineOrientation.Horizontal
                            ? LineOrientation.Vertical
                            : LineOrientation.Horizontal;
                }

            if (images != null)
                foreach (OcrImage img in images)
                {
                    double l, b, r, t;
                    MapBox(img.LeftPt, img.TopPt - img.HeightPt, img.LeftPt + img.WidthPt, img.TopPt,
                        rotation, w, h, out l, out b, out r, out t);
                    img.LeftPt = l;
                    img.TopPt = t;
                    img.WidthPt = r - l;
                    img.HeightPt = t - b;
                    img.Png = RotatePng(img.Png, rotation);
                }

            if (SwapsDimensions(rotation))
            {
                pageW = h;
                pageH = w;
            }
        }
    }
}
