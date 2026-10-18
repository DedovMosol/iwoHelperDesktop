using System.Drawing;

namespace ExcelMerger
{
    /// <summary>
    /// Отображение рамок слов «Сравнения» на отрендеренную страницу. Чистая математика
    /// без PDF и UI — под тест.
    ///
    /// Слова извлечения живут в пространстве страницы: пункты, X вправо, Y ВВЕРХ,
    /// размеры страницы БЕЗ учёта её собственного поворота (/Rotate). Системный рендерер
    /// показывает страницу повёрнутой, как в читалке, поэтому рамку сначала переводим в
    /// пространство отображения тем же маппингом, что и весь конвейер («см. PageRotation»),
    /// а затем — в пиксели растра с переворотом оси Y (у растра Y вниз).
    /// </summary>
    internal static class PdfReviewGeometry
    {
        /// <summary>
        /// Рамка слова (пространство страницы, Y вверх) → рамка в пространстве отображения
        /// (с учётом поворота страницы, всё ещё пункты и Y вверх).
        /// </summary>
        public static PdfReviewBox RawToView(double left, double bottom, double right, double top,
            int rotation, double pageW, double pageH)
        {
            double nl, nb, nr, nt;
            PageRotation.MapBox(left, bottom, right, top, rotation, pageW, pageH,
                out nl, out nb, out nr, out nt);
            return new PdfReviewBox { Left = nl, Bottom = nb, Right = nr, Top = nt };
        }

        /// <summary>Размеры страницы в пространстве отображения (поворот меняет местами).</summary>
        public static void ViewSize(double pageW, double pageH, int rotation,
            out double viewW, out double viewH)
        {
            if (PageRotation.SwapsDimensions(rotation))
            {
                viewW = pageH;
                viewH = pageW;
            }
            else
            {
                viewW = pageW;
                viewH = pageH;
            }
        }

        /// <summary>
        /// Рамка в пространстве отображения (пункты, Y вверх) → пиксельный прямоугольник
        /// на растре размером bitmapW×bitmapH (Y вниз). Вырожденные входы — пустой результат,
        /// а не исключение: подсветка не должна ронять показ страницы.
        /// </summary>
        public static RectangleF ToPixelRect(PdfReviewBox box, double viewW, double viewH,
            double bitmapW, double bitmapH)
        {
            if (viewW <= 0 || viewH <= 0 || bitmapW <= 0 || bitmapH <= 0)
                return RectangleF.Empty;
            double sx = bitmapW / viewW, sy = bitmapH / viewH;
            var rect = new RectangleF(
                (float)(box.Left * sx),
                (float)((viewH - box.Top) * sy),
                (float)((box.Right - box.Left) * sx),
                (float)((box.Top - box.Bottom) * sy));
            if (rect.Width <= 0 || rect.Height <= 0)
                return RectangleF.Empty;
            return rect;
        }

        /// <summary>
        /// Одна строка текста: вертикальное перекрытие не меньше 35% меньшей высоты либо
        /// центры ближе 45% её. Единственная реализация — по ней раскладываются полосы
        /// изменений и по ней же поверхность выбирает слова строки при клике по полосе.
        /// </summary>
        public static bool SameTextLine(RectangleF line, RectangleF next)
        {
            float smallerHeight = System.Math.Min(line.Height, next.Height);
            if (smallerHeight <= 0)
                return false;
            float overlap = System.Math.Min(line.Bottom, next.Bottom) -
                System.Math.Max(line.Top, next.Top);
            if (overlap >= smallerHeight * 0.35f)
                return true;
            float lineCenter = line.Top + line.Height / 2f;
            float nextCenter = next.Top + next.Height / 2f;
            return System.Math.Abs(lineCenter - nextCenter) <= smallerHeight * 0.45f;
        }
    }
}
