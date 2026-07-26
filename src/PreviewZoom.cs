using System;
using System.Drawing;

namespace ExcelMerger
{
    /// <summary>
    /// Математика масштабирования и панорамы полноэкранного просмотра. Вынесена из формы
    /// целиком: это решения, а не рисование, и потому проверяются тестами, а не глазами.
    ///
    /// Шаги масштаба взяты те же, что в браузерах и настольных просмотрщиках — к ним
    /// привыкли, и на каждом шаге картинка меняется заметно, но не скачком.
    /// </summary>
    public static class PreviewZoom
    {
        /// <summary>Ступени масштаба, доли от натуральной величины. 1.0 — «как есть».</summary>
        public static readonly double[] Steps =
        {
            0.25, 0.33, 0.50, 0.67, 0.75, 1.00, 1.25, 1.50, 2.00, 3.00, 4.00
        };

        /// <summary>Минимальный и максимальный масштаб — края <see cref="Steps"/>.</summary>
        public static double Min { get { return Steps[0]; } }
        public static double Max { get { return Steps[Steps.Length - 1]; } }

        /// <summary>
        /// Следующая ступень в заданную сторону от произвольного текущего масштаба (он может
        /// быть между ступенями — например после подгонки по окну). Упор в край возвращает
        /// сам край. Чистая — под тест.
        /// </summary>
        public static double Next(double current, int direction)
        {
            const double eps = 1e-6;
            if (direction > 0)
            {
                foreach (double s in Steps)
                    if (s > current + eps)
                        return s;
                return Max;
            }
            for (int i = Steps.Length - 1; i >= 0; i--)
                if (Steps[i] < current - eps)
                    return Steps[i];
            return Min;
        }

        /// <summary>
        /// Масштаб «вписать целиком»: страница видна вся, пропорции сохранены. Крупнее
        /// натуральной величины не увеличиваем — мелкая страница не должна растягиваться
        /// на весь экран мыльным пятном. Чистая — под тест.
        /// </summary>
        public static double Fit(Size image, Size viewport)
        {
            if (image.Width <= 0 || image.Height <= 0 || viewport.Width <= 0 || viewport.Height <= 0)
                return 1.0;
            double byWidth = (double)viewport.Width / image.Width;
            double byHeight = (double)viewport.Height / image.Height;
            double fit = Math.Min(byWidth, byHeight);
            return fit > 1.0 ? 1.0 : fit;
        }

        /// <summary>Размер картинки при заданном масштабе, не меньше пикселя. Чистая — под тест.</summary>
        public static Size Scaled(Size image, double scale)
        {
            int w = (int)Math.Round(image.Width * scale);
            int h = (int)Math.Round(image.Height * scale);
            return new Size(w < 1 ? 1 : w, h < 1 ? 1 : h);
        }

        /// <summary>
        /// Новое смещение прокрутки, чтобы точка под курсором ОСТАЛАСЬ под курсором при смене
        /// масштаба. Именно этого ждут от Ctrl+колеса: увеличение «к тому месту, куда смотришь»,
        /// а не к центру страницы. Отрицательное смещение не имеет смысла — обрезаем нулём,
        /// верхнюю границу накладывает уже сама прокрутка. Чистая — под тест.
        /// </summary>
        public static int Anchor(int scroll, int cursor, double oldScale, double newScale)
        {
            if (oldScale <= 0)
                return scroll;
            double content = (scroll + cursor) / oldScale; // точка страницы под курсором
            int shifted = (int)Math.Round(content * newScale) - cursor;
            return shifted < 0 ? 0 : shifted;
        }

        /// <summary>
        /// Помещается ли картинка целиком (панорама не нужна и рука не показывается).
        /// Чистая — под тест.
        /// </summary>
        public static bool FitsEntirely(Size scaledImage, Size viewport)
        {
            return scaledImage.Width <= viewport.Width && scaledImage.Height <= viewport.Height;
        }

        /// <summary>
        /// Закрывает ли клик окно просмотра. Закрывает только ЛЕВАЯ кнопка (правая принадлежит
        /// меню) и только если клик не оказался перетаскиванием и картинка видна целиком.
        ///
        /// Увеличенную страницу тем же движением ТАЩАТ, поэтому там клик не закрывает ничего:
        /// иначе окно захлопывалось бы посреди панорамы. Закрыть по-прежнему можно клавишей
        /// Esc и крестиком. Чистая — под тест.
        /// </summary>
        public static bool ClosesOnClick(System.Windows.Forms.MouseButtons button, bool dragged, bool canPan)
        {
            return button == System.Windows.Forms.MouseButtons.Left && !dragged && !canPan;
        }

        /// <summary>
        /// Считать ли движение мыши перетаскиванием, а не дрожанием руки при клике. Порог —
        /// системный размер «зоны нечувствительности» перетаскивания. Чистая — под тест.
        /// </summary>
        public static bool IsDrag(Point from, Point to, Size threshold)
        {
            return Math.Abs(to.X - from.X) > threshold.Width / 2 ||
                   Math.Abs(to.Y - from.Y) > threshold.Height / 2;
        }

        /// <summary>Масштаб в процентах для подписи. Чистая — под тест.</summary>
        public static int Percent(double scale)
        {
            int p = (int)Math.Round(scale * 100);
            return p < 1 ? 1 : p;
        }
    }
}
