using System;
using System.Drawing;

namespace ExcelMerger
{
    /// <summary>
    /// Масштаб миниатюр: ширина плитки и пересчёт от колеса мыши. Чистая логика
    /// (без UI) — покрыта юнит-тестами. Страницы рендерятся один раз в ширину
    /// RenderWidth, а при зуме плитки лишь пересобираются из кэша (GDI, без WinRT).
    /// </summary>
    public static class ThumbZoom
    {
        // ImageList.ImageSize ограничен 256×256 (WinForms) — верхняя граница
        // подобрана так, чтобы высота плитки (ширина×1.30) не превышала предел.
        private const int MaxImageDim = 256;
        public const int MinWidth = 96;
        public const int MaxWidth = 190; // 190 × 1.30 = 247 ≤ 256
        public const int DefaultWidth = 132;
        public const int RenderWidth = 300; // исходная страница — обычный Bitmap, лимита нет
        private const double TileAspect = 1.30; // высота = ширина × коэффициент
        private const int WheelStep = 16;       // пикселей ширины на «щелчок» колеса

        public static int Clamp(int width)
        {
            if (width < MinWidth) return MinWidth;
            if (width > MaxWidth) return MaxWidth;
            return width;
        }

        public static Size TileSize(int width)
        {
            int w = Math.Min(Clamp(width), MaxImageDim);
            int h = Math.Min((int)Math.Round(w * TileAspect), MaxImageDim);
            return new Size(w, h); // гарантированно в пределах ImageList
        }

        /// <summary>Новая ширина плитки по повороту колеса (Ctrl+колесо).</summary>
        public static int StepFromWheel(int currentWidth, int wheelDelta)
        {
            int notches = wheelDelta / 120; // WHEEL_DELTA
            return Clamp(currentWidth + notches * WheelStep);
        }
    }
}
