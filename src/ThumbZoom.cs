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
        private const int MinRenderWidth = 300; // исходная страница — обычный Bitmap, лимита ImageList нет
        private const int MaxRenderWidth = 640; // потолок памяти: страница ~640×930×4 ≈ 2,4 МБ
        private const double RenderOversample = 1.5; // запас на даунскейл HighQualityBicubic
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

        /// <summary>
        /// Ширина рендера страницы в ФИЗИЧЕСКИХ пикселях монитора: максимальная плитка ×
        /// DPI-масштаб × запас на даунскейл. На 100% DPI остаётся прежняя (300), на
        /// 150% — резче вместо мыла от растяжения. Считается один раз на сетку
        /// (масштаб плиток пересобирает их из этого рендера без WinRT). Чистая — под тест.
        /// </summary>
        public static int RenderWidthFor(int deviceDpi)
        {
            if (deviceDpi < 96)
                deviceDpi = 96;
            int w = (int)Math.Ceiling(MaxWidth * (deviceDpi / 96.0) * RenderOversample);
            if (w < MinRenderWidth) return MinRenderWidth;
            if (w > MaxRenderWidth) return MaxRenderWidth;
            return w;
        }

        /// <summary>
        /// Ёмкость LRU-кэша отрендеренных страниц из бюджета памяти: все записи почти
        /// одноразмерны (ширина рендера фиксирована на сессию, высота ~×1.45 для A4),
        /// поэтому счётный предел честно ограничивает байты. Чистая — под тест.
        /// </summary>
        public static int PageCacheCapacity(long budgetBytes, int renderWidth)
        {
            long perPage = (long)renderWidth * (long)(renderWidth * 1.45) * 4; // 32bpp
            long capacity = perPage > 0 ? budgetBytes / perPage : 0;
            if (capacity < 24) return 24;    // видимое окно + запас должны жить в кэше
            if (capacity > 512) return 512;  // дальше выигрыша нет, а Touch в LRU — O(n)
            return (int)capacity;
        }
    }
}
