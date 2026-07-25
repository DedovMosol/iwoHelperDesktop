using System;
using System.Drawing;

namespace ExcelMerger
{
    /// <summary>
    /// Масштаб миниатюр: ширина плитки, размер ячейки сетки и пересчёт от колеса мыши.
    /// Чистая логика (без UI) — покрыта юнит-тестами. Сетка рисует плитки сама
    /// (owner-draw), поэтому лимита ImageList 256×256 больше нет: максимум — 400 px.
    /// Страницы рендерятся фоновым воркером в одну из двух ширин: базовую (хватает до
    /// «обычного» масштаба) и увеличенную (для крупного зума); плитки пересобираются
    /// из отрендеренного без повторного WinRT.
    /// </summary>
    public static class ThumbZoom
    {
        public const int MinWidth = 96;
        /// <summary>До этой ширины плитки хватает базового рендера; выше — дорендер крупнее.</summary>
        public const int BaseWidth = 190;
        public const int MaxWidth = 400;
        public const int DefaultWidth = 132;
        private const int MinRenderWidth = 300; // прежняя ширина рендера — нижняя граница
        private const int MaxRenderWidth = 640; // потолок памяти: страница ~640×930×4 ≈ 2,4 МБ
        private const double RenderOversample = 1.5; // запас на даунскейл HighQualityBicubic
        private const double TileAspect = 1.30; // высота = ширина × коэффициент
        private const int WheelStep = 16;       // пикселей ширины на «щелчок» колеса
        private const int CellPadX = 16;        // поля ячейки вокруг плитки
        private const int CellPadY = 10;
        private const int LabelBand = 20;       // полоса под номер страницы

        public static int Clamp(int width)
        {
            if (width < MinWidth) return MinWidth;
            if (width > MaxWidth) return MaxWidth;
            return width;
        }

        /// <summary>Размер самой плитки (страница с рамкой) при заданной ширине.</summary>
        public static Size TileSize(int width)
        {
            int w = Clamp(width);
            return new Size(w, (int)Math.Round(w * TileAspect));
        }

        /// <summary>
        /// Размер ячейки сетки (плитка + поля + полоса номера) — шаг LVM_SETICONSPACING
        /// в ЛОГИЧЕСКИХ пикселях; вызывающий масштабирует под DPI.
        /// </summary>
        public static Size CellSize(int width)
        {
            Size tile = TileSize(width);
            return new Size(tile.Width + CellPadX, tile.Height + CellPadY + LabelBand);
        }

        /// <summary>Новая ширина плитки по повороту колеса (Ctrl+колесо).</summary>
        public static int StepFromWheel(int currentWidth, int wheelDelta)
        {
            int notches = wheelDelta / 120; // WHEEL_DELTA
            return Clamp(currentWidth + notches * WheelStep);
        }

        /// <summary>Масштаб в процентах относительно ширины по умолчанию (DefaultWidth = 100%). Чистая — под тест.</summary>
        public static int Percent(int width)
        {
            return (int)Math.Round(100.0 * Clamp(width) / DefaultWidth);
        }

        /// <summary>
        /// Ширина рендера страницы в ФИЗИЧЕСКИХ пикселях для плитки maxTileWidth: плитка ×
        /// DPI-масштаб × запас на даунскейл, в пределах [300..640]. Базовый рендер считается
        /// от <see cref="BaseWidth"/> (быстрый, хватает до обычного масштаба), увеличенный —
        /// от <see cref="MaxWidth"/> (докачивается только при крупном зуме). Чистая — под тест.
        /// </summary>
        public static int RenderWidthFor(int deviceDpi, int maxTileWidth)
        {
            if (deviceDpi < 96)
                deviceDpi = 96;
            int w = (int)Math.Ceiling(maxTileWidth * (deviceDpi / 96.0) * RenderOversample);
            if (w < MinRenderWidth) return MinRenderWidth;
            if (w > MaxRenderWidth) return MaxRenderWidth;
            return w;
        }

        /// <summary>
        /// Ёмкость LRU-кэша отрендеренных страниц из бюджета памяти: все записи почти
        /// одноразмерны (высота ~×1.45 для A4), поэтому счётный предел честно ограничивает
        /// байты. Считается от НАИБОЛЬШЕЙ ширины рендера (консервативно). Чистая — под тест.
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
