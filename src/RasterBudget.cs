using System;

namespace ExcelMerger
{
    /// <summary>
    /// Общая политика растров: лимиты зависят от разрядности и ограничивают площадь И
    /// каждую сторону до передачи размеров в WinRT/GDI+. На x86 пределы ниже из-за
    /// 2-гигабайтного адресного пространства и одновременных decode/copy буферов.
    /// </summary>
    internal static class RasterBudget
    {
        internal static readonly long DefaultRenderPixels = IntPtr.Size == 8 ? 25000000L : 8000000L;
        internal static readonly long ThumbnailPixels = 4000000L;
        internal static readonly long PreviewPixels = IntPtr.Size == 8 ? 24000000L : 8000000L;
        internal static readonly long ExportPixels = IntPtr.Size == 8 ? 40000000L : 10000000L;
        internal static readonly long BackgroundPixels = IntPtr.Size == 8 ? 16000000L : 8000000L;
        internal static readonly long ImagePixels = IntPtr.Size == 8 ? 50000000L : 12000000L;
        internal const int MaxRenderDimension = 20000;
        internal static readonly int MaxExportDpi = IntPtr.Size == 8 ? 600 : 300;

        /// <summary>
        /// Максимальная безопасная ширина при сохранении пропорций. Высота считается через
        /// Ceiling — так WinRT не округлит теоретически точную границу на лишнюю строку.
        /// </summary>
        internal static int FitWidth(int requestedWidth, double pageWidth,
            double pageHeight, long maxPixels, int maxHeight = MaxRenderDimension)
        {
            if (requestedWidth <= 0 || !FinitePositive(pageWidth) ||
                !FinitePositive(pageHeight) || maxPixels <= 0)
                return 0;
            double aspect = pageHeight / pageWidth;
            if (!FinitePositive(aspect))
                return 0;
            int heightLimit = maxHeight > 0
                ? Math.Min(maxHeight, MaxRenderDimension)
                : MaxRenderDimension;
            int width = Math.Min(requestedWidth, MaxRenderDimension);
            double byHeight = heightLimit / aspect;
            if (byHeight < width)
                width = Math.Max(1, (int)Math.Floor(byHeight));
            double byPixels = Math.Sqrt(maxPixels / aspect);
            if (byPixels < width)
                width = Math.Max(1, (int)Math.Floor(byPixels));
            while (width > 0)
            {
                int height = ExpectedHeight(width, aspect);
                if (height > 0 && height <= heightLimit &&
                    (long)width * height <= maxPixels)
                    return width;
                width--;
            }
            return 0;
        }

        internal static int ExpectedHeight(int width, double aspect)
        {
            if (width <= 0 || !FinitePositive(aspect))
                return 0;
            double value = Math.Ceiling(width * aspect);
            return value > int.MaxValue ? int.MaxValue : (int)Math.Max(1, value);
        }

        internal static bool IsWithin(int width, int height, long maxPixels)
        {
            return width > 0 && height > 0 && width <= MaxRenderDimension &&
                height <= MaxRenderDimension && maxPixels > 0 &&
                (long)width * height <= maxPixels;
        }

        internal static int SafeDpi(double widthPt, double heightPt, int requestedDpi,
            long maxPixels)
        {
            if (!FinitePositive(widthPt) || !FinitePositive(heightPt) ||
                requestedDpi <= 0 || maxPixels <= 0)
                return 0;
            double byWidth = MaxRenderDimension * 72.0 / widthPt;
            double byHeight = MaxRenderDimension * 72.0 / heightPt;
            double byPixels = Math.Sqrt(maxPixels * 72.0 * 72.0 / (widthPt * heightPt));
            double safe = Math.Min(requestedDpi, Math.Min(byPixels,
                Math.Min(byWidth, byHeight)));
            return safe < 1 ? 0 : (int)Math.Floor(safe);
        }

        internal static bool IsValidImageDimensions(int width, int height)
        {
            return IsWithin(width, height, ImagePixels);
        }

        internal static bool IsValidExportDpi(int dpi)
        {
            return dpi > 0 && dpi <= MaxExportDpi;
        }

        internal static long BitmapWorkingSetBytes(int width, int height, int copies)
        {
            long one = PdfMemoryBudget.EstimateBitmapBytes(width, height);
            if (one <= 0 || copies <= 0)
                return 0;
            return one > long.MaxValue / copies ? long.MaxValue : one * copies;
        }

        private static bool FinitePositive(double value)
        {
            return value > 0 && !double.IsNaN(value) && !double.IsInfinity(value);
        }
    }
}
