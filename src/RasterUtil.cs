using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Мелкие проверки растров, общие для разбора PDF и сборки слайдов. Вынесены отдельно,
    /// потому что вопрос «есть ли на картинке хоть что-нибудь» возникает в двух разных местах:
    /// извлекатель так отбраковывает не декодировавшиеся картинки, а писатель слайдов — пустые
    /// подложки страниц (страница из одного текста после снятия текста превращается в чистый
    /// лист, и класть его в файл незачем).
    /// </summary>
    internal static class RasterUtil
    {
        /// <summary>
        /// Одноцветна ли картинка (в пределах выборки). Проверка выборочная — сетка 16×16
        /// точек: сплошной обход мегапиксельного растра стоил бы дороже самой отрисовки, а
        /// пропустить он может разве что одинокую тонкую линию, потеря которой незаметна.
        /// Не декодировалась — считаем НЕ одноцветной: лучше положить лишнее, чем потерять.
        /// </summary>
        public static bool IsSolidColor(byte[] png)
        {
            try
            {
                using (var ms = new MemoryStream(png))
                using (var bmp = new Bitmap(ms))
                    return IsSolidColor(bmp);
            }
            catch { return false; }
        }

        /// <summary>Крупную картинку стоит попробовать пережать: фотография в JPEG втрое легче.</summary>
        private const int JpegTryOverBytes = 120 * 1024;
        private const long JpegWinsPercent = 70;  // берём JPEG, только если он меньше этой доли PNG
        private const long JpegQuality = 82;

        /// <summary>
        /// Выбрать для картинки более лёгкую упаковку: остаться PNG или пережать в JPEG.
        /// Мелкие картинки не трогаем вовсе, прозрачные — тем более (JPEG не умеет прозрачность,
        /// и подпись на прозрачном фоне стала бы подписью на чёрном). Рисунок из плашек и линий
        /// PNG жмёт лучше JPEG, поэтому правило «берём, только если заметно меньше» само
        /// оставляет логотипы и схемы в PNG, а фотографии переводит в JPEG.
        /// </summary>
        public static byte[] PreferSmaller(byte[] png, out bool isJpeg)
        {
            isJpeg = false;
            if (png == null || png.Length < JpegTryOverBytes)
                return png;
            try
            {
                using (var ms = new MemoryStream(png))
                using (var bmp = new Bitmap(ms))
                {
                    if (HasTransparency(bmp))
                        return png; // прозрачность важнее веса
                    byte[] jpeg = SaveJpeg(bmp, JpegQuality);
                    if (jpeg == null || jpeg.Length * 100L >= png.Length * JpegWinsPercent)
                        return png;
                    isJpeg = true;
                    return jpeg;
                }
            }
            catch { return png; } // не декодировалась — оставляем как есть
        }

        /// <summary>
        /// Есть ли в растре хоть один непрозрачный не до конца пиксель. Проверять надо
        /// СОДЕРЖИМОЕ, а не формат: копия растра и любая картинка, прошедшая через
        /// System.Drawing, почти всегда оказывается 32-битной с альфой, даже если ничего
        /// прозрачного в ней нет, — и проверка по формату запрещала бы пережатие всем подряд.
        /// Читаем построчно через LockBits: обход мегапиксельного растра по одному пикселю
        /// стоил бы дороже самой перекодировки.
        /// </summary>
        private static bool HasTransparency(Bitmap bmp)
        {
            if (!Image.IsAlphaPixelFormat(bmp.PixelFormat))
                return false;
            BitmapData data = null;
            try
            {
                data = bmp.LockBits(new Rectangle(0, 0, bmp.Width, bmp.Height),
                    ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                var row = new byte[Math.Abs(data.Stride)];
                for (int y = 0; y < data.Height; y++)
                {
                    IntPtr line = new IntPtr(data.Scan0.ToInt64() + (long)y * data.Stride);
                    System.Runtime.InteropServices.Marshal.Copy(line, row, 0, row.Length);
                    // Байт альфы — четвёртый в пикселе. Предел берём по МЕНЬШЕМУ из ширины и
                    // длины строки: у 32-битного растра они совпадают, но полагаться на это
                    // при чужом растре незачем.
                    int limit = Math.Min(data.Width * 4, row.Length);
                    for (int x = 3; x < limit; x += 4)
                        if (row[x] != 255)
                            return true;
                }
                return false;
            }
            catch { return true; } // не смогли посмотреть — считаем прозрачной и не трогаем
            finally
            {
                if (data != null)
                    try { bmp.UnlockBits(data); } catch { }
            }
        }

        /// <summary>Растр в JPEG заданного качества; null — кодировщика нет или сбой.</summary>
        public static byte[] SaveJpeg(Bitmap bmp, long quality)
        {
            try
            {
                ImageCodecInfo codec = null;
                foreach (ImageCodecInfo c in ImageCodecInfo.GetImageEncoders())
                    if (c.FormatID == ImageFormat.Jpeg.Guid)
                    {
                        codec = c;
                        break;
                    }
                if (codec == null)
                    return null;
                using (var ms = new MemoryStream())
                using (var parameters = new EncoderParameters(1))
                using (var param = new EncoderParameter(Encoder.Quality, quality))
                {
                    parameters.Param[0] = param;
                    bmp.Save(ms, codec, parameters);
                    return ms.ToArray();
                }
            }
            catch { return null; }
        }

        /// <summary>Точно одноцветен ли растр; ни одна тонкая линия не теряется из-за выборки.</summary>
        public static bool IsSolidColor(Bitmap bmp)
        {
            if (bmp == null)
                return true;
            BitmapData data = null;
            try
            {
                int width = bmp.Width, height = bmp.Height;
                if (width < 2 || height < 2)
                    return true;
                int sampled = bmp.GetPixel(0, 0).ToArgb();
                int stepX = Math.Max(1, width / 16), stepY = Math.Max(1, height / 16);
                for (int y = 0; y < height; y += stepY)
                    for (int x = 0; x < width; x += stepX)
                        if (bmp.GetPixel(x, y).ToArgb() != sampled)
                            return false; // definitive negative; full proof is unnecessary
                int bits = Image.GetPixelFormatSize(bmp.PixelFormat);
                int bytesPerPixel = bits / 8;
                if (bytesPerPixel != 3 && bytesPerPixel != 4)
                {
                    int firstArgb = bmp.GetPixel(0, 0).ToArgb();
                    for (int y = 0; y < height; y++)
                        for (int x = 0; x < width; x++)
                            if (bmp.GetPixel(x, y).ToArgb() != firstArgb)
                                return false;
                    return true;
                }

                data = bmp.LockBits(new Rectangle(0, 0, width, height),
                    ImageLockMode.ReadOnly, bmp.PixelFormat);
                int rowBytes = checked(width * bytesPerPixel);
                var row = new byte[rowBytes];
                int first = 0;
                bool haveFirst = false;
                for (int y = 0; y < height; y++)
                {
                    IntPtr start = IntPtr.Add(data.Scan0, y * data.Stride);
                    System.Runtime.InteropServices.Marshal.Copy(start, row, 0, rowBytes);
                    for (int x = 0; x < rowBytes; x += bytesPerPixel)
                    {
                        int argb = (bytesPerPixel >= 4 ? row[x + 3] : 255) << 24 |
                            row[x + 2] << 16 | row[x + 1] << 8 | row[x];
                        if (!haveFirst)
                        {
                            first = argb;
                            haveFirst = true;
                        }
                        else if (argb != first)
                            return false;
                    }
                }
                return true;
            }
            catch { return false; }
            finally
            {
                if (data != null)
                    try { bmp.UnlockBits(data); } catch { }
            }
        }
    }
}
