using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.CompilerServices;

namespace ExcelMerger
{
    /// <summary>
    /// Картинка → страницы PDF: каждый кадр ложится на лист A4 с полями, вписанный БЕЗ
    /// искажения пропорций и по центру, а сам лист встаёт в ориентацию картинки (широкая —
    /// альбомный). «Вписанный» значит именно во весь лист за полями: раз выбран A4, страница
    /// не должна выглядеть запиской с маленькой картинкой в середине.
    ///
    /// Здесь же собраны ловушки формата, каждая из которых уже кому-то испортила документ:
    /// EXIF-ориентация (GDI+ её НЕ применяет — снимок с телефона ложится боком), прозрачность
    /// (без белой подложки PNG уходит в PDF чёрными пятнами — та же беда, что была со SMask),
    /// чтение ЧЕРЕЗ ПАМЯТЬ (Image.FromFile держит файл замапленным, и его нельзя ни удалить,
    /// ни перезаписать), многостраничный TIFF (у сканов это норма) и JPEG, который переносится
    /// КАК ЕСТЬ, без перекодирования: иначе снимок на 12 Мп раздувает PDF и теряет качество
    /// на ровном месте.
    /// </summary>
    public static class ImageToPdfService
    {
        /// <summary>Что читает GDI+. HEIC/HEIF он не умеет — их и не предлагаем открыть.</summary>
        public static readonly string[] Extensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff" };

        // Лист A4 в пунктах и поля 1 см: те же поля, что у документов, к которым такой PDF подшивают.
        internal const double A4WidthPt = 595.276, A4HeightPt = 841.89, MarginPt = 28.35;

        private const int ExifOrientationId = 0x0112; // тег «как снято» — единственный, который нам нужен
        private const long JpegQuality = 92L;         // ниже видны кольца вокруг букв на снимках документов

        /// <summary>Путь похож на картинку, которую мы берёмся прочитать.</summary>
        public static bool IsImage(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;
            string ext = Path.GetExtension(path);
            foreach (string known in Extensions)
                if (string.Equals(ext, known, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        /// <summary>Фильтр для диалога выбора: подпись переводится, список расширений — один на оба языка.</summary>
        public static string DialogFilter()
        {
            var masks = new List<string>();
            foreach (string ext in Extensions)
                masks.Add("*" + ext);
            string list = string.Join(";", masks.ToArray());
            return Loc.T("img.filter") + " (" + list + ")|" + list;
        }

        /// <summary>
        /// Записать картинку в PDF: один кадр — одна страница (у многостраничного TIFF их
        /// несколько). Возвращает число записанных страниц. Битый, чужой или недоступный файл —
        /// понятное <see cref="MergeException"/>; нехватку памяти НЕ маскируем (см. ShouldWrap):
        /// огромная картинка — это не «файл повреждён», и вести по ложному следу нельзя.
        /// </summary>
        public static int WritePages(string imagePath, string outPdfPath)
        {
            EmbeddedAssemblies.Ensure(); // PdfSharp вшит ресурсом — подгрузить до первого его типа
            return WriteCore(imagePath, outPdfPath);
        }

        // NoInlining: в теле типы PdfSharp, и без этого они потребовались бы вызывающему методу
        // ещё до Ensure() (та же причина, что в PdfMergeService).
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static int WriteCore(string imagePath, string outPdfPath)
        {
            // Потоки живут до Save: PdfSharp читает их отложенно, и закрытый поток означал бы
            // пустую страницу в готовом файле.
            var pageStreams = new List<MemoryStream>();
            try
            {
                byte[] bytes = File.ReadAllBytes(imagePath);
                using (var source = new MemoryStream(bytes))
                using (Image image = Image.FromStream(source, true, false))
                using (var doc = new PdfSharp.Pdf.PdfDocument())
                {
                    int frames = FrameCount(image);
                    int orientation = ExifOrientation(image);
                    for (int i = 0; i < frames; i++)
                    {
                        if (frames > 1)
                            image.SelectActiveFrame(FrameDimension.Page, i);
                        MemoryStream encoded = Encode(image, bytes, imagePath, orientation, frames);
                        pageStreams.Add(encoded);
                        AddPage(doc, encoded);
                    }
                    doc.Save(outPdfPath);
                    return frames;
                }
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                throw new MergeException(string.Format(Loc.T("err.img.cantRead"),
                    Path.GetFileName(imagePath), ex.Message));
            }
            finally
            {
                foreach (MemoryStream s in pageStreams)
                    s.Dispose();
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void AddPage(PdfSharp.Pdf.PdfDocument doc, MemoryStream encoded)
        {
            using (PdfSharp.Drawing.XImage image = PdfSharp.Drawing.XImage.FromStream(encoded))
            {
                double pageW, pageH, x, y, w, h;
                Layout(image.PixelWidth, image.PixelHeight, out pageW, out pageH, out x, out y, out w, out h);
                PdfSharp.Pdf.PdfPage page = doc.AddPage();
                page.Width = pageW;
                page.Height = pageH;
                using (PdfSharp.Drawing.XGraphics g = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
                    g.DrawImage(image, x, y, w, h);
            }
        }

        /// <summary>
        /// Лист и место картинки на нём: ориентация листа — по картинке, картинка вписана в поля
        /// без искажения пропорций и по центру. Пропорции считаются в пикселях: dpi в
        /// метаданных врёт слишком часто, чтобы на него опираться. Чистая — под тест.
        /// </summary>
        internal static void Layout(int pixelWidth, int pixelHeight,
            out double pageW, out double pageH, out double x, out double y, out double w, out double h)
        {
            int pw = pixelWidth > 0 ? pixelWidth : 1;   // вырожденный кадр не должен делить на ноль
            int ph = pixelHeight > 0 ? pixelHeight : 1;
            bool landscape = pw > ph;
            pageW = landscape ? A4HeightPt : A4WidthPt;
            pageH = landscape ? A4WidthPt : A4HeightPt;
            double boxW = pageW - 2 * MarginPt, boxH = pageH - 2 * MarginPt;
            double scale = Math.Min(boxW / pw, boxH / ph);
            w = pw * scale;
            h = ph * scale;
            x = (pageW - w) / 2;
            y = (pageH - h) / 2;
        }

        /// <summary>
        /// Поворот кадра по тегу EXIF «как снято»: GDI+ его не применяет, и снимок с телефона
        /// ложится в PDF боком. Чистая — под тест.
        /// </summary>
        internal static RotateFlipType ExifRotation(int orientation)
        {
            switch (orientation)
            {
                case 2: return RotateFlipType.RotateNoneFlipX;
                case 3: return RotateFlipType.Rotate180FlipNone;
                case 4: return RotateFlipType.RotateNoneFlipY;
                case 5: return RotateFlipType.Rotate90FlipX;
                case 6: return RotateFlipType.Rotate90FlipNone;
                case 7: return RotateFlipType.Rotate270FlipX;
                case 8: return RotateFlipType.Rotate270FlipNone;
                default: return RotateFlipType.RotateNoneFlipNone; // 1 и всё непонятное — как есть
            }
        }

        /// <summary>Сколько кадров в файле (страницы TIFF); у обычной картинки — один.</summary>
        private static int FrameCount(Image image)
        {
            try
            {
                int count = image.GetFrameCount(FrameDimension.Page);
                return count > 0 ? count : 1;
            }
            catch { return 1; } // формат без измерения «страница» — один кадр
        }

        private static int ExifOrientation(Image image)
        {
            try
            {
                foreach (int id in image.PropertyIdList)
                    if (id == ExifOrientationId)
                    {
                        PropertyItem item = image.GetPropertyItem(ExifOrientationId);
                        if (item != null && item.Value != null && item.Value.Length >= 2)
                            return BitConverter.ToUInt16(item.Value, 0);
                    }
            }
            catch { } // тега нет или он битый — считаем «как есть»
            return 1;
        }

        /// <summary>
        /// Байты страницы. JPEG без поворота и без второго кадра уходит в PDF КАК ЕСТЬ — это и
        /// быстро, и без потерь, и файл остаётся размером с исходный снимок. Всё остальное
        /// приходится пересобрать: снять прозрачность на белое и применить EXIF-поворот.
        /// </summary>
        private static MemoryStream Encode(Image frame, byte[] original, string path, int orientation, int frames)
        {
            bool jpeg = HasExtension(path, ".jpg") || HasExtension(path, ".jpeg");
            if (jpeg && frames == 1 && ExifRotation(orientation) == RotateFlipType.RotateNoneFlipNone)
                return new MemoryStream(original);

            using (Bitmap flat = Flatten(frame, orientation))
            {
                var ms = new MemoryStream();
                // Снимок пересохраняем снимком (JPEG), а рисунок и снимок экрана — без потерь:
                // на тексте и тонких линиях JPEG даёт кольца, и именно такие картинки чаще
                // всего и подшивают к документам.
                if (jpeg || HasExtension(path, ".tif") || HasExtension(path, ".tiff"))
                    SaveJpeg(flat, ms);
                else
                    flat.Save(ms, ImageFormat.Png);
                ms.Position = 0; // PdfSharp читает поток с текущего места
                return ms;
            }
        }

        /// <summary>Кадр на белом фоне, в 24 битах и с применённым EXIF-поворотом.</summary>
        private static Bitmap Flatten(Image frame, int orientation)
        {
            var flat = new Bitmap(frame.Width, frame.Height, PixelFormat.Format24bppRgb);
            try
            {
                using (Graphics g = Graphics.FromImage(flat))
                {
                    g.Clear(Color.White); // прозрачное становится белым, а не чёрным
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(frame, new Rectangle(0, 0, flat.Width, flat.Height));
                }
                RotateFlipType rotation = ExifRotation(orientation);
                if (rotation != RotateFlipType.RotateNoneFlipNone)
                    flat.RotateFlip(rotation);
                return flat;
            }
            catch
            {
                flat.Dispose(); // не отдали наружу — освобождаем сами, иначе утечёт GDI-объект
                throw;
            }
        }

        private static void SaveJpeg(Bitmap bitmap, Stream to)
        {
            ImageCodecInfo codec = null;
            foreach (ImageCodecInfo c in ImageCodecInfo.GetImageEncoders())
                if (c.FormatID == ImageFormat.Jpeg.Guid)
                    codec = c;
            if (codec == null)
            {
                bitmap.Save(to, ImageFormat.Jpeg); // без кодировщика — качеством по умолчанию
                return;
            }
            using (var ps = new EncoderParameters(1))
            {
                ps.Param[0] = new EncoderParameter(Encoder.Quality, JpegQuality);
                bitmap.Save(to, codec, ps);
            }
        }

        private static bool HasExtension(string path, string ext)
        {
            return string.Equals(Path.GetExtension(path), ext, StringComparison.OrdinalIgnoreCase);
        }
    }
}
