using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Растеризация области страницы PDF через Ghostscript (отдельный процесс — без WinRT/COM
    /// и проблем с апартаментами на STA-потоке конвертации). Фолбэк для картинок, которые
    /// извлекатель не смог декодировать (напр. штрих-код выходит битым/одноцветным): рендерим
    /// страницу и вырезаем изображение по его рамке — так переносится ЛЮБАЯ картинка, как она
    /// выглядит. Требует Ghostscript; без него RenderPage вернёт null, и картинка пропускается.
    /// </summary>
    internal static class PageRasterizer
    {
        private const int Dpi = 200;             // чётко для штрих-кода и мелкой графики
        private const int RenderTimeoutMs = 60000;
        internal const int MaxRangePages = 512;
        private const int RangeTimeoutMs = 300000; // диапазон страниц — один запуск на весь файл

        /// <summary>
        /// Отрендерить страницу (1-based) в Bitmap через Ghostscript. null — GS недоступен или
        /// рендер не удался. Возвращённый Bitmap принадлежит вызывающему (обязан Dispose).
        /// </summary>
        internal static BudgetedBitmap RenderPage(string pdfPath, int pageNumber,
            double pageWidthPt, double pageHeightPt)
        {
            if (!Ghostscript.Available || string.IsNullOrEmpty(pdfPath) || pageNumber < 1)
                return null;
            string outPng = Path.Combine(Path.GetTempPath(), "iwo_pg_" +
                Guid.NewGuid().ToString("N") + ".png");
            string decryptedPdf = null;
            string renderPath = pdfPath;
            int renderPage = pageNumber;
            PdfMemoryLease working = null;
            Bitmap result = null;
            try
            {
                int pixelWidth, pixelHeight;
                int dpi = SafePageDpi(pageWidthPt, pageHeightPt, Dpi,
                    out pixelWidth, out pixelHeight);
                if (dpi <= 0 || !PdfMemoryBudget.TryAcquire(
                    RasterBudget.BitmapWorkingSetBytes(pixelWidth, pixelHeight, 2),
                    out working))
                    return null;
                if (!string.IsNullOrEmpty(PdfPasswords.For(pdfPath)))
                {
                    decryptedPdf = Path.Combine(Path.GetTempPath(), "iwo_pg_dec_" +
                        Guid.NewGuid().ToString("N") + ".pdf");
                    PdfMergeService.WriteUnpublished(new[]
                    {
                        new PdfPageRef { SourcePath = pdfPath, PageIndex = pageNumber - 1 }
                    }, decryptedPdf);
                    renderPath = decryptedPdf;
                    renderPage = 1;
                }
                string stderr;
                int exit = Ghostscript.Run(BuildArgs(renderPath, renderPage, renderPage, dpi,
                    outPng, false), RenderTimeoutMs, out stderr);
                if (!GsRewrite.EngineSucceeded(exit, stderr) || !File.Exists(outPng))
                    return null;
                using (var stream = File.OpenRead(outPng))
                using (var decoded = new Bitmap(stream))
                {
                    if (!RasterBudget.IsWithin(decoded.Width, decoded.Height,
                        RasterBudget.BackgroundPixels))
                        return null;
                    result = new Bitmap(decoded);
                }
                working.ReduceTo(PdfMemoryBudget.EstimateBitmapBytes(
                    result.Width, result.Height));
                var owned = new BudgetedBitmap(result, working);
                result = null;
                working = null;
                return owned;
            }
            catch (OutOfMemoryException)
            {
                if (result != null) result.Dispose();
                throw;
            }
            catch
            {
                if (result != null) result.Dispose();
                return null;
            }
            finally
            {
                if (working != null) working.Dispose();
                try { File.Delete(outPng); } catch { }
                try { if (decryptedPdf != null) File.Delete(decryptedPdf); } catch { }
            }
        }

        /// <summary>
        /// Вырезать область (PDF pt, ось Y вверх; topPt — верхняя граница) из отрендеренной
        /// страницы и вернуть PNG. null при вырожденной рамке/сбое. pageBitmap — из
        /// <see cref="RenderPage"/> ТОЙ ЖЕ страницы.
        /// </summary>
        public static byte[] CropRegion(Bitmap pageBitmap, double pageWidthPt, double pageHeightPt,
            double leftPt, double topPt, double widthPt, double heightPt)
        {
            if (pageBitmap == null || pageWidthPt <= 0 || pageHeightPt <= 0)
                return null;
            Rectangle rect = CropRect(pageBitmap.Width, pageBitmap.Height, pageWidthPt, pageHeightPt,
                leftPt, topPt, widthPt, heightPt);
            if (rect.Width < 1 || rect.Height < 1)
                return null;
            PdfMemoryLease working;
            if (!PdfMemoryBudget.TryAcquire(
                RasterBudget.BitmapWorkingSetBytes(rect.Width, rect.Height, 2), out working))
                return null;
            try
            {
                using (Bitmap crop = pageBitmap.Clone(rect, pageBitmap.PixelFormat))
                using (var ms = new MemoryStream())
                {
                    crop.Save(ms, ImageFormat.Png);
                    return ms.ToArray();
                }
            }
            catch (OutOfMemoryException) { throw; }
            catch { return null; }
            finally { working.Dispose(); }
        }

        /// <summary>
        /// PDF-прямоугольник (pt, ось Y вверх: topPt — верхняя граница) → пиксельный прямоугольник
        /// в отрендеренной странице bmpW×bmpH, обрезанный по её границам. Чистая — под тест.
        /// </summary>
        internal static Rectangle CropRect(int bmpW, int bmpH, double pageWidthPt, double pageHeightPt,
            double leftPt, double topPt, double widthPt, double heightPt)
        {
            if (pageWidthPt <= 0 || pageHeightPt <= 0 || bmpW <= 0 || bmpH <= 0)
                return Rectangle.Empty; // защита от деления на ноль/вырожденной страницы
            double sx = bmpW / pageWidthPt, sy = bmpH / pageHeightPt;
            int x = (int)Math.Floor(leftPt * sx);
            int y = (int)Math.Floor((pageHeightPt - topPt) * sy); // ось вниз: верх картинки = высота − topPt
            int w = (int)Math.Ceiling(widthPt * sx);
            int h = (int)Math.Ceiling(heightPt * sy);
            if (x < 0) { w += x; x = 0; }
            if (y < 0) { h += y; y = 0; }
            if (x > bmpW) x = bmpW;
            if (y > bmpH) y = bmpH;
            if (x + w > bmpW) w = bmpW - x;
            if (y + h > bmpH) h = bmpH - y;
            if (w < 0) w = 0;
            if (h < 0) h = 0;
            return new Rectangle(x, y, w, h);
        }

        /// <summary>
        /// Отрендерить диапазон страниц БЕЗ ТЕКСТА — подложку для слайдов: остаётся всё, что
        /// текстом не является (фон, рамки, диаграммы, логотипы), а сам текст ляжет поверх
        /// редактируемыми надписями. Файлы называются bg-0001.png и далее ПО ПОРЯДКУ
        /// отрисованных страниц (Ghostscript нумерует вывод с единицы независимо от того, с
        /// какой страницы начали, — проверено), поэтому вызывающий сопоставляет их сам.
        /// Возвращает пути к созданным файлам по порядку; пустой список — не получилось.
        /// </summary>
        public static List<string> RenderPagesWithoutText(string pdfPath, int firstPage,
            int lastPage, int dpi, string outDir, Func<bool> cancelled)
        {
            var result = new List<string>();
            if (!Ghostscript.Available || string.IsNullOrEmpty(pdfPath) || firstPage < 1 ||
                lastPage < firstPage || (long)lastPage - firstPage + 1L > MaxRangePages)
                return result;
            string decrypted = null;
            try
            {
                Cancellation.ThrowIf(cancelled);
                Directory.CreateDirectory(outDir);
                string renderPath = pdfPath;
                int renderFirst = firstPage, renderLast = lastPage;
                if (!string.IsNullOrEmpty(PdfPasswords.For(pdfPath)))
                {
                    decrypted = Path.Combine(outDir, "decrypted.pdf");
                    var pages = new List<PdfPageRef>();
                    for (int page = firstPage; page <= lastPage; page++)
                        pages.Add(new PdfPageRef
                        {
                            SourcePath = pdfPath,
                            PageIndex = page - 1
                        });
                    PdfMergeService.WriteUnpublished(pages, decrypted, null, cancelled);
                    renderPath = decrypted;
                    renderFirst = 1;
                    renderLast = pages.Count;
                }
                string pattern = Path.Combine(outDir, "bg-%04d.png");
                string stderr;
                int exit = Ghostscript.Run(BuildArgs(renderPath, renderFirst, renderLast, dpi,
                    pattern, true), RangeTimeoutMs, out stderr, cancelled);
                Cancellation.ThrowIf(cancelled);
                if (!GsRewrite.EngineSucceeded(exit, stderr))
                    return result;
                long tempBytes = 0;
                long tempLimit = IntPtr.Size == 8 ? 256L << 20 : 96L << 20;
                for (int i = 1; i <= lastPage - firstPage + 1; i++)
                {
                    Cancellation.ThrowIf(cancelled);
                    string file = Path.Combine(outDir, "bg-" +
                        i.ToString("D4", System.Globalization.CultureInfo.InvariantCulture) + ".png");
                    var info = new FileInfo(file);
                    if (!info.Exists)
                        break;
                    tempBytes += info.Length;
                    if (tempBytes > tempLimit)
                    {
                        result.Clear();
                        break;
                    }
                    result.Add(file);
                }
            }
            catch (OperationCanceledException) { throw; }
            catch { result.Clear(); }
            finally
            {
                try { if (decrypted != null) File.Delete(decrypted); } catch { }
            }
            return result;
        }

        private static int SafePageDpi(double widthPt, double heightPt, int requestedDpi,
            out int pixelWidth, out int pixelHeight)
        {
            pixelWidth = pixelHeight = 0;
            int dpi = RasterBudget.SafeDpi(widthPt, heightPt, requestedDpi,
                RasterBudget.BackgroundPixels);
            if (dpi <= 0)
                return 0;
            pixelWidth = Math.Max(1, (int)Math.Ceiling(widthPt / 72.0 * dpi));
            pixelHeight = Math.Max(1, (int)Math.Ceiling(heightPt / 72.0 * dpi));
            return dpi;
        }

        private static string BuildArgs(string input, int firstPage, int lastPage, int dpi, string output, bool withoutText)
        {
            var sb = new StringBuilder();
            sb.Append("-q -dNOPAUSE -dBATCH -dSAFER -sDEVICE=png16m");
            if (withoutText)
                sb.Append(" -dFILTERTEXT"); // текст не рисуем: он придёт надписями поверх
            sb.Append(" -r").Append(dpi);
            sb.Append(" -dFirstPage=").Append(firstPage).Append(" -dLastPage=").Append(lastPage);
            string root = Ghostscript.BundledRoot; // вшитый GS — явные -I на его lib/Resource
            if (!string.IsNullOrEmpty(root))
            {
                sb.Append(" -I ").Append(Quote(Path.Combine(root, "lib")));
                sb.Append(" -I ").Append(Quote(Path.Combine(root, "Resource", "Init")));
            }
            sb.Append(" -sOutputFile=").Append(Quote(output));
            sb.Append(' ').Append(Quote(input));
            return sb.ToString();
        }

        private static string Quote(string s) { return "\"" + s + "\""; }
    }
}
