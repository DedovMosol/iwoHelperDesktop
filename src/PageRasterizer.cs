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
        private const int RangeTimeoutMs = 300000; // диапазон страниц — один запуск на весь файл

        /// <summary>
        /// Отрендерить страницу (1-based) в Bitmap через Ghostscript. null — GS недоступен или
        /// рендер не удался. Возвращённый Bitmap принадлежит вызывающему (обязан Dispose).
        /// </summary>
        public static Bitmap RenderPage(string pdfPath, int pageNumber)
        {
            if (!Ghostscript.Available || string.IsNullOrEmpty(pdfPath) || pageNumber < 1)
                return null;
            string outPng = Path.Combine(Path.GetTempPath(), "iwo_pg_" + Guid.NewGuid().ToString("N") + ".png");
            try
            {
                string stderr;
                int exit = Ghostscript.Run(BuildArgs(pdfPath, pageNumber, pageNumber, Dpi, outPng, false),
                    RenderTimeoutMs, out stderr);
                if (exit != 0 || !File.Exists(outPng))
                    return null;
                using (var fs = File.OpenRead(outPng))
                using (var decoded = new Bitmap(fs))
                    return new Bitmap(decoded); // копия, независимая от временного файла
            }
            catch { return null; }
            finally { try { File.Delete(outPng); } catch { } }
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
            try
            {
                using (Bitmap crop = pageBitmap.Clone(rect, pageBitmap.PixelFormat))
                using (var ms = new MemoryStream())
                {
                    crop.Save(ms, ImageFormat.Png);
                    return ms.ToArray();
                }
            }
            catch { return null; }
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
        public static List<string> RenderPagesWithoutText(string pdfPath, int firstPage, int lastPage, int dpi, string outDir)
        {
            var result = new List<string>();
            if (!Ghostscript.Available || string.IsNullOrEmpty(pdfPath) || firstPage < 1 || lastPage < firstPage)
                return result;
            try
            {
                Directory.CreateDirectory(outDir);
                string pattern = Path.Combine(outDir, "bg-%04d.png");
                string stderr;
                int exit = Ghostscript.Run(BuildArgs(pdfPath, firstPage, lastPage, dpi, pattern, true),
                    RangeTimeoutMs, out stderr);
                if (exit != 0)
                    return result;
                for (int i = 1; i <= lastPage - firstPage + 1; i++)
                {
                    string file = Path.Combine(outDir, "bg-" + i.ToString("D4", System.Globalization.CultureInfo.InvariantCulture) + ".png");
                    if (!File.Exists(file))
                        break; // движок оборвался на середине — берём то, что успел
                    result.Add(file);
                }
            }
            catch { result.Clear(); }
            return result;
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
