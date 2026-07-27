using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;

namespace ExcelMerger
{
    /// <summary>
    /// Печать страниц PDF. Рисуем те же растры, что и предпросмотр (тот же
    /// <see cref="PdfThumbnailRenderer"/>), поэтому на бумаге получается ровно то, что видно
    /// на экране, и новых зависимостей печать не приносит.
    ///
    /// Страница вписывается в область печати целиком с сохранением пропорций: растянуть лист
    /// по краям — значит обрезать поля, где у документов обычно и стоят подписи и номера.
    ///
    /// UI здесь нет: настройки принтера выбирает вызывающий и передаёт готовыми, прогресс и
    /// отмена приходят колбэками. Вызывать с фонового потока — рендер страниц не мгновенный.
    /// </summary>
    public static class PdfPrintService
    {
        /// <summary>
        /// Разрешение рендера под печать. 200 dpi — принятый компромисс: на глаз неотличимо от
        /// 300 на офисном лазернике, а памяти и времени вдвое меньше (лист A4 при 300 dpi это
        /// ~35 млн точек). Текст и линии остаются чёткими, потому что это растр страницы, а не
        /// пересжатие изображений.
        /// </summary>
        private const int RenderDpi = 200;

        /// <summary>
        /// Прямоугольник, в который встаёт картинка страницы: вписана целиком, пропорции
        /// сохранены, по центру области печати. Чистая — под тест.
        /// </summary>
        internal static Rectangle FitToPage(Size image, Rectangle printable)
        {
            if (image.Width <= 0 || image.Height <= 0 || printable.Width <= 0 || printable.Height <= 0)
                return printable;
            double scale = Math.Min((double)printable.Width / image.Width,
                                    (double)printable.Height / image.Height);
            int w = Math.Max(1, (int)Math.Round(image.Width * scale));
            int h = Math.Max(1, (int)Math.Round(image.Height * scale));
            return new Rectangle(printable.X + (printable.Width - w) / 2,
                                 printable.Y + (printable.Height - h) / 2, w, h);
        }

        /// <summary>
        /// Напечатать указанные страницы. settings уже выбраны пользователем в диалоге печати.
        /// Отмена проверяется между страницами: начатый лист дорисовывается, следующий не идёт.
        /// </summary>
        public static void Print(string sourcePath, IList<int> pageIndexes, PrinterSettings settings,
            Action<int, int> progress = null, Func<bool> cancelled = null)
        {
            if (pageIndexes == null || pageIndexes.Count == 0)
                throw new MergeException(Loc.T("err.export.noPages"));

            using (var renderer = new PdfThumbnailRenderer())
            using (var doc = new PrintDocument())
            {
                doc.PrinterSettings = settings;
                doc.DocumentName = System.IO.Path.GetFileName(sourcePath);
                doc.OriginAtMargins = false; // печатаем по всей доступной области, а не по полям

                int next = 0; // индекс в pageIndexes, а не номер страницы документа
                doc.PrintPage += delegate(object s, PrintPageEventArgs e)
                {
                    // Ширину рендера считаем от РЕАЛЬНОГО размера листа принтера, а не страницы
                    // PDF: так одинаково печатаются и A4, и нестандартные вставки.
                    int width = Math.Max(1, e.PageBounds.Width * RenderDpi / 100);
                    using (Bitmap page = renderer.Render(sourcePath, pageIndexes[next], width))
                    {
                        if (page != null)
                            e.Graphics.DrawImage(page, FitToPage(page.Size, e.PageBounds));
                    }
                    next++;
                    if (progress != null)
                        progress(next, pageIndexes.Count);
                    // Отмена — только между листами: прерывать наполовину напечатанный лист нечем.
                    e.HasMorePages = next < pageIndexes.Count &&
                                     (cancelled == null || !cancelled());
                };
                doc.Print();
            }
        }
    }
}
