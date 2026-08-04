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
        public static void Print(IList<PdfPageRef> pages, PrinterSettings settings,
            Action<int, int> progress = null, Func<bool> cancelled = null)
        {
            IDisposable renderer;
            using (PrintDocument doc = CreateDocument(pages, settings, progress, cancelled, out renderer))
            using (renderer)
                doc.Print();
        }

        /// <summary>
        /// Собрать задание печати, ничего не отправляя на принтер. Нужно предпросмотру: он
        /// показывает ровно тот <see cref="PrintDocument"/>, который потом и печатается, —
        /// иначе показанное и напечатанное могли бы разойтись, а это худший вид предпросмотра.
        ///
        /// renderer отдаётся наружу, потому что страницы рисуются лениво, по событию: освободи
        /// его здесь — и предпросмотру нечем будет рисовать. Освобождает вызывающий, вместе с
        /// документом. Окно предпросмотра строит тоже вызывающий: интерфейса в сервисе нет.
        /// </summary>
        public static PrintDocument CreateDocument(IList<PdfPageRef> pages, PrinterSettings settings,
            Action<int, int> progress, Func<bool> cancelled, out IDisposable renderer)
        {
            if (pages == null || pages.Count == 0)
                throw new MergeException(Loc.T("err.export.noPages"));

            var owned = new PdfThumbnailRenderer();
            renderer = owned;
            var doc = new PrintDocument();
            try
            {
                if (settings != null)
                    doc.PrinterSettings = settings;
                doc.DocumentName = System.IO.Path.GetFileName(pages[0].SourcePath);
                doc.OriginAtMargins = false; // печатаем по всей доступной области, а не по полям

                int next = 0; // индекс в списке страниц, а не номер страницы документа
                doc.PrintPage += delegate(object s, PrintPageEventArgs e)
                {
                    PdfPageRef ref_ = pages[next];
                    // Ширину рендера считаем от РЕАЛЬНОГО размера листа принтера, а не страницы
                    // PDF: так одинаково печатаются и A4, и нестандартные вставки.
                    int width = Math.Max(1, e.PageBounds.Width * RenderDpi / 100);
                    using (Bitmap page = owned.Render(ref_.SourcePath, ref_.PageIndex, width))
                    {
                        if (page != null)
                        {
                            // Поворот, назначенный пользователем в сетке, обязан попасть и на
                            // бумагу: иначе на экране страница стоит правильно, а печатается боком.
                            if (ref_.Rotation != 0)
                                page.RotateFlip(PageRotation.FlipFor(ref_.Rotation));
                            e.Graphics.DrawImage(page, FitToPage(page.Size, e.PageBounds));
                        }
                    }
                    next++;
                    if (progress != null)
                        progress(next, pages.Count);
                    // Отмена — только между листами: прерывать наполовину напечатанный лист нечем.
                    e.HasMorePages = next < pages.Count && (cancelled == null || !cancelled());
                };
                // Предпросмотр листает документ по кругу и начинает заново — счётчик обязан
                // сбрасываться, иначе второй показ выводит хвост списка вместо начала.
                doc.BeginPrint += delegate { next = 0; };
                return doc;
            }
            catch
            {
                doc.Dispose();
                owned.Dispose();
                throw;
            }
        }
    }
}
