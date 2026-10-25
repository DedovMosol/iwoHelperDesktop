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
        /// Область, в которую вписывается страница, в координатах Graphics события печати.
        ///
        /// При OriginAtMargins = false начало координат у ПРИНТЕРА — угол его печатаемой
        /// области (сдвинут на жёсткие поля HardMarginX/Y от края листа), а у ПРЕДПРОСМОТРА —
        /// угол листа (PreviewPrintController рисует в метафайл листа без сдвига). Раньше
        /// страница вписывалась в весь лист (PageBounds) в обоих случаях: на бумаге она уезжала
        /// вправо-вниз на жёсткое поле и обрезалась справа и снизу, а предпросмотр этого не
        /// показывал.
        ///
        /// Теперь лист сужается симметрично на большую из непечатаемых полос по каждой оси —
        /// страница стоит по центру листа и целиком попадает в печатаемую область, — и
        /// переводится в координаты Graphics вычитанием начала: жёстких полей для принтера,
        /// нуля для предпросмотра. Предпросмотр так показывает ровно то, что выйдет на бумагу.
        /// PrintableArea некоторые драйверы отдают «книжной» и для альбомного листа — тогда
        /// стороны меняются местами. Чистая — под тест.
        /// </summary>
        internal static Rectangle PrintTarget(Rectangle pageBounds, RectangleF printableArea,
            float hardMarginX, float hardMarginY, bool preview)
        {
            float pw = printableArea.Width, ph = printableArea.Height;
            float px = printableArea.X, py = printableArea.Y;
            if (pw <= 0 || ph <= 0 || pageBounds.Width <= 0 || pageBounds.Height <= 0)
                return pageBounds; // данных о полях нет — весь лист, как раньше
            if ((pageBounds.Width > pageBounds.Height) != (pw > ph) && pw != ph)
            {
                float t = pw; pw = ph; ph = t;
                t = px; px = py; py = t;
            }
            float left = Math.Max(0, px), top = Math.Max(0, py);
            float right = Math.Max(0, pageBounds.Width - left - pw);
            float bottom = Math.Max(0, pageBounds.Height - top - ph);
            float insetX = Math.Max(left, right), insetY = Math.Max(top, bottom);
            float width = pageBounds.Width - 2 * insetX, height = pageBounds.Height - 2 * insetY;
            if (width < 1 || height < 1)
                return pageBounds; // поля больше листа — данные драйвера не годятся
            float originX = preview ? 0 : hardMarginX, originY = preview ? 0 : hardMarginY;
            return Rectangle.Round(new RectangleF(pageBounds.X + insetX - originX,
                pageBounds.Y + insetY - originY, width, height));
        }

        /// <summary>
        /// <see cref="PrintTarget"/> для текущего листа. Поля читаются у листа события (у каждого
        /// листа могут быть свои настройки), а не у документа. Нет принтера или драйвер не
        /// отвечает — весь лист, как было до учёта полей: напечатать важнее, чем отцентровать.
        /// </summary>
        private static Rectangle PrintTargetOf(PrintDocument doc, PrintPageEventArgs e)
        {
            try
            {
                PageSettings page = e.PageSettings;
                return PrintTarget(e.PageBounds, page.PrintableArea, page.HardMarginX, page.HardMarginY, IsPreview(doc));
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return e.PageBounds;
            }
        }

        /// <summary>
        /// Рисует ли документ сейчас в предпросмотр. PrintPreviewControl оборачивает свой
        /// PreviewPrintController в PrintControllerWithStatusDialog, и обёртка отвечает за
        /// вложенный — поэтому спрашиваем свойство, а не тип контроллера.
        /// </summary>
        internal static bool IsPreview(PrintDocument doc)
        {
            return doc != null && doc.PrintController != null && doc.PrintController.IsPreview;
        }

        /// <summary>Сколько ждать освобождения памяти перед повторным рендером листа.</summary>
        private const int MemoryWaitMs = 3000;

        /// <summary>
        /// Напечатать указанные страницы. settings уже выбраны пользователем в диалоге печати.
        /// Отмена — своя (cancelled, между листами) или самой системы печати — прерывает всё
        /// задание OperationCanceledException; страница, которую не удалось отрисовать, —
        /// MergeException. Пустой лист вместо неё и «успех» в итоге были бы враньём.
        ///
        /// Передан cancelled — у вызывающего свои прогресс и «Отмена», и штатное окно состояния
        /// печати не показывается: вторая «Отмена» на ту же операцию только путала бы. Без него
        /// окно состояния остаётся — другого способа прервать печать у человека нет.
        /// names — имена страниц для человека (см. <see cref="CreateJob"/>).
        /// </summary>
        public static void Print(IList<PdfPageRef> pages, PrinterSettings settings,
            Action<int, int> progress = null, Func<bool> cancelled = null, IList<string> names = null)
        {
            using (var owned = new PdfThumbnailRenderer())
            using (PrintDocument doc = CreateJob(pages, settings, progress, cancelled, RenderWith(owned), true, names))
            {
                if (cancelled != null)
                    doc.PrintController = new StandardPrintController();
                // Исключение из обработчика листа (страница не отрисовалась, отмена) выходит
                // отсюда, а контроллер печати при этом прерывает задание (AbortDoc): в очереди
                // принтера не остаётся ни пустых листов, ни хвоста отменённого задания.
                RunJob(doc);
            }
        }

        /// <summary>
        /// Выполнить задание контроллером документа и не выдать прерванное за удавшееся:
        /// задание, закрытое контроллером как отменённое, — OperationCanceledException.
        /// </summary>
        internal static void RunJob(PrintDocument doc)
        {
            var watch = new CancelWatch(doc.PrintController);
            doc.PrintController = watch;
            doc.Print();
            if (watch.Canceled)
                throw new OperationCanceledException();
        }

        /// <summary>
        /// Обёртка контроллера печати, запоминающая, закрыто ли задание как отменённое.
        /// PrintDocument.Print об отмене молчит: «Отмена» в штатном окне состояния
        /// (PrintControllerWithStatusDialog) и отказ в окне «Сохранить как» у печати в файл
        /// (StartDoc → ERROR_CANCELLED у StandardPrintController) лишь ставят e.Cancel —
        /// задание стирается (AbortDoc), а Print возвращается как после успеха. Событие
        /// документа EndPrint для этого не годится: его зовут раньше, чем в e.Cancel попадает
        /// отмена из цикла листов. Цикл листов ведёт ВНЕШНИЙ контроллер, поэтому обёртка
        /// зовёт и базовые OnStartPrint/OnEndPrint (по ним размечаются листы) — так же устроен
        /// и сам PrintControllerWithStatusDialog.
        /// </summary>
        private sealed class CancelWatch : PrintController
        {
            private readonly PrintController _inner;

            internal CancelWatch(PrintController inner)
            {
                _inner = inner;
            }

            internal bool Canceled { get; private set; }

            public override bool IsPreview { get { return _inner.IsPreview; } }

            public override void OnStartPrint(PrintDocument document, PrintEventArgs e)
            {
                base.OnStartPrint(document, e);
                _inner.OnStartPrint(document, e);
            }

            public override Graphics OnStartPage(PrintDocument document, PrintPageEventArgs e)
            {
                return _inner.OnStartPage(document, e);
            }

            public override void OnEndPage(PrintDocument document, PrintPageEventArgs e)
            {
                _inner.OnEndPage(document, e);
            }

            public override void OnEndPrint(PrintDocument document, PrintEventArgs e)
            {
                // Читаем ДО внутреннего контроллера: по этому же флагу он выбирает AbortDoc или
                // EndDoc. Окно состояния ставит e.Cancel и после EndDoc («Отмена», когда всё уже
                // ушло в очередь) — такое задание напечатано, и отменой оно не считается.
                Canceled = e.Cancel;
                try { _inner.OnEndPrint(document, e); }
                finally { base.OnEndPrint(document, e); }
            }
        }

        /// <summary>Рендер листа под печать этим рендерером (общий для печати и предпросмотра).</summary>
        private static Func<PdfPageRef, int, BudgetedBitmap> RenderWith(PdfThumbnailRenderer renderer)
        {
            return delegate(PdfPageRef page, int width)
            {
                return renderer.RenderOwned(page.SourcePath, page.PageIndex, width, 0,
                    RasterBudget.DefaultRenderPixels);
            };
        }

        /// <summary>
        /// Собрать задание печати, ничего не отправляя на принтер. Нужно предпросмотру: он
        /// показывает ровно тот <see cref="PrintDocument"/>, который потом и печатается, —
        /// иначе показанное и напечатанное могли бы разойтись, а это худший вид предпросмотра.
        ///
        /// renderer отдаётся наружу, потому что страницы рисуются лениво, по событию: освободи
        /// его здесь — и предпросмотру нечем будет рисовать. Освобождает вызывающий, вместе с
        /// документом. Окно предпросмотра строит тоже вызывающий: интерфейса в сервисе нет.
        ///
        /// Это задание НЕ строгое: неотрисованная страница остаётся пустым листом. Штатный
        /// контрол предпросмотра рисует его из обработчика отрисовки окна, и исключение оттуда
        /// уронило бы приложение, а не показало ошибку. Строгая печать — в <see cref="Print"/>.
        /// </summary>
        public static PrintDocument CreateDocument(IList<PdfPageRef> pages, PrinterSettings settings,
            Action<int, int> progress, Func<bool> cancelled, out IDisposable renderer)
        {
            var owned = new PdfThumbnailRenderer();
            try
            {
                PrintDocument doc = CreateJob(pages, settings, progress, cancelled, RenderWith(owned), false);
                renderer = owned;
                return doc;
            }
            catch
            {
                owned.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Задание печати поверх произвольного рендера листа (render: страница и ширина →
        /// растр с бюджетом или null). Рендер отдан параметром, чтобы строгость печати
        /// проверялась тестом без принтера и без WinRT.
        ///
        /// strict — настоящая печать: страница, которую не удалось отрисовать и со второй
        /// попытки, прерывает задание MergeException, а отмена между листами —
        /// OperationCanceledException. Прежде на месте такой страницы выходил пустой лист, а
        /// отменённое задание заканчивалось «✓ отправлено на печать».
        ///
        /// names — как страницы называет окно (вровень с pages): вместо обёртки вставленного
        /// листа или картинки — «пустая страница» или сама картинка. По ним названы задание в
        /// очереди принтера и страница в сообщении об ошибке. null — имена файлов страниц.
        /// </summary>
        internal static PrintDocument CreateJob(IList<PdfPageRef> pages, PrinterSettings settings,
            Action<int, int> progress, Func<bool> cancelled, Func<PdfPageRef, int, BudgetedBitmap> render,
            bool strict, IList<string> names = null)
        {
            if (pages == null || pages.Count == 0)
                throw new MergeException(Loc.T("err.export.noPages"));
            if (render == null)
                throw new ArgumentNullException("render");

            var doc = new PrintDocument();
            try
            {
                if (settings != null)
                    doc.PrinterSettings = settings;
                doc.DocumentName = NameOf(pages, names, 0);
                doc.OriginAtMargins = false; // печатаем по всей доступной области, а не по полям (см. PrintTarget)

                int next = 0; // индекс в списке страниц, а не номер страницы документа
                doc.PrintPage += delegate(object s, PrintPageEventArgs e)
                {
                    PdfPageRef ref_ = pages[next];
                    Rectangle target = PrintTargetOf(doc, e);
                    // Ширину рендера считаем от РЕАЛЬНОЙ области печати листа, а не страницы
                    // PDF: так одинаково печатаются и A4, и нестандартные вставки.
                    int width = Math.Max(1, target.Width * RenderDpi / 100);
                    BudgetedBitmap rendered = render(ref_, width);
                    if (rendered == null && strict)
                    {
                        // Чаще всего не хватило общего бюджета памяти: его держат миниатюры
                        // открытых окон. Дадим им отпустить растр и попробуем ещё раз — лишь
                        // потом признаём, что страницу не напечатать.
                        WaitForMemoryRelease(MemoryWaitMs);
                        rendered = render(ref_, width);
                        if (rendered == null) // номер страницы — в её файле: он и назван в сообщении
                            throw new MergeException(string.Format(Loc.T("common.err.printPageFailed"),
                                ref_.PageIndex + 1, NameOf(pages, names, next)));
                    }
                    using (rendered)
                    {
                        Bitmap page = rendered == null ? null : rendered.Bitmap;
                        if (page != null)
                        {
                            // Поворот, назначенный пользователем в сетке, обязан попасть и на
                            // бумагу: иначе на экране страница стоит правильно, а печатается боком.
                            if (ref_.Rotation != 0)
                                page.RotateFlip(PageRotation.FlipFor(ref_.Rotation));
                            e.Graphics.DrawImage(page, FitToPage(page.Size, target));
                        }
                    }
                    next++;
                    if (progress != null)
                        progress(next, pages.Count);
                    e.HasMorePages = next < pages.Count;
                    // Отмена — только между листами: прерывать наполовину напечатанный лист нечем.
                    // Исключение, а не HasMorePages = false: иначе задание закрывалось штатно и
                    // отменённая печать отчитывалась как удавшаяся.
                    if (e.HasMorePages && cancelled != null && cancelled())
                        throw new OperationCanceledException();
                };
                // Предпросмотр листает документ по кругу и начинает заново — счётчик обязан
                // сбрасываться, иначе второй показ выводит хвост списка вместо начала.
                doc.BeginPrint += delegate { next = 0; };
                return doc;
            }
            catch
            {
                doc.Dispose();
                throw;
            }
        }

        /// <summary>Имя страницы index для человека: из names, если окно его дало, иначе имя её файла.</summary>
        private static string NameOf(IList<PdfPageRef> pages, IList<string> names, int index)
        {
            string name = names != null && index < names.Count ? names[index] : null;
            return System.IO.Path.GetFileName(name ?? pages[index].SourcePath);
        }

        /// <summary>
        /// Подождать, пока общий бюджет памяти кто-нибудь пополнит (не дольше timeoutMs).
        /// Никто ничего не держит — ждать нечего, возвращаемся сразу. Только с фонового потока.
        /// </summary>
        private static void WaitForMemoryRelease(int timeoutMs)
        {
            if (PdfMemoryBudget.Used <= 0)
                return;
            using (var released = new System.Threading.ManualResetEventSlim(false))
            {
                Action onReleased = delegate { released.Set(); };
                PdfMemoryBudget.MemoryReleased += onReleased;
                try { released.Wait(timeoutMs); }
                finally { PdfMemoryBudget.MemoryReleased -= onReleased; }
            }
        }
    }
}
