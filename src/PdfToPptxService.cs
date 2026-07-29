using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Оркестрация «цифровой PDF → PowerPoint»: разбор страниц (общий слой
    /// <see cref="PdfPageExtraction"/>), отрисовка подложек без текста
    /// (<see cref="PageBackgrounds"/>) и запись .pptx (<see cref="PptxWriter"/>).
    ///
    /// Без UI и без COM — PowerPoint для конвертации не нужен, поэтому и STA-поток не
    /// требуется, в отличие от пути в Word. Результат пишется во временный файл рядом с
    /// целью и переименовывается последним действием: отмена или сбой не оставляют
    /// половину презентации под именем готовой.
    /// </summary>
    public static class PdfToPptxService
    {
        /// <summary>
        /// Переносит выбранные страницы одного или НЕСКОЛЬКИХ born-digital PDF в один .pptx
        /// в заданном порядке. Скан без текстового слоя, битый или зашифрованный файл, занятый
        /// выход — <see cref="MergeException"/>. progress — «сделано/всего» единиц работы:
        /// разбор, подложки и запись делят шкалу натрое. onCommitting — перед записью файла
        /// (точка невозврата: дальше отмена уже ничего не изменит).
        /// </summary>
        public static ConvertResult Convert(IList<PdfPageRef> order, string outputPath,
            Action<int, int> progress = null, Func<bool> cancelled = null, Action onCommitting = null)
        {
            int units = order == null ? 0 : order.Count;

            int withText;
            List<PdfPageText> pages = PdfPageExtraction.Load(order,
                Scaled(progress, 0, units), cancelled, out withText, PageLayoutMode.Slide);

            // Подложка — необязательная роскошь: без Ghostscript слой молча пропускается, и
            // презентация получается из одного текста (см. PageBackgrounds).
            Background[] backgrounds = PageBackgrounds.Render(order, Scaled(progress, 1, units), cancelled);

            string temp = outputPath + ".tmp";
            try
            {
                PptxWriter.Write(pages, temp, DateTime.UtcNow, Scaled(progress, 2, units), cancelled,
                    onCommitting, backgrounds);
                Replace(temp, outputPath);
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                throw new MergeException(string.Format(Loc.T("err.pptx.writeFailed"),
                    Path.GetFileName(outputPath), DiskSpace.Describe(ex, outputPath)));
            }
            finally
            {
                try { if (File.Exists(temp)) File.Delete(temp); } catch { }
            }
            return new ConvertResult { Pages = pages.Count, PagesWithText = withText };
        }

        /// <summary>
        /// Колбэк прогресса для одной трети шкалы. Каждый слой считает свои единицы работы от
        /// нуля и не знает, какой он по счёту, — укладывает их в общую шкалу вызывающий.
        /// </summary>
        private static Action<int, int> Scaled(Action<int, int> progress, int stage, int units)
        {
            if (progress == null || units <= 0)
                return null;
            return delegate(int done, int total)
            {
                double frac = total > 0 ? (double)done / total : 1.0;
                progress(stage * units + (int)(frac * units), 3 * units);
            };
        }

        /// <summary>
        /// Поставить готовый файл на место результата. Перезапись существующего — через
        /// удаление: File.Move на занятый файл откажет, а понятное сообщение об этом даст
        /// обёртка вызывающего.
        /// </summary>
        private static void Replace(string temp, string outputPath)
        {
            if (File.Exists(outputPath))
                File.Delete(outputPath);
            File.Move(temp, outputPath);
        }
    }
}
