using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>Итог конвертации «PDF → документ».</summary>
    public class ConvertResult
    {
        public int Pages;
        public int PagesWithText;
        internal BackgroundRenderReport BackgroundReport;
    }

    /// <summary>
    /// Оркестрация «цифровой PDF → Word»: разбор страниц (общий слой
    /// <see cref="PdfPageExtraction"/>) и запись .docx (Word COM). Без UI; вызывать в
    /// STA-потоке (требование Word COM).
    /// </summary>
    public static class PdfToWordService
    {
        /// <summary>
        /// Извлекает текст выбранных страниц одного или НЕСКОЛЬКИХ born-digital PDF и пишет их
        /// в один .docx в заданном порядке. Скан без текстового слоя, битый/зашифрованный файл
        /// или занятый выход — <see cref="MergeException"/>. order — страницы (источник + индекс
        /// с нуля) в нужном порядке; страницы могут идти из разных файлов. progress — «сделано/всего»
        /// единиц работы (разбор — первая половина шкалы, запись — вторая); может быть null.
        /// onCommitting — вызывается один раз перед сохранением .docx (точка невозврата: Word наполнен,
        /// дальше отмена уже не сработает); может быть null.
        /// </summary>
        public static ConvertResult Convert(IList<PdfPageRef> order, string outputPath, Action<int, int> progress = null,
            Func<bool> cancelled = null, Action onCommitting = null)
        {
            int writeUnits = order == null ? 0 : order.Count; // страниц к записи — вторая половина шкалы
            if (order != null)
                foreach (PdfPageRef page in order)
                    if (page != null && OutputFile.IsSameFile(page.SourcePath, outputPath))
                        throw new MergeException(Loc.T("err.output.sameSource"));
            Action<int, int> extractCb = progress == null ? null : (Action<int, int>)delegate(int d, int t)
            {
                progress(d, 2 * t); // разбор — первая половина шкалы
            };

            int withText;
            List<PdfPageText> pages = PdfPageExtraction.Load(order, extractCb, cancelled, out withText);

            Action<int, int> writeCb = progress == null ? null : (Action<int, int>)delegate(int d, int t)
            {
                double frac = t > 0 ? (double)d / t : 1.0;
                progress(writeUnits + (int)(frac * writeUnits), 2 * writeUnits);
            };
            using (var output = new AtomicOutput(outputPath))
            {
                WordDocxWriter.Write(pages, output.TempPath, writeCb, cancelled, onCommitting);
                output.Commit();
            }
            return new ConvertResult { Pages = pages.Count, PagesWithText = withText };
        }
    }
}
