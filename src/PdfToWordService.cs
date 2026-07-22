using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>Итог конвертации «PDF → Word».</summary>
    public class ConvertResult
    {
        public int Pages;
        public int PagesWithText;
    }

    /// <summary>
    /// Оркестрация «цифровой PDF → Word»: извлечение текстового слоя (PdfPig) и запись
    /// .docx (Word COM). Без UI; вызывать в STA-потоке (требование Word COM). PDF без
    /// извлекаемого текста (скан) отсекается понятной ошибкой — распознавание сканов
    /// (OCR) появится позже; это единая точка, где позже добавится ветвь «скан → OCR».
    /// </summary>
    public static class PdfToWordService
    {
        /// <summary>
        /// Извлекает текст born-digital PDF и пишет .docx. Скан без текстового слоя,
        /// битый/зашифрованный файл или занятый выход — <see cref="MergeException"/>.
        /// pageOrder — индексы страниц (с нуля) в нужном порядке/подмножестве; null — весь
        /// документ в исходном порядке. progress — «сделано/всего» единиц работы (проход
        /// извлечения по всем страницам + проход записи по выбранным), для полосы; может быть null.
        /// </summary>
        public static ConvertResult Convert(string sourcePath, string outputPath, IList<int> pageOrder = null, Action<int, int> progress = null)
        {
            // Прогресс в одну непрерывную шкалу: извлечение всех M страниц + запись выбранных N.
            int selCount = pageOrder != null ? pageOrder.Count : -1; // -1 — «все» (N станет = M)
            int extractTotal = 0;                                    // M, узнаётся в ходе извлечения
            Action<int, int> extract = progress == null ? null : (Action<int, int>)delegate(int d, int t)
            {
                extractTotal = t;
                progress(d, t + (selCount < 0 ? t : selCount));
            };
            Action<int, int> write = progress == null ? null : (Action<int, int>)delegate(int d, int t)
            {
                progress(extractTotal + d, extractTotal + t);
            };

            List<PdfPageText> all = PdfTextExtract.Extract(sourcePath, extract);
            List<PdfPageText> pages = SelectPages(all, pageOrder);

            int withText = 0;
            foreach (PdfPageText page in pages)
                if (page.Paragraphs != null && page.Paragraphs.Count > 0)
                    withText++;

            if (withText == 0)
                throw new MergeException(
                    "В этом PDF нет извлекаемого текста — похоже, это отсканированный документ (изображение). " +
                    "Поддержка отсканированных документов в настоящее время недоступна.");

            WordDocxWriter.Write(pages, outputPath, write);
            return new ConvertResult { Pages = pages.Count, PagesWithText = withText };
        }

        /// <summary>
        /// Отобрать/переупорядочить извлечённые страницы по индексам (с нуля). null — вернуть все
        /// как есть. Индексы вне диапазона пропускаются (защита). Чистая логика — под тест.
        /// </summary>
        internal static List<PdfPageText> SelectPages(List<PdfPageText> all, IList<int> order)
        {
            if (order == null)
                return all;
            var result = new List<PdfPageText>(order.Count);
            foreach (int i in order)
                if (i >= 0 && i < all.Count)
                    result.Add(all[i]);
            return result;
        }
    }
}
