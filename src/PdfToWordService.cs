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
        /// progress — «сделано/всего» единиц работы (извлечение и запись считаются двумя
        /// проходами по страницам: всего 2×страниц), для полосы прогресса; может быть null.
        /// </summary>
        public static ConvertResult Convert(string sourcePath, string outputPath, Action<int, int> progress = null)
        {
            // Два прохода (извлечение + запись) в одну непрерывную шкалу 0..2N.
            Action<int, int> extract = progress == null ? null : (Action<int, int>)delegate(int d, int t) { progress(d, 2 * t); };
            Action<int, int> write = progress == null ? null : (Action<int, int>)delegate(int d, int t) { progress(t + d, 2 * t); };

            List<PdfPageText> pages = PdfTextExtract.Extract(sourcePath, extract);
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
    }
}
