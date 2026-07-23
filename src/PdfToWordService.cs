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
        /// Извлекает текст выбранных страниц одного или НЕСКОЛЬКИХ born-digital PDF и пишет их
        /// в один .docx в заданном порядке. Скан без текстового слоя, битый/зашифрованный файл
        /// или занятый выход — <see cref="MergeException"/>. order — страницы (источник + индекс
        /// с нуля) в нужном порядке; страницы могут идти из разных файлов. progress — «сделано/всего»
        /// единиц работы (извлечение источников — первая половина шкалы, запись — вторая); может быть null.
        /// </summary>
        public static ConvertResult Convert(IList<PdfPageRef> order, string outputPath, Action<int, int> progress = null)
        {
            if (order == null || order.Count == 0)
                throw new MergeException(Loc.T("err.ocr.noPages"));

            // Уникальные источники в порядке первого появления (каждый извлекаем ОДИН раз).
            var sources = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PdfPageRef r in order)
                if (r != null && r.SourcePath != null && seen.Add(r.SourcePath))
                    sources.Add(r.SourcePath);

            int writeUnits = order.Count;    // страниц к записи — вторая половина шкалы прогресса
            int totalSources = sources.Count;

            // Быстрая проба: если НИ В ОДНОМ источнике нет текста — «скан» отклоняется сразу,
            // без дорогого извлечения (полностраничные растры скана не декодируются вовсе).
            // Смешанный ввод (текст + скан-файл или скан-страницы) идёт полным путём: страницы
            // без текста переносятся полностраничными картинками, как и раньше.
            bool anyText = false;
            foreach (string src in sources)
                if (PdfTextExtract.AnyPageHasText(src))
                {
                    anyText = true;
                    break;
                }
            if (!anyText)
                throw new MergeException(Loc.T("err.ocr.scanned"));

            // Извлечь текст каждого источника (весь файл); кэш по пути. Прогресс извлечения —
            // первая половина шкалы, по долям источников (внутри источника — по его страницам).
            var bySource = new Dictionary<string, List<PdfPageText>>(StringComparer.OrdinalIgnoreCase);
            for (int si = 0; si < totalSources; si++)
            {
                string src = sources[si];
                int idx = si;
                Action<int, int> extractCb = progress == null ? null : (Action<int, int>)delegate(int d, int t)
                {
                    double frac = t > 0 ? (double)d / t : 1.0;
                    double overall = totalSources > 0 ? (idx + frac) / totalSources : 1.0;
                    progress((int)(overall * writeUnits), 2 * writeUnits);
                };
                bySource[src] = PdfTextExtract.Extract(src, extractCb);
            }

            List<PdfPageText> pages = Assemble(bySource, order);

            int withText = 0;
            foreach (PdfPageText page in pages)
                if (HasExtractableContent(page))
                    withText++;

            if (withText == 0)
                throw new MergeException(Loc.T("err.ocr.scanned"));

            Action<int, int> writeCb = progress == null ? null : (Action<int, int>)delegate(int d, int t)
            {
                double frac = t > 0 ? (double)d / t : 1.0;
                progress(writeUnits + (int)(frac * writeUnits), 2 * writeUnits);
            };
            WordDocxWriter.Write(pages, outputPath, writeCb);
            return new ConvertResult { Pages = pages.Count, PagesWithText = withText };
        }

        /// <summary>Есть ли на странице извлекаемый текст: абзацы вне таблиц ИЛИ текст в ячейках таблиц.</summary>
        internal static bool HasExtractableContent(PdfPageText page)
        {
            if (page.Paragraphs != null && page.Paragraphs.Count > 0)
                return true;
            if (page.Tables != null)
                foreach (OcrTable table in page.Tables)
                    foreach (OcrTableRow row in table.Rows)
                        foreach (OcrTableCell cell in row.Cells)
                            if (cell.Paragraphs != null && cell.Paragraphs.Count > 0)
                                return true;
            return false;
        }

        /// <summary>
        /// Собрать страницы в заданном порядке из извлечённых по источникам (SourcePath → страницы).
        /// Каждая ссылка order берёт страницу своего файла по индексу; несуществующие источник/индекс
        /// пропускаются (защита). Страницы могут чередоваться из разных файлов. Чистая — под тест.
        /// </summary>
        internal static List<PdfPageText> Assemble(Dictionary<string, List<PdfPageText>> bySource, IList<PdfPageRef> order)
        {
            var result = new List<PdfPageText>(order.Count);
            foreach (PdfPageRef r in order)
            {
                List<PdfPageText> src;
                if (r != null && r.SourcePath != null && bySource.TryGetValue(r.SourcePath, out src)
                    && r.PageIndex >= 0 && r.PageIndex < src.Count)
                    result.Add(src[r.PageIndex]);
            }
            return result;
        }
    }
}
