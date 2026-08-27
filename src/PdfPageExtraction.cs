using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Первая половина любой конвертации «PDF → документ»: набор ссылок на страницы (возможно
    /// из РАЗНЫХ файлов, в заданном пользователем порядке) → список разобранных страниц. Что
    /// именно писать потом — .docx, .pptx или что-то ещё — этот слой не знает и знать не должен.
    ///
    /// Здесь же живёт единственный отказ, общий для всех выводов: файл без извлекаемого текста
    /// (скан) отклоняется понятной ошибкой, потому что переносить из него нечего — распознавание
    /// сканов (OCR) появится позже, и это та самая точка, где добавится ветвь «скан → OCR».
    /// Без UI; чистые части (<see cref="BuildRotations"/>, <see cref="Assemble"/>,
    /// <see cref="HasExtractableContent"/>) — под юнит-тестами.
    /// </summary>
    internal static class PdfPageExtraction
    {
        /// <summary>
        /// Разобрать выбранные страницы одного или НЕСКОЛЬКИХ born-digital PDF и вернуть их в
        /// заданном порядке. order — страницы (источник + индекс с нуля); страницы могут идти из
        /// разных файлов и повторяться. progress — «разобрано/всего» единиц работы, где всего =
        /// order.Count (вызывающий сам укладывает это в свою шкалу); может быть null. cancelled —
        /// кооперативная отмена между страницами. pagesWithText — сколько страниц несут текст
        /// (у остальных на выходе будут только картинки). Скан без текстового слоя, битый или
        /// зашифрованный файл — <see cref="MergeException"/>.
        /// </summary>
        public static List<PdfPageText> Load(IList<PdfPageRef> order, Action<int, int> progress,
            Func<bool> cancelled, out int pagesWithText, PageLayoutMode mode = PageLayoutMode.Document)
        {
            if (order == null || order.Count == 0)
                throw new MergeException(Loc.T("err.ocr.noPages"));

            // Уникальные источники в порядке первого появления (каждый извлекаем ОДИН раз).
            var sources = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PdfPageRef r in order)
                if (r != null && r.SourcePath != null && seen.Add(r.SourcePath))
                    sources.Add(r.SourcePath);

            int units = order.Count;          // единиц работы этого слоя — страниц к разбору
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

            // Повороты страниц из сетки: карта по источнику (индекс страницы → градусы).
            // Дубли одной страницы делят ОДНО извлечение — поворот берётся у первого
            // экземпляра в порядке (устройство кэша ниже, второй дубль наследует).
            Dictionary<string, int[]> rotationsBySource = BuildRotations(order);

            // Извлечь текст каждого источника (весь файл); кэш по пути. Прогресс — по долям
            // источников (внутри источника — по его страницам).
            var bySource = new Dictionary<string, List<PdfPageText>>(StringComparer.OrdinalIgnoreCase);
            for (int si = 0; si < totalSources; si++)
            {
                string src = sources[si];
                int idx = si;
                Action<int, int> extractCb = progress == null ? null : (Action<int, int>)delegate(int d, int t)
                {
                    double frac = t > 0 ? (double)d / t : 1.0;
                    double overall = totalSources > 0 ? (idx + frac) / totalSources : 1.0;
                    progress((int)(overall * units), units);
                };
                int[] rotations = null;
                if (rotationsBySource != null)
                    rotationsBySource.TryGetValue(src, out rotations);
                bySource[src] = PdfTextExtract.Extract(src, extractCb, rotations, cancelled, mode);
            }

            var pages = new List<PdfPageText>(order.Count);
            foreach (PdfPageRef reference in order)
            {
                List<PdfPageText> sourcePages;
                if (reference == null || string.IsNullOrEmpty(reference.SourcePath) ||
                    !bySource.TryGetValue(reference.SourcePath, out sourcePages) ||
                    reference.PageIndex < 0 || reference.PageIndex >= sourcePages.Count)
                {
                    string name = reference == null ? "?" :
                        Path.GetFileName(reference.SourcePath ?? "?");
                    int number = reference == null ? 0 : reference.PageIndex + 1;
                    throw new MergeException(string.Format(Loc.T("err.pdf.pageGone"),
                        name, number));
                }
                pages.Add(sourcePages[reference.PageIndex]);
            }

            pagesWithText = 0;
            foreach (PdfPageText page in pages)
                if (HasExtractableContent(page))
                    pagesWithText++;

            if (pagesWithText == 0)
                throw new MergeException(Loc.T("err.ocr.scanned"));
            return pages;
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
        /// Карта поворотов по источникам из ссылок порядка: SourcePath → массив по индексу
        /// страницы (градусы по часовой). Первый экземпляр страницы в порядке решает (в том
        /// числе решает «без поворота»); null — поворотов нет вовсе. Чистая — под тест.
        /// </summary>
        internal static Dictionary<string, int[]> BuildRotations(IList<PdfPageRef> order)
        {
            Dictionary<string, int[]> result = null;
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // «src|idx»: первый экземпляр решает
            foreach (PdfPageRef r in order)
            {
                if (r == null || r.SourcePath == null || r.PageIndex < 0)
                    continue;
                if (!seen.Add(r.SourcePath.ToLowerInvariant() + "|" + r.PageIndex))
                    continue;
                if (r.Rotation == 0)
                    continue;
                if (result == null)
                    result = new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase);
                int[] map;
                if (!result.TryGetValue(r.SourcePath, out map) || map.Length <= r.PageIndex)
                {
                    var grown = new int[r.PageIndex + 1];
                    if (map != null)
                        Array.Copy(map, grown, map.Length);
                    result[r.SourcePath] = map = grown;
                }
                map[r.PageIndex] = r.Rotation;
            }
            return result;
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
