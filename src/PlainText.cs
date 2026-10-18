using System.Collections.Generic;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Сборка простого текста страницы из результата разбора PDF.
    ///
    /// Наивное «склеить абзацы» теряет таблицы: разбор УБИРАЕТ слова таблиц из потока
    /// абзацев и складывает их в ячейки (см. PdfTextExtract), поэтому абзацы и таблицы
    /// приходится сводить обратно — в порядке чтения страницы, как их видит читатель.
    ///
    /// Ячейки строки разделяются табуляцией: так строка вставляется в Excel или Calc
    /// готовой таблицей, а не одной склеенной ячейкой. Страницы разделяются символом
    /// перевода страницы (U+000C) — так же, как это делает pdftotext, поэтому результат
    /// понимают и текстовые редакторы, и скрипты, привыкшие к этому разделителю.
    ///
    /// Чистые функции без ввода-вывода — под тестами.
    /// </summary>
    public static class PlainText
    {
        /// <summary>Разделитель страниц — перевод страницы, как у pdftotext.</summary>
        public const string PageSeparator = "\f";

        /// <summary>
        /// Текст одной страницы: абзацы и таблицы в порядке чтения страницы
        /// (<see cref="PageBlocks.OrderedItems"/>: XY-разрез, левая колонка целиком раньше
        /// правой). Простая сортировка по верху чередовала строки двух колонок
        /// («Л1, П1, Л2, П2») — так же, как порядок чтения, текст обязан идти и в .txt, и в
        /// Сравнении. Пустые блоки и картинки отбрасываются. Чистая — под тест.
        /// </summary>
        public static string Page(PdfPageText page)
        {
            if (page == null)
                return "";
            var parts = new List<string>();
            foreach (PageBlocks.PageItem item in PageBlocks.OrderedItems(page))
            {
                if (!item.IsBand)
                {
                    AddBlock(parts, item.Single);
                    continue;
                }
                // Полоса «бок о бок» — колонки по очереди, как их читает человек.
                foreach (List<PageBlocks.Block> column in item.Columns)
                    foreach (PageBlocks.Block block in column)
                        AddBlock(parts, block);
            }
            return string.Join("\n\n", parts.ToArray());
        }

        private static void AddBlock(List<string> parts, PageBlocks.Block block)
        {
            string text;
            if (block.Table != null)
                text = Table(block.Table);
            else if (block.Paragraph != null)
                text = block.Paragraph.Text ?? "";
            else
                return; // картинка — в тексте её нет
            if (text.Trim().Length > 0)
                parts.Add(text);
        }

        /// <summary>
        /// Текст документа: страницы через <see cref="PageSeparator"/>. Пустые страницы
        /// сохраняются как пустые — нумерация страниц в результате не должна съезжать.
        /// Чистая — под тест.
        /// </summary>
        public static string Document(IList<PdfPageText> pages)
        {
            if (pages == null || pages.Count == 0)
                return "";
            var sb = new StringBuilder();
            for (int i = 0; i < pages.Count; i++)
            {
                if (i > 0)
                    sb.Append('\n').Append(PageSeparator).Append('\n');
                sb.Append(Page(pages[i]));
            }
            return sb.ToString();
        }

        /// <summary>
        /// Таблица строками, ячейки через табуляцию. Накрытые объединением ячейки дают пустое
        /// поле, иначе колонки разъехались бы и вставка в таблицу стала бы бесполезной.
        /// Переводы строк внутри ячейки заменяются пробелом — строка таблицы обязана
        /// остаться одной строкой текста. Чистая — под тест.
        /// </summary>
        internal static string Table(OcrTable table)
        {
            if (table == null || table.Rows == null)
                return "";
            var lines = new List<string>(table.Rows.Count);
            foreach (OcrTableRow row in table.Rows)
            {
                var cells = new List<string>(row.Cells.Count);
                foreach (OcrTableCell cell in row.Cells)
                    cells.Add(cell.Covered ? "" : CellText(cell));
                lines.Add(string.Join("\t", cells.ToArray()).TrimEnd('\t'));
            }
            return string.Join("\n", lines.ToArray());
        }

        private static string CellText(OcrTableCell cell)
        {
            if (cell == null || cell.Paragraphs == null || cell.Paragraphs.Count == 0)
                return "";
            var parts = new List<string>(cell.Paragraphs.Count);
            foreach (OcrParagraph p in cell.Paragraphs)
            {
                string t = p.Text;
                if (!string.IsNullOrEmpty(t))
                    parts.Add(t.Replace('\r', ' ').Replace('\n', ' ').Trim());
            }
            return string.Join(" ", parts.ToArray()).Trim();
        }
    }
}
