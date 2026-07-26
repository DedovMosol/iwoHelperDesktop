using System.Collections.Generic;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Сборка простого текста страницы из результата разбора PDF.
    ///
    /// Наивное «склеить абзацы» теряет таблицы: разбор УБИРАЕТ слова таблиц из потока
    /// абзацев и складывает их в ячейки (см. PdfTextExtract), поэтому абзацы и таблицы
    /// приходится сводить обратно — по вертикали страницы, как их видит читатель.
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

        /// <summary>Блок страницы в порядке чтения: либо абзац, либо таблица.</summary>
        private struct Block
        {
            public double Top;
            public double Left;
            public OcrParagraph Paragraph;
            public OcrTable Table;
        }

        /// <summary>
        /// Текст одной страницы: абзацы и таблицы в порядке чтения (сверху вниз, при равном
        /// верхе — слева направо). Пустые блоки отбрасываются. Чистая — под тест.
        /// </summary>
        public static string Page(PdfPageText page)
        {
            if (page == null)
                return "";
            var blocks = new List<Block>();
            if (page.Paragraphs != null)
                foreach (OcrParagraph p in page.Paragraphs)
                    blocks.Add(new Block { Top = p.TopPt, Left = p.LeftPt, Paragraph = p });
            if (page.Tables != null)
                foreach (OcrTable t in page.Tables)
                    blocks.Add(new Block { Top = t.TopPt, Left = t.LeftPt, Table = t });

            // Ось Y в PDF направлена ВВЕРХ: больший Top — выше на странице, значит раньше.
            blocks.Sort(delegate(Block a, Block b)
            {
                int byTop = b.Top.CompareTo(a.Top);
                return byTop != 0 ? byTop : a.Left.CompareTo(b.Left);
            });

            var parts = new List<string>(blocks.Count);
            foreach (Block b in blocks)
            {
                string text = b.Table != null ? Table(b.Table) : (b.Paragraph.Text ?? "");
                if (text.Trim().Length > 0)
                    parts.Add(text);
            }
            return string.Join("\n\n", parts.ToArray());
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
