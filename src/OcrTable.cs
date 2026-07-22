using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Ячейка восстановленной таблицы: содержимое (абзацы в порядке чтения, как в обычном
    /// тексте — <see cref="OcrLayout"/>) и объединение. Ячейка-«хозяин» несёт ColSpan/RowSpan
    /// и текст; накрытые объединением позиции помечены <see cref="Covered"/> и при записи в
    /// Word пропускаются (их поглощает Merge хозяина).
    /// </summary>
    public sealed class OcrTableCell
    {
        public List<OcrParagraph> Paragraphs = new List<OcrParagraph>();
        public int ColSpan = 1;
        public int RowSpan = 1;
        public bool Covered; // накрыта объединением слева/сверху — не пишется отдельно
    }

    /// <summary>Строка таблицы: ровно ColumnCount ячеек (включая накрытые-заглушки).</summary>
    public sealed class OcrTableRow
    {
        public List<OcrTableCell> Cells = new List<OcrTableCell>();
    }

    /// <summary>
    /// Восстановленная таблица: полная сетка Rows×ColumnCount (накрытые ячейки — заглушки с
    /// <see cref="OcrTableCell.Covered"/>), ширины колонок из геометрии линовки и верх для
    /// порядка чтения вместе с абзацами и изображениями страницы.
    /// </summary>
    public sealed class OcrTable
    {
        public List<OcrTableRow> Rows = new List<OcrTableRow>();
        public List<double> ColumnWidthsPt = new List<double>();
        public double TopPt;  // верх таблицы (Y, ось вверх) — для сортировки блоков страницы
        public double LeftPt; // левый край таблицы — вторичный порядок (таблицы бок о бок: левее раньше)

        public int ColumnCount { get { return ColumnWidthsPt.Count; } }
    }
}
