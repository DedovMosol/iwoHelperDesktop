using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>Итог поиска сеток: безграничные таблицы + слова вне них.</summary>
    internal sealed class GridDetectResult
    {
        public List<OcrTable> Tables = new List<OcrTable>();
        public List<PdfWord> RemainingWords = new List<PdfWord>();
    }

    /// <summary>
    /// Восстановление «сеток» — табличной вёрстки БЕЗ линовки (чеки и формы «метка … значение»,
    /// зоны билетов). Признак сетки: несколько строк подряд, разделённых ШИРОКИМИ
    /// внутристрочными зазорами на сегменты, чьи левые края выровнены по общим колонкам.
    /// Такой блок переносится настоящей таблицей Word без границ: пары «метка — значение»
    /// остаются на своих строках и выровнены; разрез страницы на колонки иначе растаскивал их
    /// («сначала все метки, потом все значения»). Пороги нарочно строгие: у justified-текста
    /// нет ни широких зазоров, ни выровненных сегментов, у «(подпись) … (дата)» не хватает
    /// строк — при сомнении блок остаётся обычным текстом. Чистая логика — под тест.
    /// </summary>
    internal static class GridDetector
    {
        private const double SegmentGapEmFactor = 1.4; // внутристрочный зазор шире этой доли кегля — граница сегментов
        private const double ColumnTolEmFactor = 0.7;  // левые края в пределах этой доли кегля — одна колонка
        private const int MinRows = 3;                 // окно короче — не сетка («подпись … дата» не считается)
        private const double MinPairShare = 0.5;       // доля строк окна с двумя и более сегментами
        private const double RowGapEmFactor = 4.5;  // вертикальный зазор больше этой доли кегля рвёт окно (группы полей чека разделены до ~4 кеглей)
        private const int MinColumns = 2;
        private const int MinColumnSupport = 2;        // колонку образуют минимум два сегмента
        private const double MaxStraySegmentShare = 0.2; // больше выбросов мимо колонок — не сетка
        private const double RowSpacingMinFactor = 0.4;   // зазор сверх типичного шага больше этой доли -> интервал группы
        private const double RowSpacingMaxFactor = 3.0;   // кап интервала (доля типичного шага)

        /// <summary>Слова страницы → сетки-таблицы (Borderless) и остаток. Чистая — под тест.</summary>
        public static GridDetectResult Detect(IList<PdfWord> words)
        {
            var result = new GridDetectResult();
            if (words == null || words.Count == 0)
                return result;

            List<OcrLayout.Line> lines = OcrLayout.ToLines(words);
            var heights = new List<double>(lines.Count);
            foreach (OcrLayout.Line ln in lines)
                heights.Add(ln.Height);
            double em = MathUtil.Median(heights);
            if (em <= 0) em = 1;

            // Сегменты каждой строки — по широким внутристрочным зазорам.
            var segments = new List<List<Segment>>(lines.Count);
            foreach (OcrLayout.Line ln in lines)
                segments.Add(SplitSegments(ln, SegmentGapEmFactor * em));

            var consumed = new HashSet<PdfWord>();
            int start = 0;
            while (start < lines.Count)
            {
                // Окно — максимальная цепочка строк без больших вертикальных разрывов.
                int end = start;
                while (end + 1 < lines.Count
                    && lines[end].MidY - lines[end + 1].MidY <= RowGapEmFactor * em)
                    end++;
                OcrTable table = TryBuildGrid(lines, segments, start, end, em, consumed);
                if (table != null)
                    result.Tables.Add(table);
                start = end + 1;
            }

            foreach (PdfWord w in words)
                if (!consumed.Contains(w))
                    result.RemainingWords.Add(w);
            return result;
        }

        /// <summary>Отрезок строки без широких зазоров: рамка и слова.</summary>
        private sealed class Segment
        {
            public readonly List<PdfWord> Words = new List<PdfWord>();
            public double Left;
            public double Right;
        }

        private static List<Segment> SplitSegments(OcrLayout.Line line, double gapTol)
        {
            var result = new List<Segment>();
            Segment cur = null;
            foreach (PdfWord w in line.Words)
            {
                if (cur == null || w.Left - cur.Right > gapTol)
                {
                    cur = new Segment { Left = w.Left, Right = w.Right };
                    result.Add(cur);
                }
                if (w.Right > cur.Right)
                    cur.Right = w.Right;
                cur.Words.Add(w);
            }
            return result;
        }

        /// <summary>
        /// Собрать сетку из окна строк [start..end]: достаточно строк, достаточно «пар»,
        /// сегменты ложатся на общие колонки. Не сетка — null (слова остаются текстом).
        /// </summary>
        private static OcrTable TryBuildGrid(List<OcrLayout.Line> lines, List<List<Segment>> segments,
            int start, int end, double em, HashSet<PdfWord> consumed)
        {
            int rows = end - start + 1;
            if (rows < MinRows)
                return null;
            int pairRows = 0;
            for (int i = start; i <= end; i++)
                if (segments[i].Count >= 2)
                    pairRows++;
            if (pairRows < (int)Math.Ceiling(rows * MinPairShare))
                return null;

            // Колонки — кластеры левых краёв сегментов с достаточной поддержкой.
            var lefts = new List<double>();
            int totalSegments = 0;
            for (int i = start; i <= end; i++)
                foreach (Segment s in segments[i])
                {
                    lefts.Add(s.Left);
                    totalSegments++;
                }
            lefts.Sort();
            double tol = ColumnTolEmFactor * em;
            var columns = new List<double>(); // левая граница колонки (среднее кластера)
            var support = new List<int>();
            int ci = 0;
            while (ci < lefts.Count)
            {
                int cj = ci;
                double sum = 0;
                while (cj < lefts.Count && lefts[cj] - lefts[ci] <= tol)
                {
                    sum += lefts[cj];
                    cj++;
                }
                columns.Add(sum / (cj - ci));
                support.Add(cj - ci);
                ci = cj;
            }
            for (int i = columns.Count - 1; i >= 0; i--)
                if (support[i] < MinColumnSupport)
                {
                    columns.RemoveAt(i);
                    support.RemoveAt(i);
                }
            if (columns.Count < MinColumns)
                return null;

            // Каждый сегмент обязан лечь в свою колонку; редкие выбросы допустимы, массовые — не сетка.
            int stray = 0;
            for (int i = start; i <= end; i++)
                foreach (Segment s in segments[i])
                    if (ColumnIndex(columns, s.Left, tol) < 0)
                        stray++;
            if (stray > totalSegments * MaxStraySegmentShare)
                return null;

            // Рамки таблицы и ширины колонок (последняя — до правого края содержимого).
            double right = double.MinValue, top = double.MinValue, bottom = double.MaxValue;
            for (int i = start; i <= end; i++)
                foreach (Segment s in segments[i])
                {
                    if (s.Right > right) right = s.Right;
                    foreach (PdfWord w in s.Words)
                    {
                        if (w.Top > top) top = w.Top;
                        if (w.Bottom < bottom) bottom = w.Bottom;
                    }
                }
            var table = new OcrTable
            {
                Borderless = true,
                TopPt = top,
                LeftPt = columns[0],
                RightPt = right,
                BottomPt = bottom
            };
            for (int c = 0; c < columns.Count; c++)
            {
                double colRight = c + 1 < columns.Count ? columns[c + 1] : right;
                table.ColumnWidthsPt.Add(Math.Max(1, colRight - columns[c]));
            }

            // Типичный шаг строк окна — чтобы вернуть группировку: где исходный зазор заметно
            // больше типичного (пустая строка между группами полей чека), добавим интервал после
            // строки. Иначе безлиновочная сетка выходит равномерно-плотной, теряя разбивку на группы.
            var pitches = new List<double>();
            for (int i = start; i < end; i++)
                pitches.Add(lines[i].MidY - lines[i + 1].MidY);
            double typicalPitch = MathUtil.Median(pitches);

            // Строки: сегменты раскладываются по колонкам; одиночный сегмент, накрывающий
            // соседние колонки (итоговая строка чека во всю ширину), получает colspan.
            for (int i = start; i <= end; i++)
            {
                var row = new OcrTableRow();
                var cellWords = new List<PdfWord>[columns.Count];
                foreach (Segment s in segments[i])
                {
                    int c = ColumnIndex(columns, s.Left, tol);
                    if (c < 0)
                        c = NearestColumn(columns, s.Left); // выброс — в ближайшую (их мало, см. порог)
                    if (cellWords[c] == null)
                        cellWords[c] = new List<PdfWord>();
                    cellWords[c].AddRange(s.Words);
                }
                int spanFrom = -1, spanLen = 1;
                if (segments[i].Count == 1)
                {
                    // Насколько далеко одиночный сегмент заходит вправо по колонкам.
                    Segment s = segments[i][0];
                    spanFrom = ColumnIndex(columns, s.Left, tol);
                    if (spanFrom < 0) spanFrom = NearestColumn(columns, s.Left);
                    spanLen = 1;
                    while (spanFrom + spanLen < columns.Count && s.Right > columns[spanFrom + spanLen] + tol)
                        spanLen++;
                }
                for (int c = 0; c < columns.Count; c++)
                {
                    if (spanFrom >= 0 && c > spanFrom && c < spanFrom + spanLen)
                    {
                        row.Cells.Add(new OcrTableCell { Covered = true });
                        continue;
                    }
                    var cell = new OcrTableCell();
                    if (spanFrom >= 0 && c == spanFrom)
                        cell.ColSpan = spanLen;
                    if (cellWords[c] != null)
                    {
                        cell.Paragraphs = OcrLayout.Analyze(cellWords[c], false).Paragraphs;
                        foreach (PdfWord w in cellWords[c])
                            consumed.Add(w);
                    }
                    row.Cells.Add(cell);
                }
                // Лишний интервал перед следующей строкой (сверх типичного шага) = пустой промежуток
                // группы в исходнике. Порог и кап — от шума и патологий.
                if (i < end && typicalPitch > 0)
                {
                    double extra = (lines[i].MidY - lines[i + 1].MidY) - typicalPitch;
                    if (extra > RowSpacingMinFactor * typicalPitch)
                        row.SpaceAfterPt = Math.Min(extra, RowSpacingMaxFactor * typicalPitch);
                }
                table.Rows.Add(row);
            }
            return table;
        }

        /// <summary>Индекс колонки, чья левая граница совпадает с left в пределах tol; -1 — мимо всех.</summary>
        private static int ColumnIndex(List<double> columns, double left, double tol)
        {
            for (int c = 0; c < columns.Count; c++)
                if (Math.Abs(columns[c] - left) <= tol)
                    return c;
            return -1;
        }

        private static int NearestColumn(List<double> columns, double left)
        {
            int best = 0;
            double bestDist = double.MaxValue;
            for (int c = 0; c < columns.Count; c++)
            {
                double d = Math.Abs(columns[c] - left);
                if (d < bestDist)
                {
                    bestDist = d;
                    best = c;
                }
            }
            return best;
        }
    }
}
