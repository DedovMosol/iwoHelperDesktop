using System;
using System.Collections.Generic;
using System.Text;

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
        // Форма «метка — значение»: в сегменте 1–3 слова. Строка газетной колонки — целая фраза.
        // Медиана ≥ этого числа → две колонки текста, не сетка (RC-8).
        private const int MaxFormSegmentWords = 3;
        private const double RunningTextFill = 0.75;

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
                // Абзац над формой входит в то же вертикальное окно. Сетку начинаем с первой
                // строки-пары; хвост не режем — итоговая строка во всю ширину (один сегмент, ColSpan).
                int a = start, b = end;
                while (a <= b && segments[a].Count < 2) a++;
                OcrTable table = a <= b ? TryBuildGrid(lines, segments, a, b, em, consumed) : null;
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
            int pairRows = 0, tripleRows = 0;
            for (int i = start; i <= end; i++)
            {
                if (segments[i].Count >= 2)
                    pairRows++;
                if (segments[i].Count >= 3)
                    tripleRows++;
            }
            if (pairRows < (int)Math.Ceiling(rows * MinPairShare))
                return null;
            // Три сегмента в строке — газетные колонки, не форма из двух полей (снимок «три колонки»).
            // Медиана слов тут бесполезна: хвосты колонок однословные и опускают её до 1.
            if (tripleRows >= MinRows)
                return null;
            int medianWords = MedianSegmentWords(segments, start, end);
            if (medianWords > MaxFormSegmentWords)
                return null;
            if (LeftColumnLooksLikeList(segments, start, end))
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
            // Газета: каждая колонка заполнена почти целиком. Форма «метка — значение» —
            // короткие метки в широкой колонке (снимок «две колонки» глотался как таблица).
            if (columns.Count >= 2 && RunningTextColumns(segments, start, end, columns, table.ColumnWidthsPt, tol))
                return null;

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

        private static int MedianSegmentWords(List<List<Segment>> segments, int start, int end)
        {
            var counts = new List<double>();
            for (int i = start; i <= end; i++)
                foreach (Segment s in segments[i])
                    counts.Add(s.Words.Count);
            if (counts.Count == 0)
                return 0;
            return (int)Math.Round(MathUtil.Median(counts));
        }

        private static bool RunningTextColumns(List<List<Segment>> segments, int start, int end,
            List<double> columns, List<double> widths, double tol)
        {
            var fills = new List<double>[columns.Count];
            for (int c = 0; c < columns.Count; c++)
                fills[c] = new List<double>();
            for (int i = start; i <= end; i++)
                foreach (Segment s in segments[i])
                {
                    int c = ColumnIndex(columns, s.Left, tol);
                    if (c < 0)
                        continue;
                    double w = widths[c];
                    if (w <= 0)
                        continue;
                    fills[c].Add((s.Right - s.Left) / w);
                }
            for (int c = 0; c < columns.Count; c++)
            {
                if (fills[c].Count < MinColumnSupport)
                    return false;
                if (MathUtil.Median(fills[c]) < RunningTextFill)
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Левая колонка — маркеры списка («•», «1.»), правая — текст пункта: это не форма.
        /// </summary>
        private static bool LeftColumnLooksLikeList(List<List<Segment>> segments, int start, int end)
        {
            int pair = 0, listed = 0;
            for (int i = start; i <= end; i++)
            {
                if (segments[i].Count < 2)
                    continue;
                pair++;
                Segment left = segments[i][0];
                for (int k = 1; k < segments[i].Count; k++)
                    if (segments[i][k].Left < left.Left)
                        left = segments[i][k];
                var text = new StringBuilder();
                foreach (PdfWord w in left.Words)
                {
                    if (text.Length > 0) text.Append(' ');
                    text.Append(w.Text ?? "");
                }
                string s = text.ToString();
                if (ListMarker.LooksLikeMarker(s))
                    listed++;
            }
            return pair >= MinRows && listed * 2 >= pair;
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
