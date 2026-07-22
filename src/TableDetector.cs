using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>Итог поиска таблиц на странице: сами таблицы + слова, не попавшие ни в одну.</summary>
    internal sealed class TableDetectResult
    {
        public List<OcrTable> Tables = new List<OcrTable>();
        public List<PdfWord> RemainingWords = new List<PdfWord>();
    }

    /// <summary>
    /// Восстановление таблиц born-digital PDF по линовке (границам ячеек) и словам. Чистая
    /// логика без типов PdfPig — покрыта юнит-тестами. Подход:
    ///  1. линии группируются в связные компоненты (пересекающиеся/смежные рёбра одной таблицы);
    ///     это заодно отсекает одиночную графику — подчёркивания, разделители (свои компоненты);
    ///  2. компонент с сеткой ≥2×2 и начерченным периметром становится таблицей;
    ///  3. объединённые ячейки узнаются по ОТСУТСТВУЮЩИМ внутренним границам (colspan/rowspan);
    ///  4. слова раскладываются по ячейкам геометрически, текст ячейки — через <see cref="OcrLayout"/>.
    /// При любом сомнении компонент отбрасывается, а его слова остаются обычным текстом — вывод
    /// не бывает хуже прежнего. Таблицы без границ и многоколоночная вёрстка сюда не входят.
    /// </summary>
    internal static class TableDetector
    {
        private const double TouchTol = 3.0;       // рёбра «связаны», если пересекаются/смыкаются в пределах (pt)
        private const double ClusterTol = 3.0;     // близкие координаты границ сливаются в одну
        private const double CoverFraction = 0.7;  // ребро «начерчено», если линовка покрывает эту долю его длины
        private const double MinGridStep = 4.0;    // меньший столбец/строка — артефакт, не ячейка
        private const double EdgeTol = 2.0;        // допуск попадания центра слова внутрь таблицы

        /// <summary>Линии + слова страницы → таблицы и оставшиеся вне них слова. Чистая — под тест.</summary>
        public static TableDetectResult Detect(IList<PdfLine> lines, IList<PdfWord> words, double pageWidth, double pageHeight)
        {
            var result = new TableDetectResult();
            var allWords = words != null ? new List<PdfWord>(words) : new List<PdfWord>();
            if (lines == null || lines.Count == 0)
            {
                result.RemainingWords = allWords;
                return result;
            }

            var consumed = new HashSet<PdfWord>();
            foreach (List<PdfLine> comp in ConnectedComponents(lines))
            {
                OcrTable table = TryBuildTable(comp, allWords, consumed);
                if (table != null)
                    result.Tables.Add(table);
            }
            result.Tables.Sort(delegate(OcrTable a, OcrTable b) { return b.TopPt.CompareTo(a.TopPt); }); // сверху вниз
            foreach (PdfWord w in allWords)
                if (!consumed.Contains(w))
                    result.RemainingWords.Add(w);
            return result;
        }

        // ---------- связные компоненты линий ----------

        private static List<List<PdfLine>> ConnectedComponents(IList<PdfLine> lines)
        {
            int n = lines.Count;
            var parent = new int[n];
            for (int i = 0; i < n; i++) parent[i] = i;
            for (int i = 0; i < n; i++)
                for (int j = i + 1; j < n; j++)
                    if (Touch(lines[i], lines[j]))
                        Union(parent, i, j);

            var groups = new Dictionary<int, List<PdfLine>>();
            for (int i = 0; i < n; i++)
            {
                int root = Find(parent, i);
                List<PdfLine> g;
                if (!groups.TryGetValue(root, out g)) { g = new List<PdfLine>(); groups[root] = g; }
                g.Add(lines[i]);
            }
            return new List<List<PdfLine>>(groups.Values);
        }

        private static int Find(int[] parent, int i)
        {
            while (parent[i] != i) { parent[i] = parent[parent[i]]; i = parent[i]; }
            return i;
        }

        private static void Union(int[] parent, int a, int b)
        {
            int ra = Find(parent, a), rb = Find(parent, b);
            if (ra != rb) parent[ra] = rb;
        }

        /// <summary>Связаны ли рёбра: H и V пересекаются, либо две одноосные линии коллинеарны и смыкаются.</summary>
        private static bool Touch(PdfLine a, PdfLine b)
        {
            if (a.Orientation != b.Orientation)
            {
                PdfLine h = a.Orientation == LineOrientation.Horizontal ? a : b;
                PdfLine v = a.Orientation == LineOrientation.Horizontal ? b : a;
                return v.Position >= h.MinX - TouchTol && v.Position <= h.MaxX + TouchTol
                    && h.Position >= v.MinY - TouchTol && h.Position <= v.MaxY + TouchTol;
            }
            if (Math.Abs(a.Position - b.Position) > TouchTol)
                return false;
            return a.Orientation == LineOrientation.Horizontal
                ? Overlap(a.MinX, a.MaxX, b.MinX, b.MaxX)
                : Overlap(a.MinY, a.MaxY, b.MinY, b.MaxY);
        }

        private static bool Overlap(double lo1, double hi1, double lo2, double hi2)
        {
            return lo1 <= hi2 + TouchTol && lo2 <= hi1 + TouchTol;
        }

        // ---------- построение таблицы из компонента ----------

        private static OcrTable TryBuildTable(List<PdfLine> comp, List<PdfWord> allWords, HashSet<PdfWord> consumed)
        {
            var vlines = new List<PdfLine>();
            var hlines = new List<PdfLine>();
            foreach (PdfLine l in comp)
                (l.Orientation == LineOrientation.Vertical ? vlines : hlines).Add(l);
            if (vlines.Count < 2 || hlines.Count < 2)
                return null;

            List<double> xs = ClusterPositions(vlines);   // границы колонок (по возрастанию X)
            List<double> ys = ClusterPositions(hlines);    // границы строк (по возрастанию Y)
            xs = Filtered(xs);
            ys = Filtered(ys);
            if (xs.Count < 3 || ys.Count < 3) // ≥2 колонки И ≥2 строки — иначе не таблица
                return null;
            ys.Reverse(); // сверху вниз: ys[0] — верхняя граница (наибольший Y)

            int cols = xs.Count - 1, rows = ys.Count - 1;
            double left = xs[0], right = xs[cols], top = ys[0], bottom = ys[rows];

            // Периметр должен быть начерчен — иначе это не обрамлённая таблица (наш случай).
            if (!HasHBorder(hlines, top, left, right) || !HasHBorder(hlines, bottom, left, right)
                || !HasVBorder(vlines, left, bottom, top) || !HasVBorder(vlines, right, bottom, top))
                return null;

            OcrTableCell[,] origin = BuildCells(xs, ys, vlines, hlines, cols, rows);
            AssignWords(origin, xs, ys, cols, rows, allWords, consumed);

            var table = new OcrTable { TopPt = top, LeftPt = left };
            for (int c = 0; c < cols; c++)
                table.ColumnWidthsPt.Add(xs[c + 1] - xs[c]);
            for (int r = 0; r < rows; r++)
            {
                var row = new OcrTableRow();
                for (int c = 0; c < cols; c++)
                    row.Cells.Add(origin[r, c] ?? new OcrTableCell { Covered = true });
                table.Rows.Add(row);
            }
            return table;
        }

        /// <summary>Сетка ячеек: хозяева с ColSpan/RowSpan по отсутствующим внутренним границам; накрытые — null.</summary>
        private static OcrTableCell[,] BuildCells(List<double> xs, List<double> ys, List<PdfLine> vlines, List<PdfLine> hlines, int cols, int rows)
        {
            var cells = new OcrTableCell[rows, cols];
            var covered = new bool[rows, cols];
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    if (covered[r, c])
                        continue;
                    double cellTop = ys[r], cellBottom = ys[r + 1];
                    int cs = 1;
                    while (c + cs < cols && !HasVBorder(vlines, xs[c + cs], cellBottom, cellTop))
                        cs++;
                    int rs = 1;
                    while (r + rs < rows && !HasHBorder(hlines, ys[r + rs], xs[c], xs[c + cs]))
                        rs++;
                    cells[r, c] = new OcrTableCell { ColSpan = cs, RowSpan = rs };
                    for (int dr = 0; dr < rs; dr++)
                        for (int dc = 0; dc < cs; dc++)
                            if (dr != 0 || dc != 0)
                                covered[r + dr, c + dc] = true;
                }
            }
            return cells;
        }

        /// <summary>Разложить слова по ячейкам-хозяевам; текст ячейки — через OcrLayout (DRY).</summary>
        private static void AssignWords(OcrTableCell[,] origin, List<double> xs, List<double> ys, int cols, int rows, List<PdfWord> allWords, HashSet<PdfWord> consumed)
        {
            // Владелец каждой позиции сетки: индекс ячейки-хозяина, накрывшей её.
            var ownerR = new int[rows, cols];
            var ownerC = new int[rows, cols];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                {
                    OcrTableCell cell = origin[r, c];
                    if (cell == null) continue;
                    for (int dr = 0; dr < cell.RowSpan; dr++)
                        for (int dc = 0; dc < cell.ColSpan; dc++)
                        {
                            ownerR[r + dr, c + dc] = r;
                            ownerC[r + dr, c + dc] = c;
                        }
                }

            var byOwner = new Dictionary<int, List<PdfWord>>();
            foreach (PdfWord w in allWords)
            {
                double cx = (w.Left + w.Right) / 2, cy = (w.Top + w.Bottom) / 2;
                if (cx < xs[0] - EdgeTol || cx > xs[cols] + EdgeTol || cy < ys[rows] - EdgeTol || cy > ys[0] + EdgeTol)
                    continue;
                int c = ColIndex(xs, cx), r = RowIndex(ys, cy);
                if (c < 0 || r < 0)
                    continue;
                int or_ = ownerR[r, c], oc = ownerC[r, c];
                int key = or_ * cols + oc;
                List<PdfWord> list;
                if (!byOwner.TryGetValue(key, out list)) { list = new List<PdfWord>(); byOwner[key] = list; }
                list.Add(w);
                consumed.Add(w);
            }

            foreach (KeyValuePair<int, List<PdfWord>> kv in byOwner)
            {
                OcrTableCell cell = origin[kv.Key / cols, kv.Key % cols];
                if (cell != null)
                    cell.Paragraphs = OcrLayout.Analyze(kv.Value).Paragraphs;
            }
        }

        // ---------- геометрия ----------

        /// <summary>Индекс колонки по X центра слова (между xs[i] и xs[i+1]); -1 — вне сетки.</summary>
        private static int ColIndex(List<double> xs, double x)
        {
            for (int i = 0; i < xs.Count - 1; i++)
                if (x >= xs[i] - EdgeTol && x <= xs[i + 1] + EdgeTol)
                    return i;
            return -1;
        }

        /// <summary>Индекс строки по Y центра слова (ys по убыванию: ys[i] — верх строки i).</summary>
        private static int RowIndex(List<double> ys, double y)
        {
            for (int i = 0; i < ys.Count - 1; i++)
                if (y <= ys[i] + EdgeTol && y >= ys[i + 1] - EdgeTol)
                    return i;
            return -1;
        }

        /// <summary>Границы из позиций линий: сортировка + слияние близких (в пределах ClusterTol).</summary>
        private static List<double> ClusterPositions(List<PdfLine> group)
        {
            var ps = new List<double>(group.Count);
            foreach (PdfLine l in group) ps.Add(l.Position);
            ps.Sort();
            var clusters = new List<double>();
            var bucket = new List<double>();
            foreach (double p in ps)
            {
                if (bucket.Count > 0 && p - bucket[bucket.Count - 1] > ClusterTol)
                {
                    clusters.Add(Average(bucket));
                    bucket.Clear();
                }
                bucket.Add(p);
            }
            if (bucket.Count > 0)
                clusters.Add(Average(bucket));
            return clusters;
        }

        /// <summary>Убрать границы, стоящие вплотную (шаг меньше MinGridStep — не отдельная ячейка).</summary>
        private static List<double> Filtered(List<double> sortedAsc)
        {
            var result = new List<double>();
            foreach (double v in sortedAsc)
                if (result.Count == 0 || v - result[result.Count - 1] >= MinGridStep)
                    result.Add(v);
            return result;
        }

        /// <summary>Есть ли горизонтальная линовка у Y=y, покрывающая [xL,xR] не меньше чем на CoverFraction.</summary>
        private static bool HasHBorder(List<PdfLine> hlines, double y, double xL, double xR)
        {
            var intervals = new List<double[]>();
            foreach (PdfLine l in hlines)
                if (Math.Abs(l.Position - y) <= ClusterTol)
                {
                    double a = Math.Max(l.MinX, xL), b = Math.Min(l.MaxX, xR);
                    if (b > a) intervals.Add(new[] { a, b });
                }
            return UnionLength(intervals) >= (xR - xL) * CoverFraction;
        }

        /// <summary>Есть ли вертикальная линовка у X=x, покрывающая [yB,yT] не меньше чем на CoverFraction.</summary>
        private static bool HasVBorder(List<PdfLine> vlines, double x, double yB, double yT)
        {
            var intervals = new List<double[]>();
            foreach (PdfLine l in vlines)
                if (Math.Abs(l.Position - x) <= ClusterTol)
                {
                    double a = Math.Max(l.MinY, yB), b = Math.Min(l.MaxY, yT);
                    if (b > a) intervals.Add(new[] { a, b });
                }
            return UnionLength(intervals) >= (yT - yB) * CoverFraction;
        }

        /// <summary>Суммарная длина объединения интервалов (перекрытия не считаются дважды).</summary>
        private static double UnionLength(List<double[]> intervals)
        {
            if (intervals.Count == 0)
                return 0;
            intervals.Sort(delegate(double[] a, double[] b) { return a[0].CompareTo(b[0]); });
            double total = 0, curLo = intervals[0][0], curHi = intervals[0][1];
            for (int i = 1; i < intervals.Count; i++)
            {
                if (intervals[i][0] <= curHi)
                {
                    if (intervals[i][1] > curHi) curHi = intervals[i][1];
                }
                else
                {
                    total += curHi - curLo;
                    curLo = intervals[i][0];
                    curHi = intervals[i][1];
                }
            }
            return total + (curHi - curLo);
        }

        private static double Average(List<double> values)
        {
            double sum = 0;
            for (int i = 0; i < values.Count; i++) sum += values[i];
            return sum / values.Count;
        }
    }
}
