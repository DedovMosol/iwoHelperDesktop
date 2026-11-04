using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>Прямоугольник содержимого для XY-разреза (координаты PDF, ось Y вверх) с меткой вызывающего.</summary>
    internal struct CutBox
    {
        public double Left;
        public double Right;
        public double Bottom;
        public double Top;
        public int Tag; // индекс элемента у вызывающего; листья возвращают метки в порядке чтения
    }

    /// <summary>
    /// Лист XY-разреза: метки боксов одного связного блока и рамка КОЛОНКИ, в которой блок
    /// лежит (ближайший вертикальный разрез-предок; без колонок — рамка всего ввода).
    /// Рамка колонки — опора геометрии текста: центрирование и красная строка считаются
    /// от полей колонки, а не от полей узкого блока (иначе одиночная центрированная строка,
    /// отрезанная в свой блок, «теряет» центрирование — не от чего отступать).
    /// </summary>
    internal sealed class CutLeaf
    {
        public readonly List<int> Tags = new List<int>();
        public double ColumnLeft;
        public double ColumnRight;
        // Правый предел ДОСТУПНОГО места: левый край контента следующей колонки (или правая
        // граница родительской области). ColumnRight — рамка по СОДЕРЖИМОМУ колонки, и для
        // вопроса «влезло бы слово в строку» (умышленный перевод строки) она не годится:
        // у блока подписи из двух коротких строк собственная рамка всегда «заполнена».
        public double AvailRight;
    }

    /// <summary>
    /// Узел дерева XY-разреза. Лист (<see cref="IsLeaf"/>) несёт метки боксов и рамку колонки;
    /// внутренний узел несёт детей — либо ЭТАЖИ (стек сверху вниз, <see cref="SideBySide"/>=false),
    /// либо КОЛОНКИ (бок о бок слева направо, SideBySide=true). Дерево нужно рендеру блоков, чтобы
    /// выводить side-by-side колонки таблицей (шапка письма), а не стеком; порядок чтения (плоские
    /// листья) из него получается обходом в порядке детей — совпадает с <see cref="XyCut.Order"/>.
    /// </summary>
    internal sealed class CutNode
    {
        public List<int> Tags;          // лист: метки боксов в порядке чтения; null — внутренний узел
        public double ColumnLeft;
        public double ColumnRight;
        public double AvailRight;       // лист: правый предел доступного места (см. CutLeaf.AvailRight)
        public List<CutNode> Children;  // внутренний узел: дети; null — лист
        public bool SideBySide;         // внутренний: true — колонки (бок о бок), false — этажи (стек)

        public bool IsLeaf { get { return Tags != null; } }
    }

    /// <summary>
    /// Рекурсивный XY-разрез — порядок чтения многоколоночной вёрстки. Область режется по
    /// пустым полосам: сначала ГОРИЗОНТАЛЬНЫЕ (пустые полосы во всю ширину делят её на этажи
    /// сверху вниз), затем в каждом этаже — ВЕРТИКАЛЬНЫЕ (пустые просветы во всю высоту этажа
    /// делят его на колонки слева направо), и так рекурсивно. Листья в порядке обхода — порядок
    /// чтения: левая колонка ЦЕЛИКОМ раньше правой. Простая сортировка по Top этого не даёт:
    /// строки соседних колонок перемежаются (двухколоночная шапка письма читалась «через строку»).
    /// Горизонтальный разрез сам по себе порядка не меняет (совпадает с «сверху вниз») — он
    /// нужен, чтобы открыть вертикальные просветы внутри этажа, не пересечённые содержимым
    /// соседних этажей. Чистая логика — под тест.
    /// </summary>
    internal static class XyCut
    {
        // Страховка от патологической рекурсии: реальная вложенность вёрстки 2–3 уровня,
        // каждый уровень разреза строго уменьшает области, но ограничим явно.
        private const int MaxDepth = 8;

        /// <summary>Параметры одного разреза — чтобы не тянуть шесть аргументов по рекурсии.</summary>
        private sealed class CutContext
        {
            public IList<CutBox> Boxes;
            public double HGapMin;          // минимальная высота пустой полосы для этажа
            public double VGapMin;          // минимальная ширина пустого просвета для колонки
            public double MinColumnExtent;  // колонка ниже этого по высоте — не колонка (подпись/дата в одной строке)
            public int MinColumnItems;      // и колонка из меньшего числа элементов — не колонка
            public double MinColumnWidth;   // 0 — не проверять; иначе узкая полоса маркеров списка — не колонка
        }

        /// <summary>
        /// Разрезать боксы на блоки в порядке чтения (плоские листья). Флэттен дерева
        /// <see cref="OrderTree"/> в порядке детей — совпадает с прежним обходом. Параметры — см. OrderTree.
        /// </summary>
        public static List<CutLeaf> Order(IList<CutBox> boxes, double hGapMin, double vGapMin,
            double minColumnExtent, int minColumnItems)
        {
            return Order(boxes, hGapMin, vGapMin, minColumnExtent, minColumnItems, 0);
        }

        public static List<CutLeaf> Order(IList<CutBox> boxes, double hGapMin, double vGapMin,
            double minColumnExtent, int minColumnItems, double minColumnWidth)
        {
            var leaves = new List<CutLeaf>();
            CutNode root = OrderTree(boxes, hGapMin, vGapMin, minColumnExtent, minColumnItems, minColumnWidth);
            if (root != null)
                Flatten(root, leaves);
            return leaves;
        }

        /// <summary>
        /// Разрезать боксы в ДЕРЕВО (<see cref="CutNode"/>): этажи (стек) и колонки (бок о бок).
        /// hGapMin/vGapMin — минимальные пустые полосы (pt) для горизонтального/вертикального
        /// разреза; minColumnExtent и minColumnItems — защита от ложных колонок (широкий пробел
        /// одной строки — «подпись … дата» — не делит её на столбики). Пустой ввод — null.
        /// </summary>
        public static CutNode OrderTree(IList<CutBox> boxes, double hGapMin, double vGapMin,
            double minColumnExtent, int minColumnItems)
        {
            return OrderTree(boxes, hGapMin, vGapMin, minColumnExtent, minColumnItems, 0);
        }

        public static CutNode OrderTree(IList<CutBox> boxes, double hGapMin, double vGapMin,
            double minColumnExtent, int minColumnItems, double minColumnWidth)
        {
            if (boxes == null || boxes.Count == 0)
                return null;
            var all = new List<int>(boxes.Count);
            for (int i = 0; i < boxes.Count; i++)
                all.Add(i);
            var ctx = new CutContext
            {
                Boxes = boxes,
                HGapMin = hGapMin,
                VGapMin = vGapMin,
                MinColumnExtent = minColumnExtent,
                MinColumnItems = minColumnItems,
                MinColumnWidth = minColumnWidth
            };
            double left, right;
            HorizontalBounds(boxes, all, out left, out right);
            return CutFloors(ctx, all, MaxDepth, left, right, right, true);
        }

        /// <summary>Листья дерева в порядке чтения (обход в порядке детей).</summary>
        private static void Flatten(CutNode node, List<CutLeaf> acc)
        {
            if (node.IsLeaf)
            {
                var leaf = new CutLeaf { ColumnLeft = node.ColumnLeft, ColumnRight = node.ColumnRight, AvailRight = node.AvailRight };
                leaf.Tags.AddRange(node.Tags);
                acc.Add(leaf);
                return;
            }
            foreach (CutNode child in node.Children)
                Flatten(child, acc);
        }

        /// <summary>
        /// Этажи сверху вниз. На корне сначала пробуем КОЛОНКИ: газета с отбивкой абзацев
        /// иначе режется на этажи «левая+правая строка» и читается зигзагом (RC-14, снимок #3).
        /// Нет вертикального канала — этажи как раньше; заголовок над каналом снимает peel.
        /// </summary>
        private static CutNode CutFloors(CutContext ctx, List<int> items, int depth, double colLeft, double colRight, double availRight, bool splitX)
        {
            if (splitX && depth > 0)
            {
                List<List<int>> cols = SplitColumns(ctx, items);
                // Колонки разной высоты: шапка/подвал на всю ширину слева + узкая колонка
                // справа — не газета. Иначе «верх / середина / низ» читается как левая
                // колонка целиком (PlainText reading order).
                if (cols != null && ColumnsComparableHeight(ctx, cols))
                    return AssembleColumns(ctx, cols, depth, availRight);
                List<List<int>> peeled = PeelBlockingFloor(ctx, items);
                if (peeled != null)
                {
                    var peeledNode = new CutNode { SideBySide = false, Children = new List<CutNode>(peeled.Count) };
                    foreach (List<int> floor in peeled)
                        peeledNode.Children.Add(CutFloors(ctx, floor, depth - 1, colLeft, colRight, availRight, true));
                    return peeledNode;
                }
            }
            List<List<int>> floors = depth > 0 ? SplitAtGaps(ctx.Boxes, items, true, ctx.HGapMin) : null;
            if (floors == null)
                return splitX ? CutColumns(ctx, items, depth, colLeft, colRight, availRight) : MakeLeaf(ctx, items, colLeft, colRight, availRight);
            var node = new CutNode { SideBySide = false, Children = new List<CutNode>(floors.Count) };
            foreach (List<int> floor in floors)
                node.Children.Add(splitX
                    ? CutColumns(ctx, floor, depth - 1, colLeft, colRight, availRight)
                    : CutFloors(ctx, floor, depth - 1, colLeft, colRight, availRight, false));
            return node;
        }

        /// <summary>Колонки этажа (слева направо): узел side-by-side. Нет колонок — лист.</summary>
        private static CutNode CutColumns(CutContext ctx, List<int> items, int depth, double colLeft, double colRight, double availRight)
        {
            List<List<int>> columns = depth > 0 ? SplitColumns(ctx, items) : null;
            if (columns == null)
                return MakeLeaf(ctx, items, colLeft, colRight, availRight);
            return AssembleColumns(ctx, columns, depth, availRight);
        }

        private static CutNode AssembleColumns(CutContext ctx, List<List<int>> columns, int depth, double availRight)
        {
            var node = new CutNode { SideBySide = true, Children = new List<CutNode>(columns.Count) };
            for (int c = 0; c < columns.Count; c++)
            {
                List<int> column = columns[c];
                double left, right;
                HorizontalBounds(ctx.Boxes, column, out left, out right);
                double childAvail = availRight;
                if (c + 1 < columns.Count)
                {
                    double nextLeft, nextRight;
                    HorizontalBounds(ctx.Boxes, columns[c + 1], out nextLeft, out nextRight);
                    childAvail = nextLeft;
                }
                node.Children.Add(CutFloors(ctx, column, depth - 1, left, right, childAvail, false));
            }
            return node;
        }

        /// <summary>
        /// Вертикальный разрез с защитой: хотя бы ОДНА колонка существенна (наполнена и высока),
        /// иначе null. Все колонки — крошки («(подпись) … (дата)» одной строки) — не колонки,
        /// страница не должна читаться столбиками одиночных слов. Но существенный блок с мелким
        /// соседом — настоящая вёрстка: блок слева и одинокая пометка справа режутся,
        /// иначе пометка вклинивается в строку блока.
        /// </summary>
        private static List<List<int>> SplitColumns(CutContext ctx, List<int> items)
        {
            List<List<int>> columns = SplitAtGaps(ctx.Boxes, items, false, ctx.VGapMin);
            if (columns == null)
                return null;
            bool anySubstantial = false;
            foreach (List<int> column in columns)
            {
                if (ColumnSubstantial(ctx, column))
                {
                    anySubstantial = true;
                    break;
                }
            }
            if (!anySubstantial)
                return null;
            // Сначала маркеры списка: иначе collapse склеит «• | текст | вложенный» в две
            // газетные колонки и сломает уровень (RC-9). Правая узкая пометка — режется.
            if (ctx.MinColumnWidth > 0 && columns.Count >= 2)
            {
                double left, right;
                HorizontalBounds(ctx.Boxes, columns[0], out left, out right);
                if (right - left < ctx.MinColumnWidth)
                    return null;
            }
            columns = CollapseJustifyRivers(ctx, columns);
            return columns.Count >= 2 ? columns : null;
        }

        /// <summary>
        /// Выключка Courier/газеты даёт «реки» — пустые полосы внутри колонки, и разрез
        /// видит 3+ столбика. Настоящий канал — у середины набора. Равные по ширине
        /// три колонки (бюллетень) не трогаем.
        /// </summary>
        private static List<List<int>> CollapseJustifyRivers(CutContext ctx, List<List<int>> columns)
        {
            if (columns.Count < 3)
                return columns;
            var lefts = new double[columns.Count];
            var rights = new double[columns.Count];
            double minW = double.MaxValue, maxW = 0;
            for (int i = 0; i < columns.Count; i++)
            {
                HorizontalBounds(ctx.Boxes, columns[i], out lefts[i], out rights[i]);
                double w = rights[i] - lefts[i];
                if (w < minW) minW = w;
                if (w > maxW) maxW = w;
            }
            if (minW > 0 && maxW <= 2.0 * minW)
                return columns;
            double mid = (lefts[0] + rights[columns.Count - 1]) / 2.0;
            int best = 0;
            double bestDist = double.MaxValue;
            for (int i = 0; i < columns.Count - 1; i++)
            {
                double gutter = (rights[i] + lefts[i + 1]) / 2.0;
                double dist = gutter > mid ? gutter - mid : mid - gutter;
                if (dist < bestDist)
                {
                    bestDist = dist;
                    best = i;
                }
            }
            var left = new List<int>();
            var right = new List<int>();
            for (int i = 0; i < columns.Count; i++)
            {
                if (i <= best) left.AddRange(columns[i]);
                else right.AddRange(columns[i]);
            }
            return new List<List<int>> { left, right };
        }

        /// <summary>
        /// Разрезать элементы по пустым полосам вдоль оси. Проекции боксов на ось объединяются;
        /// разрыв объединения шире minGap — граница частей. Части идут в порядке чтения оси:
        /// по Y — сверху вниз (ось PDF направлена вверх, поэтому координата инвертируется),
        /// по X — слева направо. Меньше двух частей — null (резать негде).
        /// </summary>
        private static List<List<int>> SplitAtGaps(IList<CutBox> boxes, List<int> items, bool byY, double minGap)
        {
            if (items.Count < 2 || minGap == double.MaxValue)
                return null;
            var order = new List<int>(items);
            if (byY)
                order.Sort(delegate(int a, int b)
                {
                    int c = boxes[b].Top.CompareTo(boxes[a].Top);
                    return c != 0 ? c : boxes[a].Tag.CompareTo(boxes[b].Tag); // стабильность — детерминизм
                });
            else
                order.Sort(delegate(int a, int b)
                {
                    int c = boxes[a].Left.CompareTo(boxes[b].Left);
                    return c != 0 ? c : boxes[a].Tag.CompareTo(boxes[b].Tag);
                });

            var parts = new List<List<int>>();
            var current = new List<int>();
            double reach = 0;
            bool first = true;
            foreach (int i in order)
            {
                // Координата «вперёд по чтению»: вниз для Y (инверсия знака), вправо для X.
                double lo = byY ? -boxes[i].Top : boxes[i].Left;
                double hi = byY ? -boxes[i].Bottom : boxes[i].Right;
                if (hi < lo) hi = lo; // вырожденный бокс не ломает объединение
                // Inclusive: a 12 pt newspaper gutter with vGapMin=12 must split
                // (PageBlocks.BlockColumnGapPt). Strict > left a 12 pt channel as one
                // leaf and zigzagged L/R paragraphs (RC-14, two-col odd variants).
                if (!first && lo - reach >= minGap)
                {
                    parts.Add(current);
                    current = new List<int>();
                }
                current.Add(i);
                if (first || hi > reach)
                    reach = hi;
                first = false;
            }
            parts.Add(current);
            return parts.Count >= 2 ? parts : null;
        }

        /// <summary>
        /// Заголовок над каналом колонок: пустой полосы до HGapMin нет, вертикальный
        /// разрез всего набора тоже пуст. Снимаем верхнюю полосу, только если:
        /// остаток — две настоящие колонки, снятое не само колонки, оно перекрывает
        /// канал, тонкое, и НИ один бокс вне полосы в канал не заходит (иначе первая
        /// строка формы «метка … значение» стала бы этажом).
        /// </summary>
        private static List<List<int>> PeelBlockingFloor(CutContext ctx, List<int> items)
        {
            if (ctx.VGapMin == double.MaxValue)
                return null;
            double small = ctx.HGapMin * 0.25;
            if (small > 2.0) small = 2.0;
            if (small < 0.5) small = 0.5;
            List<List<int>> bands = SplitAtGaps(ctx.Boxes, items, true, small);
            if (bands == null || bands.Count < 2)
                return null;
            var prefix = new List<int>();
            for (int i = 0; i < bands.Count - 1; i++)
            {
                prefix.AddRange(bands[i]);
                var rest = new List<int>();
                for (int j = i + 1; j < bands.Count; j++)
                    rest.AddRange(bands[j]);
                List<List<int>> cols = SplitColumns(ctx, rest);
                if (cols == null || SubstantialColumnCount(ctx, cols) < 2)
                    continue;
                if (SplitAtGaps(ctx.Boxes, prefix, false, ctx.VGapMin) != null)
                    continue;
                double gL, gR;
                if (!ColumnGutter(ctx, cols[0], cols[1], out gL, out gR))
                    continue;
                if (!OverlapsInterval(ctx, prefix, gL, gR))
                    continue;
                if (OverlapsInterval(ctx, rest, gL, gR))
                    continue;
                if (VerticalExtent(ctx, prefix) > 1.2 * ctx.MinColumnExtent)
                    continue;
                return new List<List<int>> { new List<int>(prefix), rest };
            }
            return null;
        }

        private static int SubstantialColumnCount(CutContext ctx, List<List<int>> columns)
        {
            int n = 0;
            foreach (List<int> column in columns)
                if (ColumnSubstantial(ctx, column))
                    n++;
            return n;
        }

        private static bool ColumnGutter(CutContext ctx, List<int> leftCol, List<int> rightCol, out double gL, out double gR)
        {
            double l0, r0, l1, r1;
            HorizontalBounds(ctx.Boxes, leftCol, out l0, out r0);
            HorizontalBounds(ctx.Boxes, rightCol, out l1, out r1);
            gL = r0;
            gR = l1;
            return gR > gL;
        }

        private static bool OverlapsInterval(CutContext ctx, List<int> items, double lo, double hi)
        {
            foreach (int i in items)
                if (ctx.Boxes[i].Left < hi && ctx.Boxes[i].Right > lo)
                    return true;
            return false;
        }

        private static bool ColumnsComparableHeight(CutContext ctx, List<List<int>> columns)
        {
            if (columns.Count < 2)
                return false;
            double minE = double.MaxValue, maxE = 0;
            foreach (List<int> column in columns)
            {
                double e = VerticalExtent(ctx, column);
                if (e < minE) minE = e;
                if (e > maxE) maxE = e;
            }
            return maxE > 0 && minE >= 0.4 * maxE;
        }

        private static double VerticalExtent(CutContext ctx, List<int> items)
        {
            double top = double.MinValue, bottom = double.MaxValue;
            foreach (int i in items)
            {
                if (ctx.Boxes[i].Top > top) top = ctx.Boxes[i].Top;
                if (ctx.Boxes[i].Bottom < bottom) bottom = ctx.Boxes[i].Bottom;
            }
            return top - bottom;
        }

        private static bool ColumnSubstantial(CutContext ctx, List<int> column)
        {
            if (column.Count < ctx.MinColumnItems)
                return false;
            double top = double.MinValue, bottom = double.MaxValue;
            foreach (int i in column)
            {
                if (ctx.Boxes[i].Top > top) top = ctx.Boxes[i].Top;
                if (ctx.Boxes[i].Bottom < bottom) bottom = ctx.Boxes[i].Bottom;
            }
            return top - bottom >= ctx.MinColumnExtent;
        }

        private static CutNode MakeLeaf(CutContext ctx, List<int> items, double colLeft, double colRight, double availRight)
        {
            var order = new List<int>(items);
            order.Sort(delegate(int a, int b)
            {
                int c = ctx.Boxes[b].Top.CompareTo(ctx.Boxes[a].Top);
                return c != 0 ? c : ctx.Boxes[a].Left.CompareTo(ctx.Boxes[b].Left);
            });
            var leaf = new CutNode { Tags = new List<int>(order.Count), ColumnLeft = colLeft, ColumnRight = colRight, AvailRight = availRight };
            foreach (int i in order)
                leaf.Tags.Add(ctx.Boxes[i].Tag);
            return leaf;
        }

        /// <summary>Левая/правая граница набора боксов.</summary>
        private static void HorizontalBounds(IList<CutBox> boxes, List<int> items, out double left, out double right)
        {
            left = double.MaxValue;
            right = double.MinValue;
            foreach (int i in items)
            {
                if (boxes[i].Left < left) left = boxes[i].Left;
                if (boxes[i].Right > right) right = boxes[i].Right;
            }
        }
    }
}
