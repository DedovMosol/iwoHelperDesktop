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

        /// <summary>Параметры и результат одного разреза — чтобы не тянуть шесть аргументов по рекурсии.</summary>
        private sealed class CutContext
        {
            public IList<CutBox> Boxes;
            public double HGapMin;          // минимальная высота пустой полосы для этажа
            public double VGapMin;          // минимальная ширина пустого просвета для колонки
            public double MinColumnExtent;  // колонка ниже этого по высоте — не колонка (подпись/дата в одной строке)
            public int MinColumnItems;      // и колонка из меньшего числа элементов — не колонка
            public List<CutLeaf> Leaves;
        }

        /// <summary>
        /// Разрезать боксы на блоки в порядке чтения. hGapMin/vGapMin — минимальные пустые
        /// полосы (pt) для горизонтального/вертикального разреза; minColumnExtent и
        /// minColumnItems — защита от ложных колонок (широкий пробел одной строки — «подпись …
        /// дата» — не делит её на столбики: каждая колонка обязана быть достаточно высокой и
        /// содержать несколько элементов, иначе вертикальный разрез отклоняется целиком).
        /// </summary>
        public static List<CutLeaf> Order(IList<CutBox> boxes, double hGapMin, double vGapMin,
            double minColumnExtent, int minColumnItems)
        {
            var leaves = new List<CutLeaf>();
            if (boxes == null || boxes.Count == 0)
                return leaves;
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
                Leaves = leaves
            };
            double left, right;
            HorizontalBounds(boxes, all, out left, out right);
            CutFloors(ctx, all, MaxDepth, left, right);
            return leaves;
        }

        /// <summary>Этажи области (сверху вниз); в каждом — попытка колонок. Нет этажей — сразу колонки.</summary>
        private static void CutFloors(CutContext ctx, List<int> items, int depth, double colLeft, double colRight)
        {
            List<List<int>> floors = depth > 0 ? SplitAtGaps(ctx.Boxes, items, true, ctx.HGapMin) : null;
            if (floors == null)
            {
                CutColumns(ctx, items, depth, colLeft, colRight);
                return;
            }
            foreach (List<int> floor in floors)
                CutColumns(ctx, floor, depth - 1, colLeft, colRight);
        }

        /// <summary>Колонки этажа (слева направо); внутри колонки — снова этажи. Нет колонок — лист.</summary>
        private static void CutColumns(CutContext ctx, List<int> items, int depth, double colLeft, double colRight)
        {
            List<List<int>> columns = depth > 0 ? SplitColumns(ctx, items) : null;
            if (columns == null)
            {
                var leaf = new CutLeaf { ColumnLeft = colLeft, ColumnRight = colRight };
                leaf.Tags.Capacity = items.Count;
                foreach (int i in items)
                    leaf.Tags.Add(ctx.Boxes[i].Tag);
                ctx.Leaves.Add(leaf);
                return;
            }
            foreach (List<int> column in columns)
            {
                // Новая колонка — новая опорная рамка для геометрии её содержимого.
                double left, right;
                HorizontalBounds(ctx.Boxes, column, out left, out right);
                CutFloors(ctx, column, depth - 1, left, right);
            }
        }

        /// <summary>Вертикальный разрез с защитой: каждая колонка существенна, иначе null (не режем).</summary>
        private static List<List<int>> SplitColumns(CutContext ctx, List<int> items)
        {
            List<List<int>> columns = SplitAtGaps(ctx.Boxes, items, false, ctx.VGapMin);
            if (columns == null)
                return null;
            foreach (List<int> column in columns)
            {
                if (column.Count < ctx.MinColumnItems)
                    return null;
                double top = double.MinValue, bottom = double.MaxValue;
                foreach (int i in column)
                {
                    if (ctx.Boxes[i].Top > top) top = ctx.Boxes[i].Top;
                    if (ctx.Boxes[i].Bottom < bottom) bottom = ctx.Boxes[i].Bottom;
                }
                if (top - bottom < ctx.MinColumnExtent)
                    return null;
            }
            return columns;
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
                if (!first && lo > reach + minGap)
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
