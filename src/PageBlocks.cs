using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Геометрия страницы: разобранное содержимое (абзацы, таблицы, изображения) → элементы
    /// в порядке чтения. Чистая математика без Word, PowerPoint и PdfPig, покрыта юнит-тестами.
    ///
    /// Две группы, разделённые по назначению:
    /// 1. <see cref="OrderedItems"/> — порядок чтения (XY-разрез, <see cref="XyCut"/>). Нужен
    ///    ЛЮБОМУ писателю: в потоковом выводе он задаёт порядок абзацев, в выводе с абсолютными
    ///    координатами — порядок фигур (z-порядок и обход с клавиатуры).
    /// 2. <see cref="CoalesceRowBands"/>, <see cref="TypicalItemGap"/>, <see cref="ExtraGapPt"/> —
    ///    подпорки ПОТОКОВОЙ вёрстки: воспроизвести колонки таблицей и вернуть вертикальные
    ///    зазоры интервалами. Писателю, который ставит фигуры по координатам, они не нужны —
    ///    у него зазоры получаются сами. Живут здесь, потому что делят с первой группой рамки
    ///    блоков и пороги канала: растащить по файлам значило бы размножить константы.
    /// </summary>
    internal static class PageBlocks
    {
        // Порядок блоков страницы (XY-разрез, XyCut): этаж — пустая полоса заметно шире
        // межстрочной (иначе этажи режут каждый абзац, что безвредно, но просветы колонок
        // ищутся в слишком мелких этажах); колонка — просвет уже канала двухколоночной шапки
        // (≈22pt с учётом картинки, пересекающей канал), но шире любых межсловных зазоров.
        // Колонка обязана быть существенной — см. параметры XyCut.Order.
        private const double BlockFloorGapPt = 14;
        // Канал 12 pt: между колонкой подписи и штампом рядом бывает всего ~17 pt, а
        // межсловные зазоры текста на порядок уже — ложных колонок канал не даёт.
        private const double BlockColumnGapPt = 12;
        private const double BlockColumnMinExtentPt = 24;
        // На уровне БЛОКОВ достаточно ОДНОГО существенного блока в колонке: «подпись слева —
        // Ф.И.О. справа» и «Приложения: — список справа» — это одиночные (но высокие) блоки
        // рядом. Одиночную короткую строку («(подпись) … (дата)») по-прежнему отсекает высотный
        // гейт BlockColumnMinExtentPt. Словный уровень (OcrLayout) держит свой порог в 2 слова.
        private const int BlockColumnMinItems = 1;

        // Воспроизведение ВЕРТИКАЛЬНЫХ зазоров между блоками: исходник разделяет зоны пустыми
        // строками (пометка ниже верхнего блока, подпись ниже текста), а вывод впритык терял
        // этот ритм. Лишний зазор (сверх типичного межблочного) добавляется интервалом перед
        // блоком; порог отсеивает шум обычной вёрстки. Кап пропускает и нижний блок у нижнего
        // края почти пустой страницы (сотни pt), а от патологий страхует демпфер писателя:
        // интервалы никогда не выталкивают лишнюю страницу.
        private const double MinBlockGapExtraPt = 6;
        private const double BlockGapCapPt = 400;

        /// <summary>Блок содержимого страницы: абзац, таблица или изображение (одно из полей задано).</summary>
        internal sealed class Block
        {
            public OcrParagraph Paragraph;
            public OcrTable Table;
            public OcrImage Image;
            public double Top;    // верх блока (Y, ось вверх) — основной порядок чтения внутри листа разреза
            public double Left;   // левый край — вторичный порядок для блоков в одной строке-полосе
            public double Right;  // правый/нижний края рамки — для XY-разреза страницы на этажи и колонки
            public double Bottom;
        }

        /// <summary>Элемент вывода страницы: одиночный блок ИЛИ side-by-side полоса колонок.</summary>
        internal sealed class PageItem
        {
            public Block Single;               // одиночный блок (Columns == null)
            public List<List<Block>> Columns;  // полоса: колонки блоков слева направо, внутри — сверху вниз
            public double[] ColLeft;           // левые/правые границы колонок (pt) — для ширин ячеек
            public double[] ColRight;
            public double Top;                 // рамка элемента (Y, ось вверх) — для межблочных зазоров
            public double Bottom;
            public bool IsBand { get { return Columns != null; } }
        }

        /// <summary>
        /// Элементы страницы в порядке чтения (XY-дерево, <see cref="XyCut"/>). Обычный узел даёт
        /// одиночные блоки; узел «бок о бок» (колонки этажа) с пригодным содержимым — одну полосу
        /// (её рендер — безграничная таблица, колонки сидят рядом, как двухколоночная шапка), иначе
        /// раскрывается последовательно. Одноколоночная страница — цепочка одиночных блоков, порядок
        /// эквивалентен прежнему. Чистая — под тест.
        /// </summary>
        internal static List<PageItem> OrderedItems(PdfPageText page)
        {
            var blocks = new List<Block>();
            if (page.Paragraphs != null)
                foreach (OcrParagraph par in page.Paragraphs)
                    blocks.Add(new Block { Paragraph = par, Top = par.TopPt, Left = par.LeftPt, Right = par.RightPt, Bottom = par.BottomPt });
            if (page.Tables != null)
                foreach (OcrTable table in page.Tables)
                    blocks.Add(new Block { Table = table, Top = table.TopPt, Left = table.LeftPt, Right = table.RightPt, Bottom = table.BottomPt });
            if (page.Images != null)
                foreach (OcrImage img in page.Images)
                    blocks.Add(new Block { Image = img, Top = img.TopPt, Left = img.LeftPt, Right = img.LeftPt + img.WidthPt, Bottom = img.TopPt - img.HeightPt });

            var items = new List<PageItem>(blocks.Count);
            if (blocks.Count == 0)
                return items;

            var boxes = new CutBox[blocks.Count];
            for (int i = 0; i < blocks.Count; i++)
            {
                Block b = blocks[i];
                // Защита от некорректной рамки (правее левого, ниже верхнего) — не ломаем разрез.
                double right = b.Right > b.Left ? b.Right : b.Left;
                double bottom = b.Bottom <= b.Top ? b.Bottom : b.Top;
                boxes[i] = new CutBox { Left = b.Left, Right = right, Bottom = bottom, Top = b.Top, Tag = i };
            }
            CutNode root = XyCut.OrderTree(boxes, BlockFloorGapPt, BlockColumnGapPt,
                BlockColumnMinExtentPt, BlockColumnMinItems);
            WalkNode(root, blocks, items);
            return items;
        }

        /// <summary>
        /// Собрать подряд идущие ОДИНОЧНЫЕ блоки одной «строки» (взаимное вертикальное
        /// перекрытие ≥ половины меньшей высоты, зазор по X не меньше блочного канала, без
        /// пересечений по X) в side-by-side полосу: поля слева и боковая метка справа,
        /// «(подпись) … (дата)», наложенная картинка рядом с правым блоком. XyCut такие полосы не
        /// строит — у одиночной строки не хватает высоты на колонку. Требуется хотя бы один
        /// абзац в полосе (двум соседним картинкам полоса не нужна). Чистая — под тест.
        /// </summary>
        internal static List<PageItem> CoalesceRowBands(List<PageItem> items)
        {
            var result = new List<PageItem>(items.Count);
            int i = 0;
            while (i < items.Count)
            {
                int j = i;
                while (j + 1 < items.Count && RowNeighbours(items[j], items[j + 1]))
                    j++;
                if (j == i)
                {
                    result.Add(items[i]);
                    i++;
                    continue;
                }
                bool anyText = false;
                for (int k = i; k <= j && !anyText; k++)
                    anyText = items[k].Single.Paragraph != null;
                if (!anyText)
                {
                    for (int k = i; k <= j; k++)
                        result.Add(items[k]);
                    i = j + 1;
                    continue;
                }
                var band = new PageItem
                {
                    Columns = new List<List<Block>>(j - i + 1),
                    ColLeft = new double[j - i + 1],
                    ColRight = new double[j - i + 1],
                    Top = double.MinValue,
                    Bottom = double.MaxValue
                };
                for (int k = i; k <= j; k++)
                {
                    Block b = items[k].Single;
                    band.Columns.Add(new List<Block> { b });
                    band.ColLeft[k - i] = b.Left;
                    band.ColRight[k - i] = Math.Max(b.Right, b.Left);
                    if (items[k].Top > band.Top) band.Top = items[k].Top;
                    if (items[k].Bottom < band.Bottom) band.Bottom = items[k].Bottom;
                }
                result.Add(band);
                i = j + 1;
            }
            return result;
        }

        /// <summary>Соседи ли по «строке»: одиночные абзац/картинка, сильное вертикальное перекрытие, канал по X.</summary>
        private static bool RowNeighbours(PageItem left, PageItem right)
        {
            if (left.IsBand || right.IsBand || left.Single.Table != null || right.Single.Table != null)
                return false;
            double overlap = Math.Min(left.Top, right.Top) - Math.Max(left.Bottom, right.Bottom);
            double minH = Math.Min(left.Top - left.Bottom, right.Top - right.Bottom);
            if (minH <= 0 ? overlap < 0 : overlap < 0.5 * minH)
                return false; // не одна «строка»
            return right.Single.Left - Math.Max(left.Single.Right, left.Single.Left) >= BlockColumnGapPt;
        }

        /// <summary>
        /// Типичный вертикальный зазор между соседними блоками страницы (нижняя медиана
        /// положительных): пары соседних элементов потока + пары соседних блоков внутри колонок
        /// полос. Перекрытия и «обратные» пары (колонки, раскрытые последовательно) не в счёт.
        /// 0 — зазоров нет (один блок). Чистая — под тест.
        /// </summary>
        internal static double TypicalItemGap(List<PageItem> items)
        {
            var gaps = new List<double>();
            for (int i = 1; i < items.Count; i++)
            {
                double g = items[i - 1].Bottom - items[i].Top;
                if (g > 0 && g < BlockGapCapPt)
                    gaps.Add(g);
            }
            foreach (PageItem item in items)
            {
                if (!item.IsBand)
                    continue;
                foreach (List<Block> col in item.Columns)
                    for (int i = 1; i < col.Count; i++)
                    {
                        double g = Math.Min(col[i - 1].Bottom, col[i - 1].Top) - col[i].Top;
                        if (g > 0 && g < BlockGapCapPt)
                            gaps.Add(g);
                    }
            }
            return MathUtil.Median(gaps);
        }

        /// <summary>
        /// Интервал перед блоком (pt): лишний зазор сверх типичного, если он заметен
        /// (не меньше <see cref="MinBlockGapExtraPt"/>), с капом. Иначе 0. Чистая — под тест.
        /// </summary>
        internal static double ExtraGapPt(double gap, double typicalGap)
        {
            double extra = gap - typicalGap;
            if (extra < MinBlockGapExtraPt)
                return 0;
            return extra > BlockGapCapPt ? BlockGapCapPt : extra;
        }

        /// <summary>Обход XY-дерева в элементы: лист → блоки; узел «бок о бок» → полоса (если пригоден); иначе вглубь.</summary>
        private static void WalkNode(CutNode node, List<Block> blocks, List<PageItem> items)
        {
            if (node.IsLeaf)
            {
                foreach (Block b in LeafBlocks(node, blocks))
                    items.Add(new PageItem { Single = b, Top = b.Top, Bottom = Math.Min(b.Bottom, b.Top) });
                return;
            }
            if (node.SideBySide)
            {
                PageItem band = TryBuildBand(node, blocks);
                if (band != null)
                {
                    items.Add(band);
                    return;
                }
            }
            foreach (CutNode child in node.Children)
                WalkNode(child, blocks, items);
        }

        /// <summary>Блоки листа в порядке чтения (та же логика полос, что и раньше).</summary>
        private static List<Block> LeafBlocks(CutNode leaf, List<Block> blocks)
        {
            var group = new List<Block>(leaf.Tags.Count);
            foreach (int tag in leaf.Tags)
                group.Add(blocks[tag]);
            var ordered = new List<Block>(group.Count);
            OrderWithinLeaf(group, ordered);
            return ordered;
        }

        /// <summary>Все блоки поддерева в порядке чтения (обход в порядке детей: этажи сверху вниз, колонки слева направо).</summary>
        private static List<Block> CollectBlocks(CutNode node, List<Block> blocks)
        {
            if (node.IsLeaf)
                return LeafBlocks(node, blocks);
            var result = new List<Block>();
            foreach (CutNode child in node.Children)
                result.AddRange(CollectBlocks(child, blocks));
            return result;
        }

        /// <summary>
        /// Построить side-by-side полосу из узла-колонок, если она пригодна: ≥2 колонок, ни в одной
        /// нет ТАБЛИЦЫ (вложенную рамочную таблицу в ячейку шапки не тащим) и минимум в двух колонках
        /// есть абзац (чисто-картиночные пары — не шапка). Иначе null (обход раскроет последовательно).
        /// </summary>
        private static PageItem TryBuildBand(CutNode node, List<Block> blocks)
        {
            int n = node.Children.Count;
            if (n < 2)
                return null;
            var cols = new List<List<Block>>(n);
            var colLeft = new double[n];
            var colRight = new double[n];
            int columnsWithText = 0;
            double top = double.MinValue, bottom = double.MaxValue;
            for (int c = 0; c < n; c++)
            {
                List<Block> col = CollectBlocks(node.Children[c], blocks);
                if (col.Count == 0)
                    return null;
                bool hasText = false;
                double left = double.MaxValue, right = double.MinValue;
                foreach (Block b in col)
                {
                    if (b.Table != null)
                        return null; // рамочная таблица в колонке — полосу не строим
                    if (b.Paragraph != null)
                        hasText = true;
                    if (b.Left < left) left = b.Left;
                    if (b.Right > right) right = b.Right;
                    if (b.Top > top) top = b.Top;
                    double bb = Math.Min(b.Bottom, b.Top);
                    if (bb < bottom) bottom = bb;
                }
                if (hasText)
                    columnsWithText++;
                cols.Add(col);
                colLeft[c] = left;
                colRight[c] = right;
            }
            if (columnsWithText < 2)
                return null;
            return new PageItem { Columns = cols, ColLeft = colLeft, ColRight = colRight, Top = top, Bottom = bottom };
        }

        /// <summary>
        /// Порядок блоков внутри листа разреза: сверху вниз; блоки одной «строки-полосы» —
        /// слева направо. Полоса собирается по РЕАЛЬНОМУ вертикальному перекрытию рамок не
        /// меньше половины меньшей высоты (как строки из слов в OcrLayout): таблицы/печати бок
        /// о бок перекрыты сильно и читаются слева направо, а соседние СТРОКИ текста не
        /// перекрываются вовсе — полоса «по близости верхов» здесь переставляла бы их местами
        /// (поля шапки читались снизу вверх). Вырожденная высота (линейка) — по знаку
        /// перекрытия. Добавляет блоки листа в ordered.
        /// </summary>
        private static void OrderWithinLeaf(List<Block> group, List<Block> ordered)
        {
            group.Sort(delegate(Block a, Block b)
            {
                int c = b.Top.CompareTo(a.Top); // ось Y вверх: больший Top — выше — раньше
                return c != 0 ? c : a.Left.CompareTo(b.Left);
            });
            var band = new List<Block>();
            double bandTop = 0, bandBottom = 0;
            foreach (Block blk in group)
            {
                double top = blk.Top;
                double bottom = blk.Bottom <= top ? blk.Bottom : top; // защита от некорректной рамки
                if (band.Count > 0 && BandOverlaps(bandTop, bandBottom, top, bottom))
                {
                    if (top > bandTop) bandTop = top;
                    if (bottom < bandBottom) bandBottom = bottom;
                }
                else
                {
                    FlushBand(band, ordered);
                    bandTop = top;
                    bandBottom = bottom;
                }
                band.Add(blk);
            }
            FlushBand(band, ordered);
        }

        /// <summary>Перекрывается ли рамка [bottom..top] с полосой не меньше чем на половину меньшей высоты.</summary>
        private static bool BandOverlaps(double bandTop, double bandBottom, double top, double bottom)
        {
            double overlap = Math.Min(bandTop, top) - Math.Max(bandBottom, bottom);
            double minH = Math.Min(bandTop - bandBottom, top - bottom);
            if (minH <= 0)
                return overlap > 0; // вырожденная высота — достаточно пересечения по знаку
            return overlap >= 0.5 * minH;
        }

        /// <summary>Высыпать полосу в порядок чтения: слева направо (при равенстве — выше раньше).</summary>
        private static void FlushBand(List<Block> band, List<Block> ordered)
        {
            if (band.Count == 0)
                return;
            band.Sort(delegate(Block a, Block b)
            {
                int c = a.Left.CompareTo(b.Left);
                return c != 0 ? c : b.Top.CompareTo(a.Top);
            });
            ordered.AddRange(band);
            band.Clear();
        }
    }
}
