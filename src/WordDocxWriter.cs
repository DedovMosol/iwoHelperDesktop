using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Запись извлечённого born-digital PDF в .docx через COM Word: абзацы и изображения в
    /// порядке чтения (сверху вниз), разрыв страницы между страницами PDF. Каркас Word
    /// (открытие/сохранение/закрытие) — общий <see cref="WordCom"/> (DRY). Вызывать в STA-потоке.
    /// </summary>
    public static class WordDocxWriter
    {
        private const int WdAlignLeft = 0;
        private const int WdAlignCenter = 1;
        private const int WdAlignJustify = 3;
        private const int WdSectionBreakNextPage = 2; // каждая PDF-страница — свой раздел (свой размер листа)
        private const int WdCollapseStart = 1;
        private const int WdCollapseEnd = 0;
        private const int WdStyleNormal = -1;      // wdStyleNormal — базовый стиль абзаца документа
        private const int WdLineSpaceSingle = 0;   // одинарный межстрочный интервал
        private const double MinColWidthPt = 6;   // защита от вырожденной колонки
        private const string DefaultFontName = "Times New Roman";
        private const double DefaultFontSize = 12;
        private const double MinFontSize = 5;   // защита от мусорного кегля из PDF
        private const double MaxFontSize = 72;
        private const double MinPagePt = 72;    // 1"; разумные пределы размера страницы
        private const double MaxPagePt = 1584;  // 22" — максимум Word

        /// <summary>Пишет .docx из абзацев и изображений страниц. Занятый файл/нет Word — MergeException.</summary>
        public static void Write(IList<PdfPageText> pages, string path, Action<int, int> progress = null)
        {
            if (pages == null)
                throw new ArgumentNullException("pages");

            double firstLineIndent = DocumentIndent(pages); // pt; 0 — документ без красной строки
            string tempDir = Path.Combine(Path.GetTempPath(), "iwo_img_" + Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(tempDir);
                int imgIndex = 0;
                WordCom.WriteDocx(path, Loc.T("word.label.docx"), delegate(object wordObj, object docObj)
                {
                    dynamic word = wordObj;
                    dynamic doc = docObj;
                    dynamic sel = word.Selection;
                    ApplyDocumentDefaults(doc); // единый интервал — детерминизм и плотность born-digital оригинала
                    ListTemplates lists = ListTemplates.Load(word); // галереи нумерованного/маркированного списка
                    var listState = new ListState();

                    for (int p = 0; p < pages.Count; p++)
                    {
                        if (p > 0)
                            sel.InsertBreak(WdSectionBreakNextPage); // новый раздел = свой размер листа
                        ApplySectionSetup(sel, pages[p]); // размер и поля страницы из источника
                        // Текстовая область страницы (pt): в неё Word укладывает абзацы, от неё
                        // считаются отступы конфайна центрированной колонки (см. WriteParagraphInto).
                        double textLeft = pages[p].LeftMarginPt;
                        double textRight = pages[p].WidthPt - pages[p].RightMarginPt;
                        List<PageItem> items = OrderedItems(pages[p]);
                        double typicalGap = TypicalItemGap(items);
                        PageItem prev = null;
                        foreach (PageItem item in items)
                        {
                            // Лишний вертикальный зазор до предыдущего блока → интервал перед этим.
                            double spaceBefore = prev == null ? 0 : ExtraGapPt(prev.Bottom - item.Top, typicalGap);
                            prev = item;
                            if (item.IsBand)
                            {
                                ClearInheritedList(sel, listState); // таблица обрывает список
                                WriteColumnBand(word, doc, sel, item, textLeft, textRight, pages[p].WidthPt,
                                    tempDir, ref imgIndex, spaceBefore, typicalGap);
                                continue;
                            }
                            Block blk = item.Single;
                            if (blk.Paragraph != null)
                                WriteParagraph(sel, doc, blk.Paragraph, firstLineIndent, lists, listState, textLeft, textRight, spaceBefore);
                            else
                            {
                                ClearInheritedList(sel, listState); // таблица/картинка обрывают список
                                if (blk.Table != null)
                                {
                                    if (spaceBefore > 0)
                                        InsertSpacer(sel, spaceBefore); // таблице SpaceBefore не задать — пустой абзац той же высоты
                                    WriteTable(word, doc, sel, blk.Table);
                                }
                                else
                                    InsertImage(sel, blk.Image, pages[p].WidthPt, tempDir, ref imgIndex, spaceBefore);
                            }
                        }
                        if (progress != null)
                            progress(p + 1, pages.Count);
                    }
                    ClearInheritedList(sel, listState); // хвостовой пустой абзац не должен унаследовать маркер
                    FitSpacingToPages(doc, pages.Count); // интервалы не должны выталкивать лишнюю страницу
                });
            }
            finally
            {
                try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
            }
        }

        /// <summary>
        /// Единые интервалы документа через стиль «Обычный»: одинарный межстрочный, без отбивки
        /// до и после абзаца. Иначе .docx наследует умолчания Normal.dotm пользователя (в «офисном»
        /// шаблоне — 1.08 строки + 8 pt после КАЖДОГО абзаца), из-за чего плотный born-digital
        /// оригинал (где абзацы разделены красной строкой, а не пустотой) раздувается на лишние
        /// страницы, а сама разбивка становится машинозависимой. Косметика: сбой стиля не срывает
        /// сохранение — тогда остаются умолчания шаблона.
        /// </summary>
        private static void ApplyDocumentDefaults(dynamic doc)
        {
            try
            {
                dynamic normal = doc.Styles.Item(WdStyleNormal).ParagraphFormat;
                normal.SpaceBefore = 0;
                normal.SpaceAfter = 0;
                normal.LineSpacingRule = WdLineSpaceSingle;
            }
            catch { } // интервалы косметические — при сбое стиля просто наследуем шаблон
        }

        // Порядок блоков страницы (XY-разрез, XyCut): этаж — пустая полоса заметно шире
        // межстрочной (иначе этажи режут каждый абзац, что безвредно, но просветы колонок
        // ищутся в слишком мелких этажах); колонка — просвет уже канала двухколоночной шапки
        // (≈22pt с учётом картинки, пересекающей канал), но шире любых межсловных зазоров.
        // Колонка обязана быть существенной — см. параметры XyCut.Order.
        private const double BlockFloorGapPt = 14;
        private const double BlockColumnGapPt = 18;
        private const double BlockColumnMinExtentPt = 24;
        // На уровне БЛОКОВ достаточно ОДНОГО существенного блока в колонке: «подпись слева —
        // Ф.И.О. справа» и «Приложения: — список справа» — это одиночные (но высокие) блоки
        // рядом. Одиночную короткую строку («(подпись) … (дата)») по-прежнему отсекает высотный
        // гейт BlockColumnMinExtentPt. Словный уровень (OcrLayout) держит свой порог в 2 слова.
        private const int BlockColumnMinItems = 1;

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

        // Воспроизведение ВЕРТИКАЛЬНЫХ зазоров между блоками: исходник разделяет зоны пустыми
        // строками («(по списку)» ниже адресата, подпись ниже текста), а вывод впритык терял
        // этот ритм. Лишний зазор (сверх типичного межблочного) добавляется интервалом перед
        // блоком; порог отсеивает шум обычной вёрстки, кап — страховка от «пустой полустраницы».
        private const double MinBlockGapExtraPt = 6;
        private const double BlockGapCapPt = 120;

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
        /// (реквизиты бланка читались снизу вверх). Вырожденная высота (линейка) — по знаку
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

        private static void WriteParagraph(dynamic sel, dynamic doc, OcrParagraph paragraph, double firstLineIndent, ListTemplates lists, ListState state, double textLeftPt, double textRightPt, double spaceBeforePt = 0)
        {
            bool asList = lists.Available && paragraph.ListKind != ListKind.None;
            if (asList)
            {
                // Маркер («1.», «•») снимаем — Word рисует свой; отступ задаёт шаблон списка (indent=0).
                WriteParagraphInto(sel, doc, paragraph, 0, false, paragraph.ListContentStart, textLeftPt, textRightPt, spaceBeforePt);
                ApplyList(sel, lists, paragraph, state);
            }
            else
            {
                ClearInheritedList(sel, state); // после пункта списка следующий абзац не должен унаследовать маркер
                WriteParagraphInto(sel, doc, paragraph, firstLineIndent, false, 0, textLeftPt, textRightPt, spaceBeforePt);
            }
            sel.TypeParagraph();
        }

        /// <summary>
        /// Пустой абзац-прокладка перед таблицей/полосой (им SpaceBefore не назначить):
        /// кегль 1 (почти нулевая собственная высота) + SpaceBefore = gapPt. Через SpaceBefore,
        /// а не кегль, — чтобы прокладки ужимались демпфером страниц (<see cref="FitSpacingToPages"/>)
        /// наравне с интервалами абзацев. Формат следующего абзаца очищается от наследства.
        /// </summary>
        private static void InsertSpacer(dynamic sel, double gapPt)
        {
            try
            {
                sel.ParagraphFormat.SpaceBefore = gapPt;
                sel.ParagraphFormat.FirstLineIndent = 0;
                sel.ParagraphFormat.LeftIndent = 0;
                sel.ParagraphFormat.RightIndent = 0;
                sel.Font.Size = 1;
                sel.TypeParagraph();
                sel.ParagraphFormat.SpaceBefore = 0; // наследство прокладки не должно уехать за таблицу
            }
            catch { } // прокладка — косметика, не срываем документ
        }

        /// <summary>
        /// Демпфер пагинации: если добавленные интервалы вытолкнули документ за число страниц
        /// источника, все положительные SpaceBefore ужимаются вдвое (до двух раз), затем
        /// обнуляются. Без интервалов вывод равен прежнему (до фичи) — хуже не становится;
        /// документ, не влезавший и раньше, просто теряет интервалы. Сбой статистики/прохода
        /// не срывает сохранение.
        /// </summary>
        private static void FitSpacingToPages(dynamic doc, int sourcePages)
        {
            const int WdStatisticPages = 2;
            try
            {
                for (int attempt = 0; attempt < 3; attempt++)
                {
                    int got = (int)doc.ComputeStatistics(WdStatisticPages);
                    if (got <= sourcePages)
                        return;
                    double factor = attempt < 2 ? 0.5 : 0;
                    foreach (dynamic p in doc.Paragraphs)
                    {
                        double sb = (double)p.Format.SpaceBefore;
                        if (sb > 0)
                            p.Format.SpaceBefore = sb * factor;
                    }
                }
            }
            catch { } // интервалы — косметика; при сбое остаются как есть
        }

        /// <summary>Шаблоны нумерованного и маркированного списка Word (одни на документ). Available=false — списки не применяем.</summary>
        private sealed class ListTemplates
        {
            public dynamic Number;
            public dynamic Bullet;
            public bool Available { get { return Number != null && Bullet != null; } }

            private const int WdBulletGallery = 1;
            private const int WdNumberGallery = 2;

            /// <summary>Взять первый шаблон из галерей нумерации и маркеров. Сбой — Available=false (пишем как обычный текст).</summary>
            public static ListTemplates Load(dynamic word)
            {
                var t = new ListTemplates();
                try { t.Number = word.ListGalleries.Item(WdNumberGallery).ListTemplates.Item(1); } catch { t.Number = null; }
                try { t.Bullet = word.ListGalleries.Item(WdBulletGallery).ListTemplates.Item(1); } catch { t.Bullet = null; }
                return t;
            }
        }

        /// <summary>Состояние последовательности списка: продолжать нумерацию или начать заново.</summary>
        private sealed class ListState
        {
            public ListKind PrevKind = ListKind.None; // вид непосредственно предыдущего абзаца (для маркированного и очистки)
            public int LastNumber;                    // номер последнего нумерованного пункта; НЕ сбрасывается несписочным
                                                      // абзацем — нумерованный список продолжается сквозь вложенный текст
                                                      // (пункт может содержать внутри обычные абзацы, затем идёт следующий пункт)
        }

        private const int WdListApplyToWholeList = 0;
        private const int WdWord10ListBehavior = 2;

        /// <summary>
        /// Применить нативный список к текущему абзацу.
        ///  • Нумерованный — продолжаем ту же нумерацию, если номер ровно на 1 больше предыдущего
        ///    нумерованного пункта, ДАЖЕ если между ними были обычные абзацы (пункт с вложенным
        ///    обычным текстом внутри → следующий пункт продолжает счёт 2,3,4). Иначе начинаем
        ///    заново — так второй список с пунктами 1,2,3 стартует с 1 после первого (1–4).
        ///  • Маркированный — продолжается, только пока буллеты идут подряд (для галочек нумерация
        ///    не важна — вид одинаков).
        /// </summary>
        private static void ApplyList(dynamic sel, ListTemplates lists, OcrParagraph p, ListState state)
        {
            bool continuePrev;
            if (p.ListKind == ListKind.Numbered)
            {
                continuePrev = state.LastNumber > 0 && p.ListNumber == state.LastNumber + 1;
                state.LastNumber = p.ListNumber;
            }
            else
            {
                continuePrev = state.PrevKind == ListKind.Bulleted;
            }
            dynamic tmpl = p.ListKind == ListKind.Numbered ? lists.Number : lists.Bullet;
            try { sel.Range.ListFormat.ApplyListTemplateWithLevel(tmpl, continuePrev, WdListApplyToWholeList, WdWord10ListBehavior, 1); }
            catch { }
            state.PrevKind = p.ListKind;
        }

        /// <summary>
        /// Снять маркер списка с текущего абзаца, унаследованный от предыдущего пункта (после
        /// пункта TypeParagraph создаёт новый абзац с тем же форматом списка). Заодно убираем
        /// отступ шаблона списка. LastNumber НЕ трогаем — нумерованный список должен продолжиться
        /// после вложенных обычных абзацев. Вызываем только если список был активен — лишних
        /// COM-вызовов нет.
        /// </summary>
        private static void ClearInheritedList(dynamic sel, ListState state)
        {
            if (state.PrevKind == ListKind.None)
                return;
            try
            {
                sel.Range.ListFormat.RemoveNumbers();
                sel.ParagraphFormat.LeftIndent = 0; // снять «висячий» отступ шаблона списка
            }
            catch { }
            state.PrevKind = ListKind.None;
        }

        // Абзац центрируется в СВОЕЙ колонке (а не по странице), если колонка заметно уже
        // текстовой области — это правая колонка шапки (адресат) или левый бланк. Только для
        // центрированных: у левых/выключенных горизонталь уже задана красной строкой и полем,
        // а центрированный на всю страницу титул при конфайне в свою рамку центрируется там же.
        private const double ColumnConfineFraction = 0.72;
        private const double MinConfineIndentPt = 6; // отступ мельче — колонка практически во всю ширину

        /// <summary>
        /// Отступы для конфайна центрированного абзаца в его колонку (pt). Возвращает false и
        /// нулевые отступы, если конфайн не нужен: не центрированный/в ячейке (eligible=false),
        /// вырожденная область, или колонка почти во всю ширину (полноширинный текст — титул на
        /// всю страницу центрируется как есть). Иначе левый/правый отступ = смещение колонки от
        /// краёв текстовой области (мельче MinConfineIndentPt — обнуляется). Чистая — под тест.
        /// </summary>
        internal static bool ColumnConfineIndents(bool eligible, double blockLeft, double blockRight,
            double textLeft, double textRight, out double leftIndent, out double rightIndent)
        {
            leftIndent = 0;
            rightIndent = 0;
            if (!eligible || textRight <= textLeft)
                return false;
            double areaWidth = textRight - textLeft;
            double colWidth = blockRight - blockLeft;
            if (colWidth <= 0 || colWidth >= ColumnConfineFraction * areaWidth)
                return false;
            double li = blockLeft - textLeft, ri = textRight - blockRight;
            if (li > MinConfineIndentPt) leftIndent = li;
            if (ri > MinConfineIndentPt) rightIndent = ri;
            return leftIndent > 0 || rightIndent > 0;
        }

        /// <summary>
        /// Записать абзац в текущую позицию БЕЗ завершающего перевода строки (ядро — чтобы
        /// переиспользовать и в потоке текста, и в ячейках таблицы, DRY): выравнивание,
        /// красная строка и формат пословно (шрифт, кегль, начертание, над/подстрочный, цвет).
        /// В ячейке (inCell) выключка по ширине заменяется на левый край: короткий текст ячейки
        /// иначе Word растягивает уродливыми пробелами; центрирование (шапки) сохраняется.
        /// textLeftPt/textRightPt — рамка текстовой области страницы для конфайна колонки.
        /// </summary>
        private static void WriteParagraphInto(dynamic sel, dynamic doc, OcrParagraph paragraph, double firstLineIndent, bool inCell, int skipChars, double textLeftPt, double textRightPt, double spaceBeforePt = 0)
        {
            // Выравнивание из источника; центрированное — без красной строки.
            int align; double indent;
            switch (paragraph.Alignment)
            {
                case OcrAlignment.Center: align = WdAlignCenter; indent = 0; break;
                case OcrAlignment.Left: align = WdAlignLeft; indent = firstLineIndent; break;
                default: align = inCell ? WdAlignLeft : WdAlignJustify; indent = inCell ? 0 : firstLineIndent; break;
            }
            sel.ParagraphFormat.Alignment = align;
            sel.ParagraphFormat.FirstLineIndent = indent;
            // Интервал перед абзацем ставим ВСЕГДА (0 — обычный случай): Selection наследует
            // прямое форматирование предыдущего абзаца, и без сброса зазор «поехал» бы дальше.
            sel.ParagraphFormat.SpaceBefore = spaceBeforePt;

            // Конфайн центрированного абзаца в его колонку: адресат шапки уходит вправо, бланк —
            // влево, вместо ложного центра по всей странице. Отступы всегда переустанавливаем
            // (Selection несёт состояние от предыдущего абзаца), 0 — обычный полноширинный случай.
            double leftIndent, rightIndent;
            ColumnConfineIndents(paragraph.Alignment == OcrAlignment.Center && !inCell,
                paragraph.BlockLeftPt, paragraph.BlockRightPt, textLeftPt, textRightPt,
                out leftIndent, out rightIndent);
            sel.ParagraphFormat.LeftIndent = leftIndent;
            sel.ParagraphFormat.RightIndent = rightIndent;

            // Формат пословно (ран за раном): шрифт, кегль, полужирный, курсив, над/подстрочный, цвет.
            // skipChars — снять ведущий маркер списка, растянутый по первым ранам (Word рисует свой).
            int skip = skipChars;
            foreach (OcrRun run in paragraph.Runs)
            {
                string text = run.Text;
                if (skip > 0)
                {
                    if (skip >= text.Length) { skip -= text.Length; continue; } // весь ран — часть маркера
                    text = text.Substring(skip);
                    skip = 0;
                }
                sel.Font.Name = ResolveFont(run.FontName, text);
                sel.Font.Size = FontSize(run.FontSizePt);
                sel.Font.Bold = run.Bold ? 1 : 0;
                sel.Font.Italic = run.Italic ? 1 : 0;
                sel.Font.Superscript = run.Super ? 1 : 0;
                sel.Font.Subscript = run.Sub ? 1 : 0;
                sel.Font.Underline = run.Underline ? 1 : 0; // wdUnderlineSingle / None
                sel.Font.Color = ToBgr(run.ColorArgb);
                if (string.IsNullOrEmpty(run.Uri))
                {
                    sel.TypeText(text);
                }
                else
                {
                    int start = (int)sel.Range.End;
                    sel.TypeText(text);
                    try { doc.Hyperlinks.Add(doc.Range(start, (int)sel.Range.End), run.Uri); }
                    catch { } // не удалось оформить ссылку — текст всё равно на месте
                }
            }
        }

        /// <summary>
        /// Вставить восстановленную таблицу в текущую позицию: сетка Rows×ColumnCount с
        /// границами, ширинами колонок из геометрии линовки и форматированным текстом ячеек
        /// (тем же <see cref="WriteParagraphInto"/>, DRY). Объединение ячеек пока не переносится
        /// (накрытые позиции пишутся пустыми) — структура и текст верны в любом случае. После
        /// таблицы ставится абзац-разделитель: без него две смежные таблицы Word слил бы в одну.
        /// Сбой построения не срывает документ И не теряет текст: слова ячеек уже изъяты из
        /// потока страницы, поэтому недостроенная таблица удаляется, а содержимое выводится
        /// плоскими абзацами (см. <see cref="WriteTableFlat"/>).
        /// </summary>
        private static void WriteTable(dynamic word, dynamic doc, dynamic sel, OcrTable table)
        {
            int rows = table.Rows.Count, cols = table.ColumnCount;
            if (rows == 0 || cols == 0)
                return;
            dynamic wtable = null;
            try
            {
                wtable = doc.Tables.Add(sel.Range, rows, cols);
                wtable.AllowAutoFit = false;
                wtable.Borders.Enable = table.Borderless ? 0 : 1; // сетка без линовки — без границ

                for (int c = 0; c < cols; c++)
                {
                    try { wtable.Columns[c + 1].Width = ColWidth(table.ColumnWidthsPt[c]); }
                    catch { } // ширина косметическая; сбой одной колонки не критичен
                }

                // Сначала заполняем (все ячейки на месте — прямая адресация r,c), потом сливаем.
                for (int r = 0; r < rows; r++)
                {
                    OcrTableRow row = table.Rows[r];
                    for (int c = 0; c < cols; c++)
                    {
                        OcrTableCell cell = row.Cells[c];
                        if (cell.Covered || cell.Paragraphs == null || cell.Paragraphs.Count == 0)
                            continue;
                        WriteCell(word, doc, wtable, r + 1, c + 1, cell, row.SpaceAfterPt); // Word адресует ячейки с 1
                    }
                }

                MergeSpans(wtable, table, rows, cols);

                // Курсор — за таблицу, отделить абзацем (иначе следующая таблица сольётся с этой).
                sel.Start = wtable.Range.End;
                sel.Collapse(WdCollapseEnd);
                sel.TypeParagraph();
            }
            catch
            {
                // Частично построенную таблицу убираем (иначе плоский вывод задвоил бы уже
                // записанные ячейки); если и удаление не удалось — выводим как есть, это
                // сбой внутри сбоя, хуже прежнего поведения не станет.
                try { if (wtable != null) wtable.Delete(); } catch { }
                try { WriteTableFlat(sel, doc, table); } catch { }
            }
        }

        /// <summary>
        /// Аварийный вывод таблицы плоскими абзацами (по строкам, ячейки слева направо):
        /// Word не построил таблицу, а слова её ячеек уже изъяты из потока страницы —
        /// молча потерять их нельзя. Форматирование ранов сохраняется, сетка — нет.
        /// </summary>
        private static void WriteTableFlat(dynamic sel, dynamic doc, OcrTable table)
        {
            foreach (OcrTableRow row in table.Rows)
                foreach (OcrTableCell cell in row.Cells)
                {
                    if (cell.Covered || cell.Paragraphs == null)
                        continue;
                    foreach (OcrParagraph p in cell.Paragraphs)
                    {
                        WriteParagraphInto(sel, doc, p, 0, true, 0, 0, 0); // как в ячейке: без выключки и конфайна
                        sel.TypeParagraph();
                    }
                }
        }

        /// <summary>
        /// Объединить ячейки по ColSpan/RowSpan уже заполненной таблицы. Идём с конца (снизу
        /// вверх, справа налево): слияние блока перенумеровывает ячейки НИЖЕ и ПРАВЕЕ, а они уже
        /// обработаны, поэтому адреса ещё не слитых блоков (выше/левее) не сбиваются. Пустых
        /// абзацев слияние не плодит: накрытые ячейки не заполняются, а Merge с ПУСТОЙ ячейкой
        /// содержимого не добавляет (проверено живым Word: HEAD + пустая → один абзац «HEAD»).
        /// </summary>
        private static void MergeSpans(dynamic wtable, OcrTable table, int rows, int cols)
        {
            for (int r = rows - 1; r >= 0; r--)
            {
                for (int c = cols - 1; c >= 0; c--)
                {
                    OcrTableCell cell = table.Rows[r].Cells[c];
                    if (cell.Covered || (cell.ColSpan <= 1 && cell.RowSpan <= 1))
                        continue;
                    if (r + cell.RowSpan > rows || c + cell.ColSpan > cols)
                        continue; // спан за пределами сетки — не наш инвариант, но защищаемся здесь
                    try
                    {
                        dynamic head = wtable.Cell(r + 1, c + 1);
                        head.Merge(wtable.Cell(r + cell.RowSpan, c + cell.ColSpan));
                    }
                    catch { } // одно неудавшееся объединение не роняет таблицу
                }
            }
        }

        /// <summary>
        /// Записать абзацы ячейки в её начало (не трогая маркер конца ячейки). spaceAfterPt —
        /// доп. интервал после последнего абзаца ячейки: у безлиновочной сетки так возвращается
        /// пустой промежуток между группами полей (см. GridDetector); 0 — обычный случай.
        /// </summary>
        private static void WriteCell(dynamic word, dynamic doc, dynamic wtable, int row, int col, OcrTableCell cell, double spaceAfterPt)
        {
            dynamic cellRange = wtable.Cell(row, col).Range;
            cellRange.Collapse(WdCollapseStart); // в начало ячейки, чтобы не съесть маркер ячейки
            cellRange.Select();
            dynamic sel = word.Selection;
            for (int i = 0; i < cell.Paragraphs.Count; i++)
            {
                if (i > 0)
                    sel.TypeParagraph();
                WriteParagraphInto(sel, doc, cell.Paragraphs[i], 0, true, 0, 0, 0); // в ячейке: без красной строки, без выключки, без списков, без конфайна
            }
            if (spaceAfterPt > 0)
                try { sel.ParagraphFormat.SpaceAfter = spaceAfterPt; } catch { } // интервал группы после строки-сетки
        }

        /// <summary>Ширина колонки в pt с нижней защитой (вырожденную колонку Word рисует криво).</summary>
        private static double ColWidth(double pt)
        {
            return pt < MinColWidthPt ? MinColWidthPt : pt;
        }

        /// <summary>
        /// Вставляет изображение inline в текущую позицию и переводит строку; размер — по рамке
        /// PDF (pt), с защитой пределов. Горизонтально центрированное на странице изображение (герб,
        /// логотип) выводится по центру, как в оригинале, иначе — по левому краю. Сбой одной картинки
        /// не срывает документ. PNG кладётся во временный файл (встраивается в .docx при вставке),
        /// временная папка чистится в Write.
        /// </summary>
        private static void InsertImage(dynamic sel, OcrImage img, double pageWidthPt, string tempDir, ref int index, double spaceBeforePt = 0)
        {
            if (InsertImageCore(sel, img, tempDir, ref index, IsImageCentered(img.LeftPt, img.WidthPt, pageWidthPt), 0, 0, spaceBeforePt))
                try { sel.TypeParagraph(); } catch { } // изображение на своей строке
        }

        /// <summary>
        /// Ядро вставки inline-картинки в текущую позицию (БЕЗ завершающего перевода строки — им
        /// управляет вызывающий: поток текста ставит абзац, ячейка полосы — сама). centered —
        /// выравнивание абзаца картинки (в ячейке герб/логотип центрируется). Возвращает true, если
        /// картинка вставлена. Сбой одной картинки не срывает документ. DRY: общее ядро для потока и ячеек.
        /// </summary>
        private static bool InsertImageCore(dynamic sel, OcrImage img, string tempDir, ref int index, bool centered,
            double leftIndent = 0, double rightIndent = 0, double spaceBeforePt = 0)
        {
            if (img == null || img.Png == null || img.Png.Length == 0)
                return false;
            string file = Path.Combine(tempDir, "img_" + index + ".png");
            index++;
            try
            {
                File.WriteAllBytes(file, img.Png);
                sel.ParagraphFormat.Alignment = centered ? WdAlignCenter : WdAlignLeft;
                sel.ParagraphFormat.FirstLineIndent = 0;
                sel.ParagraphFormat.LeftIndent = leftIndent;   // конфайн в колонку (герб над бланком), 0 — обычно
                sel.ParagraphFormat.RightIndent = rightIndent;
                sel.ParagraphFormat.SpaceBefore = spaceBeforePt; // сброс наследования (0 — обычный случай)
                dynamic shape = sel.InlineShapes.AddPicture(file, false, true); // встроить в документ
                shape.LockAspectRatio = 0; // msoFalse — задаём оба размера
                shape.Width = ClampSize(img.WidthPt);
                shape.Height = ClampSize(img.HeightPt);
                return true;
            }
            catch { return false; } // одна картинка не должна сорвать конвертацию
        }

        /// <summary>
        /// Вывести side-by-side полосу колонок безграничной таблицей 1×N: колонки сидят рядом (как
        /// двухколоночная шапка — бланк слева, адресат справа), герб/логотип центрируется в своей
        /// ячейке над бланком. Ширины ячеек — по границам колонок (середина зазора между соседними),
        /// от полей текстовой области. В ячейке абзацы центрируются в её ширине (DRY: тот же
        /// <see cref="WriteParagraphInto"/> с inCell), картинки — <see cref="InsertImageCore"/>.
        /// Сбой построения таблицы не срывает документ: колонки выводятся последовательно (фолбэк).
        /// </summary>
        private static void WriteColumnBand(dynamic word, dynamic doc, dynamic sel, PageItem band,
            double textLeftPt, double textRightPt, double pageWidthPt, string tempDir, ref int index,
            double spaceBeforePt = 0, double typicalGapPt = 0)
        {
            int n = band.Columns.Count;

            // Ведущая картинка колонки, стоящая ВЫШЕ всего текста полосы (герб над бланком), выносится
            // НАД таблицей и центрируется над своей колонкой — тогда строки колонок в таблице
            // выравниваются по верху текста (адресат встаёт вровень с бланком, а не с гербом сверху).
            double textTop = double.MinValue;
            foreach (List<Block> col0 in band.Columns)
                foreach (Block bb in col0)
                    if (bb.Paragraph != null && bb.Top > textTop)
                        textTop = bb.Top;
            double pendingSpace = spaceBeforePt; // интервал полосы несёт её первый вывод (картинка или прокладка)
            for (int c = 0; c < n; c++)
            {
                List<Block> col = band.Columns[c];
                while (col.Count > 0 && col[0].Image != null && col[0].Bottom >= textTop)
                {
                    double li, ri;
                    ColumnConfineIndents(true, band.ColLeft[c], band.ColRight[c], textLeftPt, textRightPt, out li, out ri);
                    if (InsertImageCore(sel, col[0].Image, tempDir, ref index, true, li, ri, pendingSpace))
                    {
                        pendingSpace = 0;
                        try { sel.TypeParagraph(); } catch { }
                    }
                    col.RemoveAt(0);
                }
            }
            if (pendingSpace > 0)
                InsertSpacer(sel, pendingSpace); // полоса — таблица, SpaceBefore ей не задать

            double[] widths = BandColumnWidths(band, textLeftPt, textRightPt);
            dynamic wtable = null;
            try
            {
                wtable = doc.Tables.Add(sel.Range, 1, n);
                wtable.AllowAutoFit = false;
                wtable.Borders.Enable = 0; // полоса без видимых границ
                for (int c = 0; c < n; c++)
                {
                    try { wtable.Columns[c + 1].Width = ColWidth(widths[c]); }
                    catch { }
                }
                for (int c = 0; c < n; c++)
                {
                    dynamic cellRange = wtable.Cell(1, c + 1).Range;
                    cellRange.Collapse(WdCollapseStart);
                    cellRange.Select();
                    dynamic cellSel = word.Selection;
                    List<Block> col = band.Columns[c];
                    for (int i = 0; i < col.Count; i++)
                    {
                        if (i > 0)
                            cellSel.TypeParagraph(); // каждый блок колонки — своим абзацем
                        Block b = col[i];
                        // Пустой промежуток исходника внутри колонки («(по списку)» ниже адресата)
                        // возвращаем интервалом перед блоком — той же формулой, что и в потоке.
                        double extra = i == 0 ? 0
                            : ExtraGapPt(Math.Min(col[i - 1].Bottom, col[i - 1].Top) - b.Top, typicalGapPt);
                        if (b.Paragraph != null)
                            WriteParagraphInto(cellSel, doc, b.Paragraph, 0, true, 0, 0, 0, extra); // центрируется в ячейке
                        else if (b.Image != null)
                            InsertImageCore(cellSel, b.Image, tempDir, ref index, true, 0, 0, extra); // герб по центру ячейки
                    }
                }
                sel.Start = wtable.Range.End;
                sel.Collapse(WdCollapseEnd);
                sel.TypeParagraph(); // отделить полосу от следующего блока
            }
            catch
            {
                // Таблицу построить не удалось — недостроенную убираем (иначе плоский вывод
                // задвоил бы уже записанные ячейки) и выводим колонки просто по очереди.
                try { if (wtable != null) wtable.Delete(); } catch { }
                foreach (List<Block> col in band.Columns)
                    foreach (Block b in col)
                    {
                        try
                        {
                            if (b.Paragraph != null) { WriteParagraphInto(sel, doc, b.Paragraph, 0, false, 0, textLeftPt, textRightPt); sel.TypeParagraph(); }
                            else if (b.Image != null) InsertImage(sel, b.Image, pageWidthPt, tempDir, ref index);
                        }
                        catch { }
                    }
            }
        }

        /// <summary>
        /// Ширины ячеек полосы (pt): граница между соседними колонками — середина зазора между их
        /// рамками, крайние — по полям текстовой области. Так центрирование в ячейке совпадает с
        /// исходным центром колонки. Вырожденные поля — фолбэк на рамки самих колонок. Чистая — под тест.
        /// </summary>
        internal static double[] BandColumnWidths(PageItem band, double textLeftPt, double textRightPt)
        {
            int n = band.Columns.Count;
            double left = textLeftPt, right = textRightPt;
            if (right <= left)
            {
                left = band.ColLeft[0];
                right = band.ColRight[n - 1];
                if (right <= left) right = left + n * MinColWidthPt;
            }
            var bound = new double[n + 1];
            bound[0] = left;
            bound[n] = right;
            for (int c = 1; c < n; c++)
                bound[c] = (band.ColRight[c - 1] + band.ColLeft[c]) / 2;
            var widths = new double[n];
            for (int c = 0; c < n; c++)
            {
                double w = bound[c + 1] - bound[c];
                widths[c] = w > MinColWidthPt ? w : MinColWidthPt;
            }
            return widths;
        }

        // Изображение считаем центрированным, если зазоры до краёв страницы заметны (> этой доли
        // ширины) и почти равны (разница <= этой доли) — герб/логотип сверху по центру. Иначе левый
        // край: врезки у поля (leftGap ≈ 0) и печати сбоку не центрируются.
        private const double ImageCenterMinGapFraction = 0.05;
        private const double ImageCenterBalanceFraction = 0.06;

        /// <summary>
        /// Горизонтально ли центрировано изображение на странице: оба зазора до краёв заметны и
        /// почти равны. leftPt — левый край рамки (pt), widthPt — ширина, pageWidthPt — ширина
        /// страницы. Вырожденные размеры → не центрируем. Чистая — под тест.
        /// </summary>
        internal static bool IsImageCentered(double leftPt, double widthPt, double pageWidthPt)
        {
            if (pageWidthPt <= 0 || widthPt <= 0 || widthPt >= pageWidthPt)
                return false;
            double leftGap = leftPt;
            double rightGap = pageWidthPt - (leftPt + widthPt);
            if (leftGap < 0 || rightGap < 0)
                return false;
            double minGap = ImageCenterMinGapFraction * pageWidthPt;
            double balance = ImageCenterBalanceFraction * pageWidthPt;
            return leftGap > minGap && rightGap > minGap && Math.Abs(leftGap - rightGap) <= balance;
        }

        private static double ClampSize(double pt)
        {
            return pt < 1 ? 1 : (pt > MaxPagePt ? MaxPagePt : pt);
        }

        /// <summary>
        /// Имя шрифта для рана: если шрифт источника установлен в системе — оставляем его, иначе
        /// подставляем установленный по умолчанию. КЛЮЧЕВОЕ: когда Word получает НЕустановленный
        /// шрифт, он уводит кириллицу в восточноазиатский фолбэк-слот (rFonts hint="eastAsia"),
        /// и при выключке по ширине раздвигает буквы по правилам CJK — получается «р а з р я д к а»
        /// (а латиница остаётся слитной). Установленный шрифт держит кириллицу в hAnsi — обычная
        /// выключка. Но и УСТАНОВЛЕННОГО мало: «не родные» для Word семейства (Liberation Serif
        /// и т.п.) получают тот же hint="eastAsia" при вводе кириллицы — поэтому кириллический
        /// текст пишется только шрифтами из сейф-листа Word-родных семейств, прочие уводятся в
        /// fallback (Liberation Serif — метрический клон Times New Roman, подмена нейтральна).
        /// Список шрифтов читается один раз.
        /// </summary>
        private static string ResolveFont(string requested, string text)
        {
            return ResolveFontName(requested, text, InstalledFonts, DefaultFontName);
        }

        // Word-родные семейства с полной кириллицей, за которыми Word не замечен в eastAsia-хинте.
        private static readonly HashSet<string> CyrillicSafeFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Times New Roman", "Arial", "Calibri", "Calibri Light", "Courier New",
            "Cambria", "Georgia", "Verdana", "Tahoma", "Segoe UI", "Consolas"
        };

        /// <summary>Чистая логика подстановки (под тест): установленный — оставить; кириллице — только сейф-лист.</summary>
        internal static string ResolveFontName(string requested, string text, ICollection<string> installed, string fallback)
        {
            if (string.IsNullOrEmpty(requested))
                return fallback;
            if (installed == null || !installed.Contains(requested))
                return fallback;
            return HasCyrillic(text) && !CyrillicSafeFonts.Contains(requested) ? fallback : requested;
        }

        /// <summary>Есть ли в тексте кириллица (U+0400–U+04FF).</summary>
        internal static bool HasCyrillic(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;
            for (int i = 0; i < text.Length; i++)
                if (text[i] >= 'Ѐ' && text[i] <= 'ӿ')
                    return true;
            return false;
        }

        private static readonly HashSet<string> InstalledFonts = LoadInstalledFonts();

        /// <summary>Семейства установленных шрифтов (без учёта регистра). Сбой чтения — пустой набор (всё уйдёт в fallback).</summary>
        private static HashSet<string> LoadInstalledFonts()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var col = new System.Drawing.Text.InstalledFontCollection())
                    foreach (System.Drawing.FontFamily fam in col.Families)
                        set.Add(fam.Name);
            }
            catch { }
            return set;
        }

        /// <summary>Кегль рана в допустимых пределах; иначе — по умолчанию.</summary>
        private static double FontSize(double sizePt)
        {
            return sizePt >= MinFontSize && sizePt <= MaxFontSize ? sizePt : DefaultFontSize;
        }

        /// <summary>
        /// Размер и поля ТЕКУЩЕГО раздела из своей страницы источника (у каждой PDF-страницы —
        /// свой раздел, поэтому книжные и альбомные страницы уживаются в одном .docx, а широкая
        /// таблица не обрезается). Размер вне разумных пределов — оставляем шаблон Word. Поля
        /// косметические: сбой PageSetup не срывает конвертацию.
        /// </summary>
        private static void ApplySectionSetup(dynamic sel, PdfPageText page)
        {
            double pw = page.WidthPt, ph = page.HeightPt;
            if (pw < MinPagePt || pw > MaxPagePt || ph < MinPagePt || ph > MaxPagePt)
                return;
            try
            {
                dynamic ps = sel.PageSetup;
                // Явные размеры задают и ориентацию (ширина > высоты — альбомная); поля из рамок текста.
                ps.PageWidth = pw;
                ps.PageHeight = ph;
                ps.LeftMargin = ClampMargin(page.LeftMarginPt, pw);
                ps.RightMargin = ClampMargin(page.RightMarginPt, pw);
                ps.TopMargin = ClampMargin(page.TopMarginPt, ph);
                ps.BottomMargin = ClampMargin(page.BottomMarginPt, ph);
            }
            catch { } // поля — косметика; сбой PageSetup не должен срывать сохранение
        }

        private static double ClampMargin(double m, double pageDim)
        {
            double max = 0.45 * pageDim;
            return m < 0 ? 0 : (m > max ? max : m);
        }

        /// <summary>0xRRGGBB → WdColor (BGR-порядок), как ожидает Word.Font.Color.</summary>
        private static int ToBgr(int argb)
        {
            int r = (argb >> 16) & 0xFF, g = (argb >> 8) & 0xFF, b = argb & 0xFF;
            return r | (g << 8) | (b << 16);
        }

        /// <summary>
        /// Единый отступ красной строки документа: медиана положительных постраничных
        /// отступов (обычно одинаковы). 0 — если ни одна страница не была с отступами.
        /// </summary>
        private static double DocumentIndent(IList<PdfPageText> pages)
        {
            var vals = new List<double>();
            foreach (PdfPageText page in pages)
                if (page.FirstLineIndentPt > 0)
                    vals.Add(page.FirstLineIndentPt);
            return MathUtil.Median(vals);
        }
    }
}
