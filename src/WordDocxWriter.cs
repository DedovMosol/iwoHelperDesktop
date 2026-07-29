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
        private const double MinPagePt = 72;    // 1"; разумные пределы размера страницы
        private const double MaxPagePt = 1584;  // 22" — максимум Word

        /// <summary>
        /// Пишет .docx из абзацев и изображений страниц. Занятый файл/нет Word — MergeException.
        /// cancelled — кооперативная отмена между страницами: OperationCanceledException; так как
        /// сохранение идёт лишь ПОСЛЕ наполнения, при отмене файл не создаётся (Word закрывается
        /// без сохранения в finally общего каркаса). onCommitting — вызывается один раз после
        /// наполнения и перед сохранением (точка невозврата: дальше отмена уже не сработает).
        /// </summary>
        public static void Write(IList<PdfPageText> pages, string path, Action<int, int> progress = null,
            Func<bool> cancelled = null, Action onCommitting = null)
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
                        Cancellation.ThrowIf(cancelled); // отмена между страницами; файл ещё не сохранён
                        if (p > 0)
                            sel.InsertBreak(WdSectionBreakNextPage); // новый раздел = свой размер листа
                        ApplySectionSetup(sel, pages[p]); // размер и поля страницы из источника
                        // Текстовая область страницы (pt): в неё Word укладывает абзацы, от неё
                        // считаются отступы конфайна центрированной колонки (см. WriteParagraphInto).
                        double textLeft = pages[p].LeftMarginPt;
                        double textRight = pages[p].WidthPt - pages[p].RightMarginPt;
                        List<PageBlocks.PageItem> items = PageBlocks.CoalesceRowBands(PageBlocks.OrderedItems(pages[p]));
                        double typicalGap = PageBlocks.TypicalItemGap(items);
                        PageBlocks.PageItem prev = null;
                        foreach (PageBlocks.PageItem item in items)
                        {
                            // Лишний вертикальный зазор до предыдущего блока → интервал перед этим.
                            double spaceBefore = prev == null ? 0 : PageBlocks.ExtraGapPt(prev.Bottom - item.Top, typicalGap);
                            prev = item;
                            if (item.IsBand)
                            {
                                ClearInheritedList(sel, listState); // таблица обрывает список
                                WriteColumnBand(word, doc, sel, item, textLeft, textRight, pages[p].WidthPt,
                                    tempDir, ref imgIndex, spaceBefore, typicalGap);
                                continue;
                            }
                            PageBlocks.Block blk = item.Single;
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
                    // Наполнение завершено, последний Cancellation.ThrowIf позади: дальше WordCom
                    // сохраняет .docx (Ghostscript-подобная точка невозврата) — снять кнопку отмены.
                    if (onCommitting != null)
                        onCommitting();
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


        private static void WriteParagraph(dynamic sel, dynamic doc, OcrParagraph paragraph, double firstLineIndent, ListTemplates lists, ListState state, double textLeftPt, double textRightPt, double spaceBeforePt = 0)
        {
            bool asList = lists.Available && paragraph.ListKind != ListKind.None;
            if (asList)
            {
                // Маркер («1.», «•») снимаем — Word рисует свой; отступ задаёт шаблон списка
                // (indent=0, и привязка по факту пункту не нужна — иначе сдвиг задвоился бы).
                WriteParagraphInto(sel, doc, paragraph, 0, false, paragraph.ListContentStart, textLeftPt, textRightPt, spaceBeforePt, false);
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
                continuePrev = state.LastNumber > 0 && p.ListNumber == state.LastNumber + 1;
            else
                continuePrev = state.PrevKind == ListKind.Bulleted;

            // Пункт СРАЗУ за пунктом того же вида уже унаследовал список от TypeParagraph —
            // переприменение шаблона не нужно: оно и лишние COM-вызовы, и сбивает StartAt
            // рестартованного не с «1.» списка. Шаблон применяется только на старте списка и
            // при возврате к нему после вложенного обычного текста.
            if (continuePrev && state.PrevKind == p.ListKind)
            {
                if (p.ListKind == ListKind.Numbered)
                    state.LastNumber = p.ListNumber;
                return;
            }

            dynamic tmpl = p.ListKind == ListKind.Numbered ? lists.Number : lists.Bullet;
            bool applied = false;
            try
            {
                sel.Range.ListFormat.ApplyListTemplateWithLevel(tmpl, continuePrev, WdListApplyToWholeList, WdWord10ListBehavior, 1);
                // Список, начинающийся НЕ с «1.» («5. 6. 7.» — продолжение из другого
                // документа): Word при рестарте всегда нумерует с единицы, поэтому стартовый
                // номер задаётся явно. Правится шаблон ПРИМЕНЁННОГО списка (его копия в
                // документе), а не галерея пользователя.
                if (!continuePrev && p.ListKind == ListKind.Numbered && p.ListNumber > 1)
                    try { sel.Range.ListFormat.ListTemplate.ListLevels.Item(1).StartAt = p.ListNumber; }
                    catch { } // не удалось задать старт — список остаётся с «1.», текст цел
                applied = true;
            }
            catch { }
            if (applied)
            {
                if (p.ListKind == ListKind.Numbered)
                    state.LastNumber = p.ListNumber;
                state.PrevKind = p.ListKind;
                return;
            }
            // Шаблон не применился, а маркер уже снят с текста — вернуть «1.»/«•» в начало
            // абзаца, иначе пункт молча теряет номер. Состояние списка не трогаем: этот абзац
            // остался обычным текстом.
            try { sel.Paragraphs.Item(1).Range.InsertBefore(p.Text.Substring(0, p.ListContentStart)); }
            catch { } // и вернуть не вышло — хуже прежнего (маркер терялся молча) не стало
            state.PrevKind = ListKind.None;
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
        // текстовой области — это правая колонка шапки или левый блок. Только для
        // центрированных: у левых/выключенных горизонталь уже задана красной строкой и полем,
        // а центрированный на всю страницу титул при конфайне в свою рамку центрируется там же.
        private const double ColumnConfineFraction = 0.72;
        private const double MinConfineIndentPt = 6; // отступ мельче — колонка практически во всю ширину

        // Горизонтальная привязка НЕцентрированного абзаца по ФАКТУ его первой строки:
        // глубже четверти области — это позиция узкой колонки (боковая метка справа) → левый
        // отступ абзаца; мельче — вопрос красной строки: документный отступ ставится, только
        // если первая строка РЕАЛЬНО отступала (сноски и контакты стоят с края — без
        // отступа), а в документе без общего отступа берётся фактический отступ абзаца.
        private const double AnchorLeftMinFraction = 0.25;
        private const double MinFactIndentPt = 6; // мельче — шум измерения, не отступ

        /// <summary>
        /// Привязка по факту (см. константы выше): factInset — отступ первой строки абзаца от
        /// левого края текстовой области (pt), documentIndent — красная строка документа (0 —
        /// нет). Возвращает красную строку и левый отступ абзаца. Чистая — под тест.
        /// </summary>
        internal static void AnchorIndents(double factInset, double areaWidth, double documentIndent,
            out double firstLineIndent, out double leftIndent)
        {
            firstLineIndent = 0;
            leftIndent = 0;
            if (areaWidth <= 0)
                return;
            if (factInset >= AnchorLeftMinFraction * areaWidth)
            {
                leftIndent = factInset; // столбик/метка на своей позиции
                return;
            }
            if (documentIndent > 0)
            {
                if (factInset >= 0.5 * documentIndent)
                    firstLineIndent = documentIndent;
                return;
            }
            if (factInset >= MinFactIndentPt)
                firstLineIndent = factInset; // документ без общего отступа — по факту абзаца
        }

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
        private static void WriteParagraphInto(dynamic sel, dynamic doc, OcrParagraph paragraph, double firstLineIndent, bool inCell, int skipChars, double textLeftPt, double textRightPt, double spaceBeforePt = 0, bool anchor = true)
        {
            // Выравнивание из источника; центрированное — без красной строки.
            int align; double indent;
            switch (paragraph.Alignment)
            {
                case OcrAlignment.Center: align = WdAlignCenter; indent = 0; break;
                case OcrAlignment.Left: align = WdAlignLeft; indent = firstLineIndent; break;
                default: align = inCell ? WdAlignLeft : WdAlignJustify; indent = inCell ? 0 : firstLineIndent; break;
            }

            // Конфайн центрированного абзаца в его колонку: правый блок шапки уходит вправо, левый —
            // влево, вместо ложного центра по всей странице. Отступы всегда переустанавливаем
            // (Selection несёт состояние от предыдущего абзаца), 0 — обычный полноширинный случай.
            double leftIndent, rightIndent;
            ColumnConfineIndents(paragraph.Alignment == OcrAlignment.Center && !inCell,
                paragraph.BlockLeftPt, paragraph.BlockRightPt, textLeftPt, textRightPt,
                out leftIndent, out rightIndent);

            // Привязка НЕцентрированного абзаца по факту первой строки: красная строка — только
            // если она реально отступала, глубокий старт — позиция колонки (боковая метка). В ячейке
            // (textLeft/Right нулевые) и у пунктов списка (anchor=false) не применяется.
            if (anchor && !inCell && paragraph.Alignment != OcrAlignment.Center)
            {
                double anchorFli, anchorLeft;
                AnchorIndents(paragraph.LeftPt - textLeftPt, textRightPt - textLeftPt, indent,
                    out anchorFli, out anchorLeft);
                if (textRightPt > textLeftPt)
                {
                    indent = anchorFli;
                    leftIndent = anchorLeft;
                }
            }

            sel.ParagraphFormat.Alignment = align;
            sel.ParagraphFormat.FirstLineIndent = indent;
            // Интервал перед абзацем ставим ВСЕГДА (0 — обычный случай): Selection наследует
            // прямое форматирование предыдущего абзаца, и без сброса зазор «поехал» бы дальше.
            sel.ParagraphFormat.SpaceBefore = spaceBeforePt;
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
                sel.Font.Name = FontResolver.Resolve(run.FontName, text);
                sel.Font.Size = FontResolver.ClampSizePt(run.FontSizePt);
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
        /// PDF (pt), с защитой пределов. Горизонтально центрированное на странице изображение
        /// (логотип) выводится по центру, как в оригинале, иначе — по левому краю. Сбой одной картинки
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
        /// выравнивание абзаца картинки (в ячейке логотип центрируется). Возвращает true, если
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
                sel.ParagraphFormat.LeftIndent = leftIndent;   // конфайн в колонку (логотип над шапкой), 0 — обычно
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
        /// двухколоночная шапка — левый блок слева, правый справа), логотип центрируется в своей
        /// ячейке над шапкой. Ширины ячеек — по границам колонок (середина зазора между соседними),
        /// от полей текстовой области. В ячейке абзацы центрируются в её ширине (DRY: тот же
        /// <see cref="WriteParagraphInto"/> с inCell), картинки — <see cref="InsertImageCore"/>.
        /// Сбой построения таблицы не срывает документ: колонки выводятся последовательно (фолбэк).
        /// </summary>
        private static void WriteColumnBand(dynamic word, dynamic doc, dynamic sel, PageBlocks.PageItem band,
            double textLeftPt, double textRightPt, double pageWidthPt, string tempDir, ref int index,
            double spaceBeforePt = 0, double typicalGapPt = 0)
        {
            int n = band.Columns.Count;

            // Ведущая картинка колонки, стоящая ВЫШЕ всего текста полосы (логотип над шапкой), выносится
            // НАД таблицей и центрируется над своей колонкой — тогда строки колонок в таблице
            // выравниваются по верху текста (правый блок встаёт вровень с левым, а не с логотипом сверху).
            double textTop = double.MinValue;
            foreach (List<PageBlocks.Block> col0 in band.Columns)
                foreach (PageBlocks.Block bb in col0)
                    if (bb.Paragraph != null && bb.Top > textTop)
                        textTop = bb.Top;
            double pendingSpace = spaceBeforePt; // интервал полосы несёт её первый вывод (картинка или прокладка)
            for (int c = 0; c < n; c++)
            {
                List<PageBlocks.Block> col = band.Columns[c];
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
                    List<PageBlocks.Block> col = band.Columns[c];
                    // Колонка, начинающаяся НИЖЕ верха текстов полосы (Ф.И.О. напротив последней
                    // строки многострочной подписи), опускается интервалом перед первым блоком —
                    // ячейки выравниваются по верху, и без этого она всплывала на первую строку.
                    double colTop = double.MinValue;
                    foreach (PageBlocks.Block b in col)
                        if (b.Top > colTop) colTop = b.Top;
                    // Порог и кап — те же, что у межблочных зазоров (типичный зазор здесь ноль).
                    double topInset = PageBlocks.ExtraGapPt(textTop - colTop, 0);
                    for (int i = 0; i < col.Count; i++)
                    {
                        if (i > 0)
                            cellSel.TypeParagraph(); // каждый блок колонки — своим абзацем
                        PageBlocks.Block b = col[i];
                        // Пустой промежуток исходника внутри колонки (пометка ниже верхнего блока)
                        // возвращаем интервалом перед блоком — той же формулой, что и в потоке.
                        double extra = i == 0 ? topInset
                            : PageBlocks.ExtraGapPt(Math.Min(col[i - 1].Bottom, col[i - 1].Top) - b.Top, typicalGapPt);
                        if (b.Paragraph != null)
                            WriteParagraphInto(cellSel, doc, b.Paragraph, 0, true, 0, 0, 0, extra); // центрируется в ячейке
                        else if (b.Image != null)
                            InsertImageCore(cellSel, b.Image, tempDir, ref index, true, 0, 0, extra); // логотип по центру ячейки
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
                foreach (List<PageBlocks.Block> col in band.Columns)
                    foreach (PageBlocks.Block b in col)
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
        internal static double[] BandColumnWidths(PageBlocks.PageItem band, double textLeftPt, double textRightPt)
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
        // ширины) и почти равны (разница <= этой доли) — логотип сверху по центру. Иначе левый
        // край: врезки у поля (leftGap ≈ 0) и штампы сбоку не центрируются.
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
