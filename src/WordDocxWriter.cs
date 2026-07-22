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
                WordCom.WriteDocx(path, "Файл Word", delegate(object wordObj, object docObj)
                {
                    dynamic word = wordObj;
                    dynamic doc = docObj;
                    dynamic sel = word.Selection;

                    for (int p = 0; p < pages.Count; p++)
                    {
                        if (p > 0)
                            sel.InsertBreak(WdSectionBreakNextPage); // новый раздел = свой размер листа
                        ApplySectionSetup(sel, pages[p]); // размер и поля страницы из источника
                        foreach (Block blk in OrderedBlocks(pages[p]))
                        {
                            if (blk.Paragraph != null)
                                WriteParagraph(sel, doc, blk.Paragraph, firstLineIndent);
                            else if (blk.Table != null)
                                WriteTable(word, doc, sel, blk.Table);
                            else
                                InsertImage(sel, blk.Image, tempDir, ref imgIndex);
                        }
                        if (progress != null)
                            progress(p + 1, pages.Count);
                    }
                });
            }
            finally
            {
                try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
            }
        }

        private const double RowBandPt = 12; // блоки ближе этого по вертикали — одна «строка-полоса»

        /// <summary>Блок содержимого страницы: абзац, таблица или изображение (одно из полей задано).</summary>
        private sealed class Block
        {
            public OcrParagraph Paragraph;
            public OcrTable Table;
            public OcrImage Image;
            public double Top;  // верх блока (Y, ось вверх) — основной порядок чтения
            public double Left; // левый край — вторичный порядок для блоков в одной строке-полосе
        }

        /// <summary>Абзацы, таблицы и изображения страницы в порядке чтения (сверху вниз, затем слева направо).</summary>
        private static List<Block> OrderedBlocks(PdfPageText page)
        {
            var blocks = new List<Block>();
            if (page.Paragraphs != null)
                foreach (OcrParagraph par in page.Paragraphs)
                    blocks.Add(new Block { Paragraph = par, Top = par.TopPt, Left = par.LeftPt });
            if (page.Tables != null)
                foreach (OcrTable table in page.Tables)
                    blocks.Add(new Block { Table = table, Top = table.TopPt, Left = table.LeftPt });
            if (page.Images != null)
                foreach (OcrImage img in page.Images)
                    blocks.Add(new Block { Image = img, Top = img.TopPt, Left = img.LeftPt });
            blocks.Sort(delegate(Block a, Block b) { return CompareReadingOrder(a.Top, a.Left, b.Top, b.Left); });
            return blocks;
        }

        /// <summary>
        /// Порядок чтения: сверху вниз (больше Top — выше — раньше); блоки в одной строке-полосе
        /// (|разница Top| ≤ RowBandPt, напр. таблицы бок о бок) — слева направо. Чистая — под тест.
        /// </summary>
        internal static int CompareReadingOrder(double aTop, double aLeft, double bTop, double bLeft)
        {
            if (Math.Abs(aTop - bTop) > RowBandPt)
                return bTop.CompareTo(aTop); // ось Y вверх: больший Top — выше — раньше
            return aLeft.CompareTo(bLeft);   // одна полоса — левее раньше
        }

        private static void WriteParagraph(dynamic sel, dynamic doc, OcrParagraph paragraph, double firstLineIndent)
        {
            WriteParagraphInto(sel, doc, paragraph, firstLineIndent, false);
            sel.TypeParagraph();
        }

        /// <summary>
        /// Записать абзац в текущую позицию БЕЗ завершающего перевода строки (ядро — чтобы
        /// переиспользовать и в потоке текста, и в ячейках таблицы, DRY): выравнивание,
        /// красная строка и формат пословно (шрифт, кегль, начертание, над/подстрочный, цвет).
        /// В ячейке (inCell) выключка по ширине заменяется на левый край: короткий текст ячейки
        /// иначе Word растягивает уродливыми пробелами; центрирование (шапки) сохраняется.
        /// </summary>
        private static void WriteParagraphInto(dynamic sel, dynamic doc, OcrParagraph paragraph, double firstLineIndent, bool inCell)
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

            // Формат пословно (ран за раном): шрифт, кегль, полужирный, курсив, над/подстрочный, цвет.
            foreach (OcrRun run in paragraph.Runs)
            {
                sel.Font.Name = ResolveFont(run.FontName);
                sel.Font.Size = FontSize(run.FontSizePt);
                sel.Font.Bold = run.Bold ? 1 : 0;
                sel.Font.Italic = run.Italic ? 1 : 0;
                sel.Font.Superscript = run.Super ? 1 : 0;
                sel.Font.Subscript = run.Sub ? 1 : 0;
                sel.Font.Underline = run.Underline ? 1 : 0; // wdUnderlineSingle / None
                sel.Font.Color = ToBgr(run.ColorArgb);
                if (string.IsNullOrEmpty(run.Uri))
                {
                    sel.TypeText(run.Text);
                }
                else
                {
                    int start = (int)sel.Range.End;
                    sel.TypeText(run.Text);
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
        /// Сбой построения таблицы не срывает документ (текст ячеек уже недоступен — но остальное цело).
        /// </summary>
        private static void WriteTable(dynamic word, dynamic doc, dynamic sel, OcrTable table)
        {
            int rows = table.Rows.Count, cols = table.ColumnCount;
            if (rows == 0 || cols == 0)
                return;
            try
            {
                dynamic wtable = doc.Tables.Add(sel.Range, rows, cols);
                wtable.AllowAutoFit = false;
                wtable.Borders.Enable = 1; // все границы

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
                        WriteCell(word, doc, wtable, r + 1, c + 1, cell); // Word адресует ячейки с 1
                    }
                }

                MergeSpans(wtable, table, rows, cols);

                // Курсор — за таблицу, отделить абзацем (иначе следующая таблица сольётся с этой).
                sel.Start = wtable.Range.End;
                sel.Collapse(WdCollapseEnd);
                sel.TypeParagraph();
            }
            catch { } // не удалось построить таблицу — не срываем весь документ
        }

        /// <summary>
        /// Объединить ячейки по ColSpan/RowSpan уже заполненной таблицы. Идём с конца (снизу
        /// вверх, справа налево): слияние блока перенумеровывает ячейки НИЖЕ и ПРАВЕЕ, а они уже
        /// обработаны, поэтому адреса ещё не слитых блоков (выше/левее) не сбиваются. Слияние
        /// склеивает содержимое ячеек, добавляя пустые абзацы накрытых — их подчищаем.
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

        /// <summary>Записать абзацы ячейки в её начало (не трогая маркер конца ячейки).</summary>
        private static void WriteCell(dynamic word, dynamic doc, dynamic wtable, int row, int col, OcrTableCell cell)
        {
            dynamic cellRange = wtable.Cell(row, col).Range;
            cellRange.Collapse(WdCollapseStart); // в начало ячейки, чтобы не съесть маркер ячейки
            cellRange.Select();
            dynamic sel = word.Selection;
            for (int i = 0; i < cell.Paragraphs.Count; i++)
            {
                if (i > 0)
                    sel.TypeParagraph();
                WriteParagraphInto(sel, doc, cell.Paragraphs[i], 0, true); // в ячейке: без красной строки, без выключки
            }
        }

        /// <summary>Ширина колонки в pt с нижней защитой (вырожденную колонку Word рисует криво).</summary>
        private static double ColWidth(double pt)
        {
            return pt < MinColWidthPt ? MinColWidthPt : pt;
        }

        /// <summary>
        /// Вставляет изображение inline в текущую позицию и переводит строку; размер — по рамке
        /// PDF (pt), с защитой пределов. Сбой одной картинки не срывает документ. PNG кладётся во
        /// временный файл (встраивается в .docx при вставке), временная папка чистится в Write.
        /// </summary>
        private static void InsertImage(dynamic sel, OcrImage img, string tempDir, ref int index)
        {
            if (img == null || img.Png == null || img.Png.Length == 0)
                return;
            string file = Path.Combine(tempDir, "img_" + index + ".png");
            index++;
            try
            {
                File.WriteAllBytes(file, img.Png);
                sel.ParagraphFormat.Alignment = WdAlignLeft;
                sel.ParagraphFormat.FirstLineIndent = 0;
                dynamic shape = sel.InlineShapes.AddPicture(file, false, true); // встроить в документ
                shape.LockAspectRatio = 0; // msoFalse — задаём оба размера
                shape.Width = ClampSize(img.WidthPt);
                shape.Height = ClampSize(img.HeightPt);
                sel.TypeParagraph(); // изображение на своей строке
            }
            catch { } // одна картинка не должна сорвать конвертацию
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
        /// выключка. Список шрифтов читается один раз.
        /// </summary>
        private static string ResolveFont(string requested)
        {
            return ResolveFontName(requested, InstalledFonts, DefaultFontName);
        }

        /// <summary>Чистая логика подстановки (под тест): установленный шрифт — оставить, иначе — fallback.</summary>
        internal static string ResolveFontName(string requested, ICollection<string> installed, string fallback)
        {
            if (string.IsNullOrEmpty(requested))
                return fallback;
            return installed != null && installed.Contains(requested) ? requested : fallback;
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

        private static double Median(List<double> values)
        {
            if (values.Count == 0)
                return 0;
            var copy = new List<double>(values);
            copy.Sort();
            return copy[(copy.Count - 1) / 2];
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
            return Median(vals);
        }
    }
}
