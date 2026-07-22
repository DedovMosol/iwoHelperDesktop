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
        private const int WdPageBreak = 7;
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
                    ApplyPageSetup(docObj, pages); // размер страницы и поля из источника
                    dynamic sel = word.Selection;

                    for (int p = 0; p < pages.Count; p++)
                    {
                        if (p > 0)
                            sel.InsertBreak(WdPageBreak); // разрыв страницы между страницами PDF
                        foreach (Block blk in OrderedBlocks(pages[p]))
                        {
                            if (blk.Paragraph != null)
                                WriteParagraph(sel, doc, blk.Paragraph, firstLineIndent);
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

        /// <summary>Блок содержимого страницы: абзац или изображение (одно из полей задано).</summary>
        private sealed class Block
        {
            public OcrParagraph Paragraph;
            public OcrImage Image;
            public double Top; // верх блока — для порядка чтения
        }

        /// <summary>Абзацы и изображения страницы в порядке чтения (сверху вниз по Y).</summary>
        private static List<Block> OrderedBlocks(PdfPageText page)
        {
            var blocks = new List<Block>();
            if (page.Paragraphs != null)
                foreach (OcrParagraph par in page.Paragraphs)
                    blocks.Add(new Block { Paragraph = par, Top = par.TopPt });
            if (page.Images != null)
                foreach (OcrImage img in page.Images)
                    blocks.Add(new Block { Image = img, Top = img.TopPt });
            blocks.Sort(delegate(Block a, Block b) { return b.Top.CompareTo(a.Top); }); // ось Y вверх: больше Top — выше
            return blocks;
        }

        private static void WriteParagraph(dynamic sel, dynamic doc, OcrParagraph paragraph, double firstLineIndent)
        {
            // Выравнивание из источника; центрированное — без красной строки.
            int align; double indent;
            switch (paragraph.Alignment)
            {
                case OcrAlignment.Center: align = WdAlignCenter; indent = 0; break;
                case OcrAlignment.Left: align = WdAlignLeft; indent = firstLineIndent; break;
                default: align = WdAlignJustify; indent = firstLineIndent; break;
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
            sel.TypeParagraph();
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
        /// Размер страницы и поля из источника: размер — с первой страницы с текстом,
        /// поля — медиана по таким страницам (обычно одинаковы). Размер вне разумных пределов —
        /// оставляем шаблон Word. Поля косметические: сбой настройки не срывает конвертацию.
        /// </summary>
        private static void ApplyPageSetup(object docObj, IList<PdfPageText> pages)
        {
            double pw = 0, ph = 0;
            var l = new List<double>(); var r = new List<double>();
            var t = new List<double>(); var b = new List<double>();
            foreach (PdfPageText p in pages)
            {
                if (p.Paragraphs == null || p.Paragraphs.Count == 0)
                    continue; // страницы без текста поля не задают
                if (pw <= 0 && p.WidthPt > 0) { pw = p.WidthPt; ph = p.HeightPt; }
                l.Add(p.LeftMarginPt); r.Add(p.RightMarginPt);
                t.Add(p.TopMarginPt); b.Add(p.BottomMarginPt);
            }
            if (pw < MinPagePt || pw > MaxPagePt || ph < MinPagePt || ph > MaxPagePt)
                return;
            try
            {
                dynamic ps = ((dynamic)docObj).PageSetup;
                ps.PageWidth = pw;
                ps.PageHeight = ph;
                ps.LeftMargin = ClampMargin(Median(l), pw);
                ps.RightMargin = ClampMargin(Median(r), pw);
                ps.TopMargin = ClampMargin(Median(t), ph);
                ps.BottomMargin = ClampMargin(Median(b), ph);
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
