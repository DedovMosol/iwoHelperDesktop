using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Запись извлечённого текста born-digital PDF в .docx через COM Word: абзацы в
    /// порядке чтения, разрыв страницы между страницами PDF. Каркас Word (открытие/
    /// сохранение/закрытие) — общий <see cref="WordCom"/> (DRY). Вызывать в STA-потоке.
    /// </summary>
    public static class WordDocxWriter
    {
        private const int WdAlignLeft = 0;
        private const int WdAlignCenter = 1;
        private const int WdAlignJustify = 3;
        private const int WdPageBreak = 7;
        private const double DefaultFontSize = 12;
        private const double MinFontSize = 5;   // защита от мусорного кегля из PDF
        private const double MaxFontSize = 72;
        private const double MinPagePt = 72;    // 1"; разумные пределы размера страницы
        private const double MaxPagePt = 1584;  // 22" — максимум Word

        /// <summary>Пишет .docx из абзацев страниц. Занятый файл/нет Word — MergeException.</summary>
        public static void Write(IList<PdfPageText> pages, string path)
        {
            if (pages == null)
                throw new ArgumentNullException("pages");

            double firstLineIndent = DocumentIndent(pages); // pt; 0 — документ без красной строки

            WordCom.WriteDocx(path, "Файл Word", delegate(object wordObj, object docObj)
            {
                dynamic word = wordObj;
                ApplyPageSetup(docObj, pages); // размер страницы и поля из источника
                dynamic sel = word.Selection;
                sel.Font.Name = "Times New Roman";

                for (int p = 0; p < pages.Count; p++)
                {
                    if (p > 0)
                        sel.InsertBreak(WdPageBreak); // разрыв страницы между страницами PDF
                    List<OcrParagraph> paragraphs = pages[p].Paragraphs;
                    if (paragraphs == null)
                        continue;
                    foreach (OcrParagraph paragraph in paragraphs)
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

                        // Формат пословно (ран за раном): кегль, полужирный, курсив, цвет.
                        foreach (OcrRun run in paragraph.Runs)
                        {
                            sel.Font.Size = FontSize(run.FontSizePt);
                            sel.Font.Bold = run.Bold ? 1 : 0;
                            sel.Font.Italic = run.Italic ? 1 : 0;
                            sel.Font.Color = ToBgr(run.ColorArgb);
                            sel.TypeText(run.Text);
                        }
                        sel.TypeParagraph();
                    }
                }
            });
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
