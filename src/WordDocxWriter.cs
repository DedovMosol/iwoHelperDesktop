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
        private const int WdAlignJustify = 3;
        private const int WdPageBreak = 7;

        /// <summary>Пишет .docx из абзацев страниц. Занятый файл/нет Word — MergeException.</summary>
        public static void Write(IList<PdfPageText> pages, string path)
        {
            if (pages == null)
                throw new ArgumentNullException("pages");

            double firstLineIndent = DocumentIndent(pages); // pt; 0 — документ без красной строки

            WordCom.WriteDocx(path, "Файл Word", delegate(object wordObj, object docObj)
            {
                dynamic word = wordObj;
                dynamic sel = word.Selection;
                sel.Font.Name = "Times New Roman";
                sel.Font.Size = 12;
                sel.ParagraphFormat.Alignment = WdAlignJustify;
                sel.ParagraphFormat.FirstLineIndent = firstLineIndent; // красная строка, если была в источнике

                for (int p = 0; p < pages.Count; p++)
                {
                    if (p > 0)
                        sel.InsertBreak(WdPageBreak); // разрыв страницы между страницами PDF
                    List<string> paragraphs = pages[p].Paragraphs;
                    if (paragraphs == null)
                        continue;
                    foreach (string paragraph in paragraphs)
                    {
                        sel.TypeText(paragraph);
                        sel.TypeParagraph();
                    }
                }
            });
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
            if (vals.Count == 0)
                return 0;
            vals.Sort();
            return vals[(vals.Count - 1) / 2];
        }
    }
}
