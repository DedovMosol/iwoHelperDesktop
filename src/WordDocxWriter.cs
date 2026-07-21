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

            WordCom.WriteDocx(path, "Файл Word", delegate(object wordObj, object docObj)
            {
                dynamic word = wordObj;
                dynamic sel = word.Selection;
                sel.Font.Name = "Times New Roman";
                sel.Font.Size = 12;
                sel.ParagraphFormat.Alignment = WdAlignJustify;

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
    }
}
