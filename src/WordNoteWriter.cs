using System;

namespace ExcelMerger
{
    /// <summary>
    /// Запись сопроводительной записки в .docx через COM установленного Word.
    /// Оформление по ГОСТ Р 7.0.97-2016: поля 30/15/20/20 мм, Times New Roman 14,
    /// красная строка 1,25 см, полуторный интервал, выравнивание по ширине.
    /// Каркас Word (открытие/сохранение/закрытие) — общий <see cref="WordCom"/> (DRY).
    /// Вызывать в STA-потоке.
    /// </summary>
    public static class WordNoteWriter
    {
        private const int WdAlignCenter = 1;
        private const int WdAlignJustify = 3;
        private const int WdLineSpace1pt5 = 1;
        private const int WdStory = 6;
        private const int WdAutoFitWindow = 2;

        public static void Write(NoteContent note, string path)
        {
            if (note == null)
                throw new ArgumentNullException("note");

            WordCom.WriteDocx(path, Loc.T("word.label.note"), delegate(object wordObj, object docObj)
            {
                dynamic word = wordObj;
                dynamic doc = docObj;

                dynamic setup = doc.PageSetup;
                setup.LeftMargin = word.CentimetersToPoints(3.0);
                setup.RightMargin = word.CentimetersToPoints(1.5);
                setup.TopMargin = word.CentimetersToPoints(2.0);
                setup.BottomMargin = word.CentimetersToPoints(2.0);

                dynamic sel = word.Selection;
                sel.Font.Name = "Times New Roman";
                sel.Font.Size = 14;

                // Заголовок
                sel.ParagraphFormat.Alignment = WdAlignCenter;
                sel.ParagraphFormat.LineSpacingRule = WdLineSpace1pt5;
                sel.Font.Bold = 1;
                sel.TypeText(note.Title);
                sel.TypeParagraph();
                sel.Font.Bold = 0;
                sel.TypeText(note.Subtitle);
                sel.TypeParagraph();
                sel.TypeParagraph();

                // Основной текст
                float indent = (float)word.CentimetersToPoints(1.25);
                sel.ParagraphFormat.Alignment = WdAlignJustify;
                sel.ParagraphFormat.FirstLineIndent = indent;
                foreach (string paragraph in note.Body)
                {
                    sel.TypeText(paragraph);
                    sel.TypeParagraph();
                }

                if (note.SkippedIntro != null)
                {
                    sel.TypeText(note.SkippedIntro);
                    sel.TypeParagraph();
                    AppendSkippedTable(word, doc, note);
                    dynamic selEnd = word.Selection;
                    selEnd.EndKey(WdStory);
                    selEnd.TypeParagraph();
                    selEnd.ParagraphFormat.Alignment = WdAlignJustify;
                    selEnd.ParagraphFormat.FirstLineIndent = indent;
                    selEnd.ParagraphFormat.LineSpacingRule = WdLineSpace1pt5;
                    selEnd.Font.Name = "Times New Roman";
                    selEnd.Font.Size = 14;
                    sel = selEnd;
                }

                foreach (string paragraph in note.Tail)
                {
                    sel.TypeText(paragraph);
                    sel.TypeParagraph();
                }

                sel.TypeParagraph();
                sel.ParagraphFormat.FirstLineIndent = 0f;
                sel.TypeText(note.Signature);
            });
        }

        private static void AppendSkippedTable(dynamic word, dynamic doc, NoteContent note)
        {
            dynamic table = doc.Tables.Add(word.Selection.Range, note.SkippedRows.Count + 1, 3);
            table.Borders.Enable = 1;
            table.Range.ParagraphFormat.FirstLineIndent = 0f;
            table.Range.Font.Size = 12; // таблица компактнее основного текста

            SetCell(table, 1, 1, "№", true);
            SetCell(table, 1, 2, "Файл", true);
            SetCell(table, 1, 3, "Причина", true);
            for (int i = 0; i < note.SkippedRows.Count; i++)
            {
                string[] row = note.SkippedRows[i];
                SetCell(table, i + 2, 1, row[0], false);
                SetCell(table, i + 2, 2, row[1], false);
                SetCell(table, i + 2, 3, row[2], false);
            }
            table.AutoFitBehavior(WdAutoFitWindow);
        }

        private static void SetCell(dynamic table, int row, int col, string text, bool bold)
        {
            dynamic range = table.Cell(row, col).Range;
            range.Text = text;
            range.Font.Bold = bold ? 1 : 0;
        }
    }
}
