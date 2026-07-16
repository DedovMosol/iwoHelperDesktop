using System;

namespace ExcelMerger
{
    /// <summary>
    /// Запись сопроводительной записки в .docx через COM установленного Word.
    /// Оформление по ГОСТ Р 7.0.97-2016: поля 30/15/20/20 мм, Times New Roman 14,
    /// красная строка 1,25 см, полуторный интервал, выравнивание по ширине.
    /// Действуют те же COM-правила, что и в MergeService: после Close/Quit —
    /// никаких динамических операций (см. комментарий у ComSafe.Release).
    /// Вызывать в STA-потоке.
    /// </summary>
    public static class WordNoteWriter
    {
        private const int WdAlignCenter = 1;
        private const int WdAlignJustify = 3;
        private const int WdLineSpace1pt5 = 1;
        private const int WdFormatXmlDocument = 12; // .docx
        private const int WdStory = 6;
        private const int WdAutoFitWindow = 2;

        public static void Write(NoteContent note, string path)
        {
            if (note == null)
                throw new ArgumentNullException("note");

            Type wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
                throw new MergeException("Microsoft Word не установлен: COM-компонент Word.Application не найден.");

            string lockError = MergeService.CheckOutputWritable(path);
            if (lockError != null)
                throw new MergeException(lockError.Replace("Итоговый файл", "Файл записки"));

            dynamic word = null;
            dynamic doc = null;
            ComMessageFilter.Register();
            try
            {
                word = Activator.CreateInstance(wordType);
                word.Visible = false;
                word.DisplayAlerts = 0; // wdAlertsNone

                doc = word.Documents.Add();
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

                try
                {
                    doc.SaveAs2(path, WdFormatXmlDocument);
                }
                catch (Exception ex)
                {
                    throw new MergeException("Не удалось сохранить записку: " + ex.Message);
                }
            }
            finally
            {
                // Статические ссылки: после Close/Quit — никаких динамических операций.
                object docObj = doc;
                if (docObj != null)
                {
                    try { doc.Close(0); } catch { } // wdDoNotSaveChanges: файл уже сохранён
                    ComSafe.Release(docObj);
                }
                object wordObj = word;
                if (wordObj != null)
                {
                    try { word.Quit(0); } catch { }
                    ComSafe.Release(wordObj);
                }
                ComSafe.Collect();
                ComMessageFilter.Revoke();
            }
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
