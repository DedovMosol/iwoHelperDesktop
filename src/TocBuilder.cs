using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Строит лист «Содержание»: оглавление свода с гиперссылками на листы
    /// и статусами всех обработанных файлов (включая пропущенные).
    /// Работает с открытой итоговой книгой через COM; вызывать до SaveAs.
    /// </summary>
    internal static class TocBuilder
    {
        // Цвета в формате Excel (BGR: R + G*256 + B*65536).
        private const int HeaderFill = 4291600;   // #107C41 — акцент приложения
        private const int White = 16777215;
        private const int SkippedText = 2237106;  // #B22222

        public static void Build(object targetObj, string tocName, IList<FileResult> files)
        {
            dynamic target = targetObj;
            dynamic toc = target.Sheets.Add(target.Sheets[1]); // первым листом
            toc.Name = tocName;

            // Данные одним массивом — один COM-вызов вместо сотен поячеечных.
            int rows = files.Count + 1;
            var data = new object[rows, 4];
            data[0, 0] = "№";
            data[0, 1] = "Лист";
            data[0, 2] = "Исходный файл";
            data[0, 3] = "Примечание";
            for (int i = 0; i < files.Count; i++)
            {
                FileResult fr = files[i];
                data[i + 1, 0] = i + 1;
                data[i + 1, 1] = fr.Ok ? fr.SheetName : "—";
                data[i + 1, 2] = fr.FileName;
                data[i + 1, 3] = fr.Ok ? (fr.Note ?? "") : ("пропущен: " + fr.Note);
            }
            // Экранирование: имя файла, начинающееся с «=», Excel разобрал бы
            // как формулу, а ведущий апостроф имени — съел бы как префикс.
            toc.Range("A1").Resize(rows, 4).Value2 = CellText.EscapeValues(data);

            dynamic header = toc.Range("A1:D1");
            header.Font.Bold = true;
            header.Font.Color = White;
            header.Interior.Color = HeaderFill;

            for (int i = 0; i < files.Count; i++)
            {
                FileResult fr = files[i];
                int row = i + 2;
                if (fr.Ok && fr.Linkable)
                {
                    // Апострофы в имени листа в адресе ссылки удваиваются.
                    string subAddress = "'" + fr.SheetName.Replace("'", "''") + "'!A1";
                    toc.Hyperlinks.Add(toc.Cells[row, 2], "", subAddress, Type.Missing, fr.SheetName);
                }
                else if (!fr.Ok)
                {
                    toc.Range(toc.Cells[row, 1], toc.Cells[row, 4]).Font.Color = SkippedText;
                }
            }

            toc.Columns[1].ColumnWidth = 5;
            toc.Columns[2].ColumnWidth = 34;
            toc.Columns[3].ColumnWidth = 46;
            toc.Columns[4].ColumnWidth = 70;

            // Активировать «Содержание» и закрепить шапку — свод открывается с оглавления.
            try
            {
                toc.Activate();
                dynamic window = target.Application.ActiveWindow;
                window.SplitRow = 1;
                window.FreezePanes = true;
            }
            catch { } // закрепление — украшение, не причина ронять слияние
        }
    }
}
