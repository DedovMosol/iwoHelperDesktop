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
                    toc.Hyperlinks.Add(toc.Cells[row, 2], "", SheetRef(fr.SheetName), Type.Missing, fr.SheetName);
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

        /// <summary>Ссылка на ячейку A1 листа для гиперссылки (апострофы в имени удваиваются).</summary>
        internal static string SheetRef(string sheetName)
        {
            return "'" + (sheetName ?? "").Replace("'", "''") + "'!A1";
        }

        // Имя нашей фигуры-кнопки: чтобы находить и заменять её при дослиянии.
        private const string ReturnShapeName = "iwoTocLink";

        /// <summary>
        /// На каждый лист свода (кроме самого оглавления) — заметную кнопку-ссылку
        /// «К оглавлению»: плавающая фигура (данные не сдвигаются), синий градиент,
        /// белый текст. Идемпотентно: прежняя такая кнопка сначала удаляется —
        /// корректно при повторном дослиянии пропущенных.
        /// </summary>
        public static void AddReturnButtons(object targetObj, string tocName)
        {
            dynamic target = targetObj;
            dynamic sheets = target.Sheets;
            int count = (int)sheets.Count;
            string sub = SheetRef(tocName);
            int blue = Theme.ToBgr(Theme.HubBlue);
            int blueDark = Theme.ToBgr(Theme.HubBlueDark);
            for (int i = 1; i <= count; i++)
            {
                dynamic sheet = sheets[i];
                if (string.Equals((string)sheet.Name, tocName, StringComparison.Ordinal))
                    continue; // на самом оглавлении ссылка на оглавление не нужна
                try { AddReturnButton(sheet, sub, blue, blueDark); }
                catch { } // лист-диаграмма или иная особенность — без кнопки, свод не роняем
            }
        }

        private static void AddReturnButton(dynamic sheet, string subAddress, int blue, int blueDark)
        {
            try { sheet.Shapes(ReturnShapeName).Delete(); } catch { } // прежняя кнопка (дослияние)

            dynamic used = sheet.UsedRange; // на листе-диаграмме бросит — поймает вызывающий
            double top = (double)used.Top;
            double left = (double)used.Left + (double)used.Width + 8; // справа от данных, не перекрывая их

            dynamic btn = sheet.Shapes.AddShape(5, left, top, 160, 28); // 5 = скруглённый прямоугольник
            try
            {
                btn.Name = ReturnShapeName;
                // TextFrame2 — чистые свойства (в отличие от TextFrame.Characters);
                // проверено прямым COM-вызовом на живом Excel.
                dynamic tr = btn.TextFrame2.TextRange;
                tr.Text = "☰  К оглавлению";
                try
                {
                    dynamic fill = btn.Fill;
                    fill.TwoColorGradient(2, 1);     // 2 = msoGradientVertical
                    fill.ForeColor.RGB = blue;
                    fill.BackColor.RGB = blueDark;
                    btn.Line.Visible = 0;            // без рамки (msoFalse)
                    btn.Shadow.Visible = -1;         // мягкая тень (msoTrue) — дизайнерский вид
                    dynamic font = tr.Font;
                    font.Bold = -1;                  // msoTrue
                    font.Size = 11;
                    font.Name = "Segoe UI";
                    font.Fill.ForeColor.RGB = White;
                    tr.ParagraphFormat.Alignment = 2;   // msoAlignCenter
                    btn.TextFrame2.VerticalAnchor = 3;  // msoAnchorMiddle
                }
                catch { } // оформление — по возможности; сама ссылка важнее

                sheet.Hyperlinks.Add(btn, "", subAddress, "Перейти к оглавлению");
            }
            catch
            {
                try { btn.Delete(); } catch { } // без рабочей ссылки мёртвую кнопку не оставляем
            }
        }
    }
}
