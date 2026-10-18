using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace ExcelMerger
{
    /// <summary>Содержимое сопроводительной записки — без привязки к Word.</summary>
    public class NoteContent
    {
        public string Title;
        public string Subtitle;
        public readonly List<string> Body = new List<string>();
        public string SkippedIntro; // null — пропусков нет, таблица не нужна
        public readonly List<string[]> SkippedRows = new List<string[]>(); // №, файл, причина
        public readonly List<string> Tail = new List<string>();
        public string Signature;
    }

    /// <summary>
    /// Текст сопроводительной записки по итогам слияния. Чистая функция —
    /// покрыта юнит-тестами; оформление в docx делает WordNoteWriter.
    /// </summary>
    public static class NoteText
    {
        private static readonly CultureInfo Ru = CultureInfo.GetCultureInfo("ru-RU");

        public static NoteContent Build(MergeResult result, string inputFolder,
            MergeOptions options, DateTime startedAt)
        {
            if (result == null)
                throw new ArgumentNullException("result");

            var note = new NoteContent();
            note.Title = "СПРАВКА";
            note.Subtitle = "о формировании сводного файла отчётности";

            note.Body.Add(startedAt.ToString("d MMMM yyyy 'г.'", Ru) + " в " +
                startedAt.ToString("HH:mm", Ru) +
                " выполнено объединение файлов отчётности из папки «" + inputFolder + "».");
            // SkipCount считает записи (в режиме «все листы» — листы), а фраза говорит о файлах:
            // файл, вошедший в свод частично, пропущенным не называем — его потери отдельно.
            int skippedFiles, lostSheets;
            result.CountSkips(out skippedFiles, out lostSheets);
            string counts = "Обработано файлов: " + result.FileCount +
                ". Включено листов в сводный файл: " + result.OkCount +
                ". Пропущено файлов: " + skippedFiles + ".";
            if (lostSheets > 0)
                counts += " Не перенесено листов из частично включённых файлов: " + lostSheets + ".";
            note.Body.Add(counts);

            if (result.SkipCount > 0)
            {
                note.SkippedIntro = "Не включены в сводный файл следующие файлы:";
                int n = 0;
                foreach (FileResult fr in result.Files)
                {
                    if (fr.Ok)
                        continue;
                    n++;
                    note.SkippedRows.Add(new[] { n.ToString(), fr.FileName, fr.Note ?? "" });
                }
            }
            else
            {
                note.Body.Add("Замечания отсутствуют: все файлы включены в сводный файл.");
            }

            note.Tail.Add("Сводный файл: " + result.OutputPath +
                " (" + Path.GetExtension(result.OutputPath).TrimStart('.').ToUpperInvariant() + ").");
            note.Tail.Add("Параметры формирования: лист «Содержание» — " +
                YesNo(options.AddToc) + "; замена формул значениями — " +
                YesNo(options.ValuesOnly) + ".");
            if (result.TocError != null)
                note.Tail.Add("Внимание: " + result.TocError + ".");
            if (result.MacrosCarried)
                note.Tail.Add("Внимание: " + MergeService.MacrosWarningText + ".");

            note.Signature = "Исполнитель: ________________________";
            return note;
        }

        private static string YesNo(bool value)
        {
            return value ? "да" : "нет";
        }
    }
}
