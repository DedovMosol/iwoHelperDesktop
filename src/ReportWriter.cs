using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Текстовый отчёт о слиянии: единый формат строк для CLI и GUI
    /// и сохранение истории в папку отчётов с ротацией (старые удаляются).
    /// </summary>
    public static class ReportWriter
    {
        private const string FilePrefix = "report_";

        /// <summary>Строка результата по одному файлу — общая для CLI-вывода и отчёта.</summary>
        public static string FormatFileLine(FileResult fr)
        {
            return (fr.Ok ? "OK      " : "SKIPPED ") + fr.FileName +
                (fr.SheetName != null ? " -> [" + fr.SheetName + "]" : "") +
                (string.IsNullOrEmpty(fr.Note) ? "" : " | " + fr.Note);
        }

        /// <summary>Полный отчёт о слиянии для сохранения в историю.</summary>
        public static string BuildReport(MergeResult result, string inputFolder,
            MergeOptions options, DateTime startedAt)
        {
            var sb = new StringBuilder();
            sb.AppendLine("iwo Helper Desktop — отчёт о слиянии");
            sb.AppendLine("Дата:           " + startedAt.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine("Папка:          " + inputFolder);
            sb.AppendLine("Итоговый файл:  " + result.OutputPath);
            sb.AppendLine("Параметры:      лист «Содержание»: " + (options.AddToc ? "да" : "нет") +
                "; формулы → значения: " + (options.ValuesOnly ? "да" : "нет"));
            if (result.Cancelled)
                sb.AppendLine("Результат:      отменено пользователем — итоговый файл не создан");
            else
                sb.AppendLine("Результат:      перенесено " + result.OkCount + ", пропущено " + result.SkipCount);
            if (result.TocError != null)
                sb.AppendLine("Внимание:       " + result.TocError);
            sb.AppendLine();
            foreach (FileResult fr in result.Files)
                sb.AppendLine(FormatFileLine(fr));
            return sb.ToString();
        }

        /// <summary>
        /// Сохраняет отчёт в папку истории и удаляет старые, оставляя не более
        /// keepCount файлов. Возвращает путь сохранённого отчёта.
        /// </summary>
        public static string SaveWithRotation(string directory, string content,
            DateTime timestamp, int keepCount)
        {
            if (keepCount < 1)
                throw new ArgumentOutOfRangeException("keepCount");
            Directory.CreateDirectory(directory);

            // Имя сортируется лексикографически по времени; коллизии — суффиксом.
            string baseName = FilePrefix + timestamp.ToString("yyyy-MM-dd_HH-mm-ss");
            string path = Path.Combine(directory, baseName + ".txt");
            for (int i = 2; File.Exists(path); i++)
                path = Path.Combine(directory, baseName + "_" + i + ".txt");
            File.WriteAllText(path, content, Encoding.UTF8);

            var reports = new List<string>(Directory.GetFiles(directory, FilePrefix + "*.txt"));
            reports.Sort(StringComparer.OrdinalIgnoreCase);
            reports.Reverse(); // новые первыми
            for (int i = keepCount; i < reports.Count; i++)
            {
                try { File.Delete(reports[i]); }
                catch { } // залоченный старый отчёт не должен ломать сохранение нового
            }
            return path;
        }
    }
}
