using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Золотые тесты (этап 2): текст, который человек получает файлом, сверяется с эталоном
    /// целиком, а не выборочными Contains. Эталоны — tests\golden\*.txt, вшиты в UnitTests.exe
    /// ресурсами (прогон не зависит от расположения исходников). Изменили формат намеренно —
    /// прогон с IWO_UPDATE_GOLDEN=1 перепишет эталон в репозитории и упадёт с просьбой
    /// посмотреть diff: эталон меняется только осознанно и попадает в ревью.
    /// Каждый эталон строится под тайской культурой (буддийский календарь, свой формат): вывод
    /// обязан от культуры машины не зависеть.
    /// </summary>
    internal static partial class UnitTests
    {
        private static void AssertGolden(string name, string actual)
        {
            string normalized = actual.Replace("\r\n", "\n");
            if (Environment.GetEnvironmentVariable("IWO_UPDATE_GOLDEN") == "1")
            {
                string root = RepoRoot();
                AssertTrue(root != null, "IWO_UPDATE_GOLDEN: исходники рядом не найдены — эталон некуда записать");
                string path = Path.Combine(root, "tests", "golden", name + ".txt");
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, normalized, new UTF8Encoding(false));
                throw new Exception("эталон " + name + " переписан — проверьте diff и пересоберите тесты");
            }
            string expected = GoldenResource(name);
            if (expected == normalized)
                return;
            string[] e = expected.Split('\n');
            string[] a = normalized.Split('\n');
            int line = 0;
            while (line < e.Length && line < a.Length && e[line] == a[line])
                line++;
            throw new Exception("расходится с эталоном " + name + " в строке " + (line + 1) + ": ожидалось «" +
                (line < e.Length ? e[line] : "<конец>") + "», получено «" + (line < a.Length ? a[line] : "<конец>") + "»");
        }

        private static string GoldenResource(string name)
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("golden." + name + ".txt"))
            {
                AssertTrue(stream != null, "эталона golden." + name + ".txt нет в ресурсах тестов (IWO_UPDATE_GOLDEN=1 создаст)");
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                    return reader.ReadToEnd().Replace("\r\n", "\n");
            }
        }

        /// <summary>Выполнить build под чужой культурой и вернуть результат.</summary>
        private static string UnderCulture(string culture, Func<string> build)
        {
            CultureInfo savedCulture = Thread.CurrentThread.CurrentCulture;
            CultureInfo savedUi = Thread.CurrentThread.CurrentUICulture;
            try
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(culture);
                Thread.CurrentThread.CurrentUICulture = CultureInfo.GetCultureInfo(culture);
                return build();
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = savedCulture;
                Thread.CurrentThread.CurrentUICulture = savedUi;
            }
        }

        /// <summary>Итог свода со всеми видами записей: перенесён, пропущен, завис, лист без ссылки.</summary>
        private static MergeResult GoldenMergeResult()
        {
            var result = new MergeResult { OutputPath = @"C:\Отчёты\Свод.xlsm", MacrosCarried = true };
            result.TocError = "лист «Содержание» создать не удалось (тест)";
            result.Files.Add(new FileResult { FileName = "Январь.xlsx", FullPath = @"C:\Отчёты\Январь.xlsx", Ok = true, SheetName = "Январь", Linkable = true });
            result.Files.Add(new FileResult { FileName = "Февраль.xlsx", FullPath = @"C:\Отчёты\Февраль.xlsx", Ok = true, SheetName = "Февраль",
                Note = "файл содержит макросы (не выполнялись)", Linkable = true });
            result.Files.Add(new FileResult { FileName = "Март.xlsx", FullPath = @"C:\Отчёты\Март.xlsx", Note = "файл защищён паролем" });
            result.Files.Add(new FileResult { FileName = "Апрель.xlsx", FullPath = @"C:\Отчёты\Апрель.xlsx",
                Note = MergeService.StalledNote(WatchdogStop.Dialog, "Открытие документа", MergeService.StepLimitMs) });
            result.Files.Add(new FileResult { FileName = "Май.xlsx", FullPath = @"C:\Отчёты\Май.xlsx", Note = MergeService.WedgedNote });
            result.OkCount = 2;
            result.SkipCount = 3;
            return result;
        }

        private static readonly DateTime GoldenStartedAt = new DateTime(2026, 10, 27, 9, 5, 7);

        [Test("Золотой: отчёт о слиянии совпадает с эталоном и не зависит от культуры машины")]
        private static void GoldenMergeReport()
        {
            string report = UnderCulture("th-TH", delegate
            {
                return ReportWriter.BuildReport(GoldenMergeResult(), @"C:\Отчёты",
                    new MergeOptions { AddToc = true, ValuesOnly = true }, GoldenStartedAt);
            });
            AssertGolden("merge-report", report);
        }

        [Test("Золотой: сопроводительная записка совпадает с эталоном и не зависит от культуры машины")]
        private static void GoldenMergeNote()
        {
            string text = UnderCulture("th-TH", delegate
            {
                NoteContent note = NoteText.Build(GoldenMergeResult(), @"C:\Отчёты",
                    new MergeOptions { AddToc = true }, GoldenStartedAt);
                var sb = new StringBuilder();
                sb.Append("Заголовок: ").Append(note.Title).Append('\n');
                sb.Append("Подзаголовок: ").Append(note.Subtitle).Append('\n');
                foreach (string p in note.Body)
                    sb.Append("Абзац: ").Append(p).Append('\n');
                sb.Append("Пропуски: ").Append(note.SkippedIntro ?? "<нет>").Append('\n');
                foreach (string[] row in note.SkippedRows)
                    sb.Append("Строка: ").Append(string.Join(" | ", row)).Append('\n');
                foreach (string p in note.Tail)
                    sb.Append("Хвост: ").Append(p).Append('\n');
                sb.Append("Подпись: ").Append(note.Signature).Append('\n');
                return sb.ToString();
            });
            AssertGolden("merge-note", text);
        }

        [Test("Золотой: сверка с эталоном показывает первую расходящуюся строку")]
        private static void GoldenComparerReportsFirstDifference()
        {
            if (Environment.GetEnvironmentVariable("IWO_UPDATE_GOLDEN") == "1")
                return; // в режиме обновления сверки нет — сравнивать нечего
            string expected = GoldenResource("merge-report");
            string changed = expected.Replace("перенесено 2", "перенесено 3");
            AssertTrue(changed != expected, "в эталоне есть строка итога");
            var error = AssertThrowsExactly<Exception>("расхождение с эталоном", delegate { AssertGolden("merge-report", changed); });
            AssertTrue(error.Message.Contains("перенесено 3") && error.Message.Contains("перенесено 2"),
                "в сообщении — ожидаемая и полученная строки: " + error.Message);
        }
    }
}
