using System;
using System.Collections.Generic;
using System.IO;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Юнит-тесты без внешних фреймворков: компилируются вместе с src\*.cs
    /// в отдельный консольный exe (tests\build_tests.cmd) и не попадают
    /// в производственную сборку. Код выхода 0 — все тесты прошли.
    /// </summary>
    internal static class UnitTests
    {
        private static int _passed;
        private static int _failed;

        private static int Main()
        {
            Run("SheetNamer: запрещённые символы заменяются на «_»", TestNamerForbiddenChars);
            Run("SheetNamer: обрезка до 31 символа", TestNamerTruncation);
            Run("SheetNamer: дедупликация суффиксами _2, _3", TestNamerDedupe);
            Run("SheetNamer: дедупликация длинного имени не превышает 31", TestNamerDedupeLong);
            Run("SheetNamer: Reserve исключает имя", TestNamerReserve);
            Run("SheetNamer: пустое имя становится «Лист»", TestNamerEmpty);
            Run("SheetNamer: History зарезервировано Excel", TestNamerHistory);
            Run("SheetNamer: апострофы по краям срезаются", TestNamerApostrophes);
            Run("Natural: «Отчет 2» раньше «Отчет 10»", TestNaturalNumbers);
            Run("Natural: регистр не учитывается", TestNaturalCase);
            Run("Natural: null меньше любой строки", TestNaturalNulls);
            Run("Natural: сортировка списка как в Проводнике", TestNaturalSortList);
            Run("FindSourceFiles: фильтры, исключения, порядок", TestFindSourceFiles);
            Run("CellText: строки экранируются апострофом", TestCellTextEscape);
            Run("CellText: не-строки проходят без изменений", TestCellTextPassthrough);
            Run("CellText: массив области (включая 1-базный COM)", TestCellTextArrays);
            Run("CLI: разбор флагов --toc/--values", TestCliOptions);
            Run("CLI: неизвестный флаг — ошибка", TestCliUnknownFlag);
            Run("ReportWriter: формат строки файла", TestReportLine);
            Run("ReportWriter: содержимое полного отчёта", TestReportBuild);
            Run("ReportWriter: ротация хранит не более 3 отчётов", TestReportRotation);
            Run("ReportWriter: коллизия имён в одну секунду", TestReportNameCollision);
            Run("Retry: пропущенные заменяются свежими результатами", TestCombineRetryReplaces);
            Run("Retry: неудачный повтор обновляет причину", TestCombineRetryKeepsFailed);
            Run("Retry: порядок и успешные записи не меняются", TestCombineRetryOrder);
            Run("OutputFormats: код формата по расширению", TestOutputFormatCodes);
            Run("OutputFormats: срез введённого расширения", TestStripExtension);
            Run("CheckOutputWritable: занятый файл распознан", TestOutputLocked);
            Run("CheckOutputWritable: свободный и новый файлы", TestOutputWritable);
            Run("CheckOutputWritable: несуществующая папка", TestOutputBadFolder);

            Console.WriteLine();
            Console.WriteLine("Пройдено: " + _passed + ", провалено: " + _failed);
            return _failed == 0 ? 0 : 1;
        }

        // ---------- SheetNamer ----------

        private static void TestNamerForbiddenChars()
        {
            var n = new SheetNamer();
            AssertEqual("От_чет _март_ 1_", n.Next("От:чет [март] 1?"), "санитизация");
        }

        private static void TestNamerTruncation()
        {
            var n = new SheetNamer();
            string name = n.Next("Очень длинное имя файла отчета министерства за март");
            AssertTrue(name.Length <= 31, "длина " + name.Length + " превышает 31");
            AssertEqual("Очень длинное имя файла отчета", name, "обрезка с зачисткой хвоста");
        }

        private static void TestNamerDedupe()
        {
            var n = new SheetNamer();
            AssertEqual("Отчет", n.Next("Отчет"), "первое имя");
            AssertEqual("Отчет_2", n.Next("Отчет"), "второе имя");
            // Дубликаты ищутся без учёта регистра, но регистр исходного имени сохраняется.
            AssertEqual("отчет_3", n.Next("отчет"), "регистронезависимый дубль");
        }

        private static void TestNamerDedupeLong()
        {
            var n = new SheetNamer();
            string baseName = new string('а', 31);
            string first = n.Next(baseName);
            string second = n.Next(baseName);
            AssertEqual(31, first.Length, "длина первого");
            AssertTrue(second.Length <= 31, "длина дубля " + second.Length);
            AssertTrue(second.EndsWith("_2"), "суффикс дубля: " + second);
        }

        private static void TestNamerReserve()
        {
            var n = new SheetNamer();
            n.Reserve("Содержание");
            AssertEqual("Содержание_2", n.Next("Содержание"), "резерв обходится суффиксом");
        }

        private static void TestNamerEmpty()
        {
            var n = new SheetNamer();
            AssertEqual("Лист", n.Next("   "), "пустое имя");
        }

        private static void TestNamerHistory()
        {
            var n = new SheetNamer();
            AssertEqual("History_", n.Next("History"), "зарезервированное имя");
        }

        private static void TestNamerApostrophes()
        {
            var n = new SheetNamer();
            AssertEqual("абв", n.Next("'абв'"), "апострофы по краям");
        }

        // ---------- NaturalStringComparer ----------

        private static void TestNaturalNumbers()
        {
            AssertTrue(NaturalStringComparer.Instance.Compare("Отчет 2", "Отчет 10") < 0, "2 < 10");
            AssertTrue(NaturalStringComparer.Instance.Compare("Отчет 10", "Отчет 2") > 0, "10 > 2");
        }

        private static void TestNaturalCase()
        {
            AssertEqual(0, NaturalStringComparer.Instance.Compare("отчет 10", "ОТЧЕТ 10"), "регистр");
        }

        private static void TestNaturalNulls()
        {
            AssertTrue(NaturalStringComparer.Instance.Compare(null, "a") < 0, "null < строки");
            AssertTrue(NaturalStringComparer.Instance.Compare("a", null) > 0, "строка > null");
            AssertEqual(0, NaturalStringComparer.Instance.Compare(null, null), "null == null");
        }

        private static void TestNaturalSortList()
        {
            var items = new List<string> { "Файл 10", "Файл 2", "Файл 1", "Файл 20" };
            items.Sort(NaturalStringComparer.Instance);
            AssertEqual("Файл 1|Файл 2|Файл 10|Файл 20", string.Join("|", items.ToArray()), "порядок");
        }

        // ---------- FindSourceFiles ----------

        private static void TestFindSourceFiles()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                File.WriteAllText(Path.Combine(dir, "Отчет 10.xlsx"), "");
                File.WriteAllText(Path.Combine(dir, "Отчет 2.XLSX"), "");   // регистр расширения
                File.WriteAllText(Path.Combine(dir, "старый.xls"), "");
                File.WriteAllText(Path.Combine(dir, "макрос.xlsm"), "");
                File.WriteAllText(Path.Combine(dir, "бинарный.xlsb"), "");
                File.WriteAllText(Path.Combine(dir, "~" + "$Отчет 10.xlsx"), ""); // временный Excel
                File.WriteAllText(Path.Combine(dir, "прочее.txt"), "");
                File.WriteAllText(Path.Combine(dir, "Свод.xlsx"), "");      // итоговый файл

                List<string> files = MergeService.FindSourceFiles(dir, Path.Combine(dir, "Свод.xlsx"));

                var names = new List<string>();
                foreach (string f in files)
                    names.Add(Path.GetFileName(f));
                AssertEqual("бинарный.xlsb|макрос.xlsm|Отчет 2.XLSX|Отчет 10.xlsx|старый.xls",
                    string.Join("|", names.ToArray()), "состав и порядок");
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        // ---------- CellText ----------

        private static void TestCellTextEscape()
        {
            AssertEqual("'=SUM(A1)", CellText.EscapeForEntry("=SUM(A1)"), "формула-инъекция");
            AssertEqual("'12 345", CellText.EscapeForEntry("12 345"), "число-подобная строка");
            AssertEqual("'01.02.2026", CellText.EscapeForEntry("01.02.2026"), "дата-подобная строка");
            AssertEqual("''уже с апострофом", CellText.EscapeForEntry("'уже с апострофом"), "апостроф сохраняется");
            AssertEqual(null, CellText.EscapeForEntry(null), "null");
            AssertEqual("", CellText.EscapeForEntry(""), "пустая строка");
        }

        private static void TestCellTextPassthrough()
        {
            AssertEqual(42.5, CellText.EscapeValues(42.5), "число");
            AssertEqual(true, CellText.EscapeValues(true), "булево");
            AssertEqual(null, CellText.EscapeValues(null), "null");
        }

        private static void TestCellTextArrays()
        {
            // Обычный 0-базный массив
            var plain = new object[1, 2];
            plain[0, 0] = "=x";
            plain[0, 1] = 7.0;
            var outPlain = (object[,])CellText.EscapeValues(plain);
            AssertEqual("'=x", outPlain[0, 0], "строка в 0-базном массиве");
            AssertEqual(7.0, outPlain[0, 1], "число в 0-базном массиве");

            // 1-базный массив — именно такой возвращает Range.Value2 через COM
            var comStyle = (object[,])Array.CreateInstance(
                typeof(object), new[] { 2, 1 }, new[] { 1, 1 });
            comStyle[1, 1] = "12 345";
            comStyle[2, 1] = 30.0;
            var outCom = (object[,])CellText.EscapeValues(comStyle);
            AssertEqual("'12 345", outCom[1, 1], "строка в 1-базном массиве");
            AssertEqual(30.0, outCom[2, 1], "число в 1-базном массиве");
        }

        // ---------- CLI ----------

        private static void TestCliOptions()
        {
            MergeOptions o;
            string err;
            AssertTrue(Program.TryParseCliOptions(new[] { "--cli", "in", "out" }, 3, out o, out err), "без флагов");
            AssertTrue(!o.AddToc && !o.ValuesOnly, "умолчания CLI — выключено");

            AssertTrue(Program.TryParseCliOptions(new[] { "--cli", "in", "out", "--toc", "--values" }, 3, out o, out err), "оба флага");
            AssertTrue(o.AddToc && o.ValuesOnly, "флаги применены");

            AssertTrue(Program.TryParseCliOptions(new[] { "--cli", "in", "out", "--VALUES" }, 3, out o, out err), "регистр флага");
            AssertTrue(o.ValuesOnly, "флаг в верхнем регистре");
        }

        private static void TestCliUnknownFlag()
        {
            MergeOptions o;
            string err;
            AssertTrue(!Program.TryParseCliOptions(new[] { "--cli", "in", "out", "--wtf" }, 3, out o, out err), "неизвестный флаг отвергнут");
            AssertTrue(err != null && err.Contains("--wtf"), "текст ошибки называет флаг");
        }

        // ---------- ReportWriter ----------

        private static void TestReportLine()
        {
            var ok = new FileResult();
            ok.FileName = "а.xlsx";
            ok.Ok = true;
            ok.SheetName = "а";
            AssertEqual("OK      а.xlsx -> [а]", ReportWriter.FormatFileLine(ok), "перенесённый");

            var skip = new FileResult();
            skip.FileName = "б.xlsx";
            skip.Note = "битый";
            AssertEqual("SKIPPED б.xlsx | битый", ReportWriter.FormatFileLine(skip), "пропущенный");
        }

        private static void TestReportBuild()
        {
            var result = new MergeResult();
            result.OutputPath = @"C:\out\Свод.xlsx";
            result.OkCount = 1;
            result.SkipCount = 1;
            var ok = new FileResult(); ok.FileName = "а.xlsx"; ok.Ok = true; ok.SheetName = "а";
            var skip = new FileResult(); skip.FileName = "б.xlsx"; skip.Note = "битый";
            result.Files.Add(ok);
            result.Files.Add(skip);
            var options = new MergeOptions(); options.AddToc = true;

            string report = ReportWriter.BuildReport(result, @"C:\in", options, new DateTime(2026, 7, 16, 14, 0, 0));
            AssertTrue(report.Contains("2026-07-16 14:00:00"), "дата");
            AssertTrue(report.Contains(@"C:\in"), "входная папка");
            AssertTrue(report.Contains(@"C:\out\Свод.xlsx"), "итоговый файл");
            AssertTrue(report.Contains("лист «Содержание»: да"), "параметры");
            AssertTrue(report.Contains("перенесено 1, пропущено 1"), "итог");
            AssertTrue(report.Contains("OK      а.xlsx"), "строка файла");
            AssertTrue(report.Contains("SKIPPED б.xlsx"), "строка пропуска");
        }

        private static void TestReportRotation()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N"));
            try
            {
                var stamp = new DateTime(2026, 7, 16, 10, 0, 0);
                for (int i = 0; i < 5; i++)
                    ReportWriter.SaveWithRotation(dir, "отчёт " + i, stamp.AddMinutes(i), 3);

                string[] files = Directory.GetFiles(dir, "report_*.txt");
                AssertEqual(3, files.Length, "число отчётов после ротации");
                Array.Sort(files, StringComparer.OrdinalIgnoreCase);
                AssertTrue(files[0].Contains("10-02-00"), "старейший из оставшихся");
                AssertTrue(files[2].Contains("10-04-00"), "новейший");
                AssertEqual("отчёт 4", File.ReadAllText(files[2]).TrimStart('﻿'), "содержимое новейшего");
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        private static void TestReportNameCollision()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N"));
            try
            {
                var stamp = new DateTime(2026, 7, 16, 10, 0, 0);
                string p1 = ReportWriter.SaveWithRotation(dir, "первый", stamp, 3);
                string p2 = ReportWriter.SaveWithRotation(dir, "второй", stamp, 3);
                AssertTrue(!string.Equals(p1, p2, StringComparison.OrdinalIgnoreCase), "имена различаются");
                AssertEqual(2, Directory.GetFiles(dir, "report_*.txt").Length, "оба сохранены");
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        // ---------- CombineRetryResults ----------

        private static FileResult MakeResult(string path, bool ok, string note)
        {
            var fr = new FileResult();
            fr.FullPath = path;
            fr.FileName = Path.GetFileName(path);
            fr.Ok = ok;
            fr.Note = note;
            if (ok)
                fr.SheetName = Path.GetFileNameWithoutExtension(path);
            return fr;
        }

        private static MergeResult MakePrevious()
        {
            var prev = new MergeResult();
            prev.OutputPath = @"C:\out\Свод.xlsx";
            prev.Files.Add(MakeResult(@"C:\in\А.xlsx", true, null));
            prev.Files.Add(MakeResult(@"C:\in\Б.xlsx", false, "битый"));
            prev.Files.Add(MakeResult(@"C:\in\В.xlsx", false, "пароль"));
            prev.OkCount = 1;
            prev.SkipCount = 2;
            return prev;
        }

        private static void TestCombineRetryReplaces()
        {
            MergeResult prev = MakePrevious();
            var attempts = new List<FileResult> { MakeResult(@"C:\in\Б.xlsx", true, null) };
            MergeResult combined = MergeService.CombineRetryResults(prev, attempts);

            AssertEqual(3, combined.Files.Count, "число записей");
            AssertEqual(2, combined.OkCount, "перенесено после повтора");
            AssertEqual(1, combined.SkipCount, "осталось пропущенных");
            AssertTrue(combined.Files[1].Ok, "Б теперь перенесён");
            AssertEqual(prev.OutputPath, combined.OutputPath, "путь свода");
        }

        private static void TestCombineRetryKeepsFailed()
        {
            MergeResult prev = MakePrevious();
            var attempts = new List<FileResult> { MakeResult(@"C:\in\Б.xlsx", false, "снова битый") };
            MergeResult combined = MergeService.CombineRetryResults(prev, attempts);

            AssertEqual(1, combined.OkCount, "перенесено не изменилось");
            AssertEqual(2, combined.SkipCount, "пропущенных столько же");
            AssertEqual("снова битый", combined.Files[1].Note, "причина обновлена");
        }

        private static void TestCombineRetryOrder()
        {
            MergeResult prev = MakePrevious();
            var attempts = new List<FileResult>
            {
                MakeResult(@"C:\in\В.xlsx", true, null),
                MakeResult(@"C:\in\Б.xlsx", true, null)
            };
            MergeResult combined = MergeService.CombineRetryResults(prev, attempts);

            AssertEqual("А.xlsx|Б.xlsx|В.xlsx",
                combined.Files[0].FileName + "|" + combined.Files[1].FileName + "|" + combined.Files[2].FileName,
                "порядок исходного прогона сохранён");
            AssertEqual(3, combined.OkCount, "все перенесены");
            AssertEqual(0, combined.SkipCount, "пропущенных не осталось");
            AssertTrue(!ReferenceEquals(prev.Files[0], null) && combined.Files[0].Ok, "успешная запись не тронута");
        }

        // ---------- OutputFormats ----------

        private static void TestOutputFormatCodes()
        {
            AssertEqual(51, OutputFormats.FileFormatFor(@"C:\a\Свод.xlsx"), "xlsx");
            AssertEqual(52, OutputFormats.FileFormatFor("Свод.xlsm"), "xlsm");
            AssertEqual(50, OutputFormats.FileFormatFor("Свод.XLSB"), "xlsb в верхнем регистре");
            AssertEqual(56, OutputFormats.FileFormatFor("Свод.xls"), "xls");
            AssertEqual(0, OutputFormats.FileFormatFor("Свод.pdf"), "чужое расширение");
            AssertEqual(0, OutputFormats.FileFormatFor("Свод"), "без расширения");
        }

        private static void TestStripExtension()
        {
            AssertEqual("Свод", OutputFormats.StripKnownExtension("Свод.xlsx"), "xlsx");
            AssertEqual("Свод", OutputFormats.StripKnownExtension("Свод.XLS"), "xls в верхнем регистре");
            AssertEqual("Свод.pdf", OutputFormats.StripKnownExtension("Свод.pdf"), "чужое расширение не трогаем");
            AssertEqual("Свод", OutputFormats.StripKnownExtension("Свод"), "без расширения");
        }

        // ---------- CheckOutputWritable ----------

        private static void TestOutputLocked()
        {
            string path = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N") + ".xlsx");
            using (new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                string error = MergeService.CheckOutputWritable(path);
                AssertTrue(error != null, "занятый файл должен давать ошибку");
                AssertTrue(error.Contains("занят"), "сообщение о занятости: " + error);
            }
            File.Delete(path);
        }

        private static void TestOutputWritable()
        {
            string path = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N") + ".xlsx");

            // Несуществующий файл: проверка не должна оставлять след
            AssertEqual(null, MergeService.CheckOutputWritable(path), "новый файл");
            AssertTrue(!File.Exists(path), "пробный файл удалён");

            // Существующий свободный файл
            File.WriteAllText(path, "x");
            try
            {
                AssertEqual(null, MergeService.CheckOutputWritable(path), "свободный файл");
            }
            finally
            {
                File.Delete(path);
            }
        }

        private static void TestOutputBadFolder()
        {
            string path = Path.Combine(Path.GetTempPath(),
                "ExcelMergerTests_нет_такой_" + Guid.NewGuid().ToString("N"), "Свод.xlsx");
            string error = MergeService.CheckOutputWritable(path);
            AssertTrue(error != null && error.Contains("не существует"), "папка не существует: " + error);
        }

        // ---------- мини-раннер ----------

        private static void Run(string name, Action test)
        {
            try
            {
                test();
                _passed++;
                Console.WriteLine("PASS  " + name);
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("FAIL  " + name + " — " + ex.Message);
            }
        }

        private static void AssertEqual(object expected, object actual, string what)
        {
            if (!Equals(expected, actual))
                throw new Exception(what + ": ожидалось «" + expected + "», получено «" + actual + "»");
        }

        private static void AssertTrue(bool condition, string what)
        {
            if (!condition)
                throw new Exception(what);
        }
    }
}
