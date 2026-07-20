using System;
using System.Collections.Generic;
using System.IO;
using ExcelMerger;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

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
            Run("ToolRegistry: открыт/закрыт, авто-удаление Disposed", TestToolRegistry);
            Run("HelpMenu: структура «Справка» и вставка доп. пунктов", TestHelpMenu);
            Run("ThumbZoom: кламп ширины плитки", TestThumbZoomClamp);
            Run("ThumbZoom: колесо меняет масштаб и упирается в границы", TestThumbZoomWheel);
            Run("ThumbZoom: высота плитки пропорциональна", TestThumbZoomTile);
            Run("ListReorder: MoveUp/Down границы и обмен", TestListReorderMoves);
            Run("ListReorder: Move (перенос) и RemoveAt (набор)", TestListReorderMoveRemove);
            Run("SourceFileList: порядок, включение, IncludedInOrder", TestSourceFileList);
            Run("SourceFileList: сортировка по имени как в Проводнике", TestSourceFileSort);
            Run("PrepareSourceList: исключение свода и дубликатов", TestPrepareSourceList);
            Run("FileSignature: ZIP/OLE2 — книга, текст/пусто — нет", TestFileSignature);
            Run("LowSpaceMessage: мало места — понятная ошибка, иначе null", TestLowSpaceMessage);
            Run("PageRanges.Parse: диапазоны, пробелы, открытый конец", TestPageRangesParse);
            Run("PageRanges.Parse: неверный ввод — ошибка", TestPageRangesParseErrors);
            Run("PageRanges.EveryN: нарезка на равные части", TestPageRangesEveryN);
            Run("PageRanges.ToIndices: диапазоны -> индексы (порядок, повторы)", TestPageRangesToIndices);
            Run("UpdateChecker: разбор тега и сравнение версий", TestUpdateChecker);
            Run("UsageStats.ShouldAutoClear: период очистки", TestShouldAutoClear);
            Run("MessageForm.ButtonX: одна по центру, две по краям", TestMessageButtonX);
            Run("PdfSplitService.Sanitize: недопустимые символы", TestSanitize);
            Run("PdfSplit (живой): извлечение, диапазоны, каждые N, закладки", TestPdfSplitLive);
            Run("PdfPageGrid.ClampWindow: окно видимых с буфером", TestClampWindow);
            Run("Theme.ToBgr: упаковка цвета 0x00BBGGRR", TestThemeToBgr);
            Run("TocBuilder.SheetRef: ссылка на A1, апострофы удвоены", TestSheetRef);
            Run("WindowChrome: COLORREF упакован как 0x00BBGGRR", TestWindowChromeColorRef);
            Run("HeaderBand: строится с заголовком, двойная буферизация", TestHeaderBand);
            Run("HeaderBand.TextRightBound: текст не заходит под кнопку", TestHeaderTextBound);
            Run("MainForm.ClassifyListKey: Alt+↑/↓, Ctrl+C/A, Delete, Enter", TestClassifyListKey);
            Run("PdfMergeForm.ClassifyPageKey: Delete, Alt+←/→, Ctrl+A, Enter", TestClassifyPageKey);
            Run("PdfPageOrder: добавление и границы MoveUp/MoveDown", TestPdfOrderMoves);
            Run("PdfPageOrder: перенос drag&drop в обе стороны", TestPdfOrderDragMove);
            Run("PdfPageOrder: удаление набора строк", TestPdfOrderRemove);
            Run("NoteText: период, счётчики, файл свода", TestNoteBasics);
            Run("NoteText: таблица пропущенных", TestNoteSkippedTable);
            Run("NoteText: без пропусков — «замечания отсутствуют»", TestNoteClean);
            Run("SheetBaseName: первый лист vs все листы", TestSheetBaseName);
            Run("MergeResult.FileCount: файлы, а не листы", TestFileCount);
            Run("Retry: один пропущенный файл -> несколько листов", TestCombineRetryMultiSheet);
            Run("Retry: пропущенные заменяются свежими результатами", TestCombineRetryReplaces);
            Run("Retry: неудачный повтор обновляет причину", TestCombineRetryKeepsFailed);
            Run("Retry: порядок и успешные записи не меняются", TestCombineRetryOrder);
            Run("OutputFormats: код формата по расширению", TestOutputFormatCodes);
            Run("OutputFormats: срез введённого расширения", TestStripExtension);
            Run("CheckOutputWritable: занятый файл распознан", TestOutputLocked);
            Run("CheckOutputWritable: свободный и новый файлы", TestOutputWritable);
            Run("CheckOutputWritable: несуществующая папка", TestOutputBadFolder);
            Run("PdfCompression.Preset: уровень -> пресет Ghostscript", TestCompressionPreset);
            Run("PdfCompression.BuildArguments: кавычки, пресет, -I для бандла", TestCompressionArgs);
            Run("PdfCompression.ShouldReplace: только валидный и строго меньше", TestCompressionShouldReplace);
            Run("Ghostscript.PickFirstExisting: первый существующий из кандидатов", TestGhostscriptPick);
            Run("PdfCompression (живой): крупный PDF сжимается, страницы целы", TestCompressLive);
            Run("JustifiedLabel.Wrap: перенос слов по ширине", TestJustifyWrap);
            Run("AboutForm: реквизиты доната валидны (20 цифр, банк не пуст)", TestDonationRequisites);
            Run("LruCache: вытеснение самого несвежего, Count, порядок", TestLruEviction);
            Run("LruCache: touch через TryGet переносит вытеснение", TestLruTouchOnGet);
            Run("LruCache: замена ключа не растит размер и не вытесняет", TestLruReplace);
            Run("LruCache: ключи регистронезависимы (пути файлов)", TestLruCaseInsensitive);
            Run("LruCache: Clear освобождает все элементы", TestLruClear);
            Run("LruCache: ёмкость < 1 запрещена", TestLruCapacityGuard);
            Run("PdfPageGrid.BuildKeySet: ключи набора без дублей, null -> пусто", TestGridBuildKeySet);
            Run("PdfPageGrid.StaleKeys: вытесняются только отсутствующие в keep", TestGridStaleKeys);

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

        // ---------- ToolRegistry ----------

        private static void TestToolRegistry()
        {
            var reg = new ToolRegistry();
            System.Windows.Forms.Form f;
            AssertTrue(!reg.TryGetOpen("a", out f), "пустой реестр — не открыт");

            var form = new System.Windows.Forms.Form();
            reg.Add("a", form);
            AssertTrue(reg.TryGetOpen("a", out f) && ReferenceEquals(f, form), "добавленный — открыт");
            AssertEqual(1, reg.OpenForms().Count, "одно живое окно");

            form.Dispose();
            AssertTrue(!reg.TryGetOpen("a", out f), "после Dispose — не открыт");
            AssertEqual(0, reg.OpenForms().Count, "закрытые не считаются");

            var f2 = new System.Windows.Forms.Form();
            reg.Add("b", f2);
            reg.Remove("b");
            AssertTrue(!reg.TryGetOpen("b", out f), "после Remove — не открыт");
            f2.Dispose();
        }

        // ---------- HelpMenu (общее меню «Справка») ----------

        private static void TestHelpMenu()
        {
            var extra = new System.Windows.Forms.ToolStripMenuItem("Папка отчётов");
            using (System.Windows.Forms.MenuStrip menu = HelpMenu.Create(null, delegate { }, extra))
            {
                AssertEqual(1, menu.Items.Count, "один пункт строки меню");
                var help = (System.Windows.Forms.ToolStripMenuItem)menu.Items[0];
                AssertEqual("Справка", help.Text, "название пункта");

                var texts = new List<string>();
                foreach (System.Windows.Forms.ToolStripItem it in help.DropDownItems)
                    if (it is System.Windows.Forms.ToolStripMenuItem)
                        texts.Add(it.Text);
                AssertTrue(texts.Contains("Как пользоваться"), "есть «Как пользоваться»");
                AssertTrue(texts.Contains("Статистика"), "есть «Статистика»");
                AssertTrue(texts.Contains("Папка отчётов"), "доп. пункт вставлен");
                // «О программе» перенесена на стартовый экран — в меню её быть не должно.
                AssertTrue(!texts.Contains("О программе"), "«О программе» убрана из меню");
            }

            // Без доп. пунктов: «Как пользоваться», «Статистика».
            using (System.Windows.Forms.MenuStrip menu = HelpMenu.Create(null, delegate { }))
            {
                var help = (System.Windows.Forms.ToolStripMenuItem)menu.Items[0];
                int menuItems = 0;
                foreach (System.Windows.Forms.ToolStripItem it in help.DropDownItems)
                    if (it is System.Windows.Forms.ToolStripMenuItem)
                        menuItems++;
                AssertEqual(2, menuItems, "без extras — два пункта");
            }
        }

        // ---------- ThumbZoom ----------

        private static void TestThumbZoomClamp()
        {
            AssertEqual(ThumbZoom.MinWidth, ThumbZoom.Clamp(ThumbZoom.MinWidth - 50), "ниже минимума");
            AssertEqual(ThumbZoom.MaxWidth, ThumbZoom.Clamp(ThumbZoom.MaxWidth + 50), "выше максимума");
            AssertEqual(150, ThumbZoom.Clamp(150), "в диапазоне");
        }

        private static void TestThumbZoomWheel()
        {
            int up = ThumbZoom.StepFromWheel(132, 120);   // один щелчок вверх
            int down = ThumbZoom.StepFromWheel(132, -120); // один щелчок вниз
            AssertTrue(up > 132, "колесо вверх увеличивает: " + up);
            AssertTrue(down < 132, "колесо вниз уменьшает: " + down);
            AssertEqual(ThumbZoom.MaxWidth, ThumbZoom.StepFromWheel(ThumbZoom.MaxWidth, 1200), "не выше максимума");
            AssertEqual(ThumbZoom.MinWidth, ThumbZoom.StepFromWheel(ThumbZoom.MinWidth, -1200), "не ниже минимума");
        }

        private static void TestThumbZoomTile()
        {
            System.Drawing.Size s = ThumbZoom.TileSize(160);
            AssertEqual(160, s.Width, "ширина");
            AssertTrue(s.Height > s.Width, "высота больше ширины (портрет): " + s.Height);

            // Регрессия: плитка не должна превышать лимит ImageList (256×256)
            // даже на максимуме масштаба — иначе WinForms бросает исключение.
            System.Drawing.Size max = ThumbZoom.TileSize(ThumbZoom.MaxWidth);
            AssertTrue(max.Width <= 256 && max.Height <= 256,
                "плитка в пределах 256: " + max.Width + "x" + max.Height);
            System.Drawing.Size over = ThumbZoom.TileSize(10000);
            AssertTrue(over.Width <= 256 && over.Height <= 256, "кламп сверх лимита");
        }

        // ---------- ListReorder ----------

        private static void TestListReorderMoves()
        {
            var l = new List<string> { "a", "b", "c" };
            AssertEqual(0, ListReorder.MoveUp(l, 0), "верхний вверх — на месте");
            AssertEqual(2, ListReorder.MoveDown(l, 2), "нижний вниз — на месте");
            AssertEqual(0, ListReorder.MoveUp(l, 1), "b вверх -> индекс 0");
            AssertEqual("b|a|c", string.Join("|", l.ToArray()), "после MoveUp");
            AssertEqual(1, ListReorder.MoveDown(l, 0), "b вниз -> индекс 1");
            AssertEqual("a|b|c", string.Join("|", l.ToArray()), "MoveDown вернул порядок");
        }

        private static void TestListReorderMoveRemove()
        {
            var l = new List<string> { "a", "b", "c", "d" };
            ListReorder.Move(l, 0, 3);
            AssertEqual("b|c|a|d", string.Join("|", l.ToArray()), "перенос a перед позицией 3");
            ListReorder.Move(l, 2, 0);
            AssertEqual("a|b|c|d", string.Join("|", l.ToArray()), "перенос обратно");
            ListReorder.RemoveAt(l, new[] { 3, 1 });
            AssertEqual("a|c", string.Join("|", l.ToArray()), "удаление набора индексов");
        }

        // ---------- SourceFileList ----------

        private static string IncludedSig(SourceFileList list)
        {
            return string.Join("|", list.IncludedInOrder().ConvertAll(System.IO.Path.GetFileName).ToArray());
        }

        private static void TestSourceFileList()
        {
            var list = new SourceFileList();
            list.SetFiles(new[] { @"C:\in\А.xlsx", @"C:\in\Б.xlsx", @"C:\in\В.xlsx" });
            AssertEqual(3, list.Count, "три файла");
            AssertEqual(3, list.IncludedCount, "все включены");

            list[1].Include = false; // исключаем Б
            AssertEqual(2, list.IncludedCount, "два включённых");
            AssertEqual("А.xlsx|В.xlsx", IncludedSig(list), "исключённый не в списке");

            // Перестановка по позициям всего списка [А, Б, В] -> [В, А, Б]
            list.MoveUp(2); // [А, В, Б]
            list.MoveUp(1); // [В, А, Б]
            AssertEqual("В.xlsx|А.xlsx", IncludedSig(list), "порядок среди включённых изменился");

            list.SetAllIncluded(true);
            AssertEqual(3, list.IncludedCount, "все снова включены");
        }

        private static void TestSourceFileSort()
        {
            var list = new SourceFileList();
            list.SetFiles(new[] { @"C:\in\Отчет 10.xlsx", @"C:\in\Отчет 2.xlsx", @"C:\in\Отчет 1.xlsx" });
            list.SortByName();
            AssertEqual("Отчет 1.xlsx|Отчет 2.xlsx|Отчет 10.xlsx", IncludedSig(list), "естественный порядок");
        }

        // ---------- PrepareSourceList ----------

        private static void TestPrepareSourceList()
        {
            var files = new List<string>
            {
                @"C:\in\А.xlsx", @"C:\in\Свод.xlsx", @"C:\in\А.xlsx", @"C:\in\Б.xlsx"
            };
            List<string> prepared = MergeService.PrepareSourceList(files, @"C:\in\Свод.xlsx");
            var names = prepared.ConvertAll(System.IO.Path.GetFileName);
            AssertEqual("А.xlsx|Б.xlsx", string.Join("|", names.ToArray()),
                "свод исключён, дубликат убран, порядок сохранён");
            AssertEqual(0, MergeService.PrepareSourceList(null, "x").Count, "null -> пусто");
        }

        // ---------- FileSignature ----------

        private static void TestFileSignature()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string txt = Path.Combine(dir, "битый.xlsx");
                File.WriteAllText(txt, "this is not an excel file");
                AssertEqual(ExcelContainer.NotExcel, FileSignature.Detect(txt), "текст с расширением .xlsx");

                string empty = Path.Combine(dir, "пустой.xlsx");
                File.WriteAllBytes(empty, new byte[0]);
                AssertEqual(ExcelContainer.NotExcel, FileSignature.Detect(empty), "пустой файл");

                string zip = Path.Combine(dir, "ooxml.xlsx");
                File.WriteAllBytes(zip, new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00 });
                AssertEqual(ExcelContainer.Zip, FileSignature.Detect(zip), "ZIP-сигнатура PK — OOXML");

                string ole = Path.Combine(dir, "запарол.xlsx");
                File.WriteAllBytes(ole, new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 });
                AssertEqual(ExcelContainer.Ole2, FileSignature.Detect(ole), "OLE2/CFB — xls или шифр");

                AssertEqual(ExcelContainer.Unreadable, FileSignature.Detect(Path.Combine(dir, "нет.xlsx")),
                    "отсутствующий файл — решает Excel");
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        private static void TestLowSpaceMessage()
        {
            AssertEqual(null, MergeService.LowSpaceMessage(@"C:\", 5L * 1024 * 1024 * 1024), "5 ГБ — достаточно");
            AssertEqual(null, MergeService.LowSpaceMessage(@"C:\", 200L * 1024 * 1024), "ровно порог 200 МБ — достаточно");
            string low = MergeService.LowSpaceMessage(@"C:\", 10L * 1024 * 1024);
            AssertTrue(low != null && low.Contains(@"C:\") && low.Contains("10 МБ"), "мало места: " + low);
            string zero = MergeService.LowSpaceMessage(@"C:\", 0);
            AssertTrue(zero != null && zero.Contains("0 МБ"), "ноль байт: " + zero);
        }

        // ---------- WindowChrome / HeaderBand ----------

        private static void TestHeaderTextBound()
        {
            // Нет дочерних контролов — граница у правого края с отступом.
            AssertEqual(760, HeaderBand.TextRightBound(780, int.MaxValue), "без кнопки — правый край - 20");
            // Кнопка слева от края — текст обрезается до её левой границы минус зазор.
            AssertEqual(588, HeaderBand.TextRightBound(780, 600), "с кнопкой — левее её на 12");
            AssertEqual(488, HeaderBand.TextRightBound(700, 500), "узкое окно — левее кнопки");
            // Кнопка правее правого поля — не выходим за край панели.
            AssertEqual(760, HeaderBand.TextRightBound(780, 900), "кнопка за краем — ограничены полем");
        }

        private static void TestClassifyListKey()
        {
            var Alt = System.Windows.Forms.Keys.Alt;
            var Ctrl = System.Windows.Forms.Keys.Control;
            AssertEqual(MainForm.ListKeyAction.MoveUp, MainForm.ClassifyListKey(Alt | System.Windows.Forms.Keys.Up), "Alt+Up");
            AssertEqual(MainForm.ListKeyAction.MoveDown, MainForm.ClassifyListKey(Alt | System.Windows.Forms.Keys.Down), "Alt+Down");
            AssertEqual(MainForm.ListKeyAction.Copy, MainForm.ClassifyListKey(Ctrl | System.Windows.Forms.Keys.C), "Ctrl+C — копировать");
            AssertEqual(MainForm.ListKeyAction.SelectAll, MainForm.ClassifyListKey(Ctrl | System.Windows.Forms.Keys.A), "Ctrl+A — выделить всё");
            AssertEqual(MainForm.ListKeyAction.Exclude, MainForm.ClassifyListKey(System.Windows.Forms.Keys.Delete), "Delete — исключить");
            AssertEqual(MainForm.ListKeyAction.Swallow, MainForm.ClassifyListKey(System.Windows.Forms.Keys.Enter), "Enter — не сливать");
            AssertEqual(MainForm.ListKeyAction.None, MainForm.ClassifyListKey(System.Windows.Forms.Keys.Up), "просто ↑ — навигация");
        }

        private static void TestClassifyPageKey()
        {
            var Alt = System.Windows.Forms.Keys.Alt;
            var Ctrl = System.Windows.Forms.Keys.Control;
            AssertEqual(PdfMergeForm.PageKeyAction.Remove, PdfMergeForm.ClassifyPageKey(System.Windows.Forms.Keys.Delete), "Delete — удалить");
            AssertEqual(PdfMergeForm.PageKeyAction.MoveEarlier, PdfMergeForm.ClassifyPageKey(Alt | System.Windows.Forms.Keys.Left), "Alt+← — раньше");
            AssertEqual(PdfMergeForm.PageKeyAction.MoveLater, PdfMergeForm.ClassifyPageKey(Alt | System.Windows.Forms.Keys.Right), "Alt+→ — позже");
            AssertEqual(PdfMergeForm.PageKeyAction.SelectAll, PdfMergeForm.ClassifyPageKey(Ctrl | System.Windows.Forms.Keys.A), "Ctrl+A — выделить всё");
            AssertEqual(PdfMergeForm.PageKeyAction.Swallow, PdfMergeForm.ClassifyPageKey(System.Windows.Forms.Keys.Enter), "Enter — не сохранять");
            AssertEqual(PdfMergeForm.PageKeyAction.None, PdfMergeForm.ClassifyPageKey(System.Windows.Forms.Keys.Left), "просто ← — навигация");
        }

        private static string RangeSig(System.Collections.Generic.List<PageRange> ranges)
        {
            var parts = new System.Collections.Generic.List<string>();
            foreach (PageRange r in ranges)
                parts.Add(r.Start + "-" + r.End);
            return string.Join("|", parts.ToArray());
        }

        private static void TestPageRangesParse()
        {
            AssertEqual("0-2|4-4|7-9", RangeSig(PageRanges.Parse("1-3, 5, 8-", 10)), "1-3,5,8- при 10");
            AssertEqual("1-1", RangeSig(PageRanges.Parse("2", 5)), "одна страница");
            AssertEqual("0-2", RangeSig(PageRanges.Parse("-3", 10)), "открытое начало");
            AssertEqual("0-1|3-3", RangeSig(PageRanges.Parse("  1 - 2 , 4 ", 5)), "пробелы");
            AssertEqual("5-8", PageRanges.Parse("5-", 8)[0].Label, "открытый конец 5.. -> 5-8");
            AssertEqual("8", PageRanges.Parse("8-", 8)[0].Label, "последняя страница -> метка 8");
            AssertEqual("1-3", PageRanges.Parse("1-3", 3)[0].Label, "метка диапазона");
        }

        private static void TestPageRangesParseErrors()
        {
            AssertThrows("пусто", delegate { PageRanges.Parse("", 10); });
            AssertThrows("0-3 (ниже 1)", delegate { PageRanges.Parse("0-3", 10); });
            AssertThrows("5-2 (начало>конец)", delegate { PageRanges.Parse("5-2", 10); });
            AssertThrows("3-99 (выше pageCount)", delegate { PageRanges.Parse("3-99", 10); });
            AssertThrows("нечисло", delegate { PageRanges.Parse("abc", 10); });
        }

        private static void TestPageRangesEveryN()
        {
            AssertEqual("0-2|3-5|6-8|9-9", RangeSig(PageRanges.EveryN(10, 3)), "10 по 3");
            AssertEqual("0-1|2-3|4-5", RangeSig(PageRanges.EveryN(6, 2)), "6 по 2");
            AssertEqual("0-0|1-1|2-2", RangeSig(PageRanges.EveryN(3, 1)), "по одной");
            AssertEqual("0-4", RangeSig(PageRanges.EveryN(5, 10)), "n больше всего");
            AssertThrows("n<1", delegate { PageRanges.EveryN(5, 0); });
        }

        private static void TestPageRangesToIndices()
        {
            var idx = PageRanges.ToIndices(PageRanges.Parse("1-3, 5, 8-", 10)); // 0-2,4,7-9
            AssertEqual("0,1,2,4,7,8,9", string.Join(",", idx.ConvertAll(i => i.ToString()).ToArray()), "порядок");
            // Повторы и обратный смысл ввода сохраняются: «1-4, 1» -> ...,0
            var dup = PageRanges.ToIndices(PageRanges.Parse("1-4, 1", 10));
            AssertEqual("0,1,2,3,0", string.Join(",", dup.ConvertAll(i => i.ToString()).ToArray()), "повторы сохраняются");
        }

        private static void TestUpdateChecker()
        {
            AssertEqual(new Version(1, 11, 2), UpdateChecker.ParseTag("v1.11.2"), "тег с v");
            AssertEqual(new Version(1, 12, 0), UpdateChecker.ParseTag("1.12.0"), "тег без v");
            AssertEqual(null, UpdateChecker.ParseTag("release"), "мусор -> null");
            AssertTrue(UpdateChecker.IsNewer(new Version(1, 12, 0), new Version(1, 11, 2)), "1.12.0 новее 1.11.2");
            AssertTrue(!UpdateChecker.IsNewer(new Version(1, 11, 2), new Version(1, 11, 2)), "равные — не новее");
            AssertTrue(!UpdateChecker.IsNewer(new Version(1, 11, 0), new Version(1, 11, 2)), "старее — не новее");
            AssertTrue(!UpdateChecker.IsNewer(null, new Version(1, 0, 0)), "null latest — не новее");
        }

        private static void TestShouldAutoClear()
        {
            var now = new DateTime(2026, 7, 20, 12, 0, 0, DateTimeKind.Utc);
            AssertTrue(!UsageStats.ShouldAutoClear(now.AddDays(-3), now, 0), "период 0 — никогда");
            AssertTrue(!UsageStats.ShouldAutoClear(now.AddHours(-5), now, 1), "меньше суток — рано");
            AssertTrue(UsageStats.ShouldAutoClear(now.AddDays(-1), now, 1), "сутки прошли — пора");
            AssertTrue(!UsageStats.ShouldAutoClear(now.AddDays(-6), now, 7), "6 из 7 дней — рано");
            AssertTrue(UsageStats.ShouldAutoClear(now.AddDays(-8), now, 7), "8 дней при периоде 7 — пора");
            AssertTrue(UsageStats.ShouldAutoClear(now.AddDays(-31), now, 30), "31 день при 30 — пора");
        }

        private static void TestMessageButtonX()
        {
            AssertEqual(164, MessageForm.ButtonX(0, 1, 440, 112, 20), "одна кнопка — по центру");
            AssertEqual(20, MessageForm.ButtonX(0, 2, 440, 112, 20), "две: первая слева");
            AssertEqual(308, MessageForm.ButtonX(1, 2, 440, 112, 20), "две: вторая справа");
        }

        private static void TestSanitize()
        {
            AssertEqual("Глава 1", PdfSplitService.Sanitize("Глава 1"), "обычный заголовок");
            AssertEqual("a_b_c", PdfSplitService.Sanitize("a/b:c"), "недопустимые символы");
            AssertEqual("без_имени", PdfSplitService.Sanitize(""), "пустое имя");
            AssertEqual("без_имени", PdfSplitService.Sanitize("   "), "только пробелы");
        }

        private static int PdfPageCount(string path)
        {
            using (PdfDocument d = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                return d.PageCount;
        }

        private static string PageCounts(List<string> files)
        {
            var parts = new List<string>();
            foreach (string f in files)
                parts.Add(PdfPageCount(f).ToString());
            return string.Join(",", parts.ToArray());
        }

        private static void TestPdfSplitLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = Path.Combine(dir, "исходник.pdf");
                using (var doc = new PdfDocument())
                {
                    for (int i = 0; i < 10; i++)
                        doc.AddPage();
                    doc.Outlines.Add("Раздел А", doc.Pages[0]);
                    doc.Outlines.Add("Раздел Б", doc.Pages[4]);
                    doc.Outlines.Add("Раздел В", doc.Pages[7]);
                    doc.Save(src);
                }

                // Извлечь выбранные [0,2,4] → один файл, 3 страницы
                string extract = Path.Combine(dir, "выбранные.pdf");
                PdfSplitService.Extract(src, new List<int> { 0, 2, 4 }, extract);
                AssertEqual(3, PdfPageCount(extract), "извлечено 3 страницы");

                // По диапазонам 1-3,5,8- → 3 файла (3,1,3)
                List<string> ranges = PdfSplitService.SplitByRanges(src, PageRanges.Parse("1-3,5,8-", 10), dir, "диап");
                AssertEqual(3, ranges.Count, "диапазонов — 3 файла");
                AssertEqual("3,1,3", PageCounts(ranges), "страниц по диапазонам");
                AssertEqual("диап_1-3.pdf", Path.GetFileName(ranges[0]), "имя первого диапазона");

                // Каждые 3 страницы → 4 файла (3,3,3,1)
                List<string> everyN = PdfSplitService.SplitEveryN(src, 3, dir, "часть");
                AssertEqual(4, everyN.Count, "каждые 3 — 4 файла");
                AssertEqual("3,3,3,1", PageCounts(everyN), "страниц по частям");

                // По закладкам → 3 файла (4,3,3), имена с заголовками
                List<string> byMark = PdfSplitService.SplitByBookmarks(src, dir, "закл");
                AssertEqual(3, byMark.Count, "закладок — 3 файла");
                AssertEqual("4,3,3", PageCounts(byMark), "страниц по закладкам");
                AssertTrue(Path.GetFileName(byMark[0]).Contains("Раздел А"), "имя по закладке: " + Path.GetFileName(byMark[0]));
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        private static void TestCompressLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = Path.Combine(dir, "крупный.pdf");
                MakeImagePdf(src, 2);
                long before = new FileInfo(src).Length;
                int pagesBefore = PdfPageCount(src);
                AssertTrue(before > 300 * 1024, "исходник с изображением должен быть крупным: " + before);

                if (Ghostscript.Available)
                {
                    // «Нормально» (/screen, 72 DPI) — понижает разрешение изображения → размер падает.
                    bool applied = PdfCompression.Compress(src, CompressionLevel.Small);
                    long after = new FileInfo(src).Length;
                    AssertTrue(applied, "сжатие применено (GS есть)");
                    AssertTrue(after < before, "размер уменьшился: " + before + " -> " + after);
                    AssertTrue(PdfCompression.LooksLikePdf(src), "результат — валидный PDF");
                    // Страницы сохранены (это НЕ растр всего документа, а downsampling — структура цела).
                    AssertEqual(pagesBefore, PdfPageCount(src), "число страниц сохранено");
                }
                else
                {
                    // Без Ghostscript сжатие — безопасный no-op: файл не тронут.
                    bool applied = PdfCompression.Compress(src, CompressionLevel.Small);
                    AssertTrue(!applied, "без Ghostscript — без изменений");
                    AssertEqual(before, new FileInfo(src).Length, "файл не тронут без GS");
                }
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        /// <summary>PDF с крупным «шумовым» изображением (плохо сжимается) и текстом на каждой странице.</summary>
        private static void MakeImagePdf(string path, int pages)
        {
            string jpg = Path.Combine(Path.GetDirectoryName(path), "noise.jpg");
            const int side = 1800; // ~245 DPI на A4 → /screen (72) заведомо понизит
            using (var bmp = new System.Drawing.Bitmap(side, side, System.Drawing.Imaging.PixelFormat.Format24bppRgb))
            {
                System.Drawing.Imaging.BitmapData bd = bmp.LockBits(
                    new System.Drawing.Rectangle(0, 0, side, side),
                    System.Drawing.Imaging.ImageLockMode.WriteOnly,
                    System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                int total = Math.Abs(bd.Stride) * side;
                var buf = new byte[total];
                new Random(12345).NextBytes(buf); // шум — JPEG почти не сжимает, исходник тяжёлый
                System.Runtime.InteropServices.Marshal.Copy(buf, 0, bd.Scan0, total);
                bmp.UnlockBits(bd);
                bmp.Save(jpg, JpegEncoder(), JpegQuality(90));
            }
            try
            {
                using (var doc = new PdfDocument())
                using (XImage img = XImage.FromFile(jpg))
                {
                    var font = new XFont("Arial", 14);
                    for (int i = 0; i < pages; i++)
                    {
                        PdfPage page = doc.AddPage();
                        page.Width = 595;
                        page.Height = 842;
                        using (XGraphics gfx = XGraphics.FromPdfPage(page))
                        {
                            gfx.DrawImage(img, 0, 0, page.Width, page.Height);
                            gfx.DrawString("Страница " + (i + 1) + " — текст должен сохраниться.",
                                font, XBrushes.White, new XPoint(30, 30));
                        }
                    }
                    doc.Save(path);
                }
            }
            finally
            {
                try { File.Delete(jpg); } catch { }
            }
        }

        private static System.Drawing.Imaging.ImageCodecInfo JpegEncoder()
        {
            foreach (System.Drawing.Imaging.ImageCodecInfo c in System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders())
                if (c.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid)
                    return c;
            return null;
        }

        private static System.Drawing.Imaging.EncoderParameters JpegQuality(long quality)
        {
            var p = new System.Drawing.Imaging.EncoderParameters(1);
            p.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                System.Drawing.Imaging.Encoder.Quality, quality);
            return p;
        }

        private static void TestJustifyWrap()
        {
            Func<string, int> m = delegate(string s) { return s.Length * 10; }; // 10px/символ
            const int sp = 5;
            var lines = JustifiedLabel.Wrap("aa bb cc dd", 60, m, sp);
            AssertEqual(2, lines.Count, "две строки");
            AssertEqual("aa bb", string.Join(" ", lines[0].ToArray()), "строка 1");
            AssertEqual("cc dd", string.Join(" ", lines[1].ToArray()), "строка 2");
            AssertEqual(0, JustifiedLabel.Wrap("", 100, m, sp).Count, "пустой текст — 0 строк");
            AssertEqual(0, JustifiedLabel.Wrap("aa", 0, m, sp).Count, "нулевая ширина — 0 строк");
            // Слово шире строки не роняет разбивку — становится отдельной строкой.
            AssertEqual(1, JustifiedLabel.Wrap("aaaaaaaa", 30, m, sp).Count, "длинное слово — своя строка");
        }

        private static void TestDonationRequisites()
        {
            AssertEqual(20, AboutForm.DonationAccount.Length, "счёт — 20 цифр");
            foreach (char c in AboutForm.DonationAccount)
                AssertTrue(c >= '0' && c <= '9', "в счёте только цифры");
            AssertTrue(AboutForm.DonationBank.Length > 0, "банк не пуст");
        }

        private static void TestClampWindow()
        {
            int lo, hi;
            PdfPageGrid.ClampWindow(0, 10, 100, 16, out lo, out hi);
            AssertEqual("0,26", lo + "," + hi, "у начала — не ниже 0");
            PdfPageGrid.ClampWindow(50, 60, 100, 16, out lo, out hi);
            AssertEqual("34,76", lo + "," + hi, "в середине — буфер с обеих сторон");
            PdfPageGrid.ClampWindow(90, 99, 100, 16, out lo, out hi);
            AssertEqual("74,99", lo + "," + hi, "у конца — не выше count-1");
            PdfPageGrid.ClampWindow(0, 5, 6, 16, out lo, out hi);
            AssertEqual("0,5", lo + "," + hi, "мало элементов — весь список");
        }

        private static void TestThemeToBgr()
        {
            AssertEqual(12413967, Theme.ToBgr(System.Drawing.Color.FromArgb(15, 108, 189)), "HubBlue #0F6CBD");
            AssertEqual(4291600, Theme.ToBgr(System.Drawing.Color.FromArgb(16, 124, 65)), "Accent #107C41");
            AssertEqual(16777215, Theme.ToBgr(System.Drawing.Color.White), "белый");
            AssertEqual(255, Theme.ToBgr(System.Drawing.Color.FromArgb(255, 0, 0)), "красный — младший байт");
        }

        private static void TestSheetRef()
        {
            AssertEqual("'Отчет'!A1", TocBuilder.SheetRef("Отчет"), "обычное имя");
            AssertEqual("'it''s'!A1", TocBuilder.SheetRef("it's"), "апостроф удваивается");
            AssertEqual("''!A1", TocBuilder.SheetRef(null), "null — пустое имя");
        }

        private static void TestWindowChromeColorRef()
        {
            // COLORREF кладёт R в младший байт, B в старший (0x00BBGGRR).
            int packed = WindowChrome.ColorRef(System.Drawing.Color.FromArgb(16, 124, 65));
            AssertEqual(0x00417C10, packed, "упаковка акцентного цвета");
            AssertEqual(0x00FFFFFF, WindowChrome.ColorRef(System.Drawing.Color.White), "белый");
            AssertEqual(0x00000000, WindowChrome.ColorRef(System.Drawing.Color.Black), "чёрный");
            AssertEqual(0x000000FF, WindowChrome.ColorRef(System.Drawing.Color.FromArgb(255, 0, 0)), "красный -> младший байт");
        }

        private static void TestHeaderBand()
        {
            using (var band = new HeaderBand("Заголовок", "подпись", Theme.Accent, Theme.AccentPressed))
            {
                AssertTrue(band is System.Windows.Forms.Control, "это контрол");
                band.Width = 400;
                band.Height = 80;
                AssertEqual(400, band.Width, "ширина применяется");
                // Пустая подпись не должна ронять отрисовку логики (конструктор допускает null).
                using (var noSub = new HeaderBand("Только заголовок", null, Theme.PdfRed, Theme.PdfRedDark))
                    AssertTrue(noSub != null, "подпись null допустима");
            }
        }

        // ---------- PdfPageOrder ----------

        private static string OrderSignature(PdfPageOrder order)
        {
            var parts = new List<string>();
            for (int i = 0; i < order.Count; i++)
                parts.Add(order[i].FileName + ":" + (order[i].PageIndex + 1));
            return string.Join("|", parts.ToArray());
        }

        private static PdfPageOrder MakeOrder()
        {
            var order = new PdfPageOrder();
            order.AddDocument(@"C:\in\А.pdf", 2);
            order.AddDocument(@"C:\in\Б.pdf", 1);
            return order; // А:1 | А:2 | Б:1
        }

        private static void TestPdfOrderMoves()
        {
            PdfPageOrder order = MakeOrder();
            AssertEqual("А.pdf:1|А.pdf:2|Б.pdf:1", OrderSignature(order), "исходный порядок");

            AssertEqual(0, order.MoveUp(0), "MoveUp с верхней строки — на месте");
            AssertEqual(2, order.MoveDown(2), "MoveDown с нижней строки — на месте");

            AssertEqual(1, order.MoveUp(2), "MoveUp возвращает новый индекс");
            AssertEqual("А.pdf:1|Б.pdf:1|А.pdf:2", OrderSignature(order), "после MoveUp");
            AssertEqual(2, order.MoveDown(1), "MoveDown возвращает новый индекс");
            AssertEqual("А.pdf:1|А.pdf:2|Б.pdf:1", OrderSignature(order), "MoveDown вернул порядок");
        }

        private static void TestPdfOrderDragMove()
        {
            PdfPageOrder order = MakeOrder();
            order.Move(2, 0); // Б:1 в начало
            AssertEqual("Б.pdf:1|А.pdf:1|А.pdf:2", OrderSignature(order), "перенос вверх");
            order.Move(0, 3); // Б:1 в конец (вставка перед позицией 3)
            AssertEqual("А.pdf:1|А.pdf:2|Б.pdf:1", OrderSignature(order), "перенос вниз");
            order.Move(1, 1); // на себя — без изменений
            AssertEqual("А.pdf:1|А.pdf:2|Б.pdf:1", OrderSignature(order), "перенос на себя");
        }

        private static void TestPdfOrderRemove()
        {
            PdfPageOrder order = MakeOrder();
            order.RemoveAt(new[] { 2, 0 }); // произвольный порядок индексов
            AssertEqual("А.pdf:2", OrderSignature(order), "удаление набора");
            AssertEqual(1, order.Count, "осталась одна строка");
        }

        // ---------- NoteText ----------

        private static void TestNoteBasics()
        {
            MergeResult res = MakePrevious(); // 1 ok, 2 skip
            var options = new MergeOptions();
            options.AddToc = true;
            NoteContent note = NoteText.Build(res, @"C:\in", options,
                new DateTime(2026, 7, 16, 14, 5, 0));

            AssertEqual("СПРАВКА", note.Title, "заголовок");
            string all = string.Join("|", note.Body.ToArray()) + "|" + string.Join("|", note.Tail.ToArray());
            AssertTrue(all.Contains("16 июля 2026 г.") && all.Contains("14:05"), "период: " + note.Body[0]);
            AssertTrue(all.Contains("Обработано файлов: 3"), "всего");
            AssertTrue(all.Contains("Включено листов в сводный файл: 1"), "включено");
            AssertTrue(all.Contains("Пропущено файлов: 2"), "пропущено");
            AssertTrue(all.Contains(res.OutputPath) && all.Contains("(XLSX)"), "файл свода и формат");
            AssertTrue(all.Contains("лист «Содержание» — да"), "параметры");
            AssertTrue(note.Signature.Contains("Исполнитель"), "подпись");
        }

        private static void TestNoteSkippedTable()
        {
            MergeResult res = MakePrevious();
            NoteContent note = NoteText.Build(res, @"C:\in", new MergeOptions(), DateTime.MinValue);

            AssertTrue(note.SkippedIntro != null, "есть вводная к таблице");
            AssertEqual(2, note.SkippedRows.Count, "строк в таблице");
            AssertEqual("1", note.SkippedRows[0][0], "нумерация");
            AssertEqual("Б.xlsx", note.SkippedRows[0][1], "имя файла");
            AssertEqual("битый", note.SkippedRows[0][2], "причина");
        }

        private static void TestNoteClean()
        {
            var res = new MergeResult();
            res.OutputPath = @"C:\out\Свод.xlsb";
            res.Files.Add(MakeResult(@"C:\in\А.xlsx", true, null));
            res.OkCount = 1;
            NoteContent note = NoteText.Build(res, @"C:\in", new MergeOptions(), DateTime.MinValue);

            AssertTrue(note.SkippedIntro == null, "таблица не нужна");
            AssertEqual(0, note.SkippedRows.Count, "нет строк");
            string all = string.Join("|", note.Body.ToArray());
            AssertTrue(all.Contains("Замечания отсутствуют"), "формулировка чистого итога");
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

        private static void TestSheetBaseName()
        {
            AssertEqual("Отчет", MergeService.SheetBaseName("Отчет", "Лист1", false), "только первый — имя файла");
            AssertEqual("Отчет · Лист1", MergeService.SheetBaseName("Отчет", "Лист1", true), "все листы — файл · лист");
        }

        private static void TestFileCount()
        {
            var res = new MergeResult();
            res.Files.Add(MakeResult(@"C:\in\А.xlsx", true, null)); // два листа из одного файла
            res.Files.Add(MakeResult(@"C:\in\А.xlsx", true, null));
            res.Files.Add(MakeResult(@"C:\in\Б.xlsx", true, null));
            AssertEqual(3, res.Files.Count, "листов (строк) — три");
            AssertEqual(2, res.FileCount, "файлов — два");
        }

        private static void TestCombineRetryMultiSheet()
        {
            // Прошлый прогон: А перенесён (1 лист), Б пропущен.
            var prev = new MergeResult();
            prev.OutputPath = @"C:\out\Свод.xlsx";
            prev.Files.Add(MakeResult(@"C:\in\А.xlsx", true, null));
            prev.Files.Add(MakeResult(@"C:\in\Б.xlsx", false, "битый"));
            prev.OkCount = 1;
            prev.SkipCount = 1;

            // Повтор Б в режиме «все листы» дал два листа.
            var b1 = MakeResult(@"C:\in\Б.xlsx", true, null);
            var b2 = MakeResult(@"C:\in\Б.xlsx", true, null);
            var combined = MergeService.CombineRetryResults(prev, new List<FileResult> { b1, b2 });

            AssertEqual(3, combined.Files.Count, "А + два листа Б");
            AssertEqual(3, combined.OkCount, "все перенесены");
            AssertEqual(0, combined.SkipCount, "пропущенных нет");
            AssertEqual(2, combined.FileCount, "файлов — два (А и Б)");
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

        // ---------- PDF-сжатие (Ghostscript) ----------

        private static void TestCompressionPreset()
        {
            AssertEqual("/ebook", PdfCompression.Preset(CompressionLevel.Good), "Хорошо -> /ebook");
            AssertEqual("/screen", PdfCompression.Preset(CompressionLevel.Small), "Нормально -> /screen");
            AssertEqual(null, PdfCompression.Preset(CompressionLevel.None), "Отлично -> без пресета");
        }

        private static void TestCompressionArgs()
        {
            // Путь с пробелами обязан быть в кавычках, иначе GS примет его за два аргумента.
            string args = PdfCompression.BuildArguments(
                @"C:\Users\aid MINFIN\вход.pdf", @"C:\Users\aid MINFIN\выход.pdf",
                CompressionLevel.Good, null);
            AssertTrue(args.Contains("-sDEVICE=pdfwrite"), "устройство pdfwrite");
            AssertTrue(args.Contains("-dCompatibilityLevel=1.4"), "1.4 — читаемо старым PdfSharp");
            AssertTrue(args.Contains("-dPDFSETTINGS=/ebook"), "пресет /ebook");
            AssertTrue(args.Contains("-dSAFER"), "SAFER включён");
            AssertTrue(!args.Contains("-dNOSAFER"), "NOSAFER не должен передаваться");
            AssertTrue(args.Contains("\"C:\\Users\\aid MINFIN\\вход.pdf\""), "вход в кавычках");
            AssertTrue(args.Contains("-sOutputFile=\"C:\\Users\\aid MINFIN\\выход.pdf\""), "выход в кавычках");
            AssertTrue(!args.Contains(" -I "), "системный GS — без -I");

            // Вшитый GS: добавляется -I на lib и Resource\Init.
            string bundled = PdfCompression.BuildArguments("in.pdf", "out.pdf", CompressionLevel.Small, @"C:\app\gs");
            AssertTrue(bundled.Contains("-dPDFSETTINGS=/screen"), "пресет /screen");
            AssertTrue(bundled.Contains("-I \"C:\\app\\gs\\lib\""), "-I lib для бандла");
            AssertTrue(bundled.Contains("-I \"C:\\app\\gs\\Resource\\Init\""), "-I Resource\\Init для бандла");
        }

        private static void TestCompressionShouldReplace()
        {
            AssertTrue(PdfCompression.ShouldReplace(1000, 400, true), "валидный и меньше — заменяем");
            AssertTrue(!PdfCompression.ShouldReplace(1000, 1000, true), "равный размер — оставляем оригинал");
            AssertTrue(!PdfCompression.ShouldReplace(1000, 1500, true), "больше — оставляем оригинал");
            AssertTrue(!PdfCompression.ShouldReplace(1000, 400, false), "невалидный вывод — не заменяем");
            AssertTrue(!PdfCompression.ShouldReplace(1000, 0, true), "пустой вывод — не заменяем");
        }

        private static void TestGhostscriptPick()
        {
            var candidates = new List<string> { null, "", @"C:\нет\gs.exe", @"C:\есть\gs.exe", @"C:\тоже\gs.exe" };
            string picked = Ghostscript.PickFirstExisting(candidates,
                delegate(string p) { return p == @"C:\есть\gs.exe" || p == @"C:\тоже\gs.exe"; });
            AssertEqual(@"C:\есть\gs.exe", picked, "первый существующий, пустые пропущены");
            AssertEqual(null, Ghostscript.PickFirstExisting(candidates, delegate { return false; }), "ни одного — null");
            // Бросающий предикат не должен ронять выбор.
            string safe = Ghostscript.PickFirstExisting(new[] { @"C:\a", @"C:\b" },
                delegate(string p) { if (p == @"C:\a") throw new Exception("bad"); return true; });
            AssertEqual(@"C:\b", safe, "исключение в предикате -> пропуск кандидата");
        }

        // ---------- мини-раннер ----------

        // ---------- LruCache ----------

        private static void TestLruEviction()
        {
            var evicted = new List<string>();
            var cache = new LruCache<string>(2, delegate(string v) { evicted.Add(v); });
            cache.Add("a", "a");
            cache.Add("b", "b");
            AssertEqual(2, cache.Count, "две записи в пределах ёмкости");
            cache.Add("c", "c"); // переполнение: вытесняется самый несвежий — «a»
            AssertEqual(2, cache.Count, "ёмкость соблюдена");
            AssertEqual(1, evicted.Count, "ровно одно вытеснение");
            AssertEqual("a", evicted[0], "вытеснен наименее недавно использованный");
            string val;
            AssertTrue(!cache.TryGet("a", out val), "a вытеснен");
            AssertTrue(cache.TryGet("b", out val) && val == "b", "b остался");
            AssertTrue(cache.TryGet("c", out val) && val == "c", "c добавлен");
        }

        private static void TestLruTouchOnGet()
        {
            var evicted = new List<string>();
            var cache = new LruCache<string>(2, delegate(string v) { evicted.Add(v); });
            cache.Add("a", "a");
            cache.Add("b", "b");
            string val;
            AssertTrue(cache.TryGet("a", out val), "обращение к a делает его свежим");
            cache.Add("c", "c"); // теперь несвежий — «b»
            AssertEqual("b", evicted[0], "touch через TryGet сместил вытеснение на b");
            AssertTrue(cache.TryGet("a", out val), "a сохранён");
        }

        private static void TestLruReplace()
        {
            var evicted = new List<string>();
            var cache = new LruCache<string>(2, delegate(string v) { evicted.Add(v); });
            cache.Add("x", "x1");
            cache.Add("x", "x2"); // тот же ключ — замена, не рост
            AssertEqual(1, cache.Count, "замена ключа не увеличивает размер");
            AssertEqual(0, evicted.Count, "замена ничего не вытесняет");
            string val;
            AssertTrue(cache.TryGet("x", out val) && val == "x2", "значение обновлено");
        }

        private static void TestLruCaseInsensitive()
        {
            var cache = new LruCache<string>(2, null);
            cache.Add(@"C:\A.pdf", "doc");
            string val;
            AssertTrue(cache.TryGet(@"c:\a.pdf", out val) && val == "doc", "ключи-пути сравниваются без регистра");
            AssertEqual(1, cache.Count, "разный регистр — тот же ключ");
        }

        private static void TestLruClear()
        {
            var evicted = new List<string>();
            var cache = new LruCache<string>(3, delegate(string v) { evicted.Add(v); });
            cache.Add("a", "a");
            cache.Add("b", "b");
            cache.Clear();
            AssertEqual(0, cache.Count, "после Clear пусто");
            AssertEqual(2, evicted.Count, "Clear освобождает все оставшиеся элементы");
        }

        private static void TestLruCapacityGuard()
        {
            bool threw = false;
            try { new LruCache<string>(0, null); }
            catch (ArgumentOutOfRangeException) { threw = true; }
            AssertTrue(threw, "ёмкость < 1 должна отвергаться");
        }

        // ---------- PdfPageGrid: набор ключей и вытеснение кэша ----------

        private static void TestGridBuildKeySet()
        {
            var pages = new List<PdfPageRef>
            {
                new PdfPageRef { SourcePath = @"C:\a.pdf", PageIndex = 0 },
                new PdfPageRef { SourcePath = @"C:\a.pdf", PageIndex = 1 },
                new PdfPageRef { SourcePath = @"C:\a.pdf", PageIndex = 0 }, // дубль
            };
            HashSet<string> keys = PdfPageGrid.BuildKeySet(pages);
            AssertEqual(2, keys.Count, "дубли схлопываются");
            AssertTrue(keys.Contains(PdfPageGrid.ThumbKey(pages[0])), "ключ страницы 0 присутствует");
            AssertEqual(0, PdfPageGrid.BuildKeySet(null).Count, "null -> пустой набор");
        }

        private static void TestGridStaleKeys()
        {
            var cached = new List<string> { "a|0", "a|1", "b|0" };
            var keep = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "a|0", "b|0", "c|0" };
            List<string> stale = PdfPageGrid.StaleKeys(cached, keep);
            AssertEqual(1, stale.Count, "один устаревший ключ");
            AssertEqual("a|1", stale[0], "вытесняется отсутствующий в keep");

            // Тот же набор -> ничего не устаревает (переупорядочивание не роняет кэш).
            var same = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "a|0", "a|1", "b|0" };
            AssertEqual(0, PdfPageGrid.StaleKeys(cached, same).Count, "полное совпадение — без вытеснения");
        }

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

        private static void AssertThrows(string what, Action action)
        {
            try
            {
                action();
            }
            catch (MergeException)
            {
                return; // ожидаемая ошибка ввода
            }
            throw new Exception(what + ": ожидалось исключение, но его не было");
        }
    }
}
