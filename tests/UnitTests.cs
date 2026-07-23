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
            Run("UsageStats.Total: включает PdfToWord, исключает сжатия", TestUsageTotal);
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
            Run("PdfToolFormBase.ClassifyPageKey: Delete, Alt+←/→, Ctrl+A, Enter", TestClassifyPageKey);
            Run("PdfPageOrder: добавление и границы MoveUp/MoveDown", TestPdfOrderMoves);
            Run("PdfPageOrder: перенос drag&drop в обе стороны", TestPdfOrderDragMove);
            Run("PdfPageOrder: удаление набора строк + Clear", TestPdfOrderRemove);
            Run("PdfToWordService.Assemble: сборка страниц из нескольких PDF, границы", TestAssemble);
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
            Run("PdfPageGrid.LowerBound: бинарный поиск по монотонному предикату", TestLowerBound);
            Run("PdfPageGrid.VisibleRange: видимый диапазон по Top/Bottom", TestVisibleRange);
            Run("PdfSplitForm.ShouldSuggestCompression: без сжатия и ≥90% исходника", TestSuggestCompression);
            Run("OcrLayout: порядок чтения (сверху вниз, слева направо)", TestOcrReadingOrder);
            Run("OcrLayout: разрыв абзаца по вертикальному зазору", TestOcrParagraphs);
            Run("OcrLayout: разрыв абзаца по красной строке (justified)", TestOcrParagraphsIndent);
            Run("OcrLayout: измерен отступ красной строки", TestOcrIndentDetected);
            Run("OcrLayout: без отступов -> красная строка не навязывается", TestOcrNoIndentReported);
            Run("FontNames.Clean: нормализация имени шрифта", TestFontNames);
            Run("WordDocxWriter: неустановленный шрифт -> fallback (против eastAsia-разрядки)", TestResolveFontName);
            Run("PdfToolFormBase: проценты прогресса (сделано/всего, клампы)", TestProgressPercent);
            Run("OcrLayout: смена шрифта -> раны", TestOcrRunsFontFamily);
            Run("OcrLayout: стиль рана (курсив, кегль)", TestOcrParagraphStyle);
            Run("OcrLayout: смешанный формат -> раны", TestOcrRunsMixedFormat);
            Run("OcrLayout: надстрочный ран", TestOcrSuperscript);
            Run("OcrLayout: ран-гиперссылка", TestOcrHyperlinkRun);
            Run("OcrLayout: цвет рана сохранён", TestOcrColorRun);
            Run("OcrLayout: рваный абзац по левому краю", TestOcrLeftAligned);
            Run("OcrLayout: центрированная строка", TestOcrCentered);
            Run("OcrLayout: IsCentered — узкая/широкая/красная строка/рваная/правая", TestIsCenteredPredicate);
            Run("OcrLayout: центрированный многострочный титул -> один центрированный абзац", TestOcrCenteredBlock);
            Run("OcrLayout: фрагменты слова склеиваются по зазору", TestOcrGlueFragments);
            Run("OcrLayout: узкий настоящий пробел сохранён (не склейка)", TestOcrNarrowSpaceKept);
            Run("OcrLayout: тонкое тире остаётся в строке", TestOcrThinDashStaysOnLine);
            Run("OcrLayout: перенос с дефисом склеивает слово", TestOcrHyphenation);
            Run("OcrLayout: пустой ввод -> нет абзацев", TestOcrEmpty);
            Run("ListMarker: нумерованный «1.»/«12)»", TestListMarkerNumbered);
            Run("ListMarker: маркированный «•»/«—»", TestListMarkerBulleted);
            Run("ListMarker: не список (год, проценты, без пробела, обычный текст)", TestListMarkerNegatives);
            Run("OcrLayout: плотный нумерованный список -> отдельные пункты с ListKind", TestOcrNumberedList);
            Run("OcrLayout: маркированный список -> ListKind=Bulleted, содержимое без маркера", TestOcrBulletedList);
            Run("StampDetector: штамп ЭП -> область со всеми словами, вне полосы не берётся", TestStampDetected);
            Run("StampDetector: нет одного опорного слова -> не штамп", TestStampMissingAnchor);
            Run("StampDetector: опорные слова разбросаны по странице -> не штамп", TestStampScatteredRejected);
            Run("Loc: каталог полон — у каждого ключа непустые ru и en", TestLocCatalogComplete);
            Run("Loc: плейсхолдеры {N} у ru и en совпадают", TestLocPlaceholders);
            Run("Loc: Init/Current/Parse/Code", TestLocInit);
            Run("Loc: EN — в построенных формах нет кириллицы (кроме двуязычных меток)", TestNoCyrillicInEnglishForms);

            Run("TableDetector: сетка 2x2 -> строки/колонки, текст ячеек", TestTable2x2);
            Run("TableDetector: пропущенная гориз. граница -> rowspan", TestTableRowSpan);
            Run("TableDetector: пропущенная верт. граница -> colspan", TestTableColSpan);
            Run("TableDetector: одиночные линии (подчёркивания) -> не таблица", TestTableStrayLines);
            Run("TableDetector: рамка 1x1 без внутренних линий -> не таблица", TestTableSingleBox);
            Run("TableDetector: слова вне таблицы остаются в потоке", TestTableWordsOutside);
            Run("TableDetector: нет линий -> нет таблиц, все слова в потоке", TestTableNoLines);
            Run("PdfToWordService: страница-таблица считается текстовой (не «скан»)", TestHasExtractableContent);
            Run("UnderlineDetector: линия под словом -> подчёркнуто", TestUnderlineMarks);
            Run("UnderlineDetector: далёкая/короткая линия -> не подчёркнуто", TestUnderlineIgnores);
            Run("UnderlineDetector: линия во всю ширину (разделитель) -> не подчёркнуто", TestUnderlineWideRule);
            Run("OcrLayout: левый сайдбар отделяется от тела (не перемешиваются)", TestSidebarSeparation);
            Run("OcrLayout: одноколоночный текст не делится (сайдбар не срабатывает)", TestNoSidebarSingleColumn);
            Run("WordDocxWriter: порядок чтения (сверху вниз, бок о бок — слева направо)", TestReadingOrder);
            Run("WordDocxWriter: центрированное изображение (герб) -> по центру, врезка/печать -> нет", TestImageCentered);
            Run("PageRasterizer: рамка PDF (Y-вверх) -> пиксельный кроп, кламп по краю", TestCropRect);

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
            string name = n.Next("Очень длинное имя файла отчета за первый квартал");
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
                AssertEqual(Loc.T("menu.root"), help.Text, "название пункта");

                var texts = new List<string>();
                foreach (System.Windows.Forms.ToolStripItem it in help.DropDownItems)
                    if (it is System.Windows.Forms.ToolStripMenuItem)
                        texts.Add(it.Text);
                AssertTrue(texts.Contains(Loc.T("menu.howTo")), "есть «Как пользоваться»");
                AssertTrue(texts.Contains(Loc.T("menu.stats")), "есть «Статистика»");
                AssertTrue(texts.Contains(Loc.T("menu.language")), "есть выбор языка");
                AssertTrue(texts.Contains("Папка отчётов"), "доп. пункт вставлен");
                // «О программе» перенесена на стартовый экран — в меню её быть не должно.
                AssertTrue(!texts.Contains(Loc.T("hub.about")), "«О программе» убрана из меню");
            }

            // Без доп. пунктов: «Как пользоваться», «Статистика», «Язык / Language».
            using (System.Windows.Forms.MenuStrip menu = HelpMenu.Create(null, delegate { }))
            {
                var help = (System.Windows.Forms.ToolStripMenuItem)menu.Items[0];
                int menuItems = 0;
                foreach (System.Windows.Forms.ToolStripItem it in help.DropDownItems)
                    if (it is System.Windows.Forms.ToolStripMenuItem)
                        menuItems++;
                AssertEqual(3, menuItems, "без extras — три пункта (справка, статистика, язык)");
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
            AssertEqual(PdfToolFormBase.PageKeyAction.Remove, PdfToolFormBase.ClassifyPageKey(System.Windows.Forms.Keys.Delete), "Delete — удалить");
            AssertEqual(PdfToolFormBase.PageKeyAction.MoveEarlier, PdfToolFormBase.ClassifyPageKey(Alt | System.Windows.Forms.Keys.Left), "Alt+← — раньше");
            AssertEqual(PdfToolFormBase.PageKeyAction.MoveLater, PdfToolFormBase.ClassifyPageKey(Alt | System.Windows.Forms.Keys.Right), "Alt+→ — позже");
            AssertEqual(PdfToolFormBase.PageKeyAction.SelectAll, PdfToolFormBase.ClassifyPageKey(Ctrl | System.Windows.Forms.Keys.A), "Ctrl+A — выделить всё");
            AssertEqual(PdfToolFormBase.PageKeyAction.Swallow, PdfToolFormBase.ClassifyPageKey(System.Windows.Forms.Keys.Enter), "Enter — не сохранять");
            AssertEqual(PdfToolFormBase.PageKeyAction.None, PdfToolFormBase.ClassifyPageKey(System.Windows.Forms.Keys.Left), "просто ← — навигация");
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

        private static void TestUsageTotal()
        {
            var s = new UsageStats
            {
                ExcelDigests = 1, PdfMerges = 2, PdfExtracts = 3, PdfSplitRanges = 4,
                PdfSplitEveryN = 5, PdfSplitBookmarks = 6, PdfToWord = 7, PdfCompressions = 99
            };
            // Total — сумма ОПЕРАЦИЙ (1+…+7); сжатие — параметр, в Total не входит.
            AssertEqual(28, s.Total, "Total включает PdfToWord и исключает PdfCompressions");
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
            order.Clear();
            AssertEqual(0, order.Count, "Clear очищает список");
        }

        private static void TestAssemble()
        {
            // Два источника A (2 стр.) и B (1 стр.); собираем в порядке из разных файлов.
            var bysource = new Dictionary<string, List<PdfPageText>>(StringComparer.OrdinalIgnoreCase)
            {
                { "A.pdf", new List<PdfPageText> { new PdfPageText { PageIndex = 0 }, new PdfPageText { PageIndex = 1 } } },
                { "B.pdf", new List<PdfPageText> { new PdfPageText { PageIndex = 0 } } }
            };
            var order = new List<PdfPageRef>
            {
                new PdfPageRef { SourcePath = "B.pdf", PageIndex = 0 },
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = 1 },
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = 0 }
            };
            List<PdfPageText> r = PdfToWordService.Assemble(bysource, order);
            AssertEqual(3, r.Count, "собраны все 3 страницы из двух файлов");
            AssertTrue(ReferenceEquals(r[0], bysource["B.pdf"][0]), "первая — стр.0 из B");
            AssertTrue(ReferenceEquals(r[1], bysource["A.pdf"][1]), "вторая — стр.1 из A");
            AssertTrue(ReferenceEquals(r[2], bysource["A.pdf"][0]), "третья — стр.0 из A");
            // Несуществующий источник и индекс вне диапазона — пропускаются.
            var bad = new List<PdfPageRef>
            {
                new PdfPageRef { SourcePath = "нет.pdf", PageIndex = 0 },
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = 9 },
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = -1 },
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = 0 }
            };
            List<PdfPageText> rb = PdfToWordService.Assemble(bysource, bad);
            AssertEqual(1, rb.Count, "остаётся только валидная ссылка");
            AssertTrue(ReferenceEquals(rb[0], bysource["A.pdf"][0]), "валидная — A стр.0");
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

        private static void TestLowerBound()
        {
            AssertEqual(3, PdfPageGrid.LowerBound(5, delegate(int i) { return i >= 3; }), "первый индекс >= 3");
            AssertEqual(0, PdfPageGrid.LowerBound(5, delegate(int i) { return true; }), "все истинны -> 0");
            AssertEqual(5, PdfPageGrid.LowerBound(5, delegate(int i) { return false; }), "все ложны -> count");
            AssertEqual(0, PdfPageGrid.LowerBound(0, delegate(int i) { return true; }), "пустой -> 0");
            AssertEqual(1, PdfPageGrid.LowerBound(4, delegate(int i) { return i >= 1; }), "граница у начала");
        }

        private static void TestVisibleRange()
        {
            // 8 плиток высотой 10, один столбец, прокрутка вверх на 25; низ вьюпорта = 30.
            // Top[i]=i*10-25, Bottom[i]=Top+10 -> видимы i=2..5.
            Func<int, int> topOf = delegate(int i) { return i * 10 - 25; };
            Func<int, int> bottomOf = delegate(int i) { return i * 10 - 25 + 10; };
            int first, last;
            PdfPageGrid.VisibleRange(8, topOf, bottomOf, 30, out first, out last);
            AssertEqual(2, first, "первый видимый (Bottom >= 0)");
            AssertEqual(5, last, "последний видимый (Top <= 30)");

            // Всё выше вьюпорта -> ничего целиком не видно (first > last).
            Func<int, int> topHi = delegate(int i) { return i * 10 - 1000; };
            Func<int, int> botHi = delegate(int i) { return i * 10 - 1000 + 10; };
            PdfPageGrid.VisibleRange(8, topHi, botHi, 30, out first, out last);
            AssertTrue(first > last, "всё выше вьюпорта -> пусто");

            // Пустой список.
            PdfPageGrid.VisibleRange(0, topOf, bottomOf, 30, out first, out last);
            AssertTrue(first > last, "нет элементов -> пусто");
        }

        private static void TestSuggestCompression()
        {
            long mb = 1024L * 1024;
            AssertTrue(PdfSplitForm.ShouldSuggestCompression(CompressionLevel.None, 10 * mb, 95 * mb / 10),
                "без сжатия, 9.5МБ из 10МБ -> подсказать");
            AssertTrue(!PdfSplitForm.ShouldSuggestCompression(CompressionLevel.None, 10 * mb, 3 * mb),
                "3МБ из 10МБ -> не подсказывать");
            AssertTrue(!PdfSplitForm.ShouldSuggestCompression(CompressionLevel.Good, 10 * mb, 95 * mb / 10),
                "уже со сжатием -> не подсказывать");
            AssertTrue(!PdfSplitForm.ShouldSuggestCompression(CompressionLevel.None, 700 * 1024, 690 * 1024),
                "мелкий файл (<1МБ) -> не подсказывать");
            AssertTrue(PdfSplitForm.ShouldSuggestCompression(CompressionLevel.None, 10 * mb, 9 * mb),
                "ровно 90% -> подсказать (граница)");
            AssertTrue(!PdfSplitForm.ShouldSuggestCompression(CompressionLevel.None, 0, 5 * mb),
                "размер исходника неизвестен (0) -> не подсказывать");
        }

        // ---------- OcrLayout (порядок чтения born-digital) ----------

        /// <summary>Слово с рамкой (Y вверх): left/bottom — левый нижний угол, +ширина/высота.</summary>
        private static PdfWord W(string text, double left, double bottom, double width, double height)
        {
            return new PdfWord { Text = text, Left = left, Right = left + width, Bottom = bottom, Top = bottom + height };
        }

        private static void TestOcrReadingOrder()
        {
            // Ввод намеренно вперемешку; ожидаем «Hello world» затем «second line».
            var words = new List<PdfWord>
            {
                W("world", 35, 90, 30, 10),
                W("line", 35, 70, 20, 10),
                W("Hello", 0, 90, 30, 10),
                W("second", 0, 70, 30, 10)
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(1, p.Count, "две близкие строки -> один абзац");
            AssertEqual("Hello world second line", p[0], "порядок чтения");
        }

        private static void TestOcrParagraphs()
        {
            // Три плотные строки + одна с большим зазором -> два абзаца.
            var words = new List<PdfWord>
            {
                W("Aaa", 0, 100, 20, 10), // midY 105
                W("Bbb", 0, 88, 20, 10),  // midY 93  (зазор 12)
                W("Ccc", 0, 76, 20, 10),  // midY 81  (зазор 12)
                W("Ddd", 0, 50, 20, 10)   // midY 55  (зазор 26 -> новый абзац)
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(2, p.Count, "разрыв абзаца по большому зазору");
            AssertEqual("Aaa Bbb Ccc", p[0], "первый абзац");
            AssertEqual("Ddd", p[1], "второй абзац");
        }

        private static void TestOcrParagraphsIndent()
        {
            // Justified-документ (строки достают до правого поля 100), абзацы разделены
            // ТОЛЬКО красной строкой (первая строка с отступом) при равном интервале —
            // как в Word-экспорте. Зазор одинаков всюду, поэтому делить должны отступ и
            // «короткая» последняя строка, а не зазор.
            var words = new List<PdfWord>
            {
                W("A1a", 15, 180, 85, 10), // абзац A, 1-я строка с отступом, до правого поля (Right 100)
                W("A2", 0, 168, 100, 10),  // продолжение у левого поля, полная строка
                W("A3", 0, 156, 40, 10),   // короткая последняя строка абзаца
                W("B1", 15, 140, 85, 10),  // абзац B, красная строка (отступ) -> новый абзац
                W("B2", 0, 128, 30, 10)    // короткая последняя строка
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(2, p.Count, "красная строка + короткая строка делят абзацы без зазора");
            AssertEqual("A1a A2 A3", p[0], "первый абзац");
            AssertEqual("B1 B2", p[1], "второй абзац");
        }

        private static void TestOcrIndentDetected()
        {
            // Тот же justified-документ с красной строкой: отступ первых строк 15 pt
            // должен быть измерен (большинство абзацев с отступом).
            var words = new List<PdfWord>
            {
                W("A1a", 15, 180, 85, 10),
                W("A2", 0, 168, 100, 10),
                W("A3", 0, 156, 40, 10),
                W("B1", 15, 140, 85, 10),
                W("B2", 0, 128, 30, 10)
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(2, layout.Paragraphs.Count, "два абзаца");
            AssertEqual(15.0, layout.FirstLineIndentPt, "измерен отступ красной строки (медиана 15 pt)");
        }

        private static void TestOcrNoIndentReported()
        {
            // Документ без отступов (все строки у левого поля) -> красная строка не навязывается.
            var words = new List<PdfWord>
            {
                W("Aaa", 0, 100, 20, 10),
                W("Bbb", 0, 88, 20, 10),
                W("Ccc", 0, 76, 20, 10),
                W("Ddd", 0, 50, 20, 10)
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(0.0, layout.FirstLineIndentPt, "без отступов -> 0 (не портим документ)");
        }

        private static void TestFontNames()
        {
            AssertEqual("Times New Roman", FontNames.Clean("UBWKNX+Times New Roman,Italic"), "subset + курсив-суффикс");
            AssertEqual("Arial", FontNames.Clean("ABCDEF+Arial-BoldMT"), "subset + -BoldMT");
            AssertEqual("Times New Roman", FontNames.Clean("TimesNewRomanPSMT"), "PSMT + слитное имя");
            AssertEqual("Times New Roman", FontNames.Clean("TimesNewRomanPS-BoldMT"), "PS + -BoldMT");
            AssertEqual("Courier New", FontNames.Clean("CourierNewPS-BoldMT"), "Courier New");
            AssertEqual("Arial", FontNames.Clean("ArialMT"), "ArialMT");
            AssertEqual("Calibri", FontNames.Clean("Calibri"), "уже чистое");
            AssertEqual("PT Astra Serif", FontNames.Clean("BBHOZJ+PTAstraSerif-Regular"), "PT Astra Serif: префикс-аббревиатура + слитное");
            AssertEqual("MS Gothic", FontNames.Clean("MSGothic"), "MS Gothic: аббревиатура-префикс");
            AssertTrue(FontNames.Clean(null) == null, "null -> null");
            AssertTrue(FontNames.Clean("  ") == null, "пусто -> null");
        }

        private static void TestResolveFontName()
        {
            var installed = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Times New Roman", "Calibri Light"
            };
            // Установленный шрифт — оставить как есть (в т.ч. без учёта регистра).
            AssertEqual("Calibri Light", WordDocxWriter.ResolveFontName("Calibri Light", installed, "Times New Roman"), "установленный сохранён");
            // Совпадение без учёта регистра — установленный не уходит в fallback (Word регистронезависим).
            AssertEqual("times new roman", WordDocxWriter.ResolveFontName("times new roman", installed, "Times New Roman"), "регистр при поиске не важен");
            // НЕустановленный (напр. PT Astra Serif) -> fallback, иначе Word уводит кириллицу в eastAsia -> разрядка.
            AssertEqual("Times New Roman", WordDocxWriter.ResolveFontName("PT Astra Serif", installed, "Times New Roman"), "неустановленный -> fallback");
            AssertEqual("Times New Roman", WordDocxWriter.ResolveFontName(null, installed, "Times New Roman"), "null -> fallback");
            AssertEqual("Times New Roman", WordDocxWriter.ResolveFontName("X", null, "Times New Roman"), "нет списка -> fallback");
        }

        private static void TestProgressPercent()
        {
            AssertEqual(0, PdfToolFormBase.ProgressPercent(0, 0), "0/0 -> 0 (без деления на ноль)");
            AssertEqual(0, PdfToolFormBase.ProgressPercent(0, 10), "0/10 -> 0");
            AssertEqual(50, PdfToolFormBase.ProgressPercent(5, 10), "5/10 -> 50");
            AssertEqual(100, PdfToolFormBase.ProgressPercent(10, 10), "10/10 -> 100");
            AssertEqual(100, PdfToolFormBase.ProgressPercent(11, 10), "11/10 -> 100 (кламп сверху)");
            AssertEqual(0, PdfToolFormBase.ProgressPercent(-1, 10), "отрицательное сделано -> 0");
            AssertEqual(0, PdfToolFormBase.ProgressPercent(5, -1), "отрицательное всего -> 0");
            AssertEqual(33, PdfToolFormBase.ProgressPercent(1, 3), "1/3 -> 33 (округление вниз)");
            AssertEqual(66, PdfToolFormBase.ProgressPercent(2, 3), "2/3 -> 66");
            AssertEqual(50, PdfToolFormBase.ProgressPercent(1000000, 2000000), "большие числа без переполнения");
        }

        private static void TestOcrRunsFontFamily()
        {
            // Разные семейства -> разные раны.
            var words = new List<PdfWord>
            {
                new PdfWord { Text = "Ариал", Left = 0, Right = 40, Bottom = 0, Top = 8, FontSizePt = 12, FontName = "Arial" },
                new PdfWord { Text = "Таймс", Left = 45, Right = 85, Bottom = 0, Top = 8, FontSizePt = 12, FontName = "Times New Roman" }
            };
            List<OcrRun> runs = OcrLayout.Analyze(words).Paragraphs[0].Runs;
            AssertEqual(2, runs.Count, "смена шрифта -> новый ран");
            AssertEqual("Arial", runs[0].FontName, "первый ран — Arial");
            AssertEqual("Times New Roman", runs[1].FontName, "второй ран — Times New Roman");
        }

        private static void TestOcrParagraphStyle()
        {
            // Курсивная строка кеглем 14, единый формат -> один ран, курсив, кегль 14.
            var words = new List<PdfWord>
            {
                new PdfWord { Text = "Имя:", Left = 0, Right = 30, Bottom = 0, Top = 8, FontSizePt = 14, Italic = true },
                new PdfWord { Text = "_dmarc", Left = 35, Right = 90, Bottom = 0, Top = 8, FontSizePt = 14, Italic = true }
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(1, layout.Paragraphs.Count, "один абзац");
            AssertEqual(1, layout.Paragraphs[0].Runs.Count, "единый формат -> один ран");
            OcrRun r = layout.Paragraphs[0].Runs[0];
            AssertEqual(14.0, r.FontSizePt, "кегль рана");
            AssertTrue(r.Italic, "ран курсивный");
            AssertTrue(!r.Bold, "не полужирный");
            AssertEqual("Имя: _dmarc", r.Text, "текст рана");
        }

        private static void TestOcrRunsMixedFormat()
        {
            // Полужирное слово среди обычных -> три рана; жирным только среднее.
            var words = new List<PdfWord>
            {
                new PdfWord { Text = "обычное", Left = 0, Right = 40, Bottom = 0, Top = 8, FontSizePt = 12 },
                new PdfWord { Text = "жирное", Left = 45, Right = 85, Bottom = 0, Top = 8, FontSizePt = 12, Bold = true },
                new PdfWord { Text = "снова", Left = 90, Right = 130, Bottom = 0, Top = 8, FontSizePt = 12 }
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(1, layout.Paragraphs.Count, "один абзац");
            List<OcrRun> runs = layout.Paragraphs[0].Runs;
            AssertEqual(3, runs.Count, "три рана: обычный / жирный / обычный");
            AssertTrue(!runs[0].Bold && runs[1].Bold && !runs[2].Bold, "жирный только средний ран");
            AssertEqual("обычное жирное снова", layout.Paragraphs[0].Text, "склейка текста");
        }

        private static void TestOcrSuperscript()
        {
            // Мелкое приподнятое слово среди обычных -> надстрочный ран.
            var words = new List<PdfWord>
            {
                W("a", 0, 0, 10, 8),    // обычные, база 0, высота 8
                W("b", 12, 0, 10, 8),
                W("2", 24, 4, 6, 4)     // мельче (4) и приподнято (база 4) -> надстрочный
            };
            List<OcrRun> runs = OcrLayout.Analyze(words).Paragraphs[0].Runs;
            AssertEqual(2, runs.Count, "надстрочный — отдельный ран");
            AssertTrue(runs[1].Super && !runs[1].Sub, "'2' надстрочный");
            AssertEqual("2", runs[1].Text, "текст надстрочного рана");
        }

        private static void TestOcrHyperlinkRun()
        {
            var words = new List<PdfWord>
            {
                new PdfWord { Text = "ссылка", Left = 0, Right = 40, Bottom = 0, Top = 8, FontSizePt = 12, Uri = "https://example.com" },
                new PdfWord { Text = "обычный", Left = 45, Right = 90, Bottom = 0, Top = 8, FontSizePt = 12 }
            };
            List<OcrRun> runs = OcrLayout.Analyze(words).Paragraphs[0].Runs;
            AssertEqual(2, runs.Count, "ссылка — отдельный ран");
            AssertEqual("https://example.com", runs[0].Uri, "URI рана сохранён");
            AssertTrue(runs[1].Uri == null, "обычный ран без ссылки");
        }

        private static void TestOcrColorRun()
        {
            var words = new List<PdfWord>
            {
                new PdfWord { Text = "красный", Left = 0, Right = 40, Bottom = 0, Top = 8, FontSizePt = 12, ColorArgb = 0xFF0000 }
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(1, layout.Paragraphs[0].Runs.Count, "один ран");
            AssertEqual(0xFF0000, layout.Paragraphs[0].Runs[0].ColorArgb, "цвет рана сохранён");
        }

        private static void TestOcrLeftAligned()
        {
            // Рваный справа абзац (строки не достают до правого поля) -> по левому краю,
            // а не насильно по ширине. Полная строка ниже задаёт правое поле (100).
            var words = new List<PdfWord>
            {
                W("aaaa", 0, 100, 60, 8),   // абзац 1, строка 1: Right 60
                W("bbbb", 0, 88, 55, 8),    // абзац 1, строка 2: Right 55 (обе рваные)
                W("cccc", 0, 50, 100, 8)    // абзац 2: полная строка Right 100 (задаёт поле)
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(2, layout.Paragraphs.Count, "два абзаца");
            AssertEqual(OcrAlignment.Left, layout.Paragraphs[0].Alignment, "рваный абзац — по левому краю");
        }

        private static void TestOcrCentered()
        {
            // Короткая центрированная строка («7») над телом -> отдельный центрированный абзац.
            var words = new List<PdfWord>
            {
                W("7", 48, 180, 4, 8),      // центр: слева 48, справа 48 от полей 0..100
                W("Тело", 0, 150, 100, 8),  // полная строка (Left 0, Right 100) — задаёт поля
                W("строка", 0, 138, 100, 8)
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(2, layout.Paragraphs.Count, "два абзаца");
            AssertEqual(OcrAlignment.Center, layout.Paragraphs[0].Alignment, "номер страницы по центру");
            AssertEqual(OcrAlignment.Justify, layout.Paragraphs[1].Alignment, "тело по ширине");
        }

        private static void TestIsCenteredPredicate()
        {
            // Узкая центрированная (номер страницы): большие симметричные зазоры.
            AssertTrue(OcrLayout.IsCentered(48, 48, 8, 100), "узкая по центру");
            // ШИРОКАЯ центрированная строка титула: зазоры малы в долях ширины (~5%), но симметричны —
            // прежний порог 12% ширины её терял, новый (доля кегля) ловит.
            AssertTrue(OcrLayout.IsCentered(31, 22, 16, 453), "широкая строка титула по центру");
            // Красная строка: упирается в правое поле (правый зазор ≈ 0) — не центр.
            AssertTrue(!OcrLayout.IsCentered(35, 0, 16, 453), "красная строка — не центр");
            // Рваная левая: стоит у левого поля (левый зазор ≈ 0) — не центр.
            AssertTrue(!OcrLayout.IsCentered(0, 40, 10, 100), "рваная левая — не центр");
            // Правое выравнивание: у правого поля (правый зазор ≈ 0), большой левый — не центр.
            AssertTrue(!OcrLayout.IsCentered(200, 0, 16, 453), "правое выравнивание — не центр");
            // Несимметричная (левый много больше правого) — не центр, даже если оба > порога.
            AssertTrue(!OcrLayout.IsCentered(120, 20, 16, 453), "асимметрия — не центр");
        }

        private static void TestOcrCenteredBlock()
        {
            // Типичный центрированный титул документа: ШИРОКИЕ строки дотягиваются до обоих полей
            // (внешне как justified) и лишь короткий хвост «Sh» явно центрирован. Все строки соосны
            // (midX=50) -> ВСЕ распознаются центрированными по общей оси. Каждая — своим абзацем
            // (исходная разбивка сохранена: Word не сольёт их и не перевёрстывает), а не «широкие=тело
            // + сирота-хвост».
            var words = new List<PdfWord>
            {
                W("Wideoneaa", 2, 200, 96, 8),  // Left2 Right98 mid50: дотягивается до полей (не «плавает»)
                W("Widetwoaa", 2, 188, 96, 8),  // Left2 Right98 mid50: тоже широкая
                W("Sh",       35, 176, 30, 8),  // Left35 Right65 mid50: короткая центрированная — доказывает ось
                W("Bodyoneaa", 2, 150, 96, 8),  // тело: полные строки (задают поля, justified), большой зазор
                W("Bodytwoaa", 2, 138, 96, 8),
                W("Bodythree", 2, 126, 96, 8)
            };
            List<OcrParagraph> paras = OcrLayout.Analyze(words).Paragraphs;
            AssertEqual(4, paras.Count, "3 центрированные строки титула (каждая своим абзацем) + тело");
            AssertEqual(OcrAlignment.Center, paras[0].Alignment, "широкая строка титула — по центру (по общей оси)");
            AssertEqual(OcrAlignment.Center, paras[1].Alignment, "вторая широкая строка титула — по центру");
            AssertEqual(OcrAlignment.Center, paras[2].Alignment, "короткий хвост — по центру");
            AssertEqual("Wideoneaa", paras[0].Text, "исходная строка сохранена отдельным абзацем (не слита)");
            AssertEqual(OcrAlignment.Justify, paras[3].Alignment, "тело — выключка, не центр (нет короткой центрированной строки)");
        }

        private static void TestOcrGlueFragments()
        {
            // Почти соприкасающиеся куски одного токена (мизерный зазор < 0.08 кегля) склеиваем
            // без пробела; между словами обычный зазор — пробел.
            var words = new List<PdfWord>
            {
                W("м", 0, 0, 8, 10),      // Right 8
                W("и", 8.5, 0, 8, 10),    // зазор 0.5 (0.05 кегля) -> склеить
                W("р", 17, 0, 8, 10),     // зазор 0.5 (0.05 кегля) -> склеить  => «мир»
                W("тут", 30, 0, 20, 10)   // зазор 5 (0.5 кегля) -> пробел      => «мир тут»
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(1, p.Count, "одна строка — один абзац");
            AssertEqual("мир тут", p[0], "фрагменты склеены, между словами пробел");
        }

        private static void TestOcrNarrowSpaceKept()
        {
            // Регресс: в узких шрифтах (Calibri Light) настоящий межсловный зазор ≈ 0.18 кегля.
            // Прежний порог 0.2 ронял такой пробел и слеплял слова («СЛОВОСЛОВО»);
            // 0.08 сохраняет пробел. Зазор здесь 0.15 кегля — между старым и новым порогом.
            var words = new List<PdfWord>
            {
                W("МИНИСТЕРСТВО", 0, 0, 60, 16),      // Right 60
                W("ФИНАНСОВ", 62.4, 0, 40, 16)        // зазор 2.4 = 0.15 кегля -> пробел
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(1, p.Count, "одна строка — один абзац");
            AssertEqual("МИНИСТЕРСТВО ФИНАНСОВ", p[0], "узкий настоящий пробел сохранён (не склейка)");
        }

        private static void TestOcrThinDashStaysOnLine()
        {
            // Тонкое тире с крошечной рамкой и центром выше базовой линии текста не должно
            // отрываться в отдельную строку/абзац (перекрытие рамок, а не расстояние центров).
            var words = new List<PdfWord>
            {
                W("quarantine", 0, 0, 50, 8),   // [0..8], центр 4
                W("—", 55, 4.3, 8, 0.6),         // [4.3..4.9], центр 4.6 — выше на 0.6
                W("добавляет", 70, 0, 50, 8)     // [0..8], центр 4
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(1, p.Count, "тире не отрывается — одна строка, один абзац");
            AssertEqual("quarantine — добавляет", p[0], "тире осталось между словами");
        }

        private static void TestOcrHyphenation()
        {
            // Строка кончается «wo-», следующая «rld» -> склеить в «world».
            var words = new List<PdfWord>
            {
                W("hello", 0, 90, 30, 10),
                W("wo-", 35, 90, 20, 10),
                W("rld", 0, 75, 20, 10)
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(1, p.Count, "один абзац");
            AssertEqual("hello world", p[0], "дефис-перенос склеен");
        }

        private static void TestOcrEmpty()
        {
            AssertEqual(0, OcrLayout.ToParagraphs(new List<PdfWord>()).Count, "пусто -> нет абзацев");
            AssertEqual(0, OcrLayout.ToParagraphs(null).Count, "null -> нет абзацев");
        }

        // ---------- ListMarker (маркеры списка) ----------

        private static void TestListMarkerNumbered()
        {
            ListMarker.Result a = ListMarker.Detect("1. Внести изменения");
            AssertEqual((int)ListKind.Numbered, (int)a.Kind, "«1.» -> нумерованный");
            AssertEqual(1, a.Number, "номер 1");
            AssertEqual("Внести изменения", "1. Внести изменения".Substring(a.ContentStart), "содержимое без маркера");

            ListMarker.Result b = ListMarker.Detect("12) пункт");
            AssertEqual((int)ListKind.Numbered, (int)b.Kind, "«12)» -> нумерованный");
            AssertEqual(12, b.Number, "номер 12");
            AssertEqual("пункт", "12) пункт".Substring(b.ContentStart), "содержимое после «12)»");
        }

        private static void TestListMarkerBulleted()
        {
            ListMarker.Result a = ListMarker.Detect("• первый");
            AssertEqual((int)ListKind.Bulleted, (int)a.Kind, "«•» -> маркированный");
            AssertEqual("первый", "• первый".Substring(a.ContentStart), "содержимое без буллета");

            ListMarker.Result b = ListMarker.Detect("— тире-буллет");
            AssertEqual((int)ListKind.Bulleted, (int)b.Kind, "«—» -> маркированный");
            AssertEqual("тире-буллет", "— тире-буллет".Substring(b.ContentStart), "содержимое без тире");
        }

        private static void TestListMarkerNegatives()
        {
            AssertEqual((int)ListKind.None, (int)ListMarker.Detect("2025 год отчёта").Kind, "год не маркер");
            AssertEqual((int)ListKind.None, (int)ListMarker.Detect("12.5% роста").Kind, "проценты не маркер");
            AssertEqual((int)ListKind.None, (int)ListMarker.Detect("1.без пробела").Kind, "без пробела после точки — не маркер");
            AssertEqual((int)ListKind.None, (int)ListMarker.Detect("•безпробела").Kind, "буллет без пробела — не маркер");
            AssertEqual((int)ListKind.None, (int)ListMarker.Detect("Обычный текст").Kind, "обычный текст — не список");
            AssertEqual((int)ListKind.None, (int)ListMarker.Detect("").Kind, "пусто — не список");
        }

        private static void TestOcrNumberedList()
        {
            // Плотный одностроковый список (равный интервал, левый край, без отступа): без деления
            // по маркеру строки слиплись бы в один абзац. Ожидаем два пункта, ListKind=Numbered.
            var words = new List<PdfWord>
            {
                W("1.", 0, 100, 8, 10), W("Первый", 12, 100, 40, 10),
                W("2.", 0, 88, 8, 10),  W("Второй", 12, 88, 40, 10)
            };
            List<OcrParagraph> ps = OcrLayout.Analyze(words).Paragraphs;
            AssertEqual(2, ps.Count, "два пункта -> два абзаца");
            AssertEqual((int)ListKind.Numbered, (int)ps[0].ListKind, "пункт 1 — нумерованный");
            AssertEqual(1, ps[0].ListNumber, "номер первого пункта");
            AssertEqual((int)ListKind.Numbered, (int)ps[1].ListKind, "пункт 2 — нумерованный");
            AssertEqual(2, ps[1].ListNumber, "номер второго пункта");
            AssertEqual("Первый", ps[0].Text.Substring(ps[0].ListContentStart), "содержимое 1 без маркера");
            AssertEqual("Второй", ps[1].Text.Substring(ps[1].ListContentStart), "содержимое 2 без маркера");
        }

        private static void TestOcrBulletedList()
        {
            var words = new List<PdfWord>
            {
                W("•", 0, 100, 6, 10), W("яблоко", 10, 100, 40, 10),
                W("•", 0, 88, 6, 10),  W("груша", 10, 88, 40, 10)
            };
            List<OcrParagraph> ps = OcrLayout.Analyze(words).Paragraphs;
            AssertEqual(2, ps.Count, "два буллета -> два абзаца");
            AssertEqual((int)ListKind.Bulleted, (int)ps[0].ListKind, "буллет 1");
            AssertEqual((int)ListKind.Bulleted, (int)ps[1].ListKind, "буллет 2");
            AssertEqual("яблоко", ps[0].Text.Substring(ps[0].ListContentStart), "содержимое буллета 1 без маркера");
            AssertEqual("груша", ps[1].Text.Substring(ps[1].ListContentStart), "содержимое буллета 2 без маркера");
        }

        // ---------- StampDetector (текстовый штамп ЭП) ----------

        /// <summary>Синтетический текстовый штамп ЭП (4 строки) в левом нижнем углу + одно слово тела выше.</summary>
        private static List<PdfWord> StampWords()
        {
            return new List<PdfWord>
            {
                // строка-заголовок
                W("Документ", 100, 200, 50, 9), W("подписан", 155, 200, 50, 9),
                W("электронной", 210, 200, 60, 9), W("подписью", 275, 200, 50, 9),
                // сертификат
                W("Сертификат:", 100, 186, 60, 9), W("7f1224cd", 165, 186, 60, 9),
                // владелец
                W("Владелец", 100, 172, 50, 9), W("Иванов", 155, 172, 40, 9), W("Иван", 200, 172, 30, 9),
                // действителен
                W("Действителен", 100, 158, 70, 9), W("с", 175, 158, 8, 9), W("01.01.2025", 188, 158, 55, 9),
                W("по", 248, 158, 15, 9), W("01.01.2026", 268, 158, 55, 9),
                // слово тела заметно выше штампа — в область попасть не должно
                W("Обычныйтекст", 100, 500, 80, 9)
            };
        }

        private static void TestStampDetected()
        {
            List<PdfWord> words = StampWords();
            StampRegion s = StampDetector.Detect(words, 595, 842);
            AssertTrue(s != null, "штамп распознан");
            AssertEqual(14, s.Words.Count, "в область попали все 14 слов штампа (без слова тела)");
            bool bodyIn = false;
            foreach (PdfWord w in s.Words) if (w.Text == "Обычныйтекст") bodyIn = true;
            AssertTrue(!bodyIn, "слово тела вне полосы штампа не захвачено");
            AssertTrue(s.Left <= 100 && s.Right >= 323 && s.Bottom <= 158 && s.Top >= 209, "рамка охватывает все строки штампа");
        }

        private static void TestStampMissingAnchor()
        {
            // Без слова «Действителен» (заменено нейтральным) — одного опорного слова нет.
            List<PdfWord> words = StampWords();
            for (int i = 0; i < words.Count; i++)
                if (words[i].Text == "Действителен") words[i].Text = "Выдан";
            AssertTrue(StampDetector.Detect(words, 595, 842) == null, "нет всех четырёх опорных слов -> не штамп");
        }

        private static void TestStampScatteredRejected()
        {
            // Все четыре слова есть, но раскиданы по всей странице (проза про ЭП) — не компактно.
            var words = new List<PdfWord>
            {
                W("подписан", 50, 800, 50, 9),
                W("Сертификат", 50, 600, 60, 9),
                W("Владелец", 50, 400, 50, 9),
                W("Действителен", 50, 100, 70, 9),
                W("прочее1", 300, 700, 40, 9), W("прочее2", 300, 300, 40, 9)
            };
            AssertTrue(StampDetector.Detect(words, 595, 842) == null, "разбросанные опорные слова -> не штамп");
        }

        // ---------- Loc (локализация) ----------

        private static void TestLocCatalogComplete()
        {
            int n = 0;
            foreach (string key in Loc.Keys)
            {
                string[] p = Loc.Pair(key);
                AssertTrue(p != null && p.Length == 2, "пара для «" + key + "»");
                AssertTrue(!string.IsNullOrEmpty(p[0]), "ru пусто у «" + key + "»");
                AssertTrue(!string.IsNullOrEmpty(p[1]), "en пусто у «" + key + "»");
                n++;
            }
            AssertTrue(n > 100, "каталог непустой (ключей: " + n + ")");
        }

        private static void TestLocPlaceholders()
        {
            // {0},{1}… у ru и en должны совпадать — иначе string.Format кинет на одном из языков.
            foreach (string key in Loc.Keys)
            {
                string[] p = Loc.Pair(key);
                AssertEqual(Placeholders(p[0]), Placeholders(p[1]), "плейсхолдеры {N} расходятся у «" + key + "»");
            }
        }

        private static string Placeholders(string s)
        {
            var set = new SortedSet<int>();
            foreach (System.Text.RegularExpressions.Match x in
                     System.Text.RegularExpressions.Regex.Matches(s, "\\{(\\d+)\\}"))
                set.Add(int.Parse(x.Groups[1].Value));
            return string.Join(",", set);
        }

        private static void TestLocInit()
        {
            Lang saved = Loc.Current;
            try
            {
                Loc.Init(Lang.En); AssertEqual((int)Lang.En, (int)Loc.Current, "Init En");
                Loc.Init(Lang.Ru); AssertEqual((int)Lang.Ru, (int)Loc.Current, "Init Ru");
                AssertEqual((int)Lang.En, (int)Loc.Parse("en"), "Parse en");
                AssertEqual((int)Lang.En, (int)Loc.Parse("EN"), "Parse EN");
                AssertEqual((int)Lang.Ru, (int)Loc.Parse("ru"), "Parse ru");
                AssertEqual((int)Lang.Ru, (int)Loc.Parse("xx"), "Parse неизвестный -> ru");
                AssertEqual("en", Loc.Code(Lang.En), "Code En");
                AssertEqual("ru", Loc.Code(Lang.Ru), "Code Ru");
            }
            finally { Loc.Init(saved); }
        }

        private static void TestNoCyrillicInEnglishForms()
        {
            Lang saved = Loc.Current;
            var offenders = new List<string>();
            // «Язык / Language» — намеренно двуязычный пункт меню; кириллица там ожидаема.
            var whitelist = new HashSet<string> { Loc.Pair("menu.language")[1] };
            var th = new System.Threading.Thread(delegate()
            {
                Loc.Init(Lang.En);
                System.Windows.Forms.Form[] forms;
                try
                {
                    // Тул-формы — с непустым showHub, чтобы кнопка «Главная» (условная) создавалась
                    // и попадала в проверку (иначе захардкоженный перевод в ней остался бы незамеченным).
                    Action back = delegate { };
                    forms = new System.Windows.Forms.Form[]
                    {
                        new MainForm(back), new PdfMergeForm(back), new PdfSplitForm(back),
                        new OcrForm(back), new StartForm(), new AboutForm(), new StatsForm()
                    };
                }
                catch (Exception ex) { offenders.Add("ctor: " + ex.Message); return; }
                foreach (System.Windows.Forms.Form f in forms)
                {
                    CheckCyrillic(f.Text, "Form.Text", offenders, whitelist);
                    WalkControls(f, offenders, whitelist);
                    if (f.MainMenuStrip != null)
                        foreach (System.Windows.Forms.ToolStripItem it in f.MainMenuStrip.Items)
                            WalkMenu(it, offenders, whitelist);
                    try { f.Dispose(); } catch { }
                }
            });
            th.SetApartmentState(System.Threading.ApartmentState.STA);
            th.IsBackground = true;
            th.Start();
            th.Join();
            Loc.Init(saved);
            AssertTrue(offenders.Count == 0, "кириллица в EN: " + string.Join(" | ", offenders.ToArray()));
        }

        private static void WalkControls(System.Windows.Forms.Control c, List<string> offenders, HashSet<string> whitelist)
        {
            foreach (System.Windows.Forms.Control child in c.Controls)
            {
                // Пропускаем поля значений: TextBox (пути/имена/реквизиты), NumericUpDown.
                bool isValue = child is System.Windows.Forms.TextBoxBase || child is System.Windows.Forms.NumericUpDown;
                if (!isValue)
                {
                    CheckCyrillic(child.Text, child.GetType().Name + ".Text", offenders, whitelist);
                    var combo = child as System.Windows.Forms.ComboBox;
                    if (combo != null)
                        foreach (object it in combo.Items)
                            CheckCyrillic(Convert.ToString(it), "ComboItem", offenders, whitelist);
                    var lv = child as System.Windows.Forms.ListView;
                    if (lv != null)
                        foreach (System.Windows.Forms.ColumnHeader col in lv.Columns)
                            CheckCyrillic(col.Text, "Column", offenders, whitelist);
                }
                WalkControls(child, offenders, whitelist);
            }
        }

        private static void WalkMenu(System.Windows.Forms.ToolStripItem item, List<string> offenders, HashSet<string> whitelist)
        {
            CheckCyrillic(item.Text, "Menu", offenders, whitelist);
            var mi = item as System.Windows.Forms.ToolStripMenuItem;
            if (mi != null)
                foreach (System.Windows.Forms.ToolStripItem sub in mi.DropDownItems)
                    WalkMenu(sub, offenders, whitelist);
        }

        private static void CheckCyrillic(string text, string where, List<string> offenders, HashSet<string> whitelist)
        {
            if (string.IsNullOrEmpty(text) || whitelist.Contains(text))
                return;
            foreach (char ch in text)
                if (ch >= 'Ѐ' && ch <= 'ӿ')
                {
                    offenders.Add(where + ": «" + text + "»");
                    return;
                }
        }

        // ---------- TableDetector ----------
        // Хелперы строят линовку и слова в координатах PDF (pt, ось Y вверх).

        private static PdfLine HLine(double y, double x1, double x2)
        {
            return new PdfLine { Orientation = LineOrientation.Horizontal, X1 = x1, Y1 = y, X2 = x2, Y2 = y };
        }

        private static PdfLine VLine(double x, double y1, double y2)
        {
            return new PdfLine { Orientation = LineOrientation.Vertical, X1 = x, Y1 = y1, X2 = x, Y2 = y2 };
        }

        /// <summary>Слово шириной 10×10 с центром (cx, cy).</summary>
        private static PdfWord Word(string text, double cx, double cy)
        {
            return new PdfWord { Text = text, Left = cx - 5, Right = cx + 5, Bottom = cy - 5, Top = cy + 5, FontSizePt = 10 };
        }

        private static string TableCellText(OcrTable t, int row, int col)
        {
            OcrTableCell cell = t.Rows[row].Cells[col];
            return cell.Paragraphs.Count > 0 ? cell.Paragraphs[0].Text : "";
        }

        private static void TestTable2x2()
        {
            // Колонки X=0,50,100; строки Y=0,50,100 (полная сетка 2x2).
            var lines = new List<PdfLine>
            {
                HLine(100, 0, 100), HLine(50, 0, 100), HLine(0, 0, 100),
                VLine(0, 0, 100), VLine(50, 0, 100), VLine(100, 0, 100)
            };
            var words = new List<PdfWord>
            {
                Word("A", 25, 75), Word("B", 75, 75), // верхняя строка
                Word("C", 25, 25), Word("D", 75, 25)  // нижняя строка
            };
            TableDetectResult res = TableDetector.Detect(lines, words, 200, 200);
            AssertEqual(1, res.Tables.Count, "одна таблица");
            OcrTable t = res.Tables[0];
            AssertEqual(2, t.Rows.Count, "2 строки");
            AssertEqual(2, t.ColumnCount, "2 колонки");
            AssertEqual("A", TableCellText(t, 0, 0), "ячейка (0,0)");
            AssertEqual("B", TableCellText(t, 0, 1), "ячейка (0,1)");
            AssertEqual("C", TableCellText(t, 1, 0), "ячейка (1,0)");
            AssertEqual("D", TableCellText(t, 1, 1), "ячейка (1,1)");
            AssertEqual(0, res.RemainingWords.Count, "все слова в таблице");
        }

        private static void TestTableRowSpan()
        {
            // Внутренняя горизонталь Y=50 есть только в правой колонке -> левая ячейка на 2 строки.
            var lines = new List<PdfLine>
            {
                HLine(100, 0, 100), HLine(0, 0, 100), HLine(50, 50, 100),
                VLine(0, 0, 100), VLine(50, 0, 100), VLine(100, 0, 100)
            };
            var words = new List<PdfWord> { Word("L", 25, 50), Word("TR", 75, 75), Word("BR", 75, 25) };
            TableDetectResult res = TableDetector.Detect(lines, words, 200, 200);
            AssertEqual(1, res.Tables.Count, "одна таблица");
            OcrTable t = res.Tables[0];
            AssertEqual(2, t.Rows[0].Cells[0].RowSpan, "левая ячейка на 2 строки");
            AssertTrue(t.Rows[1].Cells[0].Covered, "накрытая позиция под объединением");
            AssertEqual("L", TableCellText(t, 0, 0), "текст объединённой ячейки");
            AssertEqual("TR", TableCellText(t, 0, 1), "правая верхняя");
            AssertEqual("BR", TableCellText(t, 1, 1), "правая нижняя");
        }

        private static void TestTableColSpan()
        {
            // Внутренняя вертикаль X=50 есть только в нижней строке -> верхняя ячейка на 2 колонки.
            var lines = new List<PdfLine>
            {
                HLine(100, 0, 100), HLine(50, 0, 100), HLine(0, 0, 100),
                VLine(0, 0, 100), VLine(100, 0, 100), VLine(50, 0, 50)
            };
            var words = new List<PdfWord> { Word("Top", 50, 75), Word("BL", 25, 25), Word("BR", 75, 25) };
            TableDetectResult res = TableDetector.Detect(lines, words, 200, 200);
            AssertEqual(1, res.Tables.Count, "одна таблица");
            OcrTable t = res.Tables[0];
            AssertEqual(2, t.Rows[0].Cells[0].ColSpan, "верхняя ячейка на 2 колонки");
            AssertTrue(t.Rows[0].Cells[1].Covered, "накрытая позиция справа");
            AssertEqual("Top", TableCellText(t, 0, 0), "текст объединённой ячейки");
            AssertEqual("BL", TableCellText(t, 1, 0), "нижняя левая");
            AssertEqual("BR", TableCellText(t, 1, 1), "нижняя правая");
        }

        private static void TestTableStrayLines()
        {
            // Два подчёркивания (горизонтали без вертикалей) — не таблица.
            var lines = new List<PdfLine> { HLine(60, 0, 40), HLine(20, 0, 40) };
            var words = new List<PdfWord> { Word("x", 20, 60) };
            TableDetectResult res = TableDetector.Detect(lines, words, 200, 200);
            AssertEqual(0, res.Tables.Count, "нет таблиц");
            AssertEqual(1, res.RemainingWords.Count, "слово осталось в потоке");
        }

        private static void TestTableSingleBox()
        {
            // Рамка без внутренних линий — 1x1, не таблица.
            var lines = new List<PdfLine>
            {
                HLine(100, 0, 100), HLine(0, 0, 100), VLine(0, 0, 100), VLine(100, 0, 100)
            };
            var words = new List<PdfWord> { Word("x", 50, 50) };
            TableDetectResult res = TableDetector.Detect(lines, words, 200, 200);
            AssertEqual(0, res.Tables.Count, "рамка 1x1 — не таблица");
            AssertEqual(1, res.RemainingWords.Count, "слово осталось в потоке");
        }

        private static void TestTableWordsOutside()
        {
            var lines = new List<PdfLine>
            {
                HLine(100, 0, 100), HLine(50, 0, 100), HLine(0, 0, 100),
                VLine(0, 0, 100), VLine(50, 0, 100), VLine(100, 0, 100)
            };
            var words = new List<PdfWord> { Word("in", 25, 75), Word("out", 300, 300) };
            TableDetectResult res = TableDetector.Detect(lines, words, 500, 500);
            AssertEqual(1, res.Tables.Count, "одна таблица");
            AssertEqual(1, res.RemainingWords.Count, "внешнее слово в потоке");
            AssertEqual("out", res.RemainingWords[0].Text, "именно внешнее слово");
        }

        private static void TestTableNoLines()
        {
            var words = new List<PdfWord> { Word("a", 10, 10), Word("b", 20, 20) };
            TableDetectResult res = TableDetector.Detect(new List<PdfLine>(), words, 200, 200);
            AssertEqual(0, res.Tables.Count, "нет линий — нет таблиц");
            AssertEqual(2, res.RemainingWords.Count, "все слова в потоке");
        }

        private static void TestUnderlineMarks()
        {
            // Слово [Left..Right]=[10..40], низ Bottom=50; линия у самой базовой линии на всю ширину.
            var w = Word("под", 25, 55); // центр 25 -> Left 20, Right 30; переопределим ниже вручную
            w.Left = 10; w.Right = 40; w.Bottom = 50; w.Top = 60;
            var lines = new List<PdfLine> { HLine(48, 10, 40) }; // на 2 pt ниже низа, вся ширина
            UnderlineDetector.Mark(new List<PdfWord> { w }, lines);
            AssertTrue(w.Underline, "линия у базовой линии на всю ширину -> подчёркнуто");
        }

        private static void TestUnderlineIgnores()
        {
            var far = new PdfWord { Text = "far", Left = 10, Right = 40, Bottom = 50, Top = 60 };
            UnderlineDetector.Mark(new List<PdfWord> { far }, new List<PdfLine> { HLine(30, 10, 40) }); // далеко внизу
            AssertTrue(!far.Underline, "далёкая линия -> не подчёркнуто");

            var shortLine = new PdfWord { Text = "sh", Left = 10, Right = 40, Bottom = 50, Top = 60 };
            UnderlineDetector.Mark(new List<PdfWord> { shortLine }, new List<PdfLine> { HLine(48, 10, 18) }); // покрытие ~27%
            AssertTrue(!shortLine.Underline, "короткая линия -> не подчёркнуто");

            var noLines = new PdfWord { Text = "no", Left = 10, Right = 40, Bottom = 50, Top = 60 };
            UnderlineDetector.Mark(new List<PdfWord> { noLines }, new List<PdfLine>());
            AssertTrue(!noLines.Underline, "нет линий -> не подчёркнуто");
        }

        private static void TestUnderlineWideRule()
        {
            // Полноширинный разделитель под меткой не должен подчёркивать её.
            var w = new PdfWord { Text = "метка", Left = 57, Right = 120, Bottom = 50, Top = 60, FontSizePt = 10 };
            var rule = HLine(48, 57, 520); // длина 463 >> ширины слова 63 (×7)
            UnderlineDetector.Mark(new List<PdfWord> { w }, new List<PdfLine> { rule });
            AssertTrue(!w.Underline, "разделитель во всю ширину -> не подчёркивание");
        }

        /// <summary>Слово с явной рамкой (для тестов колонок/подчёркивания).</summary>
        private static PdfWord WordBox(string text, double left, double right, double bottom)
        {
            return new PdfWord { Text = text, Left = left, Right = right, Bottom = bottom, Top = bottom + 10, FontSizePt = 10 };
        }

        private static void TestSidebarSeparation()
        {
            // 3 строки: слева узкая метка (X57-90), справа тело (X150-350), большой зазор между ними.
            var words = new List<PdfWord>();
            double[] ys = { 220, 200, 180, 160 };
            for (int i = 0; i < ys.Length; i++)
            {
                double y = ys[i];
                words.Add(WordBox("SIDE" + i, 57, 90, y));            // сайдбар-сегмент
                words.Add(WordBox("BODYa" + i, 150, 190, y));         // тело: слова с обычным интервалом
                words.Add(WordBox("BODYb" + i, 195, 235, y));
                words.Add(WordBox("BODYc" + i, 240, 280, y));
                words.Add(WordBox("BODYd" + i, 285, 350, y));
            }
            List<OcrParagraph> paras = OcrLayout.Analyze(words).Paragraphs;
            // Ни один абзац не должен содержать одновременно сайдбар- и тело-токен (нет перемешивания).
            bool anyMixed = false, anySidebarOnly = false;
            foreach (OcrParagraph p in paras)
            {
                bool hasSide = p.Text.Contains("SIDE");
                bool hasBody = p.Text.Contains("BODY");
                if (hasSide && hasBody) anyMixed = true;
                if (hasSide && !hasBody) anySidebarOnly = true;
            }
            AssertTrue(!anyMixed, "сайдбар и тело не смешаны в одном абзаце");
            AssertTrue(anySidebarOnly, "метка сайдбара — отдельным абзацем");
        }

        private static void TestNoSidebarSingleColumn()
        {
            // Одноколоночная строка без больших зазоров: слова обязаны остаться в одном абзаце.
            var words = new List<PdfWord>
            {
                WordBox("one", 50, 90, 200), WordBox("two", 95, 135, 200),
                WordBox("three", 140, 190, 200), WordBox("four", 195, 235, 200)
            };
            List<OcrParagraph> paras = OcrLayout.Analyze(words).Paragraphs;
            AssertEqual(1, paras.Count, "одна строка -> один абзац (сайдбар не сработал)");
            AssertTrue(paras[0].Text.Contains("one") && paras[0].Text.Contains("four"), "все слова в одном абзаце");
        }

        private static void TestCropRect()
        {
            // Страница 500×1000 pt отрендерена в 1000×2000 px (2 px/pt). Картинка Y-вверх:
            // left=100, top=800, w=200, h=100 (занимает Y 700..800). Верх в пикселях = (1000-800)*2.
            System.Drawing.Rectangle r = PageRasterizer.CropRect(1000, 2000, 500, 1000, 100, 800, 200, 100);
            AssertEqual(200, r.X, "X");
            AssertEqual(400, r.Y, "Y (ось вниз)");
            AssertEqual(400, r.Width, "ширина в px");
            AssertEqual(200, r.Height, "высота в px");
            // Кламп: картинка выходит за правый/нижний край — обрезается по границам страницы.
            System.Drawing.Rectangle c = PageRasterizer.CropRect(1000, 2000, 500, 1000, 400, 1000, 200, 100);
            AssertTrue(c.X + c.Width <= 1000, "правый край не за пределами");
            AssertTrue(c.Y >= 0 && c.Y + c.Height <= 2000, "по вертикали в пределах");
        }

        private static void TestReadingOrder()
        {
            // Выше по странице (больший Top) — раньше.
            AssertTrue(WordDocxWriter.CompareReadingOrder(200, 0, 100, 0) < 0, "верхний блок раньше нижнего");
            AssertTrue(WordDocxWriter.CompareReadingOrder(100, 0, 200, 0) > 0, "нижний блок позже верхнего");
            // Одна строка-полоса (|dTop| <= 12): левее — раньше (таблицы бок о бок).
            AssertTrue(WordDocxWriter.CompareReadingOrder(100, 50, 105, 300) < 0, "в одной полосе левый раньше правого");
            AssertTrue(WordDocxWriter.CompareReadingOrder(100, 300, 100, 50) > 0, "в одной полосе правый позже левого");
            // Разница по Top больше полосы — X не важен.
            AssertTrue(WordDocxWriter.CompareReadingOrder(200, 900, 100, 0) < 0, "верхний раньше, несмотря на больший X");
        }

        private static void TestImageCentered()
        {
            // Герб сверху по центру A4 (595 pt): left=281, width=64 -> зазоры 281/250, почти равны.
            AssertTrue(WordDocxWriter.IsImageCentered(281, 64, 595), "герб по центру -> центрируем");
            // Врезка у левого поля: left=72, width=100 -> правый зазор огромный, асимметрия -> нет.
            AssertTrue(!WordDocxWriter.IsImageCentered(72, 100, 595), "врезка слева -> не центр");
            // Печать сбоку справа: left=450, width=100 -> левый зазор огромный -> нет.
            AssertTrue(!WordDocxWriter.IsImageCentered(450, 100, 595), "печать справа -> не центр");
            // Изображение во всю ширину -> не центрируем (нет зазоров).
            AssertTrue(!WordDocxWriter.IsImageCentered(0, 595, 595), "во всю ширину -> не центр");
            // Вырожденная ширина страницы -> не центрируем (защита от деления/мусора).
            AssertTrue(!WordDocxWriter.IsImageCentered(10, 20, 0), "ширина страницы 0 -> не центр");
        }

        private static void TestHasExtractableContent()
        {
            // Пустая страница — нет текста.
            AssertTrue(!PdfToWordService.HasExtractableContent(new PdfPageText()), "пустая страница — нет текста");
            // Только абзац — есть текст.
            var withPar = new PdfPageText();
            withPar.Paragraphs.Add(new OcrParagraph());
            AssertTrue(PdfToWordService.HasExtractableContent(withPar), "абзац — есть текст");
            // Только таблица с текстом в ячейке — есть текст (иначе ложный «скан»).
            var withTable = new PdfPageText();
            var cell = new OcrTableCell();
            cell.Paragraphs.Add(new OcrParagraph());
            var row = new OcrTableRow();
            row.Cells.Add(cell);
            var table = new OcrTable();
            table.Rows.Add(row);
            withTable.Tables.Add(table);
            AssertTrue(PdfToWordService.HasExtractableContent(withTable), "таблица с текстом — есть текст");
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
