using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace ExcelMerger
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            // Язык интерфейса — до создания окон, по старшинству источника: выбор в установщике
            // (одноразовый маркер), затем сохранённый прошлым запуском, затем системная локаль
            // (первый запуск portable-версии — чтобы англоязычному не выскакивал русский).
            // Маркер здесь только ЧИТАЕТСЯ: снимет его GUI-путь ниже, иначе служебный прогон
            // (--selftest и прочие) съел бы выбор пользователя, ничего ему не показав.
            UserSettings startup = UserSettings.Load();
            string setupLanguage = SetupLanguage.Read();
            Loc.Init(setupLanguage != null ? Loc.Parse(setupLanguage)
                : startup.Language != null ? Loc.Parse(startup.Language)
                : Loc.DefaultForCulture(System.Globalization.CultureInfo.CurrentUICulture.Name));
            if (args.Length >= 1 && string.Equals(args[0], "--selftest", StringComparison.OrdinalIgnoreCase))
                return RunSelfTest();
            if (args.Length >= 1 && string.Equals(args[0], "--pdfcheck", StringComparison.OrdinalIgnoreCase))
                return RunPdfCheck();
            if (args.Length >= 1 && string.Equals(args[0], "--pdftextcheck", StringComparison.OrdinalIgnoreCase))
                return RunPdfTextCheck();
            if (args.Length >= 1 && string.Equals(args[0], "--pptxcheck", StringComparison.OrdinalIgnoreCase))
                return RunPptxCheck();
            if (args.Length >= 1 && string.Equals(args[0], "--thumbcheck", StringComparison.OrdinalIgnoreCase))
                return RunThumbCheck();
            if (args.Length >= 1 && string.Equals(args[0], "--gscheck", StringComparison.OrdinalIgnoreCase))
                return RunGsCheck();
            // PDF-команды: те же операции, что и в окнах, для пакетной обработки и сценариев.
            // Стоят рядом с режимом Excel и до создания окон — интерфейс им не нужен.
            if (PdfCli.IsCommand(args))
            {
                AttachConsole(-1);
                int code = PdfCli.Execute(PdfCli.Parse(args), WriteConsole);
                FastExit.Now(code);   // как прочие безоконные режимы: без выгрузки WinRT
                return code;
            }
            if (args.Length >= 3 && string.Equals(args[0], "--cli", StringComparison.OrdinalIgnoreCase))
            {
                MergeOptions options;
                string parseError;
                if (!TryParseCliOptions(args, 3, out options, out parseError))
                {
                    AttachConsole(-1);
                    WriteConsole("ERROR: " + parseError);
                    WriteConsole(Loc.T("cli.usage"));
                    return 1;
                }
                return RunCli(args[1], args[2], options);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            CrashReport.Install(); // последний рубеж: фирменный диалог + crash.log вместо тихого краха
            // Второй запуск ярлыка не плодит процесс (в диспетчере задач была бы вторая
            // запись «приложение»): будим работающий экземпляр и уходим. Служебные режимы
            // выше этой проверки — они обязаны запускаться всегда.
            if (!SingleInstance.TryAcquire())
            {
                SingleInstance.SignalExisting();
                return 0; // WinRT здесь не загружался, обычного возврата достаточно
            }
            if (setupLanguage != null)
            {
                // Выбор установщика применён: убираем маркер и переносим язык в настройки
                // приложения штатным путём (UTF-8, прочие поля на месте). Дальше язык живёт
                // там, и следующая смена в самой программе уже никем не перебивается.
                SetupLanguage.Consume();
                UserSettings.Load().Save();
            }
            var shell = new ShellContext(); // хаб + инструменты как независимые окна
            // Проверка обновлений идёт ФОНОМ и показа окна не задерживает: запрос уходит
            // с таймаутом 10 с, а ответ доставляется в UI-поток уже работающего хаба.
            shell.CheckForUpdatesOnStart();
            Application.Run(shell);
            // Все окна закрыты, настройки/COM уже освобождены детерминированно.
            // Форсируем выход, чтобы финализация WinRT (Windows.Data.Pdf, если
            // открывали инструмент PDF) не уронила процесс при выгрузке CLR.
            FastExit.Now(0);
            return 0; // недостижимо
        }

        /// <summary>Флаги CLI после обязательных аргументов: --toc, --values. Неизвестный флаг — ошибка.</summary>
        internal static bool TryParseCliOptions(string[] args, int startIndex, out MergeOptions options, out string error)
        {
            options = new MergeOptions();
            error = null;
            for (int i = startIndex; i < args.Length; i++)
            {
                if (string.Equals(args[i], "--toc", StringComparison.OrdinalIgnoreCase))
                    options.AddToc = true;
                else if (string.Equals(args[i], "--values", StringComparison.OrdinalIgnoreCase))
                    options.ValuesOnly = true;
                else if (string.Equals(args[i], "--allsheets", StringComparison.OrdinalIgnoreCase))
                    options.AllSheets = true;
                else
                {
                    error = string.Format(Loc.T("cli.unknownFlag"), args[i]);
                    return false;
                }
            }
            return true;
        }

        [DllImport("kernel32.dll")]
        private static extern bool AttachConsole(int processId);

        /// <summary>
        /// Смоук-тест GUI без показа окна: конструктор формы и создание хэндла.
        /// 0 = OK, 3 = окно не создалось, 4 = в exe нет руководства пользователя.
        /// </summary>
        private static int RunSelfTest()
        {
            try
            {
                // Руководство лежит ресурсом и открывается из «О программе». Проверяем его
                // здесь, а не в юнит-тестах: там своя сборка, и вшито ли оно в НАСТОЯЩИЙ exe,
                // она сказать не может. Проверяем ОБА языка: ресурс держится одной строкой в
                // csproj, и пропажа английского не выдала бы себя ничем, пока англоязычный
                // читатель не нажмёт «открыть» и не получит ошибку.
                Lang wasLanguage = Loc.Current;
                foreach (Lang language in new[] { Lang.Ru, Lang.En })
                {
                    Loc.Init(language);
                    if (!UserManual.IsPacked())
                    {
                        Loc.Init(wasLanguage);
                        return 4;
                    }
                }
                Loc.Init(wasLanguage);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Action noop = delegate { };
                using (var form = new MainForm(noop)) // с кнопкой «Назад в меню»
                {
                    IntPtr handle = form.Handle; // форсирует создание окна без показа
                    if (handle == IntPtr.Zero)
                        return 3;
                }
                using (var about = new AboutForm())
                {
                    IntPtr handle = about.Handle;
                    if (handle == IntPtr.Zero)
                        return 3;
                }
                using (var stats = new StatsForm())
                {
                    IntPtr handle = stats.Handle;
                    if (handle == IntPtr.Zero)
                        return 3;
                }
                using (var settings = new SettingsForm())
                {
                    IntPtr handle = settings.Handle;
                    if (handle == IntPtr.Zero)
                        return 3;
                }
                using (var history = new HistoryForm())
                {
                    IntPtr handle = history.Handle;
                    if (handle == IntPtr.Zero)
                        return 3;
                }
                using (var pdf = new PdfMergeForm(noop))
                {
                    IntPtr handle = pdf.Handle;
                    if (handle == IntPtr.Zero)
                        return 3;
                }
                using (var split = new PdfSplitForm(noop))
                {
                    IntPtr handle = split.Handle;
                    if (handle == IntPtr.Zero)
                        return 3;
                }
                using (var ocr = new OcrForm(noop))
                {
                    IntPtr handle = ocr.Handle;
                    if (handle == IntPtr.Zero)
                        return 3;
                }
                using (var pptx = new PptxForm(noop))
                {
                    IntPtr handle = pptx.Handle;
                    if (handle == IntPtr.Zero)
                        return 3;
                }
                using (var ops = new PdfOpsForm(noop))
                {
                    IntPtr handle = ops.Handle;
                    if (handle == IntPtr.Zero)
                        return 3;
                }
                using (var start = new StartForm())
                {
                    IntPtr handle = start.Handle;
                    if (handle == IntPtr.Zero)
                        return 3;
                }
                return 0;
            }
            catch (Exception ex)
            {
                try { File.WriteAllText("selftest.log", ex.ToString()); } catch { }
                return 3;
            }
        }

        /// <summary>
        /// Проверка резолва вшитого PdfSharp из ресурса exe. LoadPages ссылается
        /// на типы PdfSharp — если сборка не резолвится, JIT падает с
        /// FileNotFoundException ДО тела метода; если резолвится — получаем нашу
        /// MergeException про несуществующий файл. 0 = резолв работает.
        /// </summary>
        private static int RunPdfCheck()
        {
            AttachConsole(-1);
            try
            {
                PdfMergeService.LoadPages(Path.Combine(Path.GetTempPath(),
                    "нет_такого_" + Guid.NewGuid().ToString("N") + ".pdf"));
                WriteConsole("PDFCHECK: неожиданно без ошибки");
                return 2;
            }
            catch (MergeException)
            {
                WriteConsole("PDFCHECK OK"); // PdfSharp резолвнулся, дошли до нашей ошибки
                return 0;
            }
            catch (Exception ex)
            {
                WriteConsole("PDFCHECK FAIL: " + ex.GetType().Name + " — " + ex.Message);
                return 1;
            }
        }

        /// <summary>
        /// Проверка born-digital пути: вшитый PdfSharp рисует текст, вшитый PdfPig
        /// (резолвится из ресурса) извлекает его обратно. 0 = вшитый PdfPig работает.
        /// </summary>
        private static int RunPdfTextCheck()
        {
            AttachConsole(-1);
            EmbeddedAssemblies.Ensure();
            string pdf = Path.Combine(Path.GetTempPath(), "pdftextchk_" + Guid.NewGuid().ToString("N") + ".pdf");
            try
            {
                WriteTextPdf(pdf);
                List<PdfPageText> pages = PdfTextExtract.Extract(pdf);
                string text = pages.Count > 0 ? pages[0].Text : "";
                bool ok = pages.Count == 1 && text.Contains("Hello world") && text.Contains("Привет");
                WriteConsole(ok ? "PDFTEXTCHECK OK" : "PDFTEXTCHECK FAIL: pages=" + pages.Count + " text=[" + text + "]");
                return ok ? 0 : 1;
            }
            catch (Exception ex)
            {
                WriteConsole("PDFTEXTCHECK FAIL: " + ex.GetType().Name + " — " + ex.Message);
                return 1;
            }
            finally
            {
                try { File.Delete(pdf); } catch { }
            }
        }

        /// <summary>
        /// Сквозная проверка «PDF → PowerPoint» без Office: пишем PDF с известным текстом,
        /// прогоняем конвертацию и убеждаемся, что в собранном .pptx есть слайд и этот текст.
        /// Разбирается ZIP-ом — ровно так же, как это сделает читатель формата.
        /// </summary>
        private static int RunPptxCheck()
        {
            AttachConsole(-1);
            EmbeddedAssemblies.Ensure();
            string pdf = Path.Combine(Path.GetTempPath(), "pptxchk_" + Guid.NewGuid().ToString("N") + ".pdf");
            string pptx = Path.ChangeExtension(pdf, ".pptx");
            try
            {
                WriteTextPdf(pdf);
                var order = new List<PdfPageRef> { new PdfPageRef { SourcePath = pdf, PageIndex = 0 } };
                ConvertResult result = PdfToPptxService.Convert(order, pptx);
                string slide = ReadZipEntry(pptx, "ppt/slides/slide1.xml");
                bool ok = result.Pages == 1 && slide != null
                    && slide.Contains("Hello world") && slide.Contains("Привет");
                WriteConsole(ok ? "PPTXCHECK OK"
                    : "PPTXCHECK FAIL: pages=" + result.Pages + " slide=" + (slide == null ? "нет" : slide.Length.ToString()));
                return ok ? 0 : 1;
            }
            catch (Exception ex)
            {
                WriteConsole("PPTXCHECK FAIL: " + ex.GetType().Name + " — " + ex.Message);
                return 1;
            }
            finally
            {
                try { File.Delete(pdf); } catch { }
                try { File.Delete(pptx); } catch { }
            }
        }

        /// <summary>Текст части ZIP-архива (UTF-8) или null, если её нет.</summary>
        private static string ReadZipEntry(string archive, string entryName)
        {
            using (var zip = System.IO.Compression.ZipFile.OpenRead(archive))
            {
                System.IO.Compression.ZipArchiveEntry entry = zip.GetEntry(entryName);
                if (entry == null)
                    return null;
                using (Stream s = entry.Open())
                using (var reader = new StreamReader(s, System.Text.Encoding.UTF8))
                    return reader.ReadToEnd();
            }
        }

        /// <summary>Пишет 1-страничный PDF с известным текстом через вшитый PdfSharp. NoInlining: в теле типы PdfSharp.</summary>
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
        private static void WriteTextPdf(string path)
        {
            var doc = new PdfSharp.Pdf.PdfDocument();
            PdfSharp.Pdf.PdfPage page = doc.AddPage();
            using (PdfSharp.Drawing.XGraphics g = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
            {
                var font = new PdfSharp.Drawing.XFont("Times New Roman", 14, PdfSharp.Drawing.XFontStyle.Regular);
                g.DrawString("Привет, мир! Hello world.", font, PdfSharp.Drawing.XBrushes.Black,
                    new PdfSharp.Drawing.XPoint(50, 100));
            }
            doc.Save(path);
            doc.Dispose();
        }

        /// <summary>
        /// Проверка рендера миниатюр на фоновом MTA-потоке (как в приложении): (1) процесс не
        /// роняется при выходе; (2) РЕГРЕССИЯ ERROR_USER_MAPPED_FILE — пока рендерер жив (документ в
        /// кэше), исходный файл НЕ отображён в память, поэтому его можно перезаписать под тем же
        /// именем (иначе объединение/разделение не сохранит итог поверх показанного файла). Само
        /// наличие рендера не требуется приложению (на Windows Server движка PDF может не быть —
        /// штатный фолбэк), но БЕЗ рендера проверять нечего, поэтому это отдельный код выхода, а
        /// не «ок»: иначе шаг CI зеленел бы, ничего не проверив, и потеря защиты от
        /// ERROR_USER_MAPPED_FILE прошла бы незаметно.
        /// 0 = проверено и ок; 1 = отрендеренный файл остался замаплен (регрессия);
        /// 2 = проверить не удалось (движок недоступен или проба упала).
        /// </summary>
        private static int RunThumbCheck()
        {
            AttachConsole(-1);
            bool rendered = false;
            bool overwritable = true; // по умолчанию — если рендер не случился, файл не мапился
            string failure = null;    // причина, по которой проверить не вышло (иначе «тихий» ноль)
            var worker = new System.Threading.Thread(delegate()
            {
                string pdf = Path.Combine(Path.GetTempPath(), "thumbchk_" + Guid.NewGuid().ToString("N") + ".pdf");
                try
                {
                    PdfProbe.WriteOnePagePdf(pdf);
                    using (var renderer = new PdfThumbnailRenderer())
                    {
                        using (System.Drawing.Bitmap bmp = renderer.Render(pdf, 0, 100))
                            rendered = bmp != null;
                        // Пока рендерер жив, перезаписать файл на месте (ровно та операция, что делает
                        // сохранение объединения). Если WinRT держит файл замапленным — бросит
                        // ERROR_USER_MAPPED_FILE, и overwritable=false.
                        if (rendered)
                        {
                            try { PdfProbe.WriteOnePagePdf(pdf); }
                            catch { overwritable = false; }
                        }
                    }
                }
                catch (Exception ex) { failure = ex.GetType().Name + ": " + ex.Message; }
                finally { try { File.Delete(pdf); } catch { } }
            });
            worker.IsBackground = false; // дать потоку завершиться штатно
            worker.Start();
            worker.Join(); // синхронизация: rendered/overwritable/failure читаем после завершения потока
            int code = !rendered ? 2 : (overwritable ? 0 : 1);
            WriteConsole(
                code == 0 ? "THUMBCHECK OK (rendered=True, overwritable=True)" :
                code == 1 ? "THUMBCHECK FAIL: отрендеренный файл остался замаплен (нельзя сохранить итог под тем же именем)" :
                "THUMBCHECK SKIP: рендер не состоялся, проверить нечего" +
                    (failure != null ? " — " + failure : " (движок PDF недоступен)"));
            FastExit.Now(code); // избегаем сбоя WinRT-финализации при выгрузке
            return code; // недостижимо
        }

        /// <summary>
        /// Проверка движка сжатия (Ghostscript). Фича опциональна: если GS не найден —
        /// это не ошибка (exit 0). Если найден — реальный round-trip через наши
        /// аргументы: создаём 1-страничный PDF, сжимаем в temp, проверяем код выхода
        /// и заголовок %PDF-. 0 = либо GS нет, либо сжатие отработало.
        /// </summary>
        private static int RunGsCheck()
        {
            AttachConsole(-1);
            if (!Ghostscript.Available)
            {
                WriteConsole("GSCHECK OK (Ghostscript не установлен — сжатие опционально)");
                return 0;
            }
            string dir = Path.Combine(Path.GetTempPath(), "gschk_" + Guid.NewGuid().ToString("N"));
            string src = Path.Combine(dir, "in.pdf");
            string outp = Path.Combine(dir, "out.pdf");
            try
            {
                Directory.CreateDirectory(dir);
                PdfProbe.WriteOnePagePdf(src);
                string args2 = PdfCompression.BuildArguments(src, outp, CompressionLevel.Good, Ghostscript.BundledRoot);
                string stderr;
                int exit = Ghostscript.Run(args2, 60000, out stderr);
                bool ok = exit == 0 && File.Exists(outp) && PdfCompression.LooksLikePdf(outp);
                WriteConsole(ok
                    ? "GSCHECK OK (exe=" + Ghostscript.Exe + ")"
                    : "GSCHECK FAIL: exit=" + exit + " stderr=" + stderr);
                return ok ? 0 : 1;
            }
            catch (Exception ex)
            {
                WriteConsole("GSCHECK FAIL: " + ex.GetType().Name + " — " + ex.Message);
                return 1;
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        /// <summary>
        /// Режим для автотестов и скриптования:
        /// ExcelMerger.exe --cli &lt;папка&gt; &lt;итоговый.xlsx&gt; [--toc] [--values].
        /// Отчёт пишется рядом с итоговым файлом (&lt;имя&gt;.report.txt) и, если возможно, в консоль.
        /// Коды выхода: 0 — все файлы перенесены, 2 — есть пропущенные, 1 — ошибка.
        /// </summary>
        private static int RunCli(string inputFolder, string outputPath, MergeOptions options)
        {
            AttachConsole(-1);
            var lines = new List<string>();
            string fullOutput = Path.GetFullPath(outputPath);
            string summary;
            int exitCode;
            try
            {
                var service = new MergeService();
                service.Trace += delegate(string msg) { lines.Add("  trace: " + msg); };
                service.FileDone += delegate(FileResult fr)
                {
                    string line = ReportWriter.FormatFileLine(fr);
                    lines.Add(line);
                    WriteConsole(line);
                };
                // CLI берёт все файлы папки в естественном порядке.
                List<string> files = MergeService.FindSourceFiles(Path.GetFullPath(inputFolder), fullOutput);
                MergeResult result = service.Merge(files, fullOutput, options);
                if (result.TocError != null)
                {
                    lines.Add("WARN: " + result.TocError);
                    WriteConsole("WARN: " + result.TocError);
                }
                summary = "RESULT: ok=" + result.OkCount + " skipped=" + result.SkipCount +
                    " output=" + result.OutputPath;
                lines.Add(summary);
                exitCode = result.SkipCount == 0 ? 0 : 2;
            }
            catch (Exception ex)
            {
                summary = "ERROR: " + ex.Message.Replace("\r", " ").Replace("\n", " ");
                lines.Add(summary);
                lines.Add("TRACE: " + ex);
                exitCode = 1;
            }
            lines.Add("EXITCODE: " + exitCode);
            WriteConsole(summary);
            try { File.WriteAllLines(fullOutput + ".report.txt", lines, Encoding.UTF8); } catch { }
            return exitCode;
        }

        private static void WriteConsole(string line)
        {
            try { Console.WriteLine(line); } catch { }
        }
    }
}
