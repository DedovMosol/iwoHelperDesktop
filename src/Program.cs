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
            if (args.Length >= 1 && string.Equals(args[0], "--selftest", StringComparison.OrdinalIgnoreCase))
                return RunSelfTest();
            if (args.Length >= 1 && string.Equals(args[0], "--pdfcheck", StringComparison.OrdinalIgnoreCase))
                return RunPdfCheck();
            if (args.Length >= 1 && string.Equals(args[0], "--thumbcheck", StringComparison.OrdinalIgnoreCase))
                return RunThumbCheck();
            if (args.Length >= 3 && string.Equals(args[0], "--cli", StringComparison.OrdinalIgnoreCase))
            {
                MergeOptions options;
                string parseError;
                if (!TryParseCliOptions(args, 3, out options, out parseError))
                {
                    AttachConsole(-1);
                    WriteConsole("ERROR: " + parseError);
                    WriteConsole("Использование: ExcelMerger.exe --cli <папка> <итоговый.xlsx> [--toc] [--values]");
                    return 1;
                }
                return RunCli(args[1], args[2], options);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new StartForm());
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
                    error = "неизвестный параметр «" + args[i] + "»";
                    return false;
                }
            }
            return true;
        }

        [DllImport("kernel32.dll")]
        private static extern bool AttachConsole(int processId);

        /// <summary>Смоук-тест GUI без показа окна: конструктор формы и создание хэндла. 0 = OK.</summary>
        private static int RunSelfTest()
        {
            try
            {
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
                using (var pdf = new PdfMergeForm(noop))
                {
                    IntPtr handle = pdf.Handle;
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
        /// Проверка, что рендер миниатюр на фоновом MTA-потоке (как в приложении)
        /// не роняет процесс при выходе. Само наличие рендера не требуется
        /// (на Windows Server движка PDF может не быть — это штатный фолбэк);
        /// проверяется только чистое завершение. 0 = процесс завершился нормально.
        /// </summary>
        private static int RunThumbCheck()
        {
            AttachConsole(-1);
            bool rendered = false;
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
                    }
                }
                catch { }
                finally { try { File.Delete(pdf); } catch { } }
            });
            worker.IsBackground = false; // дать потоку завершиться штатно
            worker.Start();
            worker.Join();
            WriteConsole("THUMBCHECK OK (rendered=" + rendered + ")");
            FastExit.Now(0); // избегаем сбоя WinRT-финализации при выгрузке
            return 0;
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
