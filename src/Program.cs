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
            Application.Run(new MainForm());
            return 0;
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
                using (var form = new MainForm())
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
                return 0;
            }
            catch (Exception ex)
            {
                try { File.WriteAllText("selftest.log", ex.ToString()); } catch { }
                return 3;
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
                MergeResult result = service.Merge(Path.GetFullPath(inputFolder), fullOutput, options);
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
