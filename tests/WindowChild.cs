using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Дочерний процесс с настоящими окнами для тестов правил окон (страховка после Quit, сторож
    /// вызова). Это тот же UnitTests.exe с ключом <see cref="Switch"/>: стартует за доли секунды,
    /// той же разрядности, что и прогон. Раньше дочерним был Windows PowerShell с WinForms, и на
    /// холодном раннере CI (32-битный PowerShell без готовых образов сборок) он не успевал
    /// создать окна за 60 с — тесты падали без дефекта в программе.
    ///
    /// Протокол: первая строка stdout — дескриптор опорного окна. Команда «close-first» в stdin
    /// закрывает первое окно раскладки «два окна» и отвечает «closed». Конец stdin (родитель
    /// закрыл поток или умер) завершает процесс; предел жизни — 5 минут, чтобы потерянный
    /// дочерний процесс не остался навсегда.
    /// </summary>
    internal static class WindowChild
    {
        internal const string Switch = "--window-child";

        internal const string TwoVisible = "two-visible";     // первое окно и второе видимое
        internal const string TwoHidden = "two-hidden";       // первое окно и второе скрытое
        internal const string DialogHiddenFrame = "dialog-hidden-frame";   // как скрытый Office с модальным окном
        internal const string DialogVisibleFrame = "dialog-visible-frame"; // как Office человека с модальным окном

        internal const string DialogTitle = "Проверка сторожа";

        private const int LifetimeMs = 5 * 60 * 1000;

        /// <summary>Точка входа дочернего процесса; код возврата 2 — неизвестная раскладка.</summary>
        internal static int Run(string layout)
        {
            int code = 0;
            var ui = new Thread(delegate() { code = RunLayout(layout); });
            ui.SetApartmentState(ApartmentState.STA); // WinForms
            ui.Start();
            ui.Join();
            return code;
        }

        private static int RunLayout(string layout)
        {
            Form first;
            Form close = null; // окно, которое закрывает «close-first»
            switch (layout)
            {
                case TwoVisible:
                case TwoHidden:
                {
                    first = NewWindow("Первое", true);
                    first.Show();
                    Form second = NewWindow("Второе", true);
                    if (layout == TwoVisible)
                        second.Show();
                    else
                        second.CreateControl(); // окно создано, но не показано
                    close = first;
                    break;
                }
                case DialogHiddenFrame:
                case DialogVisibleFrame:
                {
                    // Скрытая рамка — без своей кнопки на панели задач она и так не видна; видимая —
                    // без кнопки, чтобы тест не мелькал на панели задач.
                    bool visible = layout == DialogVisibleFrame;
                    first = NewWindow("Рамка", visible);
                    if (visible)
                        first.Show();
                    else
                        first.CreateControl();
                    // ShowInTaskbar по умолчанию: так владельцем окна становится именно рамка
                    // (проверено перебором окон). Окно за пределами экрана.
                    Form dialog = NewWindow(DialogTitle, false);
                    dialog.Show(first);
                    first.Enabled = false; // так рамку выключает модальное окно
                    break;
                }
                default:
                    Console.Error.WriteLine("неизвестная раскладка: " + layout);
                    return 2;
            }

            Console.Out.WriteLine(first.Handle.ToInt64());
            Console.Out.Flush();

            var stdin = new Thread(delegate()
            {
                // При кодовой странице консоли UTF-8 (так на раннерах CI) Process.StandardInput
                // родителя пишет BOM перед первой строкой, а Console.In дочернего процесса без
                // консоли читает по кодовой странице ANSI: BOM превращался в «п»ї», команда не
                // узнавалась, тест ждал ответа 30 с (воспроизведено локально через chcp 65001).
                // Поэтому stdin читается как UTF-8 с распознаванием и снятием BOM.
                var input = new System.IO.StreamReader(Console.OpenStandardInput(), new System.Text.UTF8Encoding(false), true);
                string line;
                while ((line = input.ReadLine()) != null)
                {
                    // BOM может прийти и не один (StreamWriter родителя пишет свой при AutoFlush),
                    // а U+FEFF — не пробел для Trim.
                    line = line.Trim().TrimStart('\uFEFF').Trim();
                    if (line == "close-first" && close != null)
                        first.BeginInvoke((Action)delegate
                        {
                            close.Close();
                            Console.Out.WriteLine("closed");
                            Console.Out.Flush();
                        });
                }
                Environment.Exit(0); // родитель закрыл поток или завершился
            });
            stdin.IsBackground = true;
            stdin.Start();

            var lifetime = new System.Windows.Forms.Timer { Interval = LifetimeMs };
            lifetime.Tick += delegate { Environment.Exit(3); };
            lifetime.Start();
            Application.Run();
            return 0;
        }

        /// <summary>Окно за пределами экрана; ShowInTaskbar = false, если skipTaskbar.</summary>
        private static Form NewWindow(string text, bool skipTaskbar)
        {
            var form = new Form();
            form.Text = text;
            form.StartPosition = FormStartPosition.Manual;
            form.Location = new Point(-32000, -32000);
            form.ShowInTaskbar = !skipTaskbar;
            return form;
        }

        /// <summary>
        /// Запустить дочерний процесс с раскладкой layout и дождаться дескриптора опорного окна.
        /// Окон консоли у процесса нет: его верхние окна — только окна раскладки (и скрытые
        /// служебные окна WinForms).
        /// </summary>
        internal static Process Start(string layout, out long window)
        {
            var info = new ProcessStartInfo(Assembly.GetEntryAssembly().Location, Switch + " " + layout);
            info.UseShellExecute = false;
            info.CreateNoWindow = true;
            info.RedirectStandardInput = true;
            info.RedirectStandardOutput = true;
            Process child = Process.Start(info);
            try
            {
                string line = ReadLine(child);
                if (!long.TryParse(line, out window))
                    throw new Exception("дочерний процесс не сообщил окно: «" + line + "»");
                return child;
            }
            catch
            {
                Stop(child);
                child.Dispose();
                throw;
            }
        }

        /// <summary>Строка ответа дочернего процесса; не дольше 30 с (обычный ответ — доли секунды).</summary>
        internal static string ReadLine(Process child)
        {
            var line = child.StandardOutput.ReadLineAsync();
            if (!line.Wait(30000))
                throw new Exception("дочерний процесс не ответил за 30 с");
            return line.Result;
        }

        /// <summary>Завершить дочерний процесс, если он ещё жив.</summary>
        internal static void Stop(Process child)
        {
            try
            {
                if (!child.HasExited)
                {
                    child.Kill();
                    child.WaitForExit(5000);
                }
            }
            catch { } // уже вышел
        }
    }
}
