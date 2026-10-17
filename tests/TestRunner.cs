using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Тест, который подключается сам: метод <c>private static void X()</c> с этим атрибутом
    /// выполняется после явного списка Run(...) в порядке объявления. Новые тесты пишутся так —
    /// строку в Main добавлять не нужно, а забытая регистрация невозможна.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    internal sealed class TestAttribute : Attribute
    {
        public TestAttribute(string name) { Name = name; }
        public string Name { get; private set; }
    }

    /// <summary>
    /// Тест честно сообщает, что проверить нечего (нет Ghostscript, нет WinRT): это не PASS.
    /// Пропуск считается отдельно и в строгом режиме (IWO_STRICT=1) валит прогон.
    /// </summary>
    internal sealed class SkipTestException : Exception
    {
        public SkipTestException(string reason) : base(reason) { }
    }

    /// <summary>
    /// Обвязка мини-раннера: фильтр по имени, сторож зависаний, перехват исключений чужих
    /// потоков, изоляция профиля и языка между тестами, полная диагностика отказа.
    /// </summary>
    internal static partial class UnitTests
    {
        private static int _skipped;
        private static string _filter;
        private static bool _listOnly;
        private static bool _strictSkips;
        private static int _timeoutSeconds = 600;
        private static readonly HashSet<MethodInfo> _registeredMethods = new HashSet<MethodInfo>();

        /// <summary>
        /// Каталог профиля на весь прогон. Тесты по умолчанию НИКОГДА не видят настоящий
        /// %APPDATA%: живые окна читают и сохраняют настройки, и без этого прогон менял бы
        /// настройки разработчика, а его настройки — исход теста.
        /// </summary>
        private static string _sessionRoot;

        // Сторож: имя и время старта текущего теста. Пишет UI-поток раннера, читает сторож.
        private static volatile string _currentTest;
        private static long _currentStartedTicks;

        private static bool ConfigureRunner(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                if (string.Equals(a, "--filter", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                    _filter = args[++i];
                else if (string.Equals(a, "--list", StringComparison.OrdinalIgnoreCase))
                    _listOnly = true;
                else if (string.Equals(a, "--timeout", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    int seconds;
                    if (!int.TryParse(args[++i], out seconds) || seconds <= 0)
                    {
                        Console.WriteLine("ОШИБКА: --timeout ждёт положительное число секунд");
                        return false;
                    }
                    _timeoutSeconds = seconds;
                }
                else
                {
                    Console.WriteLine("Использование: UnitTests.exe [--filter подстрока] [--list] [--timeout секунды]");
                    return false;
                }
            }
            _strictSkips = Environment.GetEnvironmentVariable("IWO_STRICT") == "1";

            _sessionRoot = Path.Combine(Path.GetTempPath(), "iwo_tests_" + Guid.NewGuid().ToString("N"));
            AppPaths.SetRootForTests(_sessionRoot);

            // Исключение в обработчике события окна по умолчанию открывает модальный
            // ThreadExceptionDialog — прогон молча висел бы до таймаута CI. Пусть летит в тест.
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
            AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e)
            {
                // Падение чужого (фонового) потока: процесс уже не спасти, но причина и
                // виновный тест обязаны попасть в лог, а код возврата — быть ненулевым.
                Console.WriteLine("CRASH " + (_currentTest ?? "(вне теста)") + " — необработанное исключение потока:");
                Console.WriteLine(Indent(Convert.ToString(e.ExceptionObject)));
                Console.Out.Flush();
                FastExit.Now(3);
            };
            StartWatchdog();
            return true;
        }

        private static bool Selected(string name)
        {
            return _filter == null || (name ?? "").IndexOf(_filter, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void RunCore(string name, Action test)
        {
            if (!Selected(name))
                return;
            if (_listOnly)
            {
                Console.WriteLine(name);
                return;
            }
            Lang lang = Loc.Current;
            long budget = PdfMemoryBudget.Used;
            _currentStartedTicks = DateTime.UtcNow.Ticks;
            _currentTest = name;
            try
            {
                test();
                _passed++;
                Console.WriteLine("PASS  " + name);
            }
            catch (SkipTestException skip)
            {
                _skipped++;
                Console.WriteLine("SKIP  " + name + " — " + skip.Message);
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("FAIL  " + name + " — " + ex.Message);
                // Тип, стек и вложенные исключения: «Object reference not set» без стека не чинится.
                Console.WriteLine(Indent(ex.ToString()));
            }
            finally
            {
                _currentTest = null;
                // Изоляция: упавший на полпути тест не оставляет следующим чужой язык и профиль.
                if (Loc.Current != lang)
                    Loc.Init(lang);
                AppPaths.SetRootForTests(_sessionRoot);
                long leaked = PdfMemoryBudget.Used - budget;
                if (leaked > 0)
                    Console.WriteLine("WARN  " + name + " — резерв PdfMemoryBudget не вернулся: +" + leaked + " байт");
            }
        }

        /// <summary>Тесты с [Test] — после явного списка, в порядке объявления в исходнике.</summary>
        private static void RunDiscoveredTests()
        {
            var tests = new List<KeyValuePair<MethodInfo, TestAttribute>>();
            foreach (MethodInfo m in typeof(UnitTests).GetMethods(BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public))
            {
                var attr = (TestAttribute)Attribute.GetCustomAttribute(m, typeof(TestAttribute));
                if (attr != null)
                    tests.Add(new KeyValuePair<MethodInfo, TestAttribute>(m, attr));
            }
            tests.Sort(delegate(KeyValuePair<MethodInfo, TestAttribute> a, KeyValuePair<MethodInfo, TestAttribute> b)
            {
                return a.Key.MetadataToken.CompareTo(b.Key.MetadataToken);
            });
            foreach (KeyValuePair<MethodInfo, TestAttribute> t in tests)
            {
                MethodInfo method = t.Key;
                if (method.GetParameters().Length != 0 || method.ReturnType != typeof(void))
                {
                    _failed++;
                    Console.WriteLine("FAIL  " + t.Value.Name + " — [Test] допустим только на void-методе без параметров");
                    continue;
                }
                RunCore(t.Value.Name, delegate
                {
                    try { method.Invoke(null, null); }
                    catch (TargetInvocationException ex) when (ex.InnerException != null)
                    {
                        System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                    }
                });
            }
        }

        /// <summary>Методы вида <c>static void TestX()</c>, которых нет ни в Run(...), ни с [Test].</summary>
        private static string UnregisteredTestMethods()
        {
            var names = new List<string>();
            foreach (MethodInfo m in typeof(UnitTests).GetMethods(BindingFlags.Static | BindingFlags.NonPublic))
            {
                if (!m.Name.StartsWith("Test", StringComparison.Ordinal) || m.GetParameters().Length != 0 ||
                    m.ReturnType != typeof(void) || m.IsDefined(typeof(TestAttribute), false))
                    continue;
                if (!_registeredMethods.Contains(m))
                    names.Add(m.Name);
            }
            return string.Join(", ", names.ToArray());
        }

        /// <summary>Честный пропуск: см. <see cref="SkipTestException"/>.</summary>
        private static void Skip(string reason)
        {
            throw new SkipTestException(reason);
        }

        /// <summary>
        /// Сторож зависаний: тест, ждущий модальное окно или вечный Join, иначе держал бы CI
        /// до шестичасового лимита задачи. Состояние процесса после этого не определено —
        /// выходим сразу, назвав виновника.
        /// </summary>
        private static void StartWatchdog()
        {
            var watchdog = new Thread(delegate()
            {
                while (true)
                {
                    Thread.Sleep(1000);
                    string current = _currentTest;
                    if (current == null)
                        continue;
                    TimeSpan elapsed = DateTime.UtcNow - new DateTime(Interlocked.Read(ref _currentStartedTicks), DateTimeKind.Utc);
                    if (elapsed.TotalSeconds > _timeoutSeconds && ReferenceEquals(current, _currentTest))
                    {
                        Console.WriteLine("FAIL  " + current + " — TIMEOUT: дольше " + _timeoutSeconds + " с, прогон остановлен");
                        Console.Out.Flush();
                        FastExit.Now(4);
                    }
                }
            });
            watchdog.IsBackground = true;
            watchdog.Name = "test-watchdog";
            watchdog.Start();
        }

        private static string Indent(string text)
        {
            if (string.IsNullOrEmpty(text))
                return "";
            return "        " + text.Replace("\r\n", "\n").Replace("\n", Environment.NewLine + "        ");
        }

        /// <summary>Сравнение с типом: boxed int против long больше не даёт загадочного FAIL.</summary>
        private static void AssertEq<T>(T expected, T actual, string what)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
                throw new Exception(what + ": ожидалось «" + expected + "», получено «" + actual + "»");
        }

        /// <summary>Ожидаемое исключение ЗАДАННОГО типа (а не «любое»): NRE больше не сходит за отказ.</summary>
        private static TException AssertThrowsExactly<TException>(string what, Action action) where TException : Exception
        {
            try
            {
                action();
            }
            catch (TException ex)
            {
                return ex;
            }
            catch (Exception ex)
            {
                throw new Exception(what + ": ожидалось " + typeof(TException).Name + ", получено " +
                    ex.GetType().Name + ": " + ex.Message, ex);
            }
            throw new Exception(what + ": ожидалось " + typeof(TException).Name + ", но исключения не было");
        }

        /// <summary>Временная папка теста; удаляется вместе с содержимым по Dispose.</summary>
        private sealed class TempDir : IDisposable
        {
            public readonly string Path;

            public TempDir(string prefix)
            {
                Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), prefix + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(Path);
            }

            public string File(string name) { return System.IO.Path.Combine(Path, name); }

            public void Dispose()
            {
                try { Directory.Delete(Path, true); } catch { } // занята антивирусом — уберёт ОС
            }
        }
    }
}
