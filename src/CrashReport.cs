using System;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Последний рубеж обработки ошибок GUI: без него исключение UI-потока показывает
    /// системный ThreadExceptionDialog, а исключение фонового потока молча роняет процесс —
    /// пользователь не видит ни объяснения, ни следа для поддержки. Здесь: понятный
    /// фирменный диалог (UI-поток, работа продолжается) и crash.log в профиле приложения.
    /// Устанавливать из Main ДО создания первого окна.
    /// </summary>
    internal static class CrashReport
    {
        private const long MaxLogBytes = 256 * 1024; // разросся — начинаем заново (простая ротация)
        private static bool _showing; // диалог уже на экране: сбой в самом диалоге не должен зациклить

        public static string LogPath
        {
            get { return Path.Combine(AppPaths.Root, "crash.log"); }
        }

        /// <summary>Подписать все «последние рубежи» процесса. Звать один раз, до Application.Run.</summary>
        public static void Install()
        {
            // Все исключения UI-потока — в ThreadException (без этого в отладчике/некоторых
            // конфигурациях WinForms кидает их как необработанные, минуя наш диалог).
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += delegate(object s, ThreadExceptionEventArgs e)
            {
                Handle(e.Exception, true);
            };
            // Фоновые потоки: показать диалог не с UI-потока нельзя — только след в логе
            // (CLR после этого события процесс всё равно завершит).
            AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e)
            {
                Handle(e.ExceptionObject as Exception, false);
            };
            TaskScheduler.UnobservedTaskException += delegate(object s, UnobservedTaskExceptionEventArgs e)
            {
                Log(e.Exception);
                e.SetObserved(); // след оставлен; финализатор не должен ронять процесс
            };
        }

        private static void Handle(Exception ex, bool canShowDialog)
        {
            Log(ex);
            if (!canShowDialog || _showing)
                return;
            _showing = true;
            try
            {
                Dialogs.Error(Form.ActiveForm, "iwo Helper Desktop", Loc.T("crash.title"),
                    string.Format(Loc.T("crash.body"), Short(ex), LogPath));
            }
            catch { } // диалог не показался (нет ресурсов и т.п.) — лог уже есть
            finally { _showing = false; }
        }

        /// <summary>Дописать исключение в crash.log (с простой ротацией). Сбои записи молчат.</summary>
        public static void Log(Exception ex)
        {
            try
            {
                Directory.CreateDirectory(AppPaths.Root);
                var f = new FileInfo(LogPath);
                if (f.Exists && f.Length > MaxLogBytes)
                    f.Delete();
                string version = Assembly.GetExecutingAssembly().GetName().Version.ToString(3);
                File.AppendAllText(LogPath, Format(ex, version, DateTime.Now));
            }
            catch { } // недоступный профиль не повод падать вторично
        }

        /// <summary>Одна запись лога: метка времени, версия, полное исключение. Чистая — под тест.</summary>
        internal static string Format(Exception ex, string version, DateTime at)
        {
            return "[" + at.ToString("yyyy-MM-dd HH:mm:ss") + "] v" + version + "\r\n"
                + (ex != null ? ex.ToString() : "(null)") + "\r\n\r\n";
        }

        /// <summary>Короткая строка исключения для диалога: тип и сообщение, без стека.</summary>
        private static string Short(Exception ex)
        {
            if (ex == null)
                return "(null)";
            string m = (ex.Message ?? "").Trim();
            return ex.GetType().Name + (m.Length > 0 ? ": " + m : "");
        }
    }
}
