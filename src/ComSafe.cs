using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExcelMerger
{
    /// <summary>Общие правила безопасного освобождения COM (Excel и Word).</summary>
    internal static class ComSafe
    {
        /// <summary>
        /// ВАЖНО: передавать аргумент только со статическим типом object (сохранить
        /// dynamic-ссылку в object-переменную до Close/Quit). Динамическая привязка
        /// любой операции на уже закрытом COM-объекте (например, Workbook после Close)
        /// падает с COMException 0x80010114 ещё до входа в метод, мимо его try/catch.
        /// </summary>
        public static void Release(object o)
        {
            try
            {
                if (o != null && Marshal.IsComObject(o))
                    Marshal.FinalReleaseComObject(o);
            }
            catch { }
        }

        /// <summary>Гарантированная сборка RCW — чтобы не оставались процессы Office.</summary>
        public static void Collect()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    /// <summary>
    /// Процесс Office, который запустили МЫ: PID по окну приложения плюс время старта.
    /// Нужен последней страховкой после Quit/Release/Collect: заклинивший экземпляр
    /// (например, после перезапуска из-за зависшего файла) иначе остаётся висеть
    /// невидимым EXCEL.EXE и мешает всем следующим COM-вызовам на машине. Время старта
    /// защищает от переиспользования PID: чужой процесс с тем же номером не тронем.
    /// </summary>
    internal sealed class OwnedOfficeProcess
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        public readonly int Id;
        private readonly DateTime _startTime;

        private OwnedOfficeProcess(int id, DateTime startTime)
        {
            Id = id;
            _startTime = startTime;
        }

        /// <summary>По окну приложения (Application.Hwnd). null — определить не удалось: страховки нет.</summary>
        public static OwnedOfficeProcess FromWindow(long hwnd)
        {
            try
            {
                if (hwnd == 0)
                    return null;
                var window = new IntPtr(hwnd);
                uint pid;
                // 0 — окно недействительно, и pid тогда не заполнен (документация WinAPI).
                if (GetWindowThreadProcessId(window, out pid) == 0 || pid == 0)
                    return null;
                return FromProcessId((int)pid);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>По PID (для тестов и случаев без окна). null — процесса нет.</summary>
        internal static OwnedOfficeProcess FromProcessId(int pid)
        {
            try
            {
                using (Process p = Process.GetProcessById(pid))
                    return new OwnedOfficeProcess(pid, p.StartTime);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Добивать ли процесс: он ещё жив, это ТОТ ЖЕ процесс (время старта совпало) и ни одно
        /// его окно не показано. Видимое окно значит, что экземпляром завладел пользователь
        /// (например, открыл в нём файл двойным щелчком) — его несохранённую работу не убиваем.
        /// Чистая — под тест.
        /// </summary>
        internal static bool ShouldKill(bool stillRunning, DateTime ownedStart, DateTime actualStart, bool windowVisible)
        {
            return stillRunning && ownedStart == actualStart && !windowVisible;
        }

        /// <summary>
        /// Вызывать ПОСЛЕ Quit/Release/Collect: ждёт штатного выхода до waitMs и завершает
        /// процесс, только если он всё ещё висит. true — пришлось завершить принудительно.
        /// </summary>
        public bool KillIfLingering(int waitMs)
        {
            try
            {
                using (Process p = Process.GetProcessById(Id))
                {
                    if (p.WaitForExit(waitMs))
                        return false; // вышел сам — штатный путь
                    if (!ShouldKill(!p.HasExited, _startTime, p.StartTime, HasVisibleWindow(Id)))
                        return false;
                    p.Kill();
                    p.WaitForExit(2000);
                    return true;
                }
            }
            catch
            {
                return false; // процесса уже нет (ArgumentException) или доступ закрыт — не наш случай
            }
        }

        /// <summary>
        /// Видно ли хоть одно верхнее окно процесса — на момент решения, а не окно, запомненное при
        /// старте: в SDI-Excel рамку Application.Hwnd забирает первая книга (свод), и target.Close
        /// её уничтожает, а книга, открытая человеком, живёт в своём окне. Перебор окон не шлёт
        /// процессу сообщений — ответ есть и у заклинившего экземпляра. У скрытого Excel все
        /// верхние окна невидимы (проверено на живом Excel 15.0).
        /// </summary>
        private static bool HasVisibleWindow(int pid)
        {
            bool visible = false;
            EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
            {
                uint owner;
                if (GetWindowThreadProcessId(hWnd, out owner) != 0 && owner == (uint)pid && IsWindowVisible(hWnd))
                {
                    visible = true;
                    return false; // нашли — перебор можно остановить
                }
                return true;
            }, IntPtr.Zero);
            return visible;
        }
    }
}
