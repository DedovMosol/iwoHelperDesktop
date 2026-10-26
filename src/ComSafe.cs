using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

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
    /// Тот же процесс останавливает и сторож зависшего вызова (<see cref="ComWatchdog"/>).
    /// </summary>
    internal sealed class OwnedOfficeProcess : IWatchedProcess
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowEnabled(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, uint command);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const uint GwOwner = 4; // GW_OWNER: владелец верхнего окна

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
        /// Добивать ли процесс: он ещё жив, это ТОТ ЖЕ процесс (время старта совпало) и у него нет
        /// окна человека. Окно человека значит, что экземпляром завладел пользователь (например,
        /// открыл в нём файл двойным щелчком) — его несохранённую работу не убиваем. Чистая — под тест.
        /// </summary>
        internal static bool ShouldKill(bool stillRunning, DateTime ownedStart, DateTime actualStart, bool humanWindow)
        {
            return stillRunning && ownedStart == actualStart && !humanWindow;
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
                    return Kill(p);
                }
            }
            catch
            {
                return false; // процесса уже нет (ArgumentException) или доступ закрыт — не наш случай
            }
        }

        /// <summary>Остановить зависший процесс сразу (сторож вызова). false — нельзя или он уже вышел.</summary>
        public bool TryStop()
        {
            try
            {
                using (Process p = Process.GetProcessById(Id))
                    return Kill(p);
            }
            catch
            {
                return false; // процесса уже нет или доступ закрыт
            }
        }

        private bool Kill(Process p)
        {
            if (!ShouldKill(!p.HasExited, _startTime, p.StartTime, Scan(Id).Human))
                return false;
            p.Kill();
            p.WaitForExit(2000);
            return true;
        }

        /// <summary>
        /// Заголовок модального окна скрытого экземпляра ("" — без заголовка); null — такого окна
        /// нет или экземпляром пользуется человек. Модальное окно Office выключает своё главное
        /// окно-владельца; у скрытого экземпляра владелец ещё и невидим (проверено на живых Excel
        /// и Word 15.0: «Открытие документа» — видимое окно #32770, его владелец XLMAIN или
        /// OpusApp невидим и выключен). Занятый расчётом экземпляр окно не выключает.
        /// </summary>
        public string ModalWindowTitle()
        {
            WindowScan scan = Scan(Id);
            return scan.Human ? null : scan.ModalTitle;
        }

        private struct WindowScan
        {
            public bool Human;        // есть окно человека
            public string ModalTitle; // первое модальное окно скрытого экземпляра
        }

        /// <summary>
        /// Окна процесса на момент решения (а не окно, запомненное при старте: в SDI-Excel рамку
        /// Application.Hwnd забирает первая книга-свод, и target.Close её уничтожает, а книга
        /// человека живёт в своём окне). Видимое верхнее окно, чей корневой владелец
        /// (<see cref="RootOwner"/>) невидим и выключен, — модальное окно скрытого экземпляра: ответить на
        /// него некому. Любое другое видимое окно — окно человека (рамка Excel или Word без
        /// владельца, окна поверх видимой рамки). Перебор окон, проверки видимости и
        /// заголовок окна чужого процесса сообщений процессу не шлют — ответ есть и у
        /// заклинившего экземпляра (заголовок GetWindowText для окна другого процесса читает
        /// из системы, без WM_GETTEXT).
        /// </summary>
        private static WindowScan Scan(int pid)
        {
            var scan = new WindowScan();
            EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
            {
                uint owner;
                if (GetWindowThreadProcessId(hWnd, out owner) == 0 || owner != (uint)pid || !IsWindowVisible(hWnd))
                    return true;
                IntPtr root = RootOwner(hWnd);
                bool modalOfHidden = root != hWnd && !IsWindowVisible(root) && !IsWindowEnabled(root);
                if (!modalOfHidden)
                {
                    scan.Human = true;
                    return false; // окно человека решает всё — перебор можно остановить
                }
                if (scan.ModalTitle == null)
                    scan.ModalTitle = WindowText(hWnd);
                return true;
            }, IntPtr.Zero);
            return scan;
        }

        /// <summary>
        /// Последний владелец в цепочке GW_OWNER (у окна без владельца — оно само). Не GetAncestor
        /// (GA_ROOTOWNER): тот идёт по GetParent, а GetParent отдаёт владельца только окнам WS_POPUP —
        /// у владеемого окна без этого стиля корнем оказалось бы оно само (проверено на WinForms).
        /// Глубина ограничена: цикл владельцев система не допускает, но перебор не должен зависеть от этого.
        /// </summary>
        private static IntPtr RootOwner(IntPtr hWnd)
        {
            IntPtr root = hWnd;
            for (int depth = 0; depth < 64; depth++)
            {
                IntPtr owner = GetWindow(root, GwOwner);
                if (owner == IntPtr.Zero)
                    break;
                root = owner;
            }
            return root;
        }

        private static string WindowText(IntPtr hWnd)
        {
            var text = new StringBuilder(256);
            GetWindowText(hWnd, text, text.Capacity);
            return text.ToString();
        }
    }
}
