using System;
using System.Runtime.InteropServices;

namespace ExcelMerger
{
    /// <summary>
    /// Какой режим DPI получило окно. Режим задаёт не код и не манифест, а
    /// iwoHelperDesktop.exe.config (build\app.config): WinForms 4.7+ читает DpiAwareness и
    /// сам ставит его процессу при создании первого элемента управления. Без файла рядом с
    /// exe процесс остаётся DPI-unaware: Windows растягивает окна картинкой (размыто, но
    /// вёрстка цела). Спрашивает --selftest: только он запускает настоящий exe в том виде,
    /// в каком его поставляют, а у юнит-тестов своя сборка без этого файла.
    /// </summary>
    internal static class DpiAwareness
    {
        // DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 (windef.h)
        private static readonly IntPtr PerMonitorAwareV2 = new IntPtr(-4);
        // PROCESS_PER_MONITOR_DPI_AWARE (shellscalingapi.h)
        private const int ProcessPerMonitorDpiAware = 2;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsValidDpiAwarenessContext(IntPtr value);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowDpiAwarenessContext(IntPtr hwnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AreDpiAwarenessContextsEqual(IntPtr a, IntPtr b);

        [DllImport("shcore.dll")]
        private static extern int GetProcessDpiAwareness(IntPtr process, out int value);

        /// <summary>
        /// Окно создано в режиме, который обещает app.config: PerMonitorV2 там, где система его
        /// знает (Windows 10 1703+), иначе per-monitor процесса — так ту же строку конфигурации
        /// применяет WinForms на 8.1–1607.
        /// </summary>
        internal static bool IsPerMonitor(IntPtr window)
        {
            if (SystemSupportsPerMonitorV2())
                return AreDpiAwarenessContextsEqual(GetWindowDpiAwarenessContext(window), PerMonitorAwareV2);
            int value;
            return GetProcessDpiAwareness(IntPtr.Zero, out value) == 0 && value == ProcessPerMonitorDpiAware;
        }

        /// <summary>
        /// Спрашиваем саму систему, а не Environment.OSVersion: версию Windows процессу сообщает
        /// манифест, и без объявленной в нём Windows 10 WinForms выбрал бы per-monitor v1 —
        /// окна снова перестали бы перестраиваться, а проверка по версии этого бы не заметила.
        /// </summary>
        private static bool SystemSupportsPerMonitorV2()
        {
            try { return IsValidDpiAwarenessContext(PerMonitorAwareV2); }
            catch (EntryPointNotFoundException) { return false; } // Windows 8.1: функции нет
        }
    }
}
