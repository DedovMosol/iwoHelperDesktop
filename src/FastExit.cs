using System;
using System.Runtime.InteropServices;

namespace ExcelMerger
{
    /// <summary>
    /// Немедленное завершение процесса через TerminateProcess. Нужно после работы с
    /// Windows.Data.Pdf (WinRT): при штатной выгрузке процесс роняет access violation в
    /// скрытом окне .NET-BroadcastEventWindow. Причина — не управляемые финализаторы, а
    /// НАТИВНЫЙ DLL_PROCESS_DETACH: отсоединение рантайма WinRT при выгрузке DLL падает.
    /// ExitProcess этот detach выполняет (и потому не спасал), а TerminateProcess завершает
    /// процесс БЕЗ detach и без финализаторов. Критичная очистка (сохранение настроек, Quit
    /// COM Excel/Word) выполняется детерминированно ДО этого момента, поэтому это безопасно.
    /// </summary>
    internal static class FastExit
    {
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetCurrentProcess();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool TerminateProcess(IntPtr hProcess, uint exitCode);

        public static void Now(int code)
        {
            try { Console.Out.Flush(); }
            catch { }
            TerminateProcess(GetCurrentProcess(), (uint)code);
        }
    }
}
