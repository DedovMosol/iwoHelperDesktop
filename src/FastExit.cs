using System;
using System.Runtime.InteropServices;

namespace ExcelMerger
{
    /// <summary>
    /// Немедленное завершение процесса без запуска управляемых финализаторов.
    /// Нужно после работы с Windows.Data.Pdf (WinRT): его COM-обёртки при
    /// штатной выгрузке CLR роняют процесс (access violation). Наша критичная
    /// очистка (сохранение настроек, Quit COM Excel/Word) выполняется
    /// детерминированно до этого момента, поэтому пропуск финализаторов безопасен.
    /// </summary>
    internal static class FastExit
    {
        [DllImport("kernel32.dll")]
        private static extern void ExitProcess(uint exitCode);

        public static void Now(int code)
        {
            try { Console.Out.Flush(); }
            catch { }
            ExitProcess((uint)code);
        }
    }
}
