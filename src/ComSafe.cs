using System;
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
}
