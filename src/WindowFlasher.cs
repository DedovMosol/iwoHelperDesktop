using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Мигание кнопкой окна на панели задач, пока пользователь не вернётся в окно
    /// (FLASHW_TIMERNOFG). Стандартный способ Windows сказать «работа завершена»,
    /// не воруя фокус.
    /// </summary>
    internal static class WindowFlasher
    {
        private const uint FLASHW_ALL = 3;
        private const uint FLASHW_TIMERNOFG = 12;

        [StructLayout(LayoutKind.Sequential)]
        private struct FLASHWINFO
        {
            public uint cbSize;
            public IntPtr hwnd;
            public uint dwFlags;
            public uint uCount;
            public uint dwTimeout;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        public static void FlashUntilForeground(Form form)
        {
            try
            {
                var info = new FLASHWINFO();
                info.cbSize = (uint)Marshal.SizeOf(typeof(FLASHWINFO));
                info.hwnd = form.Handle;
                info.dwFlags = FLASHW_ALL | FLASHW_TIMERNOFG;
                info.uCount = uint.MaxValue;
                info.dwTimeout = 0;
                FlashWindowEx(ref info);
            }
            catch { } // уведомление — украшение, не причина падать
        }
    }
}
