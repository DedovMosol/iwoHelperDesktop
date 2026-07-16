using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Нативные алерты Windows (TaskDialog, comctl32 v6 — включён через манифест)
    /// с фолбэком на MessageBox, если TaskDialog недоступен.
    /// </summary>
    public static class Dialogs
    {
        [DllImport("comctl32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        private static extern int TaskDialog(
            IntPtr hwndParent, IntPtr hInstance,
            string title, string mainInstruction, string content,
            int commonButtons, IntPtr icon, out int button);

        private const int TDCBF_OK_BUTTON = 0x0001;
        private const int TDCBF_YES_BUTTON = 0x0002;
        private const int TDCBF_NO_BUTTON = 0x0004;
        private const int IDYES = 6;

        private static readonly IntPtr WarningIcon = new IntPtr(0xFFFF); // TD_WARNING_ICON
        private static readonly IntPtr ErrorIcon = new IntPtr(0xFFFE);   // TD_ERROR_ICON
        private static readonly IntPtr InfoIcon = new IntPtr(0xFFFD);    // TD_INFORMATION_ICON

        /// <summary>Вопрос с кнопками Да/Нет и предупреждающим значком. true = Да.</summary>
        public static bool ConfirmWarning(IWin32Window owner, string title, string instruction, string content)
        {
            int button;
            if (TryShowNative(owner, title, instruction, content,
                    TDCBF_YES_BUTTON | TDCBF_NO_BUTTON, WarningIcon, out button))
                return button == IDYES;
            return MessageBox.Show(owner, Fallback(instruction, content), title,
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes;
        }

        public static void Info(IWin32Window owner, string title, string instruction, string content)
        {
            int button;
            if (TryShowNative(owner, title, instruction, content, TDCBF_OK_BUTTON, InfoIcon, out button))
                return;
            MessageBox.Show(owner, Fallback(instruction, content), title,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void Error(IWin32Window owner, string title, string instruction, string content)
        {
            int button;
            if (TryShowNative(owner, title, instruction, content, TDCBF_OK_BUTTON, ErrorIcon, out button))
                return;
            MessageBox.Show(owner, Fallback(instruction, content), title,
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static bool TryShowNative(IWin32Window owner, string title, string instruction,
            string content, int buttons, IntPtr icon, out int button)
        {
            button = 0;
            try
            {
                IntPtr hwnd = owner != null ? owner.Handle : IntPtr.Zero;
                return TaskDialog(hwnd, IntPtr.Zero, title, instruction, content, buttons, icon, out button) >= 0;
            }
            catch
            {
                return false;
            }
        }

        private static string Fallback(string instruction, string content)
        {
            return string.IsNullOrEmpty(content) ? instruction : instruction + "\n\n" + content;
        }
    }
}
