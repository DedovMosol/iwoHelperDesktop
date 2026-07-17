using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Окраска системного заголовка окна (там, где кнопки свернуть/развернуть/закрыть)
    /// в акцентный цвет через DWM. Поддерживается Windows 11 (build 22000+); на более
    /// старых системах вызов безвреден — DWM игнорирует неизвестный атрибут, заголовок
    /// остаётся штатным, а брендовую шапку в клиентской области рисует <see cref="HeaderBand"/>.
    /// </summary>
    internal static class WindowChrome
    {
        private const int DWMWA_CAPTION_COLOR = 35; // Windows 11 22000+
        private const int DWMWA_TEXT_COLOR = 36;

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

        /// <summary>Красит заголовок в заданный цвет сразу и на каждое пересоздание хэндла окна.</summary>
        public static void Enable(Form form, Color captionColor)
        {
            form.HandleCreated += delegate { Apply(form.Handle, captionColor); };
            if (form.IsHandleCreated)
                Apply(form.Handle, captionColor);
        }

        private static void Apply(IntPtr handle, Color captionColor)
        {
            if (handle == IntPtr.Zero)
                return;
            try
            {
                int caption = ColorRef(captionColor);
                int text = ColorRef(Color.White);
                DwmSetWindowAttribute(handle, DWMWA_CAPTION_COLOR, ref caption, sizeof(int));
                DwmSetWindowAttribute(handle, DWMWA_TEXT_COLOR, ref text, sizeof(int));
            }
            catch { } // dwmapi недоступен — остаётся штатный заголовок
        }

        /// <summary>COLORREF (0x00BBGGRR), как ждёт DWM. internal — под юнит-тест упаковки.</summary>
        internal static int ColorRef(Color c)
        {
            return c.R | (c.G << 8) | (c.B << 16);
        }
    }
}
