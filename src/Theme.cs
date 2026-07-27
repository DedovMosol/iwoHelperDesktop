using System.Drawing;

namespace ExcelMerger
{
    /// <summary>Единая палитра приложения (согласована с assets/icon.svg).</summary>
    public static class Theme
    {
        public static readonly Color Accent = Color.FromArgb(16, 124, 65);        // #107C41
        public static readonly Color AccentHover = Color.FromArgb(14, 107, 56);   // #0E6B38
        public static readonly Color AccentPressed = Color.FromArgb(11, 86, 45);  // #0B562D

        // Цвета шапок по инструментам: свод Excel — зелёная (Accent), объединение
        // PDF — красная (в тон значку), стартовый выбор — синяя.
        public static readonly Color PdfRed = Color.FromArgb(211, 47, 47);        // #D32F2F
        public static readonly Color PdfRedDark = Color.FromArgb(154, 34, 34);    // #9A2222
        public static readonly Color HubBlue = Color.FromArgb(15, 108, 189);      // #0F6CBD
        public static readonly Color HubBlueDark = Color.FromArgb(10, 78, 134);   // #0A4E86
        // «PDF → Word» — фиолетовая (отличается от синей/красной/зелёной).
        public static readonly Color WordViolet = Color.FromArgb(91, 79, 191);    // #5B4FBF
        public static readonly Color WordVioletDark = Color.FromArgb(62, 52, 140);// #3E348C

        public static readonly Color TextPrimary = Color.FromArgb(45, 45, 45);
        public static readonly Color TextMuted = Color.FromArgb(115, 115, 115);

        public static readonly Color OkGreen = Color.FromArgb(23, 111, 44);
        public static readonly Color WarnOrange = Color.FromArgb(176, 98, 0);
        public static readonly Color ErrRed = Color.FromArgb(178, 34, 34);

        public static readonly Color Border = Color.FromArgb(200, 200, 200);
        public static readonly Color BorderDark = Color.FromArgb(150, 150, 150);
        public static readonly Color SecondaryHover = Color.FromArgb(243, 244, 246);
        public static readonly Color SecondaryPressed = Color.FromArgb(233, 234, 236);
        public static readonly Color DisabledFill = Color.FromArgb(228, 228, 228);
        public static readonly Color DisabledText = Color.FromArgb(155, 155, 155);
        // Недоступная ОБЫЧНАЯ кнопка: заливка, а не белизна. Без неё она отличалась от рабочей
        // только цветом надписи, и в панелях, где половина кнопок ждёт открытого документа,
        // «нельзя нажать» приходилось выяснять нажатием.
        public static readonly Color DisabledSecondaryFill = Color.FromArgb(245, 245, 245);
        public static readonly Color DisabledBorder = Color.FromArgb(226, 226, 226);

        // Кнопки на тёмной подложке (полоса лупы в полноэкранном просмотре).
        public static readonly Color DarkBarFill = Color.FromArgb(62, 62, 65);
        public static readonly Color DarkBarHover = Color.FromArgb(76, 76, 80);
        public static readonly Color DarkBarPressed = Color.FromArgb(92, 92, 96);
        public static readonly Color DarkBarBorder = Color.FromArgb(96, 96, 100);
        public static readonly Color DarkBarDisabledFill = Color.FromArgb(56, 56, 58);
        // Светлее, чем просится «серым по тёмному»: недоступную подпись читать не обязательно,
        // но понять, ЧТО это за кнопка, нужно — иначе полоса превращается в ряд пустых плашек.
        public static readonly Color DarkBarDisabledText = Color.FromArgb(154, 154, 158);

        /// <summary>
        /// Цвет в формате OLE/COLORREF (0x00BBGGRR) — как ждут Excel Range/Shape
        /// «.RGB» и Win32/DWM. Единое место упаковки (DRY). Чистая — под тест.
        /// </summary>
        public static int ToBgr(Color c)
        {
            return c.R | (c.G << 8) | (c.B << 16);
        }
    }
}
