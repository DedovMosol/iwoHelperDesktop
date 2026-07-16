using System.Drawing;

namespace ExcelMerger
{
    /// <summary>Единая палитра приложения (согласована с assets/icon.svg).</summary>
    public static class Theme
    {
        public static readonly Color Accent = Color.FromArgb(16, 124, 65);        // #107C41
        public static readonly Color AccentHover = Color.FromArgb(14, 107, 56);   // #0E6B38
        public static readonly Color AccentPressed = Color.FromArgb(11, 86, 45);  // #0B562D

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
    }
}
