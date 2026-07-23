using System.Drawing;
using System.Drawing.Drawing2D;

namespace ExcelMerger
{
    /// <summary>
    /// Небольшие иконки-флаги для выбора языка (триколор РФ и Union Jack Великобритании).
    /// Рисуются через GDI+ (без внешних файлов/SVG — WinForms их нативно не рендерит, а
    /// эмодзи-флаги Windows показывает буквами). Кэшируются на время жизни приложения.
    /// Размер 24×16 (иконка меню). Тонкая серая рамка отделяет белые края от белого фона меню.
    /// </summary>
    internal static class Flags
    {
        private const int W = 24, H = 16;
        private static Image _ru, _uk;

        /// <summary>Флаг для языка: EN → Великобритания, иначе → Россия.</summary>
        public static Image For(Lang lang) { return lang == Lang.En ? Uk : Russia; }

        public static Image Russia { get { return _ru ?? (_ru = DrawRussia()); } }
        public static Image Uk { get { return _uk ?? (_uk = DrawUk()); } }

        private static Bitmap DrawRussia()
        {
            var bmp = new Bitmap(W, H);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                int band = H / 3;
                using (var white = new SolidBrush(Color.White))
                using (var blue = new SolidBrush(Color.FromArgb(0, 57, 166)))
                using (var red = new SolidBrush(Color.FromArgb(213, 43, 30)))
                {
                    g.FillRectangle(white, 0, 0, W, band);
                    g.FillRectangle(blue, 0, band, W, band);
                    g.FillRectangle(red, 0, 2 * band, W, H - 2 * band);
                }
                Border(g);
            }
            return bmp;
        }

        private static Bitmap DrawUk()
        {
            var bmp = new Bitmap(W, H);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                var full = new Rectangle(0, 0, W, H);
                using (var blue = new SolidBrush(Color.FromArgb(1, 33, 105)))
                    g.FillRectangle(blue, full);

                Color white = Color.White;
                Color red = Color.FromArgb(200, 16, 46);
                var tl = new Point(0, 0); var tr = new Point(W, 0);
                var bl = new Point(0, H); var br = new Point(W, H);

                using (var region = new Region(full))
                {
                    g.Clip = region; // диагонали не вылезают за флаг
                    // Белая косая (диагонали), затем красная косая уже.
                    using (var wpen = new Pen(white, 4f)) { g.DrawLine(wpen, tl, br); g.DrawLine(wpen, tr, bl); }
                    using (var rpen = new Pen(red, 1.8f)) { g.DrawLine(rpen, tl, br); g.DrawLine(rpen, tr, bl); }
                    g.ResetClip();
                }

                // Белый прямой крест, затем красный уже.
                using (var wcross = new SolidBrush(white))
                {
                    g.FillRectangle(wcross, W / 2 - 3, 0, 6, H);
                    g.FillRectangle(wcross, 0, H / 2 - 3, W, 6);
                }
                using (var rcross = new SolidBrush(red))
                {
                    g.FillRectangle(rcross, W / 2 - 2, 0, 4, H);
                    g.FillRectangle(rcross, 0, H / 2 - 2, W, 4);
                }
                Border(g);
            }
            return bmp;
        }

        private static void Border(Graphics g)
        {
            using (var pen = new Pen(Color.FromArgb(150, 150, 150)))
                g.DrawRectangle(pen, 0, 0, W - 1, H - 1);
        }
    }
}
