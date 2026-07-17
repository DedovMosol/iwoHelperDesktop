using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Брендовая шапка окна: панель с вертикальным градиентом акцентного цвета,
    /// заголовком и подписью в белом. Занимает то же место, что прежняя пара
    /// «заголовок + подпись», поэтому тело окна не сдвигается. Общая для
    /// стартового экрана и инструментов (DRY). Кнопку «Назад в меню» кладут
    /// прямо на неё как дочерний элемент.
    /// </summary>
    public class HeaderBand : Panel
    {
        private readonly string _title;
        private readonly string _subtitle;
        private static readonly Font TitleFont = new Font("Segoe UI", 15f, FontStyle.Bold);
        private static readonly Color SubtitleColor = Color.FromArgb(214, 236, 224);

        public HeaderBand(string title, string subtitle)
        {
            _title = title ?? string.Empty;
            _subtitle = subtitle ?? string.Empty;
            DoubleBuffered = true;
            SetStyle(ControlStyles.ResizeRedraw, true); // перерисовка градиента при растяжении
            ForeColor = Color.White;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Rectangle r = ClientRectangle;
            using (var brush = new LinearGradientBrush(r, Theme.Accent, Theme.AccentPressed, LinearGradientMode.Vertical))
                e.Graphics.FillRectangle(brush, r);
            // Тонкая нижняя грань отделяет шапку от тела окна.
            using (var pen = new Pen(Color.FromArgb(60, 0, 0, 0)))
                e.Graphics.DrawLine(pen, 0, r.Height - 1, r.Width, r.Height - 1);

            int subtitleY = Height - 28;
            int titleY = subtitleY - 30;
            TextRenderer.DrawText(e.Graphics, _title, TitleFont, new Point(18, titleY),
                Color.White, TextFormatFlags.NoPrefix);
            if (_subtitle.Length > 0)
                TextRenderer.DrawText(e.Graphics, _subtitle, Font, new Point(20, subtitleY),
                    SubtitleColor, TextFormatFlags.NoPrefix);
        }
    }
}
