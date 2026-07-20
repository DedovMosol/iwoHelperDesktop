using System;
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
        private readonly Color _top;
        private readonly Color _bottom;
        private static readonly Font TitleFont = new Font("Segoe UI", 15f, FontStyle.Bold);
        private static readonly Color SubtitleColor = Color.FromArgb(233, 238, 242); // нейтрально-белая: читаема на любом фоне

        /// <summary>Центрировать заголовок и подпись по ширине (для стартового экрана без кнопок).</summary>
        public bool Centered { get; set; }

        public HeaderBand(string title, string subtitle, Color top, Color bottom)
        {
            _title = title ?? string.Empty;
            _subtitle = subtitle ?? string.Empty;
            _top = top;
            _bottom = bottom;
            DoubleBuffered = true;
            SetStyle(ControlStyles.ResizeRedraw, true); // перерисовка градиента при растяжении
            ForeColor = Color.White;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Rectangle r = ClientRectangle;
            using (var brush = new LinearGradientBrush(r, _top, _bottom, LinearGradientMode.Vertical))
                e.Graphics.FillRectangle(brush, r);
            // Тонкая нижняя грань отделяет шапку от тела окна.
            using (var pen = new Pen(Color.FromArgb(60, 0, 0, 0)))
                e.Graphics.DrawLine(pen, 0, r.Height - 1, r.Width, r.Height - 1);

            int subtitleY = Height - 28;
            int titleY = subtitleY - 30;
            TextFormatFlags flags = TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis |
                TextFormatFlags.SingleLine | TextFormatFlags.NoPadding;

            if (Centered)
            {
                flags |= TextFormatFlags.HorizontalCenter;
                TextRenderer.DrawText(e.Graphics, _title, TitleFont,
                    new Rectangle(10, titleY, Width - 20, 26), Color.White, flags);
                if (_subtitle.Length > 0)
                    TextRenderer.DrawText(e.Graphics, _subtitle, Font,
                        new Rectangle(10, subtitleY, Width - 20, 20), SubtitleColor, flags);
                return;
            }

            // Слева, не заходя под дочерние контролы (кнопку «Главная»): обрезаем многоточием.
            int leftmostChild = int.MaxValue;
            foreach (Control c in Controls)
                if (c.Visible && c.Left < leftmostChild)
                    leftmostChild = c.Left;
            int rightBound = TextRightBound(Width, leftmostChild);
            TextRenderer.DrawText(e.Graphics, _title, TitleFont,
                new Rectangle(18, titleY, rightBound - 18, 26), Color.White, flags);
            if (_subtitle.Length > 0)
                TextRenderer.DrawText(e.Graphics, _subtitle, Font,
                    new Rectangle(20, subtitleY, rightBound - 20, 20), SubtitleColor, flags);
        }

        /// <summary>
        /// Правая граница текста шапки: не заходить под левый край дочерних
        /// контролов (кнопки «Назад»), но и не за правый край панели. Чистая — под тест.
        /// </summary>
        internal static int TextRightBound(int bandWidth, int leftmostChildLeft)
        {
            int bound = leftmostChildLeft < int.MaxValue ? leftmostChildLeft - 12 : bandWidth - 20;
            return Math.Min(bound, bandWidth - 20);
        }
    }
}
