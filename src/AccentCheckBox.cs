using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Чекбокс в стиле приложения: скруглённый квадрат, во включённом состоянии —
    /// белая галочка на зелёном фоне. Поведение (клавиатура, фокус, события)
    /// наследуется от стандартного CheckBox, своя только отрисовка.
    /// </summary>
    public class AccentCheckBox : CheckBox
    {
        private const int BoxSize = 18;
        private const int TextGap = 8;

        private bool _hover;

        public AccentCheckBox()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Cursor = Cursors.Hand;
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnCheckedChanged(EventArgs e) { Invalidate(); base.OnCheckedChanged(e); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }

        public override Size GetPreferredSize(Size proposedSize)
        {
            // Всё в пикселях устройства, как и замер текста: вызывающий переводит размер целиком
            // (DpiScale.ToLogical), и логический квадрат рядом с экранным текстом ужался бы.
            Size text = TextRenderer.MeasureText(Text, Font);
            int box = Box, gap = Gap;
            return new Size(box + gap + text.Width + DpiScale.Device(this, 4),
                Math.Max(box + DpiScale.Device(this, 4), text.Height + DpiScale.Device(this, 6)));
        }

        /// <summary>Сторона квадрата в пикселях устройства (UI-25: при 150 % — 27, а не 18).</summary>
        private int Box { get { return DpiScale.Device(this, BoxSize); } }

        /// <summary>Просвет между квадратом и подписью в пикселях устройства.</summary>
        private int Gap { get { return DpiScale.Device(this, TextGap); } }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Parent != null ? Parent.BackColor : Color.White);

            // Рисование — всегда после масштаба окна: квадрат, галочка, рамки — в пикселях устройства.
            float s = DpiScale.Device(this);
            int side = Box;
            int boxTop = (Height - side) / 2;
            float line = Math.Max(1f, s);
            var rect = new RectangleF(line / 2, boxTop + line / 2, side - line, side - line);
            using (GraphicsPath box = Ui.RoundedRect(rect, 4f * s))
            {
                if (Checked)
                {
                    Color fill = !Enabled ? Theme.DisabledFill
                        : (_hover ? Theme.AccentHover : Theme.Accent);
                    using (var b = new SolidBrush(fill))
                        g.FillPath(b, box);

                    Color mark = Enabled ? Color.White : Theme.DisabledText;
                    using (var pen = new Pen(mark, 2f * s))
                    {
                        pen.StartCap = LineCap.Round;
                        pen.EndCap = LineCap.Round;
                        pen.LineJoin = LineJoin.Round;
                        float y = boxTop;
                        // Галочка в координатах квадрата 18×18, растянутых под его сторону.
                        float k = side / (float)BoxSize;
                        g.DrawLines(pen, new[]
                        {
                            new PointF(4.5f * k, y + 9.5f * k),
                            new PointF(7.5f * k, y + 12.5f * k),
                            new PointF(13.5f * k, y + 5.5f * k)
                        });
                    }
                }
                else
                {
                    using (var b = new SolidBrush(Color.White))
                        g.FillPath(b, box);
                    Color border = !Enabled ? Theme.Border
                        : (_hover ? Theme.BorderDark : Theme.Border);
                    using (var pen = new Pen(border, line))
                        g.DrawPath(pen, box);
                }

                if (Focused && Enabled)
                {
                    using (var pen = new Pen(Color.FromArgb(120, Theme.Accent), 1.5f * s))
                    {
                        var focus = new RectangleF(rect.X - line / 2, rect.Y - line / 2, rect.Width + line, rect.Height + line);
                        using (GraphicsPath ring = Ui.RoundedRect(focus, 5f * s))
                            g.DrawPath(pen, ring);
                    }
                }
            }

            Color textColor = Enabled ? Theme.TextPrimary : Theme.DisabledText;
            var textRect = new Rectangle(side + Gap, 0, Width - side - Gap, Height);
            TextRenderer.DrawText(g, Text, Font, textRect, textColor,
                // WordBreak вместо SingleLine: подпись, не влезающая в узкую колонку, переносится
                // на вторую строку вместо того, чтобы обрезаться. Коротким подписям это ничего
                // не меняет — они как были в одну строку, так и остаются.
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak);
        }

    }
}
