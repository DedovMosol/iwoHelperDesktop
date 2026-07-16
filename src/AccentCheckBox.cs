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
            Size text = TextRenderer.MeasureText(Text, Font);
            return new Size(BoxSize + TextGap + text.Width + 4, Math.Max(BoxSize + 4, text.Height + 6));
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Parent != null ? Parent.BackColor : Color.White);

            int boxTop = (Height - BoxSize) / 2;
            var rect = new RectangleF(0.5f, boxTop + 0.5f, BoxSize - 1, BoxSize - 1);
            using (GraphicsPath box = RoundedPath(rect, 4f))
            {
                if (Checked)
                {
                    Color fill = !Enabled ? Theme.DisabledFill
                        : (_hover ? Theme.AccentHover : Theme.Accent);
                    using (var b = new SolidBrush(fill))
                        g.FillPath(b, box);

                    Color mark = Enabled ? Color.White : Theme.DisabledText;
                    using (var pen = new Pen(mark, 2f))
                    {
                        pen.StartCap = LineCap.Round;
                        pen.EndCap = LineCap.Round;
                        pen.LineJoin = LineJoin.Round;
                        float y = boxTop;
                        g.DrawLines(pen, new[]
                        {
                            new PointF(4.5f, y + 9.5f),
                            new PointF(7.5f, y + 12.5f),
                            new PointF(13.5f, y + 5.5f)
                        });
                    }
                }
                else
                {
                    using (var b = new SolidBrush(Color.White))
                        g.FillPath(b, box);
                    Color border = !Enabled ? Theme.Border
                        : (_hover ? Theme.BorderDark : Theme.Border);
                    using (var pen = new Pen(border, 1f))
                        g.DrawPath(pen, box);
                }

                if (Focused && Enabled)
                {
                    using (var pen = new Pen(Color.FromArgb(120, Theme.Accent), 1.5f))
                    {
                        var focus = new RectangleF(rect.X - 0.5f, rect.Y - 0.5f, rect.Width + 1, rect.Height + 1);
                        using (GraphicsPath ring = RoundedPath(focus, 5f))
                            g.DrawPath(pen, ring);
                    }
                }
            }

            Color textColor = Enabled ? Theme.TextPrimary : Theme.DisabledText;
            var textRect = new Rectangle(BoxSize + TextGap, 0, Width - BoxSize - TextGap, Height);
            TextRenderer.DrawText(g, Text, Font, textRect, textColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
        }

        private static GraphicsPath RoundedPath(RectangleF r, float radius)
        {
            var p = new GraphicsPath();
            float d = radius * 2f;
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }
    }
}
