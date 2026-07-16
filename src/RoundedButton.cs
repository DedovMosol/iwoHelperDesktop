using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Кнопка со скруглёнными углами, сглаживанием и состояниями
    /// normal / hover / pressed / disabled / focused.
    /// primary — акцентная (зелёная), иначе — вторичная (белая с рамкой).
    /// </summary>
    public class RoundedButton : Button
    {
        private readonly bool _primary;
        private bool _hover;
        private bool _pressed;

        public RoundedButton(bool primary)
        {
            _primary = primary;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
            Font = primary
                ? new Font("Segoe UI", 10.5f, FontStyle.Bold)
                : new Font("Segoe UI", 9.75f);
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; _pressed = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { _pressed = true; Invalidate(); base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { _pressed = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Parent != null ? Parent.BackColor : Color.White);

            Color fill;
            Color text;
            Color border = Color.Empty;
            if (_primary)
            {
                if (!Enabled) { fill = Theme.DisabledFill; text = Theme.DisabledText; }
                else if (_pressed) { fill = Theme.AccentPressed; text = Color.White; }
                else if (_hover) { fill = Theme.AccentHover; text = Color.White; }
                else { fill = Theme.Accent; text = Color.White; }
            }
            else
            {
                if (!Enabled) { fill = Color.White; text = Theme.DisabledText; border = Theme.Border; }
                else if (_pressed) { fill = Theme.SecondaryPressed; text = Theme.TextPrimary; border = Theme.BorderDark; }
                else if (_hover) { fill = Theme.SecondaryHover; text = Theme.TextPrimary; border = Theme.BorderDark; }
                else { fill = Color.White; text = Theme.TextPrimary; border = Theme.Border; }
            }

            float radius = Height >= 36 ? 9f : 6f;
            var rect = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            using (GraphicsPath path = RoundedPath(rect, radius))
            {
                using (var b = new SolidBrush(fill))
                    g.FillPath(b, path);
                if (border != Color.Empty)
                    using (var p = new Pen(border, 1f))
                        g.DrawPath(p, path);
            }

            // Кольцо фокуса для управления с клавиатуры.
            if (Focused && Enabled)
            {
                var inner = new RectangleF(2.5f, 2.5f, Width - 5f, Height - 5f);
                Color ring = _primary ? Color.FromArgb(130, 255, 255, 255) : Color.FromArgb(120, Theme.Accent);
                using (GraphicsPath ringPath = RoundedPath(inner, radius - 2f))
                using (var p = new Pen(ring, 1.5f))
                    g.DrawPath(p, ringPath);
            }

            TextRenderer.DrawText(g, Text, Font, ClientRectangle, text,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
                TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);
        }

        private static GraphicsPath RoundedPath(RectangleF r, float radius)
        {
            var p = new GraphicsPath();
            float d = Math.Max(2f, radius * 2f);
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }
    }
}
