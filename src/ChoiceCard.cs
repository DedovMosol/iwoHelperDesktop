using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Какой значок рисовать на карточке выбора.</summary>
    public enum CardGlyph
    {
        Excel,
        Pdf
    }

    /// <summary>
    /// Крупная кликабельная карточка выбора инструмента: значок, заголовок,
    /// описание. Состояния normal/hover/pressed/focused, активация мышью,
    /// Enter и Пробелом. Только отрисовка — поведение стандартной кнопки.
    /// </summary>
    public class ChoiceCard : Control
    {
        private static readonly Color PdfRed = Color.FromArgb(198, 40, 40);

        private readonly CardGlyph _glyph;
        private readonly string _title;
        private readonly string _description;
        private bool _hover;
        private bool _pressed;

        public ChoiceCard(CardGlyph glyph, string title, string description)
        {
            _glyph = glyph;
            _title = title;
            _description = description;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                     ControlStyles.Selectable, true);
            TabStop = true;
            Cursor = Cursors.Hand;
            BackColor = Color.White;
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; _pressed = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { _pressed = true; Focus(); Invalidate(); base.OnMouseDown(e); }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            bool wasPressed = _pressed;
            _pressed = false;
            Invalidate();
            base.OnMouseUp(e);
            if (wasPressed && ClientRectangle.Contains(e.Location) && e.Button == MouseButtons.Left)
                OnClick(EventArgs.Empty);
        }

        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }

        protected override bool IsInputKey(Keys keyData)
        {
            if (keyData == Keys.Enter || keyData == Keys.Space)
                return true; // получать эти клавиши, а не отдавать форме
            return base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space)
            {
                _pressed = true;
                Invalidate();
            }
            base.OnKeyDown(e);
        }

        protected override void OnKeyUp(KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space) && _pressed)
            {
                _pressed = false;
                Invalidate();
                OnClick(EventArgs.Empty);
            }
            base.OnKeyUp(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            g.Clear(Parent != null ? Parent.BackColor : Color.White);

            var rect = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            using (GraphicsPath card = Rounded(rect, 12f))
            {
                Color fill = _pressed ? Theme.SecondaryPressed : (_hover ? Theme.SecondaryHover : Color.White);
                using (var b = new SolidBrush(fill))
                    g.FillPath(b, card);
                Color border = (_hover || _pressed || Focused) ? Theme.Accent : Theme.Border;
                using (var pen = new Pen(border, (_hover || Focused) ? 2f : 1f))
                    g.DrawPath(pen, card);
            }

            int glyphSize = 60;
            var glyphRect = new Rectangle((Width - glyphSize) / 2, 26, glyphSize, glyphSize);
            DrawGlyph(g, glyphRect);

            using (var titleFont = new Font("Segoe UI", 12.5f, FontStyle.Bold))
            {
                var titleRect = new Rectangle(10, glyphRect.Bottom + 12, Width - 20, 28);
                TextRenderer.DrawText(g, _title, titleFont, titleRect, Theme.TextPrimary,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.Top | TextFormatFlags.EndEllipsis);
            }

            var descRect = new Rectangle(18, glyphRect.Bottom + 46, Width - 36, Height - glyphRect.Bottom - 56);
            TextRenderer.DrawText(g, _description, Font, descRect, Theme.TextMuted,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.Top | TextFormatFlags.WordBreak);
        }

        // Значок-документ с загнутым уголком (стиль file-excel): единое семейство
        // для обоих инструментов, различаются цветом и содержимым внутри листа.
        private void DrawGlyph(Graphics g, Rectangle r)
        {
            bool excel = _glyph == CardGlyph.Excel;
            Color main = excel ? Theme.Accent : PdfRed;
            Color fold = excel ? Theme.AccentPressed : Color.FromArgb(150, 26, 26);

            // Координаты из примера (viewBox 24) масштабируются в прямоугольник значка.
            float s = r.Width / 24f;
            Func<float, float, PointF> p = delegate(float x, float y)
            {
                return new PointF(r.X + x * s, r.Y + y * s);
            };

            using (var body = new GraphicsPath())
            {
                body.AddPolygon(new[] { p(4, 2), p(14, 2), p(20, 8), p(20, 22), p(4, 22) });
                using (var b = new SolidBrush(main))
                    g.FillPath(b, body);
            }
            using (var foldPath = new GraphicsPath())
            {
                foldPath.AddPolygon(new[] { p(14, 2), p(14, 8), p(20, 8) });
                using (var b = new SolidBrush(fold))
                    g.FillPath(b, foldPath);
            }

            if (excel)
            {
                // Белая рамка «таблицы» с перекладиной — как в примере.
                using (var pen = new Pen(Color.White, 1.8f * s))
                {
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    pen.LineJoin = LineJoin.Round;
                    using (var bracket = new GraphicsPath())
                    {
                        bracket.AddLines(new[] { p(15, 11), p(9, 11), p(9, 19), p(15, 19) });
                        g.DrawPath(pen, bracket);
                    }
                    g.DrawLine(pen, p(9, 15), p(14, 15));
                }
            }
            else
            {
                using (var f = new Font("Segoe UI", 8.5f * s, FontStyle.Bold))
                {
                    var textRect = new Rectangle(r.X, r.Y + (int)(6 * s), r.Width, (int)(16 * s));
                    TextRenderer.DrawText(g, "PDF", f, textRect, Color.White,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                }
            }
        }

        private static GraphicsPath Rounded(RectangleF r, float radius)
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
