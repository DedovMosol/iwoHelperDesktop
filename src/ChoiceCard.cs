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
        Pdf,
        PdfSplit,
        Ocr
    }

    /// <summary>
    /// Крупная кликабельная карточка выбора инструмента: значок, заголовок,
    /// описание. Состояния normal/hover/pressed/focused, активация мышью,
    /// Enter и Пробелом. Только отрисовка — поведение стандартной кнопки.
    /// </summary>
    public class ChoiceCard : Control
    {
        private static readonly Color PdfRed = Color.FromArgb(211, 47, 47);      // #D32F2F
        private static readonly Color PdfFold = Color.FromArgb(154, 34, 34);     // #9A2222
        // Единый шрифт заголовка карточек: создание в каждом OnPaint — лишний GDI-чурн.
        private static readonly Font TitleFont = new Font("Segoe UI", 12.5f, FontStyle.Bold);

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
            // Доступность: экранный диктор объявляет карточку как кнопку с названием.
            AccessibleRole = AccessibleRole.PushButton;
            AccessibleName = title;
            AccessibleDescription = description;
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; _pressed = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { _pressed = true; Focus(); Invalidate(); base.OnMouseDown(e); }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            _pressed = false;
            Invalidate();
            base.OnMouseUp(e); // Click поднимает база (ControlStyles.StandardClick) — не дублируем
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

            var titleRect = new Rectangle(10, glyphRect.Bottom + 12, Width - 20, 28);
            TextRenderer.DrawText(g, _title, TitleFont, titleRect, Theme.TextPrimary,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.Top | TextFormatFlags.EndEllipsis);

            var descRect = new Rectangle(18, glyphRect.Bottom + 46, Width - 36, Height - glyphRect.Bottom - 56);
            TextRenderer.DrawText(g, _description, Font, descRect, Theme.TextMuted,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.Top | TextFormatFlags.WordBreak);
        }

        // Значок-документ с загнутым уголком (стиль file-excel): единое семейство
        // для обоих инструментов, различаются цветом и содержимым внутри листа.
        private void DrawGlyph(Graphics g, Rectangle r)
        {
            bool excel = _glyph == CardGlyph.Excel;
            Color main, fold;
            if (excel) { main = Theme.Accent; fold = Theme.AccentPressed; }
            else if (_glyph == CardGlyph.Ocr) { main = Theme.WordViolet; fold = Theme.WordVioletDark; }
            else { main = PdfRed; fold = PdfFold; }

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
            else if (_glyph == CardGlyph.PdfSplit)
            {
                // Ножницы — универсальный знак «разрезать/разделить».
                using (var pen = new Pen(Color.White, 1.4f * s))
                {
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    g.DrawLine(pen, p(9f, 16.5f), p(17f, 6.5f));   // лезвие
                    g.DrawLine(pen, p(15f, 16.5f), p(7f, 6.5f));   // лезвие
                    g.DrawEllipse(pen, p(7f, 15.5f).X, p(7f, 15.5f).Y, 3.4f * s, 3.4f * s);   // кольцо
                    g.DrawEllipse(pen, p(13.6f, 15.5f).X, p(13.6f, 15.5f).Y, 3.4f * s, 3.4f * s); // кольцо
                }
            }
            else if (_glyph == CardGlyph.Ocr)
            {
                // Буква «W» (Word) — извлечение текста цифрового PDF в .docx.
                using (var pen = new Pen(Color.White, 1.8f * s))
                {
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    pen.LineJoin = LineJoin.Round;
                    DrawPolyline(g, pen, p, new[] { new[] { 7f, 11f }, new[] { 9.5f, 18f }, new[] { 12f, 13.5f }, new[] { 14.5f, 18f }, new[] { 17f, 11f } });
                }
            }
            else
            {
                // Векторные буквы «PDF» из file-pdf.svg — крупнее и чётче текста.
                using (var pen = new Pen(Color.White, 1.5f * s))
                {
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    pen.LineJoin = LineJoin.Round;
                    DrawPolyline(g, pen, p, new[] { new[] { 6.5f, 19f }, new[] { 6.5f, 11.5f }, new[] { 9f, 11.5f }, new[] { 9f, 15f }, new[] { 6.5f, 15f } });
                    DrawPolyline(g, pen, p, new[] { new[] { 11f, 19f }, new[] { 11f, 11.5f }, new[] { 13f, 11.5f }, new[] { 14.5f, 13.5f }, new[] { 14.5f, 17f }, new[] { 13f, 19f }, new[] { 11f, 19f } });
                    DrawPolyline(g, pen, p, new[] { new[] { 16.5f, 19f }, new[] { 16.5f, 11.5f }, new[] { 19.5f, 11.5f } });
                    DrawPolyline(g, pen, p, new[] { new[] { 16.5f, 15f }, new[] { 19f, 15f } });
                }
            }
        }

        private static void DrawPolyline(Graphics g, Pen pen, Func<float, float, PointF> map, float[][] pts)
        {
            var points = new PointF[pts.Length];
            for (int i = 0; i < pts.Length; i++)
                points[i] = map(pts[i][0], pts[i][1]);
            g.DrawLines(pen, points);
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
