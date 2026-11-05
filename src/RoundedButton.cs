using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Роль кнопки в интерфейсе — от неё зависят цвета и шрифт.</summary>
    public enum ButtonLook
    {
        /// <summary>Главное действие окна: акцентная заливка, полужирная подпись. Одна на окно.</summary>
        Primary,
        /// <summary>Обычное действие: светлая кнопка с рамкой.</summary>
        Secondary,
        /// <summary>Действие на тёмной подложке (полоса лупы в полноэкранном просмотре).</summary>
        OnDark,
        /// <summary>
        /// Навигация на цветной шапке: белый текст, полупрозрачная подложка при наведении.
        /// Не путать с Secondary — белая «таблетка» на градиенте выглядит чужой.
        /// </summary>
        OnHeader
    }

    /// <summary>Шеврон на кнопке. Рисуем сами — не символом в переводе.</summary>
    internal enum ButtonChevron
    {
        None,
        Left,
        Right,
        Up,
        Down,
        /// <summary>Вправо, но слева от текста — свёрнутый раскрывающийся список.</summary>
        Open
    }

    /// <summary>16 px знак слева на светлой кнопке панели (не шрифт — Windows 8.1 без MDL2).</summary>
    internal enum ButtonMark
    {
        None,
        Plus,
        Trash,
        Print,
        Open,
        Menu
    }

    /// <summary>
    /// Кнопка со скруглёнными углами и состояниями normal / hover / pressed / disabled /
    /// focused. Рисуется вручную, поэтому выглядит одинаково на всех версиях Windows и не
    /// зависит от системной темы — это единственная кнопка приложения (DRY): цвета, радиус,
    /// кольцо фокуса и поля подписи заданы здесь один раз.
    ///
    /// Три вещи в отрисовке сделаны намеренно и их легко потерять при правке:
    /// • НЕДОСТУПНАЯ обычная кнопка получает заливку, а не только серую надпись — иначе она
    ///   отличается от рабочей так слабо, что недоступность выясняют нажатием. В панелях
    ///   инструментов половина кнопок ждёт открытого документа, так что случай не редкий.
    /// • Радиус считается от высоты одной формулой: кнопки в приложении бывают 24, 28, 30, 32,
    ///   36 и 38 пикселей высотой, и любой фиксированный радиус выглядит на краях диапазона
    ///   либо рубленым, либо «таблеткой».
    /// • Подпись рисуется с боковыми полями, поэтому длинный перевод обрывается многоточием
    ///   ДО скругления угла, а не упирается в него.
    /// </summary>
    public class RoundedButton : Button
    {
        private const int TextPad = 10;   // боковые поля подписи
        private const float MinRadius = 5f, MaxRadius = 10f;

        private readonly ButtonLook _look;
        internal ButtonChevron Chevron { get; set; }
        internal ButtonMark Mark { get; set; }
        internal ButtonLook Look { get { return _look; } }
        private bool _selected;
        private bool _hover;
        private bool _pressed;
        private Point[] _chevronPts; // переиспользуем в OnPaint, не new[] на кадр

        public bool Selected
        {
            get { return _selected; }
            set
            {
                if (_selected == value) return;
                _selected = value;
                Invalidate();
            }
        }


        /// <summary>true — главное действие окна (акцентная), false — обычная кнопка.</summary>
        public RoundedButton(bool primary) : this(primary ? ButtonLook.Primary : ButtonLook.Secondary) { }

        public RoundedButton(ButtonLook look)
        {
            _look = look;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
            // Кэшированные шрифты: кнопок много (по 4–6 на окно), свой Font у каждой
            // копил бы GDI-объекты до финализатора при каждой пересборке окон.
            Font = look == ButtonLook.Primary
                ? Ui.Font(10.5f, FontStyle.Bold)
                : Ui.Font(9.75f);
            if (look == ButtonLook.OnHeader)
            {
                SetStyle(ControlStyles.SupportsTransparentBackColor, true);
                BackColor = Color.Transparent;
            }
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; _pressed = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { _pressed = true; Invalidate(); base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { _pressed = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }

        /// <summary>
        /// Радиус скругления по высоте кнопки: пропорция вместо ступеньки. Пределы — логические
        /// и растут с масштабом экрана (scale — <see cref="DpiScale.Device(Control)"/>): иначе
        /// кнопка высотой 45 px при 150 % упиралась в те же 10 px, что и 38 px при 100 %, и
        /// выглядела рубленее. Чистая — под тест.
        /// </summary>
        internal static float RadiusFor(int height, float scale = 1f)
        {
            float r = height / 4f;
            float min = MinRadius * scale, max = MaxRadius * scale;
            return r < min ? min : (r > max ? max : r);
        }

        /// <summary>
        /// Боковое поле подписи. Постоянные 10 px на узкой кнопке съедали её содержимое:
        /// у квадратной кнопки лупы (32 px) под текст оставалось 12 px, и «+» с «−»
        /// обрезались многоточием. Поэтому поле привязано к ширине. Чистая — под тест.
        /// </summary>
        internal static int TextPadFor(int width, float scale = 1f)
        {
            int pad = width / 8;
            int cap = (int)Math.Round(TextPad * scale); // поле растёт вместе с подписью
            return pad < cap ? pad : cap;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            var band = Parent as HeaderBand;
            if (_look == ButtonLook.OnHeader && band != null)
                band.PaintBehindChild(g, this);
            else
                g.Clear(Parent != null ? Parent.BackColor : Color.White);

            Color fill, text, border;
            Colors(out fill, out text, out border);

            float s = DpiScale.Device(this);
            float line = Math.Max(1f, s);
            float radius = RadiusFor(Height, s);
            if (_look == ButtonLook.OnHeader)
                radius = Math.Max(2f, Height / 2f - line);
            var rect = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            using (GraphicsPath path = Ui.RoundedRect(rect, radius))
            {
                if (fill.A > 0)
                {
                    using (var b = new SolidBrush(fill))
                        g.FillPath(b, path);
                }
                if (border != Color.Empty && border.A > 0)
                    using (var p = new Pen(border, line))
                        g.DrawPath(p, path);
            }

            // Кольцо фокуса для управления с клавиатуры: сплошное, а не полупрозрачное —
            // сквозь прозрачность на светлой заливке его почти не видно, а именно оно и
            // отвечает на вопрос «где я сейчас нахожусь».
            if (Focused && Enabled && _look != ButtonLook.OnHeader)
            {
                float inset = 3f * s;
                var inner = new RectangleF(inset, inset, Width - 2 * inset, Height - 2 * inset);
                Color ring = Theme.HighContrast ? text
                    : _look == ButtonLook.Secondary && !_selected ? Theme.Accent : Color.White;
                using (GraphicsPath ringPath = Ui.RoundedRect(inner, radius - 2f * s))
                using (var p = new Pen(ring, 1.6f * s))
                    g.DrawPath(p, ringPath);
            }

            int pad = TextPadFor(Width, s);
            int mark = Mark == ButtonMark.None ? 0 : (int)Math.Round(18f * s);
            int glyph = Chevron == ButtonChevron.None ? 0 : (int)Math.Round(18f * s);
            bool markOnly = Mark != ButtonMark.None && string.IsNullOrEmpty(Text) && Chevron == ButtonChevron.None;
            int insetL = markOnly ? 0 : mark + (Chevron == ButtonChevron.Right ? 0 : glyph);
            int insetR = Chevron == ButtonChevron.Right ? glyph : 0;
            var textArea = new Rectangle(pad + insetL, 0, Math.Max(1, Width - 2 * pad - insetL - insetR), Height);
            if (Mark != ButtonMark.None)
                DrawMark(g, markOnly ? (int)Math.Round((Width - 16f * s) / 2f) : pad, s, text);
            if (Chevron != ButtonChevron.None)
                DrawChevron(g, pad + mark, s, text);
            if (!string.IsNullOrEmpty(Text))
            {
                TextFormatFlags tf = TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine |
                    TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix;
                if (Mark == ButtonMark.None && Chevron == ButtonChevron.None)
                    tf |= TextFormatFlags.HorizontalCenter;
                TextRenderer.DrawText(g, Text, Font, textArea, text, tf);
            }
        }

        /// <summary>Шеврон влево. Чистая геометрия — под тест.</summary>
        internal static void ChevronPoints(int left, int height, float scale, out Point a, out Point b, out Point c)
        {
            ChevronPoints(ButtonChevron.Left, left, 0, height, scale, out a, out b, out c);
        }

        internal static void ChevronPoints(ButtonChevron dir, int pad, int width, int height, float scale,
            out Point a, out Point b, out Point c)
        {
            int arm = Math.Max(4, (int)Math.Round(7f * scale));
            int o = pad + (int)Math.Round(3f * scale);
            int midY = height / 2;
            switch (dir)
            {
                case ButtonChevron.Right:
                    a = new Point(width - o - arm, midY - arm);
                    b = new Point(width - o, midY);
                    c = new Point(width - o - arm, midY + arm);
                    return;
                case ButtonChevron.Open:
                    a = new Point(o, midY - arm);
                    b = new Point(o + arm, midY);
                    c = new Point(o, midY + arm);
                    return;
                case ButtonChevron.Up:
                    a = new Point(o, midY + arm);
                    b = new Point(o + arm, midY);
                    c = new Point(o + 2 * arm, midY + arm);
                    return;
                case ButtonChevron.Down:
                    a = new Point(o, midY - arm);
                    b = new Point(o + arm, midY);
                    c = new Point(o + 2 * arm, midY - arm);
                    return;
                default:
                    a = new Point(o + arm, midY - arm);
                    b = new Point(o, midY);
                    c = new Point(o + arm, midY + arm);
                    return;
            }
        }

        private void DrawChevron(Graphics g, int pad, float scale, Color color)
        {
            if (_chevronPts == null)
                _chevronPts = new Point[3];
            Point a, b, c;
            ChevronPoints(Chevron, pad, Width, Height, scale, out a, out b, out c);
            _chevronPts[0] = a;
            _chevronPts[1] = b;
            _chevronPts[2] = c;
            using (var p = new Pen(color, Math.Max(1.5f, 1.6f * scale)))
            {
                p.StartCap = LineCap.Round;
                p.EndCap = LineCap.Round;
                p.LineJoin = LineJoin.Round;
                g.DrawLines(p, _chevronPts);
            }
        }

        private void DrawMark(Graphics g, int pad, float scale, Color color)
        {
            float size = 16f * scale;
            float x = pad;
            float y = (Height - size) / 2f;
            float w = Math.Max(1.2f, 1.4f * scale);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var p = new Pen(color, w))
            {
                p.StartCap = LineCap.Round;
                p.EndCap = LineCap.Round;
                p.LineJoin = LineJoin.Round;
                float m = size / 2f;
                switch (Mark)
                {
                    case ButtonMark.Plus:
                        g.DrawLine(p, x + 2, y + m, x + size - 2, y + m);
                        g.DrawLine(p, x + m, y + 2, x + m, y + size - 2);
                        break;
                    case ButtonMark.Trash:
                        g.DrawLine(p, x + m - 2f, y + 2f, x + m + 2f, y + 2f);
                        g.DrawLine(p, x + 2f, y + 5f, x + size - 2f, y + 5f);
                        g.DrawLine(p, x + 4f, y + 5f, x + 5.2f, y + size - 2f);
                        g.DrawLine(p, x + size - 4f, y + 5f, x + size - 5.2f, y + size - 2f);
                        g.DrawLine(p, x + 5.2f, y + size - 2f, x + size - 5.2f, y + size - 2f);
                        break;
                    case ButtonMark.Print:
                        g.DrawRectangle(p, x + 4.5f, y + 1.5f, size - 9f, 5f);
                        g.DrawRectangle(p, x + 1.5f, y + 6.5f, size - 3f, 5.5f);
                        g.DrawLine(p, x + 4f, y + size - 2f, x + size - 4f, y + size - 2f);
                        break;
                    case ButtonMark.Menu:
                        float gap = 4.5f * scale;
                        float left = x + 3;
                        float right = x + size - 3;
                        g.DrawLine(p, left, y + m - gap, right, y + m - gap);
                        g.DrawLine(p, left, y + m, right, y + m);
                        g.DrawLine(p, left, y + m + gap, right, y + m + gap);
                        break;
                    default:
                        g.DrawLine(p, x + 2f, y + 6f, x + 6f, y + 6f);
                        g.DrawLine(p, x + 6f, y + 6f, x + 8f, y + 3f);
                        g.DrawLine(p, x + 8f, y + 3f, x + size - 2f, y + 3f);
                        g.DrawLine(p, x + size - 2f, y + 3f, x + size - 2f, y + size - 2f);
                        g.DrawLine(p, x + size - 2f, y + size - 2f, x + 2f, y + size - 2f);
                        g.DrawLine(p, x + 2f, y + size - 2f, x + 2f, y + 6f);
                        break;
                }
            }
        }

        /// <summary>Цвета текущего состояния: заливка, подпись, рамка (Empty — без рамки).</summary>
        private void Colors(out Color fill, out Color text, out Color border)
        {
            if (Theme.HighContrast)
            {
                // Все три вида — как кнопка Windows в этом режиме; главная отличается местом в окне.
                Theme.HighContrastPush(Enabled, _hover || _pressed || _selected, out fill, out text, out border);
                return;
            }
            border = Color.Empty;
            switch (_look)
            {
                case ButtonLook.Primary:
                    text = Enabled ? Color.White : Theme.DisabledText;
                    fill = !Enabled ? Theme.DisabledFill
                        : _pressed ? Theme.AccentPressed
                        : _hover ? Theme.AccentHover : Theme.Accent;
                    return;

                case ButtonLook.OnDark:
                    text = Enabled ? Color.White : Theme.DarkBarDisabledText;
                    fill = !Enabled ? Theme.DarkBarDisabledFill
                        : _pressed ? Theme.DarkBarPressed
                        : _hover ? Theme.DarkBarHover : Theme.DarkBarFill;
                    border = Theme.DarkBarBorder;
                    return;

                case ButtonLook.OnHeader:
                    text = Enabled ? Color.White : Color.FromArgb(190, 255, 255, 255);
                    Theme.HeaderChipColors(Enabled, _hover || _selected || Focused, _pressed, out fill, out border);
                    return;

                default:
                    if (_selected)
                    {
                        text = Enabled ? Color.White : Theme.DisabledText;
                        fill = !Enabled ? Theme.DisabledFill
                            : _pressed ? Theme.AccentPressed
                            : _hover ? Theme.AccentHover : Theme.Accent;
                        border = Theme.AccentPressed;
                        return;
                    }
                    text = Enabled ? Theme.TextPrimary : Theme.DisabledText;
                    fill = !Enabled ? Theme.DisabledSecondaryFill
                        : _pressed ? Theme.SecondaryPressed
                        : _hover ? Theme.SecondaryHover : Color.White;
                    border = !Enabled ? Theme.DisabledBorder
                        : (_pressed || _hover) ? Theme.BorderDark : Theme.Border;
                    return;
            }
        }
    }
}
