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
        /// <summary>Шеврон слева (шапка «Назад»). Рисуем сами, не символом в переводе.</summary>
        internal bool ChevronLeft { get; set; }
        internal ButtonLook Look { get { return _look; } }
        private bool _selected;
        private bool _hover;
        private bool _pressed;
        private Point[] _chevronPts; // переиспользуем в OnPaint, не new[] на кадр
        private Region _headerShape; // живой: GetHrgn после Dispose даёт «недопустимый параметр»

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
                BackColor = Color.Transparent; // углы: Clear возьмёт BackColor шапки (акцент)
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; _pressed = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { _pressed = true; Invalidate(); base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { _pressed = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            if (!IsHandleCreated)
                return;
            ApplyHeaderShape();
        }

        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (Visible && IsHandleCreated)
                ApplyHeaderShape();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Region = null;
                if (_headerShape != null)
                {
                    _headerShape.Dispose();
                    _headerShape = null;
                }
            }
            base.Dispose(disposing);
        }

        /// <summary>Радиус пилюли на шапке. Совпадает у Region и у обводки, иначе торчат углы.</summary>
        internal static float HeaderChipRadius(int height, float scale)
        {
            float line = Math.Max(1f, scale);
            float cap = 14f * scale;
            float r = height / 2f - line;
            return r < cap ? r : cap;
        }

        private void ApplyHeaderShape()
        {
            if (_look != ButtonLook.OnHeader || !IsHandleCreated || Width < 4 || Height < 4)
                return;
            float radius = HeaderChipRadius(Height, DpiScale.Device(this));
            using (GraphicsPath path = Ui.RoundedRect(new RectangleF(0, 0, Width, Height), radius))
            {
                Region next = new Region(path);
                Region = next;
                if (_headerShape != null)
                    _headerShape.Dispose();
                _headerShape = next;
            }
            if (Parent != null)
                Parent.Invalidate(Bounds, false);
        }

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
            Color ground = Parent != null ? Parent.BackColor : Color.White;
            if (_look == ButtonLook.OnHeader && Parent != null)
                ground = Parent.BackColor;
            g.Clear(ground);

            Color fill, text, border;
            Colors(out fill, out text, out border);

            // Рисование — всегда после масштаба окна, в пикселях устройства (UI-25).
            float s = DpiScale.Device(this);
            float line = Math.Max(1f, s);
            float radius = _look == ButtonLook.OnHeader
                ? HeaderChipRadius(Height, s)
                : RadiusFor(Height, s);
            var rect = new RectangleF(line / 2, line / 2, Width - line, Height - line);
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
            if (Focused && Enabled)
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
            int chevron = ChevronLeft ? (int)Math.Round(12f * s) : 0;
            var textArea = new Rectangle(pad + chevron, 0, Math.Max(1, Width - 2 * pad - chevron), Height);
            if (ChevronLeft)
                DrawChevronLeft(g, pad, Height, s, text);
            TextRenderer.DrawText(g, Text, Font, textArea, text,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
                TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);
        }

        /// <summary>Шеврон «назад»: две линии, без юникода в строке. Чистая геометрия — под тест.</summary>
        internal static void ChevronPoints(int left, int height, float scale, out Point a, out Point b, out Point c)
        {
            int x = left + (int)Math.Round(3f * scale);
            int mid = height / 2;
            int arm = Math.Max(3, (int)Math.Round(5f * scale));
            a = new Point(x + arm, mid - arm);
            b = new Point(x, mid);
            c = new Point(x + arm, mid + arm);
        }

        private void DrawChevronLeft(Graphics g, int left, int height, float scale, Color color)
        {
            if (_chevronPts == null)
                _chevronPts = new Point[3];
            Point a, b, c;
            ChevronPoints(left, height, scale, out a, out b, out c);
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
                    // Fluent subtle: лёгкий подъём тона полосы и тонкая кромка, не молоко.
                    text = Enabled ? Color.White : Color.FromArgb(190, 255, 255, 255);
                    fill = !Enabled ? Color.FromArgb(28, 255, 255, 255)
                        : _pressed ? Color.FromArgb(72, 255, 255, 255)
                        : (_hover || _selected) ? Color.FromArgb(56, 255, 255, 255)
                        : Color.FromArgb(38, 255, 255, 255);
                    border = Enabled
                        ? Color.FromArgb(_hover || _pressed || _selected ? 160 : 110, 255, 255, 255)
                        : Color.FromArgb(70, 255, 255, 255);
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
