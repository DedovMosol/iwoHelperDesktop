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
        OnDark
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
        private bool _selected;
        private bool _hover;
        private bool _pressed;

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
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; _pressed = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { _pressed = true; Invalidate(); base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { _pressed = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }

        /// <summary>
        /// Радиус скругления по высоте кнопки: пропорция вместо ступеньки. Чистая — под тест.
        /// </summary>
        internal static float RadiusFor(int height)
        {
            float r = height / 4f;
            return r < MinRadius ? MinRadius : (r > MaxRadius ? MaxRadius : r);
        }

        /// <summary>
        /// Боковое поле подписи. Постоянные 10 px на узкой кнопке съедали её содержимое:
        /// у квадратной кнопки лупы (32 px) под текст оставалось 12 px, и «+» с «−»
        /// обрезались многоточием. Поэтому поле привязано к ширине. Чистая — под тест.
        /// </summary>
        internal static int TextPadFor(int width)
        {
            int pad = width / 8;
            return pad < TextPad ? pad : TextPad;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Parent != null ? Parent.BackColor : Color.White);

            Color fill, text, border;
            Colors(out fill, out text, out border);

            float radius = RadiusFor(Height);
            var rect = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            using (GraphicsPath path = Ui.RoundedRect(rect, radius))
            {
                using (var b = new SolidBrush(fill))
                    g.FillPath(b, path);
                if (border != Color.Empty)
                    using (var p = new Pen(border, 1f))
                        g.DrawPath(p, path);
            }

            // Кольцо фокуса для управления с клавиатуры: сплошное, а не полупрозрачное —
            // сквозь прозрачность на светлой заливке его почти не видно, а именно оно и
            // отвечает на вопрос «где я сейчас нахожусь».
            if (Focused && Enabled)
            {
                var inner = new RectangleF(3f, 3f, Width - 6f, Height - 6f);
                Color ring = _look == ButtonLook.Secondary && !_selected
                    ? Theme.Accent : Color.White;
                using (GraphicsPath ringPath = Ui.RoundedRect(inner, radius - 2f))
                using (var p = new Pen(ring, 1.6f))
                    g.DrawPath(p, ringPath);
            }

            int pad = TextPadFor(Width);
            var textArea = new Rectangle(pad, 0, Math.Max(1, Width - 2 * pad), Height);
            TextRenderer.DrawText(g, Text, Font, textArea, text,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
                TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);
        }

        /// <summary>Цвета текущего состояния: заливка, подпись, рамка (Empty — без рамки).</summary>
        private void Colors(out Color fill, out Color text, out Color border)
        {
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
