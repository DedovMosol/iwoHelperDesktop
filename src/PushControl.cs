using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Рисуемый вручную контрол, который ведёт себя как кнопка: карточка инструмента, карточка
    /// недавнего файла, кнопка-значок. Раньше каждая из трёх держала свою копию поведения, и
    /// копии разошлись (UI-21): одна нажималась по отпусканию клавиши, две — по нажатию; одна
    /// показывала фокус тем же цветом, что и наведение, а другая теряла его при уходе мыши
    /// (UI-10); у кнопки-значка не было роли кнопки, и экранный диктор называл её «клиент»
    /// (UI-4). Здесь — одно поведение на всех:
    /// <list type="bullet">
    /// <item>мышь: наведение, нажатие, фокус по нажатию, Click — от базы (StandardClick);</item>
    /// <item>клавиатура: Enter и Пробел нажимают по отпусканию, как кнопка Windows по Пробелу —
    /// удержание показывает нажатое состояние, а отпускание вне фокуса ничего не делает;</item>
    /// <item>фокус: отдельное кольцо (<see cref="DrawFocusRing"/>), не зависящее от мыши;</item>
    /// <item>доступность: роль «кнопка» и действие по умолчанию «Нажать» — им пользуется
    /// диктор (MSAA DoDefaultAction, в UIA — Invoke через LegacyIAccessible).</item>
    /// </list>
    /// </summary>
    public abstract class PushControl : Control
    {
        protected PushControl()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                     ControlStyles.Selectable | ControlStyles.StandardClick |
                     ControlStyles.SupportsTransparentBackColor, true);
            TabStop = true;
            Cursor = Cursors.Hand;
            AccessibleRole = AccessibleRole.PushButton;
        }

        /// <summary>Курсор над контролом.</summary>
        protected bool Hover { get; private set; }

        /// <summary>Нажат мышью или клавишей и ещё не отпущен.</summary>
        protected bool Pressed { get; private set; }

        /// <summary>
        /// Нажатие начала клавиша (а не мышь). Одно нажатие — один жест: пока держат мышь,
        /// клавиша не нажимает (иначе нажатие клавишей, а затем отпускание мыши открыли бы
        /// два окна), а взявшаяся за контрол мышь отменяет нажатие клавишей.
        /// </summary>
        private bool _keyPressed;

        /// <summary>Нажатие, как с клавиатуры или из диктора: только у доступного контрола.</summary>
        internal void PerformClick()
        {
            if (Enabled && Visible)
                OnClick(EventArgs.Empty);
        }

        private void SetState(bool hover, bool pressed)
        {
            if (Hover == hover && Pressed == pressed)
                return;
            Hover = hover;
            Pressed = pressed;
            Invalidate();
        }

        protected override void OnMouseEnter(EventArgs e) { SetState(true, Pressed); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { SetState(false, _keyPressed); base.OnMouseLeave(e); }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _keyPressed = false;
                Focus();
                SetState(Hover, true);
            }
            base.OnMouseDown(e);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            SetState(Hover, false);
            base.OnMouseUp(e); // Click поднимает база (StandardClick) — не дублируем
        }

        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }

        protected override void OnLostFocus(EventArgs e)
        {
            _keyPressed = false; // клавиша, отпущенная уже в другом контроле, не нажимает этот
            SetState(Hover, false);
            Invalidate();
            base.OnLostFocus(e);
        }

        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }

        protected override bool IsInputKey(Keys keyData)
        {
            return keyData == Keys.Enter || keyData == Keys.Space || base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space) && e.Modifiers == Keys.None)
            {
                if (!Pressed) // держат мышь — жест уже идёт, клавиша его не перехватывает
                {
                    _keyPressed = true;
                    SetState(Hover, true);
                }
                e.Handled = true;
            }
            base.OnKeyDown(e);
        }

        protected override void OnKeyUp(KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space) && _keyPressed)
            {
                _keyPressed = false;
                SetState(Hover, false);
                e.Handled = true;
                PerformClick();
            }
            base.OnKeyUp(e);
        }

        /// <summary>
        /// Кольцо фокуса внутри bounds: сплошное, толщиной 2 логических px, на 3 px внутрь
        /// (UI-10). Показывается, пока контрол в фокусе, — и мышь его не гасит. Цвет задаёт
        /// контрол: на белом — акцентный, на цветной шапке — белый.
        /// </summary>
        protected void DrawFocusRing(Graphics g, RectangleF bounds, float radius, Color color)
        {
            if (!Focused)
                return;
            float s = DpiScale.Device(this);
            float inset = 3f * s, width = 2f * s;
            var ring = new RectangleF(bounds.X + inset, bounds.Y + inset,
                bounds.Width - 2 * inset, bounds.Height - 2 * inset);
            if (ring.Width <= 0 || ring.Height <= 0)
                return;
            SmoothingMode was = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (GraphicsPath path = Ui.RoundedRect(ring, Math.Max(0f, radius - inset)))
            using (var pen = new Pen(color, width))
                g.DrawPath(pen, path);
            g.SmoothingMode = was;
        }

        protected override AccessibleObject CreateAccessibilityInstance()
        {
            return new PushAccessibleObject(this);
        }

        /// <summary>Для диктора: «кнопка», действие «Нажать», выполнение — то же нажатие.</summary>
        private sealed class PushAccessibleObject : ControlAccessibleObject
        {
            private readonly PushControl _owner;

            public PushAccessibleObject(PushControl owner) : base(owner) { _owner = owner; }

            /// <summary>Своё описание (AccessibleDefaultActionDescription) главнее общего.</summary>
            public override string DefaultAction
            {
                get
                {
                    string own = base.DefaultAction;
                    return string.IsNullOrEmpty(own) ? Loc.T("common.a11y.press") : own;
                }
            }

            public override void DoDefaultAction() { _owner.PerformClick(); }
        }
    }
}
