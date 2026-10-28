using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Логические пиксели (96 DPI, в них написана вся раскладка) против пикселей устройства.
    /// Окна строятся в логических единицах, а масштаб ставит <see cref="ScaledForm"/> ПОСЛЕ
    /// постройки: заданный до добавления контролов, автомасштаб WinForms срабатывал на пустом
    /// окне и дальше не делал ничего (UI-1). Два перевода нужны коду, который работает в
    /// другой системе единиц:
    /// <list type="bullet">
    /// <item><see cref="Px"/> — раскладка на ходу (OnLayout, рисование): константы логические,
    /// а окно до масштаба логическое, после — в пикселях устройства;</item>
    /// <item><see cref="ToLogical"/> — замер текста в конструкторе: шрифт уже в пикселях
    /// устройства, а окно ещё логическое, и без перевода размер вырос бы дважды.</item>
    /// </list>
    /// </summary>
    internal static class DpiScale
    {
        internal const float LogicalDpi = 96f;

        /// <summary>
        /// Шов для тестов: масштаб экрана, который изображает машина с 96 DPI (1,5 — 150 %).
        /// Шрифты (<see cref="Ui.Font"/>), автомасштаб окон и оба перевода умножаются на него;
        /// в работе он всегда 1, и масштаб берётся из настоящего DPI. Только UI-поток.
        /// </summary>
        internal static float EmulatedScale = 1f;

        private static int _systemDpi;

        /// <summary>
        /// DPI, в котором созданы шрифты (системный, по нему GDI переводит пункты в пиксели).
        /// Окно до создания хэндла сообщает тот же DPI (<see cref="Control.DeviceDpi"/>).
        /// </summary>
        private static int SystemDpi
        {
            get
            {
                if (_systemDpi == 0)
                    using (Graphics screen = Graphics.FromHwnd(IntPtr.Zero))
                        _systemDpi = Math.Max(1, (int)Math.Round(screen.DpiY));
                return _systemDpi;
            }
        }

        /// <summary>
        /// Логическая длина → единицы, в которых контрол живёт сейчас. Окно приложения до
        /// масштаба (пока строится) логическое, после — в пикселях устройства своего монитора.
        /// Контрол вне <see cref="ScaledForm"/> считается уже масштабированным.
        /// </summary>
        internal static int Px(Control control, int logical)
        {
            var form = control as ScaledForm ?? control.FindForm() as ScaledForm;
            if (form != null && !form.IsScaled)
                return logical;
            return (int)Math.Round(logical * (control.DeviceDpi / LogicalDpi * EmulatedScale));
        }

        /// <summary>
        /// Замер текста (всегда пиксели устройства) → единицы, в которых контрол живёт сейчас:
        /// до масштаба окна — логические, после — как есть. Пара к <see cref="Px"/> для
        /// раскладки, которую зовут и из конструктора, и на ходу (раскрытие секции).
        /// </summary>
        internal static int FromDevice(Control control, int device)
        {
            var form = control as ScaledForm ?? control.FindForm() as ScaledForm;
            return form != null && !form.IsScaled ? ToLogical(device) : device;
        }

        /// <summary>
        /// Ширина, на которой мерить перенос текста: шрифт уже в пикселях экрана, поэтому и
        /// ширину до масштаба окна (логическую) переводим туда же — иначе текст «переносился
        /// бы» на полторы ширины раньше и замер выходил выше настоящего.
        /// </summary>
        internal static int MeasureWidth(Control control, int current)
        {
            var form = control as ScaledForm ?? control.FindForm() as ScaledForm;
            if (form == null || form.IsScaled)
                return current;
            return (int)Math.Floor(current * (SystemDpi / LogicalDpi * EmulatedScale) + 0.001);
        }

        /// <summary>
        /// Замер текста (пиксели устройства) → логические единицы, с запасом вверх: после
        /// масштаба окна место должно оказаться не меньше замеренного.
        /// </summary>
        internal static int ToLogical(int device)
        {
            return (int)Math.Ceiling(device / (SystemDpi / LogicalDpi * EmulatedScale) - 0.001);
        }

        /// <summary>
        /// Низ контрола, чью высоту задаёт шрифт (подпись AutoSize, поле ввода, замеренный
        /// текст), в логических единицах — для раскладки в конструкторе («следующий ряд ниже»).
        /// У такого контрола Top логический, а высота уже в пикселях экрана, и сложить их как
        /// есть значило бы растянуть отступ под ним вместе с окном ещё раз.
        /// </summary>
        internal static int LogicalBottom(Control control)
        {
            return control.Top + ToLogical(control.Height);
        }

        /// <summary>То же для размера (например, GetPreferredSize подписи).</summary>
        internal static Size ToLogical(Size device)
        {
            return new Size(ToLogical(device.Width), ToLogical(device.Height));
        }

        /// <summary>Включить автомасштаб окна: построенное в логических единицах растянется под DPI.</summary>
        internal static void Apply(Form form)
        {
            float logical = LogicalDpi / EmulatedScale;
            form.AutoScaleDimensions = new SizeF(logical, logical);
            form.AutoScaleMode = AutoScaleMode.Dpi;
        }
    }
}
