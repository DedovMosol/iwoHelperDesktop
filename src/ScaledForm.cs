using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Окно приложения, которое масштабируется под DPI один раз — перед созданием хэндла,
    /// то есть когда конструкторы всей цепочки наследования уже построили его целиком в
    /// логических единицах. Раньше масштаб задавался первой строкой конструктора, на пустом
    /// окне: WinForms масштабировал пустоту, а добавленные потом контролы оставались в
    /// пикселях 96 DPI при шрифтах, выросших вместе с DPI (UI-1). Именно перед созданием
    /// хэндла, а не в Load: место окна для CenterScreen считается при его создании.
    ///
    /// Правило для наследников: раскладка, которая пересчитывается на ходу, живёт в OnLayout
    /// (не в OnResize) и переводит свои константы через <see cref="DpiScale.Px"/>.
    /// OnResize приходит посреди масштабирования — окно уже выросло, а дети ещё нет, и
    /// поставленное там WinForms тут же умножил бы ещё раз.
    /// </summary>
    public class ScaledForm : Form
    {
        private bool _scaled;

        /// <summary>Масштаб применён: дальше окно и его дети — в пикселях устройства.</summary>
        internal bool IsScaled { get { return _scaled; } }

        /// <summary>
        /// Шов для тестов: рабочая область экрана вместо настоящей (null — настоящая). Так
        /// маленький экран CI (1024×768) проверяется на любой машине.
        /// </summary>
        internal static Func<Rectangle, Rectangle> WorkingAreaForTests { get; set; }

        /// <summary>Рабочая область экрана, на котором лежит прямоугольник (не создаёт хэндлов).</summary>
        internal static Rectangle WorkingArea(Rectangle bounds)
        {
            return WorkingAreaForTests != null ? WorkingAreaForTests(bounds)
                : Screen.FromRectangle(bounds).WorkingArea;
        }



        protected override void CreateHandle()
        {
            if (!_scaled)
            {
                FitToScreen(); // ещё в логических единицах — до масштаба
                // Как в коде дизайнера: под SuspendLayout. Иначе уже первая из двух настроек
                // запускала раскладку, наследник видел IsScaled, а масштаба ещё не было, —
                // и его пиксели устройства WinForms потом умножал ещё раз (окно свода
                // получалось шире в 1,44 раза).
                SuspendLayout();
                try
                {
                    DpiScale.Apply(this);
                    _scaled = true;
                }
                finally { ResumeLayout(true); }
            }
            base.CreateHandle();
        }

        /// <summary>
        /// Окно, выросшее с масштабом, не должно уходить за край экрана: 780×785 на 150 % —
        /// это 1170×1178 пикселей при рабочей области ноутбука 1920×1032. Ужимаем клиентскую
        /// область ДО масштаба, в логических единицах: окно, которое выросло бы больше экрана,
        /// Windows обрезает посреди масштабирования — когда окно уже выросло, а контролы ещё
        /// нет, — и привязанные к низу кнопки остаются за краем. Не меньше MinimumSize: ниже
        /// него раскладка не рассчитана. Сохранённый размер (WindowPlacement) приходит позже,
        /// в Load, и проходит свою такую же проверку.
        /// </summary>
        private void FitToScreen()
        {
            // Только окна с изменяемым размером: их раскладка рассчитана на сжатие до
            // MinimumSize. Диалог с неизменяемой рамкой, ужатый по экрану, потерял бы низ —
            // кнопки остались бы за краем клиентской области без прокрутки.
            if (WindowState != FormWindowState.Normal || !TopLevel ||
                (FormBorderStyle != FormBorderStyle.Sizable && FormBorderStyle != FormBorderStyle.SizableToolWindow))
                return;
            // Не FromControl: он создаёт хэндл владельца, если его ещё нет.
            Rectangle near = Owner != null ? Owner.Bounds : new Rectangle(MousePosition, Size.Empty);
            Size frame = Size - ClientSize; // рамка и заголовок — системные, не масштабируются
            Size fit = FitClient(ClientSize, MinimumSize, frame, WorkingArea(near).Size, DpiScale.Device(this));
            if (fit != ClientSize)
                ClientSize = fit;
        }

        /// <summary>
        /// Клиентская область (логическая), которая после масштаба factor вместе с рамкой
        /// впишется в рабочую область, но не меньше минимума окна. Чистая — под тест.
        /// </summary>
        internal static Size FitClient(Size client, Size minimum, Size frame, Size area, float factor)
        {
            int maxW = (int)((area.Width - frame.Width) / factor);
            int maxH = (int)((area.Height - frame.Height) / factor);
            int minW = Math.Max(0, minimum.Width - frame.Width);
            int minH = Math.Max(0, minimum.Height - frame.Height);
            return new Size(Math.Min(client.Width, Math.Max(minW, maxW)),
                Math.Min(client.Height, Math.Max(minH, maxH)));
        }
    }
}
