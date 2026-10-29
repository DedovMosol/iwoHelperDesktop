using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Брендовая шапка окна: панель с вертикальным градиентом акцентного цвета,
    /// заголовком и подписью в белом. Занимает то же место, что прежняя пара
    /// «заголовок + подпись», поэтому тело окна не сдвигается. Общая для
    /// стартового экрана и инструментов (DRY). Кнопку «Назад в меню» кладут
    /// прямо на неё как дочерний элемент.
    /// </summary>
    public class HeaderBand : Panel
    {
        private readonly string _title;
        private string _subtitle;
        private readonly Color _top;
        private readonly Color _bottom;
        /// <summary>Шрифт заголовка — под DPI монитора шапки (см. <see cref="Ui.Font(Control, float, FontStyle)"/>).</summary>
        private Font TitleFont { get { return Ui.Font(this, 15f, FontStyle.Bold); } }
        private static readonly Color SubtitleColor = Color.FromArgb(233, 238, 242); // нейтрально-белая: читаема на любом фоне

        /// <summary>Центрировать заголовок и подпись по ширине (для стартового экрана без кнопок).</summary>
        public bool Centered { get; set; }

        /// <summary>
        /// Подпись под заголовком. Меняется на ходу: стартовый экран называет ею текущий раздел
        /// («Выберите раздел» → «Инструменты PDF»), и человек всегда видит, где находится.
        /// </summary>
        public string Subtitle
        {
            get { return _subtitle; }
            set
            {
                string text = value ?? string.Empty;
                if (_subtitle == text)
                    return;
                _subtitle = text;
                Invalidate();
            }
        }

        /// <summary>
        /// Выровнять свой элемент по центру ТЕКСТОВОГО БЛОКА (заголовок плюс подпись) — так
        /// стоят глобус выбора языка и кнопка «Назад» на стартовом экране. Шапка расставляет их
        /// сама на каждой раскладке, и это не прихоть: строки считаются от РЕАЛЬНЫХ высот
        /// шрифтов, а те растут вместе с масштабом экрана. Координата, выставленная один раз в
        /// конструкторе, разъехалась бы с текстом на 125% и выше — ровно та ошибка, из-за
        /// которой высоты строк здесь уже считаются, а не задаются литералами.
        /// </summary>
        public void AlignToText(Control child)
        {
            if (child != null && !_aligned.Contains(child))
                _aligned.Add(child);
        }

        private readonly System.Collections.Generic.List<Control> _aligned =
            new System.Collections.Generic.List<Control>();

        protected override void OnLayout(LayoutEventArgs e)
        {
            base.OnLayout(e);
            if (_aligned.Count == 0)
                return;
            // Раскладка идёт и до масштаба окна, поэтому просвет — через Px, а не Device.
            int center = TextBlockCenter(Height, TitleFont.Height, Font.Height, DpiScale.Px(this, RowGap));
            foreach (Control child in _aligned)
            {
                if (child.IsDisposed)
                    continue; // элемент могли убрать из шапки, а список о нём ещё помнит
                int top = center - child.Height / 2;
                // Сравнение обязательно: присваивание внутри раскладки запускает её заново, и
                // без этой проверки получилась бы бесконечная рекурсия.
                if (child.Top != top)
                    child.Top = top;
            }
        }

        public HeaderBand(string title, string subtitle, Color top, Color bottom)
        {
            _title = title ?? string.Empty;
            _subtitle = subtitle ?? string.Empty;
            _top = top;
            _bottom = bottom;
            DoubleBuffered = true;
            SetStyle(ControlStyles.ResizeRedraw, true); // перерисовка градиента при растяжении
            ForeColor = Color.White;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Rectangle r = ClientRectangle;
            // Высокий контраст (UI-7): шапка — системного цвета окна, отделена линией текста.
            bool hc = Theme.HighContrast;
            Color titleColor = hc ? SystemColors.ControlText : Color.White;
            Color subtitleColor = hc ? SystemColors.ControlText : SubtitleColor;
            if (hc)
                e.Graphics.Clear(SystemColors.Control);
            else
                using (var brush = new LinearGradientBrush(r, _top, _bottom, LinearGradientMode.Vertical))
                    e.Graphics.FillRectangle(brush, r);
            // Тонкая нижняя грань отделяет шапку от тела окна.
            using (var pen = new Pen(hc ? SystemColors.ControlText : Color.FromArgb(60, 0, 0, 0)))
                e.Graphics.DrawLine(pen, 0, r.Height - 1, r.Width, r.Height - 1);

            // Высоты строк — РЕАЛЬНЫЕ высоты шрифтов: они заданы в пунктах и растут вместе с
            // масштабом экрана, поэтому фиксированные прямоугольники срезали бы низ букв.
            // Рисование — после масштаба окна: поля и просветы в пикселях устройства (UI-25).
            Func<int, int> d = delegate(int v) { return DpiScale.Device(this, v); };
            Font titleFont = TitleFont;
            int titleH = titleFont.Height, subtitleH = _subtitle.Length > 0 ? Font.Height : 0;
            int rowGap = d(RowGap);
            int titleY, subtitleY;
            TextRows(Height, titleH, subtitleH, out titleY, out subtitleY, rowGap);
            TextFormatFlags flags = TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis |
                TextFormatFlags.SingleLine | TextFormatFlags.NoPadding;

            if (Centered)
            {
                flags |= TextFormatFlags.HorizontalCenter;
                TextRenderer.DrawText(e.Graphics, _title, titleFont,
                    new Rectangle(d(10), titleY, Width - d(20), titleH), titleColor, flags);
                if (_subtitle.Length > 0)
                    TextRenderer.DrawText(e.Graphics, _subtitle, Font,
                        new Rectangle(d(10), subtitleY, Width - d(20), subtitleH), subtitleColor, flags);
                return;
            }

            // Заголовок не заходит под дочерние контролы (кнопку «Главная») — обрезаем многоточием.
            int leftmostChild = int.MaxValue, childBottom = 0;
            foreach (Control c in Controls)
                if (c.Visible)
                {
                    if (c.Left < leftmostChild) leftmostChild = c.Left;
                    if (c.Bottom > childBottom) childBottom = c.Bottom;
                }
            int rightBound = TextRightBound(Width, leftmostChild, d(20), d(12));
            TextRenderer.DrawText(e.Graphics, _title, titleFont,
                new Rectangle(d(18), titleY, rightBound - d(18), titleH), titleColor, flags);

            if (_subtitle.Length > 0)
            {
                // Подпись опускаем НИЖЕ кнопки «Главная» и даём полную ширину — иначе длинный
                // текст режется у левого края кнопки. Если под кнопкой не помещается по высоте —
                // остаёмся на прежнем месте, обрезая многоточием у кнопки.
                int subY = subtitleY, subRight = rightBound;
                if (childBottom + rowGap + subtitleH <= Height)
                {
                    subY = Math.Max(subtitleY, childBottom + rowGap);
                    subRight = Width - d(20);
                }
                TextRenderer.DrawText(e.Graphics, _subtitle, Font,
                    new Rectangle(d(20), subY, subRight - d(20), subtitleH), subtitleColor, flags);
            }
        }

        private const int RowGap = 2;    // просвет между заголовком и подписью

        /// <summary>
        /// Вертикальная раскладка шапки: блок «заголовок + подпись» стоит ПО ЦЕНТРУ полосы.
        /// Раньше он был прижат к низу, и над ним оставалась пустая синева — на глаз шапка
        /// выглядела съехавшей. Высоты строк приходят снаружи, потому что это РЕАЛЬНЫЕ высоты
        /// шрифтов: они заданы в пунктах и растут с масштабом экрана, а фиксированные
        /// прямоугольники (было 26 и 20) на 125% и выше срезали низ заголовка.
        /// subtitleHeight = 0 означает «подписи нет» — тогда центрируется один заголовок.
        /// Чистая — под тест.
        /// </summary>
        internal static void TextRows(int bandHeight, int titleHeight, int subtitleHeight,
            out int titleY, out int subtitleY, int rowGap = RowGap)
        {
            int gap = subtitleHeight > 0 ? rowGap : 0;
            int blockHeight = titleHeight + gap + subtitleHeight;
            titleY = (bandHeight - blockHeight) / 2;
            if (titleY < 0)
                titleY = 0; // полоса ниже текста — верх важнее низа, иначе срежется заголовок
            subtitleY = titleY + titleHeight + gap;
        }

        /// <summary>
        /// Правая граница текста шапки: не заходить под левый край дочерних
        /// контролов (кнопки «Назад»), но и не за правый край панели. Чистая — под тест.
        /// </summary>
        internal static int TextRightBound(int bandWidth, int leftmostChildLeft, int edge = 20, int gap = 12)
        {
            int bound = leftmostChildLeft < int.MaxValue ? leftmostChildLeft - gap : bandWidth - edge;
            return Math.Min(bound, bandWidth - edge);
        }

        /// <summary>
        /// Вертикальный центр текстового блока шапки: от верха заголовка до низа подписи.
        /// По нему выравнивается <see cref="TextAlignedIcon"/> — иначе иконка висит выше
        /// текста и выглядит забытой в углу. Чистая — под тест.
        /// </summary>
        internal static int TextBlockCenter(int bandHeight, int titleHeight, int subtitleHeight, int rowGap = RowGap)
        {
            int titleY, subtitleY;
            TextRows(bandHeight, titleHeight, subtitleHeight, out titleY, out subtitleY, rowGap);
            return (titleY + subtitleY + subtitleHeight) / 2;
        }
    }
}
