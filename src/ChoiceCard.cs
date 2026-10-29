using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Какой значок рисовать на карточке выбора.</summary>
    public enum CardGlyph
    {
        Excel,
        Pdf,
        PdfSplit,
        Ocr,
        /// <summary>«PDF → PowerPoint» — буква «P» на оранжевом листе.</summary>
        Pptx,
        /// <summary>«Прочие операции» с PDF — ползунки настроек на красном листе.</summary>
        Tools,
        /// <summary>Сравнение двух цифровых PDF — две страницы бок о бок.</summary>
        Review,
        /// <summary>Раздел «Иной функционал» — плитка приложений на синем листе.</summary>
        Other
    }

    /// <summary>
    /// Крупная кликабельная карточка выбора инструмента: значок, заголовок,
    /// описание. Состояния normal/hover/pressed/focused, активация мышью,
    /// Enter и Пробелом. Только отрисовка — поведение стандартной кнопки.
    /// </summary>
    public class ChoiceCard : PushControl
    {
        /// <summary>
        /// Шрифт заголовка — под DPI монитора карточки. Статическое поле застывало бы в том
        /// масштабе, в каком тип впервые тронули: свой Font контрола WinForms пересчитывает
        /// при переходе окна на монитор с другим DPI, а нарисованный вручную заголовок — нет.
        /// </summary>
        private Font TitleFont { get { return Ui.Font(this, 12.5f, FontStyle.Bold); } }

        // Ритм содержимого: значок, отступ, заголовок в одну строку, отступ, описание с переносами.
        private const int GlyphSize = 60, GlyphGap = 12, TitleGap = 6, MinTop = 18;
        private const int TitlePad = 10, DescPad = 18;
        private const TextFormatFlags TitleFlags =
            TextFormatFlags.HorizontalCenter | TextFormatFlags.Top | TextFormatFlags.EndEllipsis;
        private const TextFormatFlags DescFlags =
            TextFormatFlags.HorizontalCenter | TextFormatFlags.Top | TextFormatFlags.WordBreak;

        private readonly CardGlyph _glyph;
        private readonly string _title;
        private readonly string _description;

        private string[] _acceptExtensions; // расширения принимаемых дропом файлов (null — не принимать)
        private bool _dragOk;               // текущая drag-сессия несёт подходящие файлы (решено в DragEnter)

        /// <summary>Раскрытые дропом на карточку файлы (уже отфильтрованы по расширению).</summary>
        public event Action<string[]> FilesDropped;

        /// <summary>Разрешить дроп файлов с этими расширениями (например «.pdf»); включает подсветку эффекта.</summary>
        public void AcceptFiles(params string[] extensions)
        {
            _acceptExtensions = extensions;
            AllowDrop = extensions != null && extensions.Length > 0;
        }

        /// <summary>Пути с одним из расширений exts (регистр не важен, файл должен существовать). Чистая — под тест.</summary>
        internal static string[] FilterByExtension(string[] paths, string[] exts)
        {
            var result = new List<string>();
            if (paths == null || exts == null)
                return result.ToArray();
            foreach (string path in paths)
            {
                if (!File.Exists(path))
                    continue;
                string ext = Path.GetExtension(path);
                foreach (string e in exts)
                    if (string.Equals(ext, e, StringComparison.OrdinalIgnoreCase))
                    {
                        result.Add(path);
                        break;
                    }
            }
            return result.ToArray();
        }

        private string[] AcceptedFrom(IDataObject data)
        {
            if (_acceptExtensions == null || data == null || !data.GetDataPresent(DataFormats.FileDrop))
                return new string[0];
            return FilterByExtension(data.GetData(DataFormats.FileDrop) as string[], _acceptExtensions);
        }

        protected override void OnDragEnter(DragEventArgs e)
        {
            // Подходящесть выясняем один раз за drag-сессию: DragOver сыплется на каждое
            // движение мыши, разбирать HDROP и звать File.Exists там было бы расточительно.
            _dragOk = AcceptedFrom(e.Data).Length > 0;
            e.Effect = _dragOk ? DragDropEffects.Copy : DragDropEffects.None;
            base.OnDragEnter(e);
        }

        protected override void OnDragOver(DragEventArgs e)
        {
            e.Effect = _dragOk ? DragDropEffects.Copy : DragDropEffects.None; // из кэша DragEnter
            base.OnDragOver(e);
        }

        protected override void OnDragLeave(EventArgs e)
        {
            _dragOk = false;
            base.OnDragLeave(e);
        }

        protected override void OnDragDrop(DragEventArgs e)
        {
            _dragOk = false;
            string[] paths = AcceptedFrom(e.Data);
            if (paths.Length > 0)
            {
                Action<string[]> h = FilesDropped;
                if (h != null) h(paths);
            }
            base.OnDragDrop(e);
        }

        public ChoiceCard(CardGlyph glyph, string title, string description)
        {
            _glyph = glyph;
            _title = title;
            _description = description;
            BackColor = Color.White;
            // Доступность: экранный диктор объявляет карточку как кнопку (PushControl) с названием.
            AccessibleName = title;
            AccessibleDescription = description;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            g.Clear(Parent != null ? Parent.BackColor : Color.White);

            // Рисование — всегда после масштаба окна: размеры значка, поля, рамки и скругление
            // в пикселях устройства (UI-25).
            float s = DpiScale.Device(this);
            Func<int, int> d = delegate(int v) { return DpiScale.Device(this, v); };
            // Наведение — рамка карточки; фокус — отдельное кольцо внутри (UI-10): раньше оба
            // рисовались одной рамкой, и с клавиатуры нельзя было понять, где фокус, а где мышь.
            float line = (Hover ? 2f : 1f) * Math.Max(1f, s);
            var rect = new RectangleF(line / 2, line / 2, Width - line, Height - line);
            Color fill, border, title = Theme.TextPrimary, description = Theme.TextMuted;
            if (Theme.HighContrast) // UI-7: системные цвета кнопки; описание — тем же текстом
                Theme.HighContrastPush(Enabled, Hover || Pressed, out fill, out title, out border);
            else
            {
                fill = Pressed ? Theme.SecondaryPressed : (Hover ? Theme.SecondaryHover : Color.White);
                border = (Hover || Pressed) ? Theme.Accent : Theme.Border;
            }
            if (Theme.HighContrast)
                description = title;
            using (GraphicsPath card = Ui.RoundedRect(rect, 12f * s))
            {
                using (var b = new SolidBrush(fill))
                    g.FillPath(b, card);
                using (var pen = new Pen(border, line))
                    g.DrawPath(pen, card);
            }
            DrawFocusRing(g, new RectangleF(0, 0, Width, Height), 12f * s, Theme.HighContrast ? title : Theme.Accent);

            // Блок «значок + заголовок + описание» ставим по центру карточки: описания
            // разной длины, и при верхней привязке карточки с короткой подписью выглядели
            // пустыми снизу. Высоту меряем теми же флагами, какими рисуем.
            int glyph = d(GlyphSize), glyphGap = d(GlyphGap), titleGap = d(TitleGap);
            int titlePad = d(TitlePad), descPad = d(DescPad);
            int titleW = Width - 2 * titlePad, descW = Width - 2 * descPad;
            int titleH = TextRenderer.MeasureText(g, _title, TitleFont,
                new Size(titleW, int.MaxValue), TitleFlags).Height;
            int descH = TextRenderer.MeasureText(g, _description, Font,
                new Size(descW, int.MaxValue), DescFlags).Height;
            int top = ContentTop(Height, glyph + glyphGap + titleH + titleGap + descH, d(MinTop));

            var glyphRect = new Rectangle((Width - glyph) / 2, top, glyph, glyph);
            DrawGlyph(g, glyphRect);

            var titleRect = new Rectangle(titlePad, glyphRect.Bottom + glyphGap, titleW, titleH);
            TextRenderer.DrawText(g, _title, TitleFont, titleRect, title, TitleFlags);

            var descRect = new Rectangle(descPad, titleRect.Bottom + titleGap, descW, descH);
            TextRenderer.DrawText(g, _description, Font, descRect, description, DescFlags);
        }

        /// <summary>
        /// Верх блока содержимого, центрированного по высоте карточки; minTop — нижний предел
        /// в тех же единицах (по умолчанию логический). Чистая — под тест.
        /// </summary>
        internal static int ContentTop(int cardHeight, int blockHeight, int minTop = MinTop)
        {
            int top = (cardHeight - blockHeight) / 2;
            return top < minTop ? minTop : top; // очень длинное описание прижимаем к верху, а не режем сверху
        }

        // Значок-документ с загнутым уголком (стиль file-excel): единое семейство
        // для обоих инструментов, различаются цветом и содержимым внутри листа.
        private void DrawGlyph(Graphics g, Rectangle r)
        {
            DrawGlyph(g, _glyph, r);
        }

        /// <summary>
        /// Тот же значок, но для любого места, где он нужен: карточка недавнего файла на
        /// стартовом экране рисует его же. Рисование вынесено, а не скопировано, — иначе
        /// значки инструмента и его результата однажды разошлись бы.
        /// </summary>
        internal static void DrawGlyph(Graphics g, CardGlyph _glyph, Rectangle r)
        {
            bool excel = _glyph == CardGlyph.Excel;
            Color main, fold;
            if (excel) { main = Theme.Accent; fold = Theme.AccentPressed; }
            else if (_glyph == CardGlyph.Ocr) { main = Theme.WordViolet; fold = Theme.WordVioletDark; }
            else if (_glyph == CardGlyph.Pptx) { main = Theme.PowerPointOrange; fold = Theme.PowerPointOrangeDark; }
            else if (_glyph == CardGlyph.Review) { main = Theme.ReviewBlue; fold = Theme.ReviewBlueDark; }
            else if (_glyph == CardGlyph.Other) { main = Theme.HubBlue; fold = Theme.HubBlueDark; }
            else { main = Theme.PdfRed; fold = Theme.PdfRedDark; }

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
            else if (_glyph == CardGlyph.Tools)
            {
                // Три ползунка — общепринятый знак «настройки и прочие действия».
                using (var pen = new Pen(Color.White, 1.6f * s))
                using (var knob = new SolidBrush(Color.White))
                {
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    float[] rows = { 11.5f, 15f, 18.5f };
                    float[] knobs = { 15f, 9.5f, 13f }; // ползунки на разных позициях — это регуляторы
                    for (int i = 0; i < rows.Length; i++)
                    {
                        g.DrawLine(pen, p(7f, rows[i]), p(17f, rows[i]));
                        PointF c = p(knobs[i], rows[i]);
                        float radius = 1.7f * s;
                        g.FillEllipse(knob, c.X - radius, c.Y - radius, 2 * radius, 2 * radius);
                    }
                }
            }
            else if (_glyph == CardGlyph.Review)
            {
                // Две версии бок о бок и связь сравнения между ними.
                using (var pen = new Pen(Color.White, 1.35f * s))
                {
                    pen.LineJoin = LineJoin.Round;
                    PointF left = p(6.5f, 10.5f), right = p(13.3f, 10.5f);
                    g.DrawRectangle(pen, left.X, left.Y, 4.2f * s, 7.8f * s);
                    g.DrawRectangle(pen, right.X, right.Y, 4.2f * s, 7.8f * s);
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    g.DrawLine(pen, p(10.8f, 14.4f), p(13.2f, 14.4f));
                }
            }
            else if (_glyph == CardGlyph.Other)
            {
                // Плитка из четырёх квадратов — «прочие инструменты», не про PDF.
                using (var brush = new SolidBrush(Color.White))
                {
                    float side = 3.6f * s, gap = 1.6f * s;
                    PointF start = p(8f, 11f);
                    for (int row = 0; row < 2; row++)
                        for (int col = 0; col < 2; col++)
                            g.FillRectangle(brush, start.X + col * (side + gap),
                                start.Y + row * (side + gap), side, side);
                }
            }
            else if (_glyph == CardGlyph.Pptx)
            {
                // Буква «P» (PowerPoint) — перенос страниц в слайды с настоящим текстом.
                using (var pen = new Pen(Color.White, 1.8f * s))
                {
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    pen.LineJoin = LineJoin.Round;
                    DrawPolyline(g, pen, p, new[] { new[] { 9f, 19f }, new[] { 9f, 11f }, new[] { 13f, 11f },
                        new[] { 14.8f, 12.8f }, new[] { 13f, 14.6f }, new[] { 9f, 14.6f } });
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

    }
}
