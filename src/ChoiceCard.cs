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
        /// <summary>Раздел «Иной функционал» — плитка приложений на синем листе.</summary>
        Other
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
        private bool _hover;
        private bool _pressed;
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
            using (GraphicsPath card = Ui.RoundedRect(rect, 12f))
            {
                Color fill = _pressed ? Theme.SecondaryPressed : (_hover ? Theme.SecondaryHover : Color.White);
                using (var b = new SolidBrush(fill))
                    g.FillPath(b, card);
                Color border = (_hover || _pressed || Focused) ? Theme.Accent : Theme.Border;
                using (var pen = new Pen(border, (_hover || Focused) ? 2f : 1f))
                    g.DrawPath(pen, card);
            }

            // Блок «значок + заголовок + описание» ставим по центру карточки: описания
            // разной длины, и при верхней привязке карточки с короткой подписью выглядели
            // пустыми снизу. Высоту меряем теми же флагами, какими рисуем.
            int titleW = Width - 2 * TitlePad, descW = Width - 2 * DescPad;
            int titleH = TextRenderer.MeasureText(g, _title, TitleFont,
                new Size(titleW, int.MaxValue), TitleFlags).Height;
            int descH = TextRenderer.MeasureText(g, _description, Font,
                new Size(descW, int.MaxValue), DescFlags).Height;
            int top = ContentTop(Height, GlyphSize + GlyphGap + titleH + TitleGap + descH);

            var glyphRect = new Rectangle((Width - GlyphSize) / 2, top, GlyphSize, GlyphSize);
            DrawGlyph(g, glyphRect);

            var titleRect = new Rectangle(TitlePad, glyphRect.Bottom + GlyphGap, titleW, titleH);
            TextRenderer.DrawText(g, _title, TitleFont, titleRect, Theme.TextPrimary, TitleFlags);

            var descRect = new Rectangle(DescPad, titleRect.Bottom + TitleGap, descW, descH);
            TextRenderer.DrawText(g, _description, Font, descRect, Theme.TextMuted, DescFlags);
        }

        /// <summary>Верх блока содержимого, центрированного по высоте карточки. Чистая — под тест.</summary>
        internal static int ContentTop(int cardHeight, int blockHeight)
        {
            int top = (cardHeight - blockHeight) / 2;
            return top < MinTop ? MinTop : top; // очень длинное описание прижимаем к верху, а не режем сверху
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
            else if (_glyph == CardGlyph.Other) { main = Theme.HubBlue; fold = Theme.HubBlueDark; }
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
