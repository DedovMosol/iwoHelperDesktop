using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelMerger;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace ExcelMerger.Tests
{
    /// <summary>Что и где нарисовано: одна запись на абзац, ячейку-таблицу или строку скана.</summary>
    internal sealed class CDrawn
    {
        public int Page;          // с нуля
        public CBlock Block;
        public int Lines;         // строк абзаца после переноса (для выключки)
        public bool Running;      // колонтитул или номер страницы
        public CStyle Style;      // стиль синтетического абзаца (колонтитул, строка скана)
        public string Text;       // текст синтетического абзаца
    }

    internal sealed class CRendered
    {
        public int Pages;
        public double PageW, PageH; // размер листа (pt), у всех страниц документа один
        public readonly List<CDrawn> Drawn = new List<CDrawn>();
        public readonly HashSet<string> Markers = new HashSet<string>(StringComparer.Ordinal);
    }

    /// <summary>
    /// Модель документа → настоящий PDF через PdfSharp, как его сохранил бы текстовый редактор:
    /// перенос по словам, выключка, отступы, списки с висячим отступом, таблицы с линовкой,
    /// колонки, заливки, колонтитулы. Каждое слово рисуется отдельно в своей позиции — так же
    /// раскладывает строку Word, и PdfPig собирает слова по тем же зазорам. Подчёркивание и
    /// зачёркивание — сплошной полосой через всю фразу, вместе с пробелами (как у Word).
    /// Абзац и таблица не делятся между страницами: не помещаются — уходят на следующую.
    /// </summary>
    internal sealed class CorpusRenderer : IDisposable
    {
        private readonly CDoc _doc;
        private readonly string _dir;
        private readonly PdfDocument _pdf = new PdfDocument();
        private readonly CRendered _out = new CRendered();
        private readonly Dictionary<string, XFont> _fonts = new Dictionary<string, XFont>(StringComparer.Ordinal);
        private readonly List<XImage> _images = new List<XImage>();
        private readonly HashSet<int> _scanPages = new HashSet<int>();
        private XGraphics _g;
        private int _page = -1;
        private double _y;

        private static readonly XPdfFontOptions Unicode = new XPdfFontOptions(PdfFontEncoding.Unicode);

        private CorpusRenderer(CDoc doc, string dir)
        {
            _doc = doc;
            _dir = dir;
            _pdf.Options.CompressContentStreams = false; // слой скана правится по тексту потока
        }

        /// <summary>Нарисовать документ в path. dir — куда класть временные картинки сканов.</summary>
        public static CRendered Write(CDoc doc, string path, string dir)
        {
            using (var r = new CorpusRenderer(doc, dir))
                return r.Run(path);
        }

        private CRendered Run(string path)
        {
            NewPage();
            foreach (CBlock block in _doc.Blocks)
                PlaceTopLevel(block);
            FinishPage();
            MakeScanLayersInvisible();
            _pdf.Save(path);
            _out.Pages = _pdf.PageCount;
            _out.PageW = _doc.PageW;
            _out.PageH = _doc.PageH;
            return _out;
        }

        public void Dispose()
        {
            if (_g != null)
                _g.Dispose();
            foreach (XImage image in _images)
                image.Dispose();
            _pdf.Dispose();
        }

        private double ContentLeft { get { return _doc.MarginL; } }
        private double ContentWidth { get { return _doc.PageW - _doc.MarginL - _doc.MarginR; } }
        private double ContentBottom { get { return _doc.PageH - _doc.MarginB; } }

        // ---------- страницы ----------

        private void NewPage()
        {
            FinishPage();
            PdfPage page = _pdf.AddPage();
            page.Width = _doc.PageW;
            page.Height = _doc.PageH;
            _g = XGraphics.FromPdfPage(page);
            _page++;
            _y = _doc.MarginT;
            _pageHasContent = false;
            if (_doc.Header != null)
                Running(_doc.Header, _doc.MarginT * 0.5);
        }

        private bool _pageHasContent;

        private void FinishPage()
        {
            if (_g == null)
                return;
            if (_doc.PageNumbers && !_scanPages.Contains(_page))
                Running((_page + 1).ToString(System.Globalization.CultureInfo.InvariantCulture),
                    _doc.PageH - _doc.MarginB * 0.45);
            _g.Dispose();
            _g = null;
        }

        /// <summary>Колонтитул: строка по центру на поле страницы.</summary>
        private void Running(string text, double baseline)
        {
            var style = new CStyle("Arial", 8);
            XFont font = Font(style);
            double w = _g.MeasureString(text, font).Width;
            _g.DrawString(text, font, XBrushes.Black, ContentLeft + (ContentWidth - w) / 2, baseline);
            _out.Drawn.Add(new CDrawn { Page = _page, Running = true, Lines = 1, Style = style, Text = text });
        }

        private void PlaceTopLevel(CBlock block)
        {
            if (block is CPageBreak)
            {
                NewPage();
                return;
            }
            if (block is CScan)
            {
                // Скан — страница целиком: невидимым делается весь текст её потока, поэтому
                // колонтитулов на ней быть не может.
                if (_doc.Header != null || _doc.PageNumbers)
                    throw new InvalidOperationException("скан и колонтитулы в одном документе корпуса");
                if (_pageHasContent)
                    NewPage();
                DrawScan((CScan)block);
                NewPage();
                return;
            }
            _pageHasContent = true;
            if (block is CRotated)
            {
                DrawRotated((CRotated)block);
                return;
            }
            double h = Measure(block, ContentWidth);
            if (h - block.SpaceAfter > ContentBottom - _doc.MarginT)
                throw new InvalidOperationException("блок корпуса выше страницы: " + _doc.Name);
            if (_y + h > ContentBottom && _y > _doc.MarginT + 1)
                NewPage();
            Draw(block, ContentLeft, ContentWidth);
        }

        // ---------- шрифты ----------

        private XFont Font(CStyle s)
        {
            XFontStyle fs = XFontStyle.Regular;
            if (s.Bold) fs |= XFontStyle.Bold;
            if (s.Italic) fs |= XFontStyle.Italic;
            string key = s.Font + "|" + s.Size.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + "|" + (int)fs;
            XFont font;
            if (!_fonts.TryGetValue(key, out font))
            {
                font = new XFont(s.Font, s.Size, fs, Unicode);
                _fonts[key] = font;
            }
            return font;
        }

        private double Width(string text, CStyle s)
        {
            return _g.MeasureString(text, Font(s)).Width;
        }

        private double SpaceWidth(CStyle s)
        {
            return Width("a a", s) - Width("aa", s);
        }

        private static XBrush Brush(int rgb)
        {
            return new XSolidBrush(XColor.FromArgb((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF));
        }

        // ---------- раскладка текста ----------

        private sealed class Token
        {
            public string Text;
            public CStyle Style;
            public double Width;
            public double X; // позиция на строке после выключки (от левого края строки)
        }

        private sealed class LaidLine
        {
            public readonly List<Token> Tokens = new List<Token>();
            public double Natural; // ширина без растяжки
            public double Height;
            public double Ascent;
        }

        /// <summary>Перенос ранов по словам в ширину width (у первой строки — firstWidth).</summary>
        private List<LaidLine> Wrap(List<CRun> runs, double firstWidth, double width, double spacing)
        {
            var tokens = new List<Token>();
            foreach (CRun run in runs)
                foreach (string part in run.Text.Split(' '))
                    if (part.Length > 0)
                        tokens.Add(new Token { Text = part, Style = run.Style, Width = Width(part, run.Style) });
            var lines = new List<LaidLine>();
            var line = new LaidLine();
            foreach (Token t in tokens)
            {
                double limit = lines.Count == 0 ? firstWidth : width;
                if (t.Width > limit)
                    throw new InvalidOperationException("слово «" + t.Text + "» шире строки: " + _doc.Name);
                double add = line.Tokens.Count == 0 ? t.Width : SpaceWidth(t.Style) + t.Width;
                if (line.Tokens.Count > 0 && line.Natural + add > limit)
                {
                    lines.Add(line);
                    line = new LaidLine();
                    add = t.Width;
                }
                line.Tokens.Add(t);
                line.Natural += add;
            }
            if (line.Tokens.Count > 0)
                lines.Add(line);
            foreach (LaidLine l in lines)
                foreach (Token t in l.Tokens)
                {
                    XFont f = Font(t.Style);
                    l.Height = Math.Max(l.Height, f.GetHeight() * spacing);
                    l.Ascent = Math.Max(l.Ascent, t.Style.Size * 0.9);
                }
            return lines;
        }

        /// <summary>Позиции слов строки по выключке. avail — ширина строки.</summary>
        private void Justify(LaidLine line, double avail, CAlign align, bool last)
        {
            double extra = Math.Max(0, avail - line.Natural);
            double start = 0, gapExtra = 0;
            if (align == CAlign.Right)
                start = extra;
            else if (align == CAlign.Center)
                start = extra / 2;
            else if (align == CAlign.Justify && !last && line.Tokens.Count > 1)
                gapExtra = extra / (line.Tokens.Count - 1);
            else if (align == CAlign.Distribute && line.Tokens.Count > 1)
            {
                gapExtra = extra / (line.Tokens.Count - 1);
                // Последняя строка из пары слов растянулась бы дырами шире колонки — такую
                // Word рисует, но это уже не «полная строка», а россыпь слов; оставляем влево.
                if (last && gapExtra > 2 * SpaceWidth(line.Tokens[1].Style))
                    gapExtra = 0;
            }
            double x = start;
            for (int i = 0; i < line.Tokens.Count; i++)
            {
                Token t = line.Tokens[i];
                if (i > 0)
                    x += SpaceWidth(t.Style) + gapExtra;
                t.X = x;
                x += t.Width;
            }
        }

        /// <summary>
        /// Висячий отступ пункта: не меньше ширины маркера и полутора пробелов. Word ставит
        /// текст по позиции табуляции, и маркер никогда не слипается со словом; узкий отступ
        /// с широким маркером («1.1.» кеглем 14) иначе склеил бы их в одно слово.
        /// </summary>
        private double Hang(CParagraph p)
        {
            if (p.List == ListKind.None)
                return 0;
            CStyle ms = p.MarkerStyle ?? p.Runs[0].Style;
            double need = p.Marker == null ? 0 : Width(p.Marker, ms) + 1.5 * SpaceWidth(p.Runs[0].Style);
            return Math.Max(p.Hang, need);
        }

        private List<LaidLine> LayoutParagraph(CParagraph p, double width)
        {
            double textLeft = p.LeftIndent + Hang(p);
            double first = width - textLeft - (p.List == ListKind.None ? p.FirstIndent : 0);
            return Wrap(p.Runs, first, width - textLeft, p.LineSpacing);
        }

        private static double LinesHeight(List<LaidLine> lines)
        {
            double h = 0;
            foreach (LaidLine l in lines)
                h += l.Height;
            return h;
        }

        // ---------- блоки ----------

        private double Measure(CBlock block, double width)
        {
            var p = block as CParagraph;
            if (p != null)
                return LinesHeight(LayoutParagraph(p, width)) + p.SpaceAfter;
            var t = block as CTable;
            if (t != null)
            {
                double h = 0;
                foreach (double rh in RowHeights(t, Scale(t, width)))
                    h += rh;
                return h + t.SpaceAfter;
            }
            var c = block as CColumns;
            if (c != null)
            {
                double colW = (width - c.Gap * (c.Columns.Count - 1)) / c.Columns.Count;
                double max = 0;
                foreach (List<CBlock> col in c.Columns)
                {
                    double h = 0;
                    foreach (CBlock b in col)
                        h += Measure(b, colW);
                    max = Math.Max(max, h);
                }
                return max + c.SpaceAfter;
            }
            return 0;
        }

        private void Draw(CBlock block, double left, double width)
        {
            var p = block as CParagraph;
            if (p != null) { DrawParagraph(p, left, width); return; }
            var t = block as CTable;
            if (t != null) { DrawTable(t, left, width); return; }
            var c = block as CColumns;
            if (c != null) { DrawColumns(c, left, width); return; }
            throw new InvalidOperationException("блок не рисуется в потоке: " + block.GetType().Name);
        }

        private void DrawParagraph(CParagraph p, double left, double width)
        {
            List<LaidLine> lines = LayoutParagraph(p, width);
            double textLeft = left + p.LeftIndent + Hang(p);
            if (p.Fill >= 0)
            {
                double h = LinesHeight(lines);
                _g.DrawRectangle(Brush(p.Fill), left + p.LeftIndent - 4, _y - 2,
                    width - p.LeftIndent + 8, h + 4);
            }
            if (p.List != ListKind.None && p.Marker != null && lines.Count > 0)
            {
                CStyle ms = p.MarkerStyle ?? p.Runs[0].Style;
                _g.DrawString(p.Marker, Font(ms), Brush(ms.Color), left + p.LeftIndent, _y + lines[0].Ascent);
                _out.Markers.Add(p.Marker);
            }
            for (int i = 0; i < lines.Count; i++)
            {
                LaidLine line = lines[i];
                double x0 = textLeft + (i == 0 && p.List == ListKind.None ? p.FirstIndent : 0);
                double avail = left + width - x0;
                Justify(line, avail, p.Align, i == lines.Count - 1);
                double baseline = _y + line.Ascent;
                DrawLine(line, x0, baseline);
                _y += line.Height;
            }
            _y += p.SpaceAfter;
            _out.Drawn.Add(new CDrawn { Page = _page, Block = p, Lines = lines.Count });
        }

        /// <summary>Слова строки и полосы подчёркивания/зачёркивания через фразу.</summary>
        private void DrawLine(LaidLine line, double x0, double baseline)
        {
            foreach (Token t in line.Tokens)
                _g.DrawString(t.Text, Font(t.Style), Brush(t.Style.Color), x0 + t.X, baseline);
            DecorationBars(line, x0, baseline, true);
            DecorationBars(line, x0, baseline, false);
        }

        private void DecorationBars(LaidLine line, double x0, double baseline, bool underline)
        {
            int i = 0;
            while (i < line.Tokens.Count)
            {
                Token t = line.Tokens[i];
                bool on = underline ? t.Style.Underline : t.Style.Strike;
                if (!on) { i++; continue; }
                int j = i;
                while (j + 1 < line.Tokens.Count &&
                       (underline ? line.Tokens[j + 1].Style.Underline : line.Tokens[j + 1].Style.Strike))
                    j++;
                Token a = line.Tokens[i], b = line.Tokens[j];
                double size = a.Style.Size;
                double thick = Math.Max(0.5, size * 0.05);
                double y = underline ? baseline + size * 0.11 : baseline - size * 0.27;
                _g.DrawRectangle(Brush(a.Style.Color), x0 + a.X, y, b.X + b.Width - a.X, thick);
                i = j + 1;
            }
        }

        // ---------- таблицы ----------

        private const double CellPad = 4;

        /// <summary>
        /// Ширины колонок: каждая не уже самого длинного слова своих ячеек (слово, вылезшее
        /// в соседнюю ячейку, слиплось бы с её словом), остаток — пропорционально заданным
        /// ширинам. Таблица, которой не хватает листа, — ошибка генератора, а не случай.
        /// </summary>
        private double[] Scale(CTable t, double width)
        {
            int n = t.ColWidths.Length;
            var need = new double[n];
            foreach (CCell[] row in t.Rows)
                for (int c = 0; c < n; c++)
                    if (row[c] != null && row[c].ColSpan == 1)
                        foreach (CRun run in row[c].Runs)
                            foreach (string word in run.Text.Split(' '))
                                if (word.Length > 0)
                                    need[c] = Math.Max(need[c], Width(word, run.Style) + 2 * CellPad + 2);
            double sumNeed = 0, sumWant = 0;
            for (int c = 0; c < n; c++)
            {
                sumNeed += need[c];
                sumWant += t.ColWidths[c];
            }
            if (sumNeed > width)
                throw new InvalidOperationException("таблица корпуса шире листа: " + _doc.Name);
            double total = Math.Min(width, Math.Max(sumWant, sumNeed));
            var result = new double[n];
            for (int c = 0; c < n; c++)
                result[c] = need[c] + (total - sumNeed) * t.ColWidths[c] / sumWant;
            return result;
        }

        private static double SpanWidth(double[] cols, int from, int span)
        {
            double w = 0;
            for (int i = from; i < from + span && i < cols.Length; i++)
                w += cols[i];
            return w;
        }

        private double[] RowHeights(CTable t, double[] cols)
        {
            var heights = new double[t.Rows.Count];
            for (int r = 0; r < t.Rows.Count; r++)
            {
                double h = 0;
                for (int c = 0; c < t.Rows[r].Length; c++)
                {
                    CCell cell = t.Rows[r][c];
                    if (cell == null || cell.RowSpan > 1)
                        continue;
                    double inner = SpanWidth(cols, c, cell.ColSpan) - 2 * CellPad;
                    h = Math.Max(h, LinesHeight(Wrap(cell.Runs, inner, inner, 1.0)) + 2 * CellPad);
                }
                heights[r] = Math.Max(h, 14);
            }
            return heights;
        }

        private void DrawTable(CTable t, double left, double width)
        {
            double[] cols = Scale(t, width);
            double[] heights = RowHeights(t, cols);
            double top = _y;
            var pen = new XPen(XColors.Black, t.LineWidth);
            double rowTop = top;
            for (int r = 0; r < t.Rows.Count; r++)
            {
                double x = left;
                for (int c = 0; c < t.Rows[r].Length; c++)
                {
                    CCell cell = t.Rows[r][c];
                    if (cell == null)
                    {
                        x += cols[c];
                        continue;
                    }
                    double w = SpanWidth(cols, c, cell.ColSpan);
                    double h = 0;
                    for (int k = r; k < r + cell.RowSpan && k < heights.Length; k++)
                        h += heights[k];
                    if (r == 0 && t.HeaderFill >= 0)
                        _g.DrawRectangle(Brush(t.HeaderFill), x, rowTop, w, h);
                    if (t.Ruled)
                        _g.DrawRectangle(pen, x, rowTop, w, h);
                    double inner = w - 2 * CellPad;
                    double y = rowTop + CellPad;
                    foreach (LaidLine line in Wrap(cell.Runs, inner, inner, 1.0))
                    {
                        Justify(line, inner, CAlign.Left, true);
                        DrawLine(line, x + CellPad, y + line.Ascent);
                        y += line.Height;
                    }
                    x += cols[c];
                }
                rowTop += heights[r];
            }
            _y = rowTop + t.SpaceAfter;
            _out.Drawn.Add(new CDrawn { Page = _page, Block = t });
        }

        // ---------- колонки ----------

        private void DrawColumns(CColumns c, double left, double width)
        {
            double colW = (width - c.Gap * (c.Columns.Count - 1)) / c.Columns.Count;
            double top = _y, bottom = _y;
            for (int i = 0; i < c.Columns.Count; i++)
            {
                _y = top;
                foreach (CBlock b in c.Columns[i])
                    Draw(b, left + i * (colW + c.Gap), colW);
                bottom = Math.Max(bottom, _y);
            }
            _y = bottom + c.SpaceAfter;
        }

        // ---------- повёрнутая надпись ----------

        private void DrawRotated(CRotated r)
        {
            XGraphicsState state = _g.Save();
            _g.RotateAtTransform(r.Angle, new XPoint(r.X, r.Y));
            _g.DrawString(r.Text, Font(r.Style), Brush(r.Style.Color), r.X, r.Y);
            _g.Restore(state);
            _out.Drawn.Add(new CDrawn { Page = _page, Block = r });
        }

        // ---------- скан с текстовым слоем ----------

        private void DrawScan(CScan scan)
        {
            const double dpi = 150;
            double w = ContentWidth, lineH = scan.FontSize * 2.4;
            double h = lineH * scan.Lines.Count + scan.FontSize;
            string png = Path.Combine(_dir, "scan_" + _page + "_" + Guid.NewGuid().ToString("N") + ".png");
            using (var bmp = new System.Drawing.Bitmap((int)(w / 72 * dpi), (int)(h / 72 * dpi)))
            {
                bmp.SetResolution((float)dpi, (float)dpi);
                using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bmp))
                using (var f = new System.Drawing.Font("Arial", (float)scan.FontSize, System.Drawing.GraphicsUnit.Point))
                {
                    g.Clear(System.Drawing.Color.FromArgb(250, 248, 240));
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                    for (int i = 0; i < scan.Lines.Count; i++)
                        g.DrawString(scan.Lines[i], f, System.Drawing.Brushes.Black,
                            0, (float)((i * lineH) / 72 * dpi));
                }
                bmp.Save(png, System.Drawing.Imaging.ImageFormat.Png);
            }
            XImage image = XImage.FromFile(png);
            _images.Add(image);
            _g.DrawImage(image, ContentLeft, _y, w, h);
            var style = new CStyle("Arial", scan.FontSize);
            XFont font = Font(style);
            for (int i = 0; i < scan.Lines.Count; i++)
            {
                _g.DrawString(scan.Lines[i], font, XBrushes.Black, ContentLeft, _y + i * lineH + scan.FontSize);
                _out.Drawn.Add(new CDrawn { Page = _page, Block = scan, Lines = 1, Style = style, Text = scan.Lines[i] });
            }
            _scanPages.Add(_page);
            _y += h;
        }

        /// <summary>
        /// Текст страниц-сканов — в режим 3 («не рисовать», ISO 32000-1, 9.3.6): так пишут
        /// текстовый слой распознавания поверх картинки. У PdfSharp нет такого режима, поэтому
        /// он ставится в несжатый поток страницы после рисования.
        /// </summary>
        private void MakeScanLayersInvisible()
        {
            foreach (int index in _scanPages)
            {
                PdfPage page = _pdf.Pages[index];
                for (int i = 0; i < page.Contents.Elements.Count; i++)
                {
                    var content = page.Contents.Elements.GetReference(i).Value as PdfDictionary;
                    if (content == null || content.Stream == null)
                        continue;
                    Encoding latin = Encoding.GetEncoding(28591); // байты как есть
                    string raw = latin.GetString(content.Stream.Value);
                    content.Stream.Value = latin.GetBytes(raw.Replace("BT\n", "BT\n3 Tr\n"));
                }
            }
        }
    }
}
