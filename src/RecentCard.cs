using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Карточка недавнего файла на стартовом экране: значок инструмента, которым он сделан,
    /// имя файла и подпись операции. Клик открывает файл — тем же способом, что и кнопка
    /// «Открыть файл» после операции.
    ///
    /// Своя маленькая карточка, а не <see cref="ChoiceCard"/>: та рассказывает про инструмент
    /// абзацем описания и занимает четверть экрана, а здесь нужна строка, которую видно и
    /// понятно с одного взгляда. Общее у них — рисование значка (<see cref="ChoiceCard.DrawGlyph"/>),
    /// и оно взято, а не переписано.
    /// </summary>
    internal sealed class RecentCard : Control
    {
        private static readonly Color Border = Color.FromArgb(226, 226, 226);
        private static readonly Color HoverBack = Color.FromArgb(245, 248, 252);
        private static readonly Color HoverBorder = Color.FromArgb(190, 205, 225);

        private readonly string _path;
        private readonly CardGlyph _glyph;
        private readonly string _operation;
        private bool _hot;
        private ToolTip _tip;

        public RecentCard(HistoryEntry entry)
        {
            _path = entry == null ? "" : (entry.Path ?? "");
            _operation = entry == null ? "" : Loc.T(entry.Operation ?? "");
            _glyph = GlyphFor(entry == null ? null : entry.Operation, _path);
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            BackColor = Color.White;
            Cursor = Cursors.Hand;
            TabStop = true;
            AccessibleName = Path.GetFileName(_path);
            _tip = new ToolTip();
            // В подсказке — что за операция и полный путь: на карточке помещается только имя.
            _tip.SetToolTip(this, string.IsNullOrEmpty(_operation)
                ? _path : _operation + Environment.NewLine + _path);
            Click += delegate { Ui.OpenPath(_path); };
        }

        /// <summary>
        /// Значок: сначала по инструменту, которым файл сделан (это точнее — файл мог быть
        /// переименован), иначе по расширению. Чистая — под тест.
        /// </summary>
        internal static CardGlyph GlyphFor(string operationKey, string path)
        {
            string key = operationKey ?? "";
            if (key.IndexOf("split", StringComparison.OrdinalIgnoreCase) >= 0) return CardGlyph.PdfSplit;
            if (key.IndexOf("excel", StringComparison.OrdinalIgnoreCase) >= 0) return CardGlyph.Excel;
            if (key.IndexOf("pptx", StringComparison.OrdinalIgnoreCase) >= 0) return CardGlyph.Pptx;
            if (key.IndexOf("word", StringComparison.OrdinalIgnoreCase) >= 0 ||
                key.IndexOf("ocr", StringComparison.OrdinalIgnoreCase) >= 0) return CardGlyph.Ocr;
            if (key.IndexOf("ops", StringComparison.OrdinalIgnoreCase) >= 0) return CardGlyph.Tools;
            if (key.IndexOf("merge", StringComparison.OrdinalIgnoreCase) >= 0 ||
                key.IndexOf("pdf", StringComparison.OrdinalIgnoreCase) >= 0) return CardGlyph.Pdf;

            string ext = "";
            try { ext = (Path.GetExtension(path) ?? "").ToLowerInvariant(); }
            catch { }
            switch (ext)
            {
                case ".xlsx": case ".xlsm": case ".xlsb": case ".xls": return CardGlyph.Excel;
                case ".pptx": return CardGlyph.Pptx;
                case ".docx": case ".doc": return CardGlyph.Ocr;
                case ".pdf": return CardGlyph.Pdf;
                default: return CardGlyph.Other;
            }
        }

        protected override void OnMouseEnter(EventArgs e) { _hot = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hot = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnEnter(EventArgs e) { _hot = true; Invalidate(); base.OnEnter(e); }
        protected override void OnLeave(EventArgs e) { _hot = false; Invalidate(); base.OnLeave(e); }

        protected override bool IsInputKey(Keys keyData)
        {
            return keyData == Keys.Enter || keyData == Keys.Space || base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space)
            {
                Ui.OpenPath(_path);      // карточка доступна и с клавиатуры, как остальные
                e.Handled = true;
            }
            base.OnKeyDown(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(_hot ? HoverBack : BackColor);
            using (var pen = new Pen(_hot ? HoverBorder : Border))
                g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);

            // Карточка стоит в нижнем ряду и потому невысока: значок и ОДНА строка — имя файла.
            // Название операции ушло в подсказку вместе с путём: в строке высотой в ряд две
            // строки текста читались бы хуже, чем одна, а имя файла здесь главное.
            int side = Math.Min(Height - 8, 24);
            var glyphBox = new Rectangle(7, (Height - side) / 2, side, side);
            ChoiceCard.DrawGlyph(g, _glyph, glyphBox);

            int textLeft = glyphBox.Right + 8;
            int textWidth = Width - textLeft - 8;
            if (textWidth <= 0)
                return;
            Font nameFont = Ui.Font(9f);
            TextRenderer.DrawText(g, Path.GetFileName(_path), nameFont,
                new Rectangle(textLeft, 0, textWidth, Height), Theme.TextPrimary,
                TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix |
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
        }

        protected override void Dispose(bool disposing)
        {
            // ToolTip — компонент, а не дочерний контрол: сам он не освободится.
            if (disposing && _tip != null)
            {
                _tip.Dispose();
                _tip = null;
            }
            base.Dispose(disposing);
        }
    }
}
