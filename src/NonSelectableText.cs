using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Текстовый блок без выделения и редактирования: рисует текст через TextRenderer,
    /// не получает фокус, не реагирует на мышь. Замена RichTextBox для read-only контента,
    /// где выделение не нужно (окно «Что нового»).
    /// </summary>
    internal sealed class NonSelectableText : Control
    {
        private string _text = string.Empty;

        public NonSelectableText()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer | ControlStyles.SupportsTransparentBackColor, true);
            SetStyle(ControlStyles.Selectable, false);
            BackColor = Color.White;
            TabStop = false;
        }

        public override string Text
        {
            get { return _text; }
            set
            {
                _text = value ?? string.Empty;
                Invalidate();
            }
        }

        /// <summary>Высота текста под заданную ширину (перенос слов, без выключки по ширине).</summary>
        public static int MeasureHeight(string text, Font font, int width)
        {
            if (string.IsNullOrEmpty(text))
                return 0;
            const int Inset = 4;
            Size measured = TextRenderer.MeasureText(text, font,
                new Size(Math.Max(width - Inset, 1), int.MaxValue), TextFormatFlags.WordBreak);
            return measured.Height + Inset;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (string.IsNullOrEmpty(_text))
                return;
            TextRenderer.DrawText(e.Graphics, _text, Font, ClientRectangle, ForeColor,
                TextFormatFlags.WordBreak | TextFormatFlags.Left | TextFormatFlags.Top);
        }
    }
}
