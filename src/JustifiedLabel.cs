using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Метка с выравниванием текста «по ширине» (justify): во всех строках, кроме
    /// последней, лишний зазор распределяется между словами до полной ширины
    /// контрола; последняя строка — по левому краю. WinForms штатно justify не умеет,
    /// поэтому рисуем сами. Перенос слов вынесен в чистую <see cref="Wrap"/> (под тест).
    /// Высота подгоняется под число строк (<see cref="GetPreferredHeight"/>).
    /// </summary>
    public sealed class JustifiedLabel : Control
    {
        public JustifiedLabel()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
        }

        public override string Text
        {
            get { return base.Text; }
            set { base.Text = value; Invalidate(); }
        }

        /// <summary>Жадная разбивка на строки по ширине. Чистая — measure/spaceWidth инжектятся.</summary>
        internal static List<List<string>> Wrap(string text, int maxWidth, Func<string, int> measure, int spaceWidth)
        {
            var lines = new List<List<string>>();
            if (string.IsNullOrEmpty(text) || maxWidth <= 0)
                return lines;
            var cur = new List<string>();
            int curW = 0;
            foreach (string w in text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int ww = measure(w);
                int add = cur.Count == 0 ? ww : ww + spaceWidth;
                if (cur.Count > 0 && curW + add > maxWidth)
                {
                    lines.Add(cur);
                    cur = new List<string>();
                    curW = 0;
                    add = ww;
                }
                cur.Add(w);
                curW += add;
            }
            if (cur.Count > 0)
                lines.Add(cur);
            return lines;
        }

        private int MeasureWord(string s)
        {
            return TextRenderer.MeasureText(s, Font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding).Width;
        }

        private int SpaceWidth()
        {
            // Ширина пробела как разница «a a» и двух «a» (NoPadding).
            return Math.Max(1, MeasureWord("a a") - 2 * MeasureWord("a"));
        }

        private int LineHeight()
        {
            return TextRenderer.MeasureText("Ag", Font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding).Height;
        }

        /// <summary>Нужная высота при текущей ширине и тексте.</summary>
        public int GetPreferredHeight()
        {
            int lh = LineHeight();
            int count = Wrap(Text, Width, MeasureWord, SpaceWidth()).Count;
            return Math.Max(lh, count * lh);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            if (string.IsNullOrEmpty(Text))
                return;
            // Кэш ширин на отрисовку: каждое слово меряется один раз, а не трижды
            // (перенос, сумма строки, позиция при выводе) — GDI-measure недёшев.
            var widths = new Dictionary<string, int>();
            Func<string, int> measure = delegate(string s)
            {
                int v;
                if (!widths.TryGetValue(s, out v))
                {
                    v = MeasureWord(s);
                    widths[s] = v;
                }
                return v;
            };
            int space = SpaceWidth();
            List<List<string>> lines = Wrap(Text, Width, measure, space);
            int lh = LineHeight();
            int y = 0;
            for (int i = 0; i < lines.Count; i++)
            {
                List<string> line = lines[i];
                bool last = i == lines.Count - 1;
                int totalWords = 0;
                foreach (string w in line)
                    totalWords += measure(w);

                int x = 0;
                int gap = space;
                if (!last && line.Count > 1)
                    gap = (Width - totalWords) / (line.Count - 1); // justify: заполнить ширину

                foreach (string w in line)
                {
                    TextRenderer.DrawText(e.Graphics, w, Font, new Point(x, y), ForeColor, TextFormatFlags.NoPadding);
                    x += measure(w) + gap;
                }
                y += lh;
            }
        }
    }
}
