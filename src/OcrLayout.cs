using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelMerger
{
    /// <summary>Слово с рамкой (координаты PDF, ось Y направлена вверх).</summary>
    internal class PdfWord
    {
        public string Text;
        public double Left;
        public double Right;
        public double Bottom;
        public double Top;

        public double MidY { get { return (Top + Bottom) / 2; } }
        public double Height { get { return Top - Bottom; } }
    }

    /// <summary>
    /// Порядок чтения born-digital PDF: слова с рамками → строки → абзацы
    /// (сверху вниз, слева направо). Рассчитан на одноколоночную вёрстку (типичный
    /// экспорт из Word); многоколоночную сюда не тащим — это отдельная задача
    /// (PdfPig DLA). Чистая логика без типов PdfPig — покрыта юнит-тестами.
    /// </summary>
    internal static class OcrLayout
    {
        // Слово в той же строке, если его вертикальный центр ближе половины высоты к
        // центру строки (при одинаковом кегле центры строки совпадают).
        private const double SameLineFactor = 0.5;
        // Новый абзац — если вертикальный зазор между строками заметно больше обычного.
        private const double ParagraphGapFactor = 1.6;
        // Отступ первой строки (красная строка) — сигнал нового абзаца: строка начинается
        // правее общего левого поля больше чем на IndentFactor кегля (или IndentWidthFraction ширины).
        private const double IndentFactor = 0.75;
        private const double IndentWidthFraction = 0.03;
        // Выключка по ширине (justified): если столько строк достаёт до правого поля — считаем
        // абзац законченным после «короткой» строки (не дотянувшей до поля на ShortLineFraction ширины).
        private const double JustifiedShare = 0.6;
        private const double ShortLineFraction = 0.15;
        // Красную строку воссоздаём в выводе, только если ею оформлено не меньше этой доли абзацев
        // (иначе документ без отступов — не портим его). Отступ крупнее MaxIndentFraction ширины —
        // это центрирование (номер страницы и т.п.), в расчёт красной строки не берём.
        private const double IndentedShare = 0.6;
        private const double MaxIndentFraction = 0.25;

        /// <summary>Итог разбора страницы: абзацы и измеренный отступ первой строки (0 — без красной строки).</summary>
        internal sealed class OcrPageLayout
        {
            public List<string> Paragraphs = new List<string>();
            public double FirstLineIndentPt; // pt (PDF = pt), 0 если документ без отступов
        }

        /// <summary>Слова (в любом порядке) → абзацы в порядке чтения (только текст). Чистая — под тест.</summary>
        public static List<string> ToParagraphs(IList<PdfWord> words)
        {
            return Analyze(words).Paragraphs;
        }

        /// <summary>
        /// Слова (в любом порядке) → абзацы в порядке чтения + отступ красной строки. Чистая — под тест.
        /// Граница абзаца определяется тремя независимыми сигналами (любой срабатывает):
        ///  • вертикальный зазор заметно больше обычного (абзацы разделены пустым местом);
        ///  • красная строка — текущая строка начата правее общего левого поля (отступ);
        ///  • в justified-тексте предыдущая строка «короткая» (не дотянулась до правого поля —
        ///    значит была последней в абзаце). Это ключ для Word-экспортов, где абзацы
        ///    разделены НЕ зазором, а отступом первой строки при равном межстрочном интервале.
        /// FirstLineIndentPt — медиана отступов первых строк, если красной строкой оформлено
        /// большинство абзацев; иначе 0 (документ без отступов не трогаем).
        /// </summary>
        public static OcrPageLayout Analyze(IList<PdfWord> words)
        {
            List<Line> lines = ToLines(words);
            var result = new OcrPageLayout();
            if (lines.Count == 0)
                return result;

            // Геометрия страницы: левое/правое поле, типичный кегль, признак выключки.
            double bodyLeft = double.MaxValue, bodyRight = double.MinValue;
            var heights = new List<double>(lines.Count);
            foreach (Line ln in lines)
            {
                if (ln.Left < bodyLeft) bodyLeft = ln.Left;
                if (ln.Right > bodyRight) bodyRight = ln.Right;
                heights.Add(ln.Height);
            }
            double em = Median(heights);
            if (em <= 0) em = 1;
            double width = bodyRight - bodyLeft;
            if (width <= 0) width = 1;

            int reaching = 0; // строк, достающих до правого поля (полные строки justified-текста)
            foreach (Line ln in lines)
                if (ln.Right >= bodyRight - 0.5 * em) reaching++;
            bool justified = reaching >= (int)Math.Ceiling(lines.Count * JustifiedShare);

            double gapThreshold = ParagraphThreshold(lines);
            double indentTol = Math.Max(IndentFactor * em, IndentWidthFraction * width);
            double shortTol = ShortLineFraction * width;

            // Разбить на группы строк-абзацев.
            var groups = new List<List<Line>>();
            var current = new List<Line>();
            for (int i = 0; i < lines.Count; i++)
            {
                if (i > 0)
                {
                    bool gapBreak = lines[i - 1].MidY - lines[i].MidY > gapThreshold;
                    bool indentBreak = lines[i].Left - bodyLeft > indentTol;
                    bool shortBreak = justified && lines[i - 1].Right < bodyRight - shortTol;
                    if (gapBreak || indentBreak || shortBreak)
                    {
                        groups.Add(current);
                        current = new List<Line>();
                    }
                }
                current.Add(lines[i]);
            }
            groups.Add(current);

            // Текст абзацев + сбор отступов первых строк (исключая центрирование).
            double maxIndent = MaxIndentFraction * width;
            var indents = new List<double>();
            foreach (List<Line> g in groups)
            {
                result.Paragraphs.Add(JoinLines(g));
                double ind = g[0].Left - bodyLeft;
                if (ind > indentTol && ind <= maxIndent)
                    indents.Add(ind);
            }
            if (groups.Count >= 2 && indents.Count >= (int)Math.Ceiling(groups.Count * IndentedShare))
                result.FirstLineIndentPt = Median(indents);

            return result;
        }

        /// <summary>Медиана (нижняя при чётном числе). Чистая.</summary>
        private static double Median(List<double> values)
        {
            if (values.Count == 0)
                return 0;
            var copy = new List<double>(values);
            copy.Sort();
            return copy[(copy.Count - 1) / 2];
        }

        private sealed class Line
        {
            public readonly List<PdfWord> Words = new List<PdfWord>();
            public double MidY; // центр строки — по слову-затравке (самому верхнему)

            // Геометрия строки. Валидна после сортировки слов слева направо в ToLines:
            // Words[0] — самое левое; строка всегда непуста.
            public double Left { get { return Words[0].Left; } }

            public double Right
            {
                get
                {
                    double r = double.MinValue;
                    for (int i = 0; i < Words.Count; i++)
                        if (Words[i].Right > r) r = Words[i].Right;
                    return r;
                }
            }

            public double Height // кегль строки (высота самого крупного слова)
            {
                get
                {
                    double h = 0;
                    for (int i = 0; i < Words.Count; i++)
                        if (Words[i].Height > h) h = Words[i].Height;
                    return h;
                }
            }
        }

        private static List<Line> ToLines(IList<PdfWord> words)
        {
            var result = new List<Line>();
            if (words == null || words.Count == 0)
                return result;

            var sorted = new List<PdfWord>(words);
            // Сверху вниз (MidY убывает), затем слева направо, затем по тексту (детерминизм).
            sorted.Sort(delegate(PdfWord a, PdfWord b)
            {
                int c = b.MidY.CompareTo(a.MidY);
                if (c != 0) return c;
                c = a.Left.CompareTo(b.Left);
                if (c != 0) return c;
                return string.CompareOrdinal(a.Text ?? "", b.Text ?? "");
            });

            Line line = null;
            foreach (PdfWord w in sorted)
            {
                double tol = SameLineFactor * Math.Max(w.Height, 1.0);
                if (line != null && Math.Abs(w.MidY - line.MidY) <= tol)
                {
                    line.Words.Add(w);
                }
                else
                {
                    line = new Line { MidY = w.MidY };
                    line.Words.Add(w);
                    result.Add(line);
                }
            }

            // Слова внутри строки — строго слева направо.
            foreach (Line ln in result)
                ln.Words.Sort(delegate(PdfWord a, PdfWord b)
                {
                    int c = a.Left.CompareTo(b.Left);
                    return c != 0 ? c : string.CompareOrdinal(a.Text ?? "", b.Text ?? "");
                });
            return result;
        }

        /// <summary>
        /// Порог зазора для разрыва абзаца: типичный межстрочный зазор × коэффициент.
        /// Типичный — нижняя медиана зазоров (обычные строки плотнее абзацных разрывов).
        /// </summary>
        private static double ParagraphThreshold(List<Line> lines)
        {
            if (lines.Count < 2)
                return double.MaxValue; // одна строка — один абзац
            var gaps = new List<double>(lines.Count - 1);
            for (int i = 1; i < lines.Count; i++)
                gaps.Add(lines[i - 1].MidY - lines[i].MidY);
            gaps.Sort();
            double typical = gaps[(gaps.Count - 1) / 2];
            return typical * ParagraphGapFactor;
        }

        /// <summary>Строки абзаца → сплошной текст (перенос склеивает слова, дефис-перенос снимается).</summary>
        private static string JoinLines(List<Line> lines)
        {
            var sb = new StringBuilder();
            foreach (Line ln in lines)
            {
                string text = LineText(ln);
                if (sb.Length == 0)
                    sb.Append(text);
                else if (EndsWithHyphen(sb))
                {
                    sb.Length -= 1;   // снять дефис
                    sb.Append(text);  // склеить перенос без пробела
                }
                else
                    sb.Append(' ').Append(text);
            }
            return sb.ToString().Trim();
        }

        private static string LineText(Line line)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < line.Words.Count; i++)
            {
                if (i > 0) sb.Append(' ');
                sb.Append(line.Words[i].Text ?? "");
            }
            return sb.ToString();
        }

        /// <summary>Строка кончается дефисом-переносом (дефис сразу после буквы).</summary>
        private static bool EndsWithHyphen(StringBuilder sb)
        {
            return sb.Length >= 2 && sb[sb.Length - 1] == '-' && char.IsLetter(sb[sb.Length - 2]);
        }
    }
}
