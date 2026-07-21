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

        /// <summary>Слова (в любом порядке) → абзацы в порядке чтения. Чистая — под тест.</summary>
        public static List<string> ToParagraphs(IList<PdfWord> words)
        {
            List<Line> lines = ToLines(words);
            var paragraphs = new List<string>();
            if (lines.Count == 0)
                return paragraphs;

            double gapThreshold = ParagraphThreshold(lines);
            var current = new List<Line>();
            for (int i = 0; i < lines.Count; i++)
            {
                if (i > 0 && lines[i - 1].MidY - lines[i].MidY > gapThreshold) // зазор вниз
                {
                    paragraphs.Add(JoinLines(current));
                    current = new List<Line>();
                }
                current.Add(lines[i]);
            }
            paragraphs.Add(JoinLines(current));
            return paragraphs;
        }

        private sealed class Line
        {
            public readonly List<PdfWord> Words = new List<PdfWord>();
            public double MidY; // центр строки — по слову-затравке (самому верхнему)
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
