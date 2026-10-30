using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Подчёркивание и зачёркивание слов. В born-digital PDF это НАРИСОВАННЫЕ горизонтальные
    /// линии (в тексте их нет). Word, LibreOffice и PdfSharp рисуют ОДНУ полосу через всю
    /// фразу, поэтому линия проверяется против слов своей строки целиком (RC-5, RC-6):
    ///  • слово «на линии» — линия в полосе своего вида относительно базовой линии слова и
    ///    накрывает не меньше 0,6 его ширины. Подчёркивание: ниже базовой линии не больше чем на
    ///    0,35 кегля (но не меньше 4 pt), выше — не больше чем на 0,12 кегля; шрифты кладут его на
    ///    0,05–0,25 кегля ниже. Зачёркивание: от 0,15 до 0,45 кегля выше (по OS/2 шрифтов ≈ 0,26–0,3,
    ///    до высоты строчных ≈ 0,5). Полосы не пересекаются — линия не бывает обоими сразу;
    ///  • если линия прилегает к своим словам (выступает за крайние не больше чем на полкегля с
    ///    каждой стороны — это захваченный пробел), помечены все эти слова;
    ///  • иначе — правило по слову: линия не длиннее трёх его ширин (поле подписи под словом).
    ///    Разделитель во всю колонку под заголовком не проходит ни то, ни другое.
    /// Вызывать по ВНЕтабличным словам и ОДИНОЧНЫМ линиям (<see cref="TableDetectResult.LoneLines"/>):
    /// граница ячейки — это рамка, даже вплотную под абзацем (RC-33). Чистая логика — под тест.
    /// </summary>
    internal static class TextDecorations
    {
        private const double UnderlineDropMinPt = 4.0; // нижний допуск не меньше этого (мелкий кегль, толстая полоса)
        private const double UnderlineDropEm = 0.35;   // насколько ниже базовой линии может лежать подчёркивание
        private const double UnderlineRiseEm = 0.12;   // и насколько выше
        private const double StrikeLowEm = 0.15;       // зачёркивание — не ниже этого над базовой линией
        private const double StrikeHighEm = 0.45;      // и не выше
        private const double CoverFraction = 0.6;      // линия должна перекрывать эту долю ширины слова
        private const double OverhangEm = 0.5;         // выступ полосы за фразу: захваченный пробел, не больше
        private const double MaxWidthFactor = 3.0;     // линия длиннее слова во столько раз — это разделитель
        private const double DefaultSizePt = 10.0;     // кегль слова неизвестен

        private enum Kind { Underline, Strike }

        /// <summary>
        /// Проставить <see cref="PdfWord.Underline"/> и <see cref="PdfWord.Strike"/> словам по
        /// линиям. Возвращает линии, ставшие подчёркиванием или зачёркиванием, — прочерками
        /// полей (<see cref="PdfTextExtract.AddRuleWords"/>) они уже не станут.
        /// </summary>
        public static HashSet<PdfLine> Mark(IList<PdfWord> words, IList<PdfLine> lines)
        {
            var used = new HashSet<PdfLine>();
            if (words == null || lines == null || words.Count == 0)
                return used;
            var on = new List<PdfWord>();
            foreach (PdfLine line in lines)
                if (line.Orientation == LineOrientation.Horizontal
                    && (MarkLine(words, line, Kind.Underline, on) || MarkLine(words, line, Kind.Strike, on)))
                    used.Add(line);
            return used;
        }

        /// <summary>Пометить слова, которым линия служит украшением вида kind; помечено ли хоть одно.</summary>
        private static bool MarkLine(IList<PdfWord> words, PdfLine line, Kind kind, List<PdfWord> on)
        {
            on.Clear();
            double spanLeft = double.MaxValue, spanRight = double.MinValue, em = 0;
            foreach (PdfWord w in words)
                if (OnLine(w, line, kind))
                {
                    on.Add(w);
                    spanLeft = Math.Min(spanLeft, w.Left);
                    spanRight = Math.Max(spanRight, w.Right);
                    em = Math.Max(em, SizeOf(w));
                }
            bool phrase = line.MinX >= spanLeft - OverhangEm * em && line.MaxX <= spanRight + OverhangEm * em;
            bool any = false;
            foreach (PdfWord w in on)
                if (phrase || line.Length <= MaxWidthFactor * (w.Right - w.Left))
                {
                    if (kind == Kind.Underline) w.Underline = true;
                    else w.Strike = true;
                    any = true;
                }
            return any;
        }

        /// <summary>
        /// Лежит ли линия в полосе вида kind относительно базовой линии слова и накрывает ли
        /// бо́льшую часть его ширины. Без базовой линии (0 — неизвестна) мерить приходится от низа рамки.
        /// </summary>
        private static bool OnLine(PdfWord w, PdfLine line, Kind kind)
        {
            double width = w.Right - w.Left;
            if (width <= 0)
                return false;
            double size = SizeOf(w);
            double rise = line.Position - (w.BaselineYPt != 0 ? w.BaselineYPt : w.Bottom);
            bool inBand = kind == Kind.Underline
                ? rise <= UnderlineRiseEm * size && rise >= -Math.Max(UnderlineDropMinPt, UnderlineDropEm * size)
                : rise >= StrikeLowEm * size && rise <= StrikeHighEm * size;
            if (!inBand)
                return false;
            double overlap = Math.Min(line.MaxX, w.Right) - Math.Max(line.MinX, w.Left);
            return overlap >= CoverFraction * width;
        }

        private static double SizeOf(PdfWord w)
        {
            return w.FontSizePt > 0 ? w.FontSizePt : DefaultSizePt;
        }
    }
}
