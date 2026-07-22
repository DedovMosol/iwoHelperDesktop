using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Пометка подчёркнутых слов: в born-digital PDF подчёркивание — это НАРИСОВАННАЯ
    /// горизонтальная линия под базовой линией слова (в тексте его нет). Слово считается
    /// подчёркнутым, если под ним (у самой базовой линии) проходит горизонтальная линовка,
    /// перекрывающая большую часть ширины слова. Чистая логика — под тест. Вызывать по
    /// ВНЕтабличным словам и линиям: границы ячеек проходят по краям ячейки, а не под словом,
    /// и порогом близости к базовой линии не считаются подчёркиванием.
    /// </summary>
    internal static class UnderlineDetector
    {
        private const double DropBelow = 4.0;   // насколько ниже низа слова может лежать линия
        private const double RiseAbove = 1.5;   // и насколько выше (линия у самой базовой линии)
        private const double CoverFraction = 0.6; // линия должна перекрывать эту долю ширины слова

        /// <summary>Проставить <see cref="PdfWord.Underline"/> словам, под которыми есть линовка.</summary>
        public static void Mark(IList<PdfWord> words, IList<PdfLine> lines)
        {
            if (words == null || lines == null || lines.Count == 0)
                return;
            var horizontals = new List<PdfLine>();
            foreach (PdfLine l in lines)
                if (l.Orientation == LineOrientation.Horizontal)
                    horizontals.Add(l);
            if (horizontals.Count == 0)
                return;

            foreach (PdfWord w in words)
            {
                double width = w.Right - w.Left;
                if (width <= 0)
                    continue;
                foreach (PdfLine line in horizontals)
                {
                    // Линия у базовой линии слова: не выше низа на RiseAbove и не ниже на DropBelow.
                    if (line.Position > w.Bottom + RiseAbove || line.Position < w.Bottom - DropBelow)
                        continue;
                    double overlap = Math.Min(line.MaxX, w.Right) - Math.Max(line.MinX, w.Left);
                    if (overlap >= CoverFraction * width)
                    {
                        w.Underline = true;
                        break;
                    }
                }
            }
        }
    }
}
