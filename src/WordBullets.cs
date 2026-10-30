using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Значки маркированных списков Word в PDF (RC-4). Word рисует их символьными шрифтами, и в
    /// текстовом слое они приходят знаком из области частного использования (U+F0xx — код
    /// символа шрифта Symbol или Wingdings плюс 0xF000) или буквой «o» шрифта Courier New. Без
    /// перевода первые выбрасывались как бессмысленный PUA, вторая оставалась словом «o» —
    /// и пункт терял маркер, а с ним и список. Здесь — перевод в обычные знаки Юникода, которые
    /// распознаёт <see cref="ListMarker"/>. Чистая логика — под тест.
    ///
    /// Таблица — по соответствию знаков Wingdings и Юникода (alanwood.net/demos/wingdings.html)
    /// и кодировке Adobe Symbol (0xB7 — bullet). Взяты значки библиотеки маркеров Word и
    /// близкие к ним геометрические фигуры; остальные символы шрифтов не трогаются: угадывать
    /// значок, которого нет в таблице, хуже, чем его пропустить.
    /// </summary>
    internal static class WordBullets
    {
        private static readonly Dictionary<int, char> Wingdings = new Dictionary<int, char>
        {
            { 0x6C, '\u25CF' }, // ●
            { 0x6E, '\u25A0' }, // ■
            { 0x71, '\u2751' }, // ❑
            { 0x75, '\u25C6' }, // ◆
            { 0x76, '\u2756' }, // ❖ — библиотека Word
            { 0x9F, '\u2022' }, // •
            { 0xA7, '\u25AA' }, // ▪ — библиотека Word
            { 0xD8, '\u2B9A' }, // ⮚ — библиотека Word («стрелка»)
            { 0xFC, '\u2714' }  // ✔ — библиотека Word («галочка»)
        };

        /// <summary>Знаки, в которые переводятся значки, — для <see cref="ListMarker"/>.</summary>
        internal const string Glyphs = "●■❑◆❖•▪⮚✔◦";

        /// <summary>Знак второго уровня Word — «o» шрифта Courier New.</summary>
        internal const char Hollow = '\u25E6'; // ◦

        /// <summary>
        /// Слово из одного значка символьного шрифта → знак Юникода; иначе text как есть.
        /// family — семейство шрифта слова (после <see cref="FontNames.Clean"/>). Wingdings 2 и
        /// Wingdings 3 — другие шрифты с другой раскладкой, они не переводятся.
        /// </summary>
        internal static string FromSymbolFont(string text, string family)
        {
            if (text == null || text.Length != 1 || text[0] < 0xF020 || text[0] > 0xF0FF || family == null)
                return text;
            int code = text[0] - 0xF000;
            string f = family.Replace(" ", "");
            char mapped;
            if (string.Equals(f, "Wingdings", StringComparison.OrdinalIgnoreCase) && Wingdings.TryGetValue(code, out mapped))
                return mapped.ToString();
            if (string.Equals(f, "Symbol", StringComparison.OrdinalIgnoreCase) && code == 0xB7)
                return "\u2022";
            return text;
        }

        private static bool IsCourier(string family)
        {
            return family != null && family.Replace(" ", "").StartsWith("Courier", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Пометить маркеры второго уровня Word: слово «o» шрифта Courier New в НАЧАЛЕ строки,
        /// за которым на той же строке через пробел или табуляцию идёт текст ДРУГОГО шрифта, —
        /// это значок, а не буква: в листинге кода «o» соседствует с тем же Courier, а в
        /// обычной фразе «o» не набирают Courier. Такое слово получает текст «◦» и семейство
        /// по умолчанию (у Courier New этого знака может не быть). Геометрия — рамки слов в
        /// выправленном пространстве страницы. Возвращает число помеченных слов.
        /// </summary>
        internal static int MarkCourierBullets(List<PdfWord> words)
        {
            int marked = 0;
            foreach (PdfWord w in words)
            {
                if (w.Text != "o" || !IsCourier(w.FontName))
                    continue;
                // Кегль, а не высота рамки: рамка «o» — высота строчной буквы, вдвое меньше кегля.
                double em = w.FontSizePt > 0 ? w.FontSizePt : w.Top - w.Bottom;
                if (em <= 0)
                    continue;
                double mid = (w.Top + w.Bottom) / 2;
                PdfWord next = null;
                bool lineStart = true;
                foreach (PdfWord x in words)
                {
                    if (ReferenceEquals(x, w) || Math.Abs((x.Top + x.Bottom) / 2 - mid) > 0.5 * em)
                        continue;
                    if (x.Right <= w.Left && w.Left - x.Right < 2 * em)
                        lineStart = false; // слева на той же строке вплотную есть слово
                    if (x.Left >= w.Right && (next == null || x.Left < next.Left))
                        next = x;
                }
                if (!lineStart || next == null || IsCourier(next.FontName))
                    continue;
                double gap = next.Left - w.Right;
                if (gap < 0.2 * em || gap > 4 * em)
                    continue;
                w.Text = Hollow.ToString();
                w.FontName = null;
                marked++;
            }
            return marked;
        }
    }
}
