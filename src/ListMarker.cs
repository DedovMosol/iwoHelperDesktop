namespace ExcelMerger
{
    /// <summary>Вид маркера списка в начале абзаца.</summary>
    public enum ListKind { None, Bulleted, Numbered }

    /// <summary>
    /// Распознавание маркера списка в начале абзаца born-digital PDF: «1.»/«2)» — нумерованный,
    /// «•»/«—»/«–»/«-» и т.п. — маркированный. Возвращает вид, номер (для нумерованного) и индекс,
    /// с которого начинается содержимое (маркер при записи в Word снимается — Word рисует свой).
    /// Строгие условия (пунктуация + пробел + непустое содержимое) отсекают ложные срабатывания
    /// вроде «2025 г.» или «12.5%». Чистая логика — под тест.
    /// </summary>
    internal static class ListMarker
    {
        /// <summary>
        /// Символы-буллеты в начале абзаца (маркированный список): обычные и те, в которые
        /// <see cref="WordBullets"/> переводит значки Word из Symbol, Wingdings и Courier New.
        /// </summary>
        private const string Bullets = "•◦▪‣·●○*–—-" + WordBullets.Glyphs;

        /// <summary>
        /// Буллет-тире («–», «—», «-»). Отдельно от прочих: тире — ещё и знак препинания, и
        /// строка, начатая им, бывает переносом фразы, а не пунктом (см. OcrLayout).
        /// </summary>
        internal static bool IsDash(char c)
        {
            return c == '–' || c == '—' || c == '-';
        }

        internal struct Result
        {
            public ListKind Kind;
            public int Number;        // номер для Numbered; 0 иначе
            public int ContentStart;  // индекс начала текста после маркера и пробелов
        }

        /// <summary>Распознать маркер списка в начале text. Kind=None — не список.</summary>
        public static Result Detect(string text)
        {
            var none = new Result { Kind = ListKind.None, Number = 0, ContentStart = 0 };
            if (string.IsNullOrEmpty(text) || char.IsWhiteSpace(text[0]))
                return none; // абзацы уже обрезаны слева; ведущий пробел — не наш случай

            // Маркированный: одиночный буллет-символ, затем пробел, затем содержимое.
            if (Bullets.IndexOf(text[0]) >= 0)
            {
                int c = SkipSpaces(text, 1);
                if (c > 1 && c < text.Length) // после буллета был пробел и есть содержимое
                    return new Result { Kind = ListKind.Bulleted, Number = 0, ContentStart = c };
                return none;
            }

            Result numbered = TryNumbered(text);
            if (numbered.Kind != ListKind.None)
                return numbered;
            return none;
        }

        /// <summary>
        /// Маркер списка, в том числе голый токен «•» / «1.» без текста пункта (разрыв строки
        /// на слайде, левая колонка сетки).
        /// </summary>
        internal static bool LooksLikeMarker(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;
            if (Detect(text).Kind != ListKind.None)
                return true;
            return Detect(text.TrimEnd() + " x").Kind != ListKind.None;
        }

        /// <summary>
        /// Цифра номера списка — ТОЛЬКО ASCII 0–9. char.IsDigit пропускает и другие десятичные
        /// цифры Юникода (полноширинные «１», арабско-индийские «١»), а «text[k] - '0'» превращал
        /// бы их в мусор порядка 65 000: Word такой начальный номер списка отвергает, исключение
        /// глушится, и пункт молча оставался бы вообще без номера.
        /// </summary>
        private static Result TryNumbered(string text)
        {
            var none = new Result { Kind = ListKind.None, Number = 0, ContentStart = 0 };
            // «(1) пункт», «(12) пункт» — скобочная нумерация ГОСТ/инструкций (RC-3).
            if (text[0] == '(')
            {
                int i = 1, n = 0;
                while (i < text.Length && i < 4 && IsAsciiDigit(text[i]))
                {
                    n = n * 10 + (text[i] - '0');
                    i++;
                }
                if (n > 0 && i < text.Length && text[i] == ')')
                {
                    int c = ContentAfter(text, i + 1);
                    if (c >= 0)
                        return new Result { Kind = ListKind.Numbered, Number = n, ContentStart = c };
                }
                return none;
            }
            if (IsAsciiDigit(text[0]))
            {
                int i = 0, segs = 0, first = 0, last = 0;
                while (i < text.Length && segs < 4 && IsAsciiDigit(text[i]))
                {
                    int n = 0, start = i;
                    while (i < text.Length && i - start < 3 && IsAsciiDigit(text[i]))
                    {
                        n = n * 10 + (text[i] - '0');
                        i++;
                    }
                    if (segs == 0) first = n;
                    last = n;
                    segs++;
                    if (i < text.Length && text[i] == '.') i++;
                    else break;
                }
                if (segs == 1)
                {
                    // «1.» / «12)» — последняя точка уже съедена, либо «)».
                    int p = first <= 9 ? 1 : (first <= 99 ? 2 : 3);
                    if (p < text.Length && (text[p] == ')' || text[p] == ']') && p + 1 < text.Length && char.IsWhiteSpace(text[p + 1]))
                    {
                        int c = SkipSpaces(text, p + 1);
                        if (c < text.Length)
                            return new Result { Kind = ListKind.Numbered, Number = first, ContentStart = c };
                    }
                }
                if (segs >= 1)
                {
                    int c = SkipSpaces(text, i);
                    if (c > i && c < text.Length && (segs >= 2 || (i > 0 && (text[i - 1] == '.' || text[i - 1] == ')'))))
                        return new Result { Kind = ListKind.Numbered, Number = segs >= 2 ? last : first, ContentStart = c };
                }
                return none;
            }
            int r = 0;
            while (r < text.Length && r < 7 && "IVXLCDMivxlcdm".IndexOf(text[r]) >= 0)
                r++;
            if (r > 0 && r < text.Length && (text[r] == '.' || text[r] == ')') && r + 1 < text.Length && ContentAfter(text, r + 1) >= 0)
            {
                char u = text[0] >= 'a' ? (char)(text[0] - 32) : text[0];
                bool singleOk = r == 1 && (u == 'I' || u == 'V' || u == 'X');
                if (r >= 2 || singleOk)
                {
                    int c = ContentAfter(text, r + 1);
                    int val = RomanValue(text, r);
                    if (c >= 0 && val > 0)
                        return new Result { Kind = ListKind.Numbered, Number = val, ContentStart = c };
                }
            }
            if (text.Length >= 3 && (text[1] == '.' || text[1] == ')'))
            {
                int n = LetterIndex(text[0]);
                int c = ContentAfter(text, 2);
                if (n > 0 && c >= 0)
                    return new Result { Kind = ListKind.Numbered, Number = n, ContentStart = c };
            }
            return none;
        }

        private static int LetterIndex(char ch)
        {
            if (ch >= 'A' && ch <= 'Z') return ch - 'A' + 1;
            if (ch >= 'a' && ch <= 'z') return ch - 'a' + 1;
            if (ch >= 'А' && ch <= 'Я') return ch - 'А' + 1;
            if (ch >= 'а' && ch <= 'я') return ch - 'а' + 1;
            return 0;
        }

        private static int RomanValue(string text, int len)
        {
            int total = 0, prev = 0;
            for (int i = len - 1; i >= 0; i--)
            {
                char ch = text[i];
                if (ch >= 'a') ch = (char)(ch - 32);
                int v = ch == 'I' ? 1 : ch == 'V' ? 5 : ch == 'X' ? 10 : ch == 'L' ? 50 : ch == 'C' ? 100 : ch == 'D' ? 500 : ch == 'M' ? 1000 : 0;
                if (v == 0) return 0;
                total += v < prev ? -v : v;
                prev = v;
            }
            return total > 0 && total < 4000 ? total : 0;
        }

        private static bool IsAsciiDigit(char c)
        {
            return c >= '0' && c <= '9';
        }

        private static int SkipSpaces(string text, int from)
        {
            int i = from;
            while (i < text.Length && char.IsWhiteSpace(text[i]))
                i++;
            return i;
        }

        /// <summary>Начало содержимого после маркера: пробел(ы) или сразу буква («III.Пункт»).</summary>
        private static int ContentAfter(string text, int from)
        {
            if (from >= text.Length)
                return -1;
            if (char.IsWhiteSpace(text[from]))
            {
                int c = SkipSpaces(text, from);
                return c < text.Length ? c : -1;
            }
            return LetterIndex(text[from]) > 0 ? from : -1;
        }
    }
}
