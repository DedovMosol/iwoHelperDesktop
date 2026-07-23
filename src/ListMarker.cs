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
        /// <summary>Символы-буллеты в начале абзаца (маркированный список).</summary>
        private const string Bullets = "•◦▪‣·●○*–—-";

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

            // Нумерованный: 1–3 цифры, затем «.» или «)», затем пробел, затем содержимое.
            if (char.IsDigit(text[0]))
            {
                int i = 0;
                while (i < text.Length && i < 3 && char.IsDigit(text[i]))
                    i++;
                if (i < text.Length && (text[i] == '.' || text[i] == ')') && i + 1 < text.Length && char.IsWhiteSpace(text[i + 1]))
                {
                    int c = SkipSpaces(text, i + 1);
                    if (c < text.Length)
                    {
                        int number = 0;
                        for (int k = 0; k < i; k++)
                            number = number * 10 + (text[k] - '0');
                        return new Result { Kind = ListKind.Numbered, Number = number, ContentStart = c };
                    }
                }
            }

            return none;
        }

        private static int SkipSpaces(string text, int from)
        {
            int i = from;
            while (i < text.Length && char.IsWhiteSpace(text[i]))
                i++;
            return i;
        }
    }
}
