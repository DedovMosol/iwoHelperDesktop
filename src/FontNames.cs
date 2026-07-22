using System.Text.RegularExpressions;

namespace ExcelMerger
{
    /// <summary>
    /// Нормализация имени шрифта из PDF в семейство, понятное Word. Имена в PDF приходят с
    /// subset-префиксом («ABCDEF+»), стиль-суффиксами («,Bold», «-ItalicMT») и PostScript-
    /// хвостами («PSMT», «MT»), а имя без пробелов бывает слитным («TimesNewRoman»). Чистая
    /// логика — под тест. Стиль (полужирный/курсив) определяется отдельно по полному имени.
    /// </summary>
    public static class FontNames
    {
        private static readonly Regex Subset = new Regex(@"^[A-Z]{6}\+", RegexOptions.CultureInvariant);
        // Разбивка слитного имени: сначала отделяем ЗАГЛАВНЫЙ префикс-аббревиатуру от следующего
        // слова («PTAstra» → «PT Astra», «MSGothic» → «MS Gothic»), затем обычные границы
        // строчная→ЗАГЛАВНАЯ («AstraSerif» → «Astra Serif»).
        private static readonly Regex AcronymWord = new Regex(@"(?<=[A-Z])(?=[A-Z][a-z])", RegexOptions.CultureInvariant);
        private static readonly Regex DeCamel = new Regex(@"(?<=[a-z])(?=[A-Z])", RegexOptions.CultureInvariant);

        /// <summary>Имя шрифта из PDF → семейство для Word; null/пусто → null (писать по умолчанию).</summary>
        public static string Clean(string fontName)
        {
            if (string.IsNullOrWhiteSpace(fontName))
                return null;

            string s = Subset.Replace(fontName.Trim(), "");
            int cut = s.IndexOfAny(new[] { ',', '-' }); // стиль-суффикс отделяется запятой/дефисом
            if (cut > 0)
                s = s.Substring(0, cut);
            s = StripEnd(s, "PSMT");
            s = StripEnd(s, "MT");
            s = StripEnd(s, "PS");
            s = s.Trim();
            if (s.Length == 0)
                return null;

            s = AcronymWord.Replace(s, " "); // «PTAstraSerif» → «PT AstraSerif»
            s = DeCamel.Replace(s, " ");     // «PT AstraSerif» → «PT Astra Serif»
            return s;
        }

        private static string StripEnd(string s, string suffix)
        {
            return s.Length > suffix.Length && s.EndsWith(suffix)
                ? s.Substring(0, s.Length - suffix.Length)
                : s;
        }
    }
}
