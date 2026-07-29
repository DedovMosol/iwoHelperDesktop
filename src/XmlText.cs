using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Подготовка текста из PDF к укладке в XML. Два шага и строго в этом порядке:
    /// <see cref="Sanitize"/> убирает то, чего в XML 1.0 быть НЕ МОЖЕТ, затем
    /// <see cref="Escape"/> экранирует то, что в XML имеет особый смысл. Обратный порядок
    /// испортил бы уже подставленные сущности.
    ///
    /// Почему это отдельный слой, а не «строка с заменой &amp;»: born-digital PDF регулярно
    /// приносит мусор из битой таблицы ToUnicode — управляющие символы и ОДИНОЧНЫЕ суррогаты.
    /// Одиночный суррогат не является допустимым символом XML, и файл с ним читатель объявит
    /// повреждённым целиком — то есть один битый глиф уносит всю презентацию. Чистая логика,
    /// покрыта юнит-тестами.
    /// </summary>
    internal static class XmlText
    {
        /// <summary>
        /// Текст, готовый к вставке в узел или значение атрибута: очистка + экранирование.
        /// null/пусто — пустая строка (узел без текста корректен, отсутствие узла — нет).
        /// </summary>
        public static string Encode(string s)
        {
            return Escape(Sanitize(s));
        }

        /// <summary>
        /// Убрать символы, недопустимые в XML 1.0: управляющие (кроме табуляции), непарные
        /// суррогаты, U+FFFE/U+FFFF. Перевод строки и возврат каретки заменяются ПРОБЕЛОМ:
        /// внутри текстового узла они были бы видимым мусором, а настоящий перенос строки —
        /// это отдельный элемент разметки, а не символ. Корректные суррогатные пары
        /// (эмодзи, редкие иероглифы) сохраняются целиком. Чистая — под тест.
        /// </summary>
        public static string Sanitize(string s)
        {
            if (string.IsNullOrEmpty(s))
                return string.Empty;
            var sb = new StringBuilder(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '\t')
                {
                    sb.Append(c);
                    continue;
                }
                if (c == '\n' || c == '\r' || c == '\v' || c == '\f')
                {
                    sb.Append(' ');
                    continue;
                }
                // U+FFFE и U+FFFF — «не символы»: записаны escape-последовательностями,
                // потому что литералами они невидимы в исходнике.
                if (c < 0x20 || c == '\uFFFE' || c == '\uFFFF')
                    continue;
                if (char.IsHighSurrogate(c))
                {
                    // Пара имеет смысл только целиком: одиночная половина ломает документ.
                    if (i + 1 < s.Length && char.IsLowSurrogate(s[i + 1]))
                    {
                        sb.Append(c);
                        sb.Append(s[i + 1]);
                        i++;
                    }
                    continue;
                }
                if (char.IsLowSurrogate(c))
                    continue; // хвост без головы
                sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Экранировать символы с особым смыслом. Кавычка тоже: этой же функцией готовятся
        /// значения атрибутов, а они пишутся в двойных кавычках. Чистая — под тест.
        /// </summary>
        public static string Escape(string s)
        {
            if (string.IsNullOrEmpty(s))
                return string.Empty;
            var sb = new StringBuilder(s.Length + 16);
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                switch (c)
                {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    case '"': sb.Append("&quot;"); break;
                    default: sb.Append(c); break;
                }
            }
            return sb.ToString();
        }
    }
}
