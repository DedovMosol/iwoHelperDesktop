using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Значения токенов для одного выходного файла. Заполняет вызывающий сервис,
    /// неизвестные ему величины оставляет как есть (шаблон подставит пустую строку).
    /// </summary>
    public sealed class NameValues
    {
        public string BaseName = "";   // имя исходного файла без расширения
        public int FileNumber;         // номер выходного файла, с 1
        public int TotalFiles;         // всего выходных файлов
        public int CurrentPage;        // номер первой страницы этого файла в исходнике, с 1
        public string BookmarkName;    // название закладки (только разделение по закладкам)
        public DateTime Timestamp = DateTime.MinValue; // время запуска операции
    }

    /// <summary>
    /// Шаблон имени выходного файла: «[BASENAME]_часть[FILENUMBER##]» и т. п.
    ///
    /// Раньше имя было жёстким («основа_1», «основа_2»), и подогнать его под чужой
    /// регламент именования было нельзя. Синтаксис взят у распространённых PDF-утилит,
    /// чтобы не изобретать свой: имя токена в квадратных скобках, необязательные решётки
    /// задают дополнение нулями, необязательное число — СМЕЩЕНИЕ нумерации.
    ///
    ///   [FILENUMBER]      → 1, 2, 3
    ///   [FILENUMBER###]   → 001, 002, 003
    ///   [FILENUMBER10]    → 11, 12, 13          (смещение, а не начало)
    ///   [FILENUMBER###10] → 011, 012, 013
    ///   [BASENAME]        → имя исходного файла
    ///   [CURRENTPAGE]     → номер страницы в исходнике
    ///   [TOTAL_FILES]     → сколько файлов получится
    ///   [BOOKMARK]        → название закладки (иначе пусто)
    ///   [TIMESTAMP]       → 2026-07-26_21-45-03
    ///
    /// Всё, что не является известным токеном, попадает в имя как есть — включая
    /// одиночные скобки. Итог всегда прогоняется через <see cref="Sanitize"/>: имя файла
    /// не должно зависеть от того, что пользователь написал в шаблоне.
    /// Чистые функции — под тестами.
    /// </summary>
    public static class NameTemplate
    {
        /// <summary>Имена токенов — они же пункты меню подстановки (порядок показа).</summary>
        public static readonly string[] Tokens =
        {
            "BASENAME", "FILENUMBER", "CURRENTPAGE", "TOTAL_FILES", "BOOKMARK", "TIMESTAMP"
        };

        /// <summary>Шаблон по умолчанию — прежнее поведение: «основа_1», «основа_2».</summary>
        public const string Default = "[BASENAME]_[FILENUMBER]";

        /// <summary>
        /// Подставить значения в шаблон и вернуть ГОТОВОЕ имя файла без расширения.
        /// Пустой шаблон и шаблон, давший пустое имя (например «[BOOKMARK]» там, где закладки
        /// нет), откатываются к <see cref="Default"/>: файл без имени создать нельзя, а ронять
        /// из-за этого операцию незачем. Дальше отката не нужно — шаблон по умолчанию всегда
        /// даёт непустое имя, потому что номер файла в нём подставляется всегда.
        /// </summary>
        public static string Apply(string template, NameValues values)
        {
            if (values == null)
                values = new NameValues();
            string name = Sanitize(Expand(string.IsNullOrEmpty(template) ? Default : template, values));
            return name.Length > 0 ? name : Sanitize(Expand(Default, values));
        }

        /// <summary>
        /// Есть ли в шаблоне токен, различающий файлы между собой. Без такого токена все
        /// части получат ОДНО имя, и вызывающий обязан развести их сам (иначе останется
        /// один файл вместо десяти). Чистая — под тест.
        /// </summary>
        public static bool IsUniquePerFile(string template)
        {
            string t = string.IsNullOrEmpty(template) ? Default : template;
            return Find(t, "FILENUMBER") || Find(t, "CURRENTPAGE") || Find(t, "BOOKMARK");
        }

        private static bool Find(string template, string token)
        {
            return IndexOfToken(template, token) >= 0;
        }

        private static int IndexOfToken(string template, string token)
        {
            int from = 0;
            while (true)
            {
                int open = template.IndexOf('[', from);
                if (open < 0)
                    return -1;
                int close = template.IndexOf(']', open + 1);
                if (close < 0)
                    return -1;
                string body = template.Substring(open + 1, close - open - 1);
                string name;
                int pad, offset;
                if (TryParseToken(body, out name, out pad, out offset) &&
                    string.Equals(name, token, StringComparison.OrdinalIgnoreCase))
                    return open;
                from = close + 1;
            }
        }

        /// <summary>Подстановка токенов без очистки имени (внутренняя половина <see cref="Apply"/>).</summary>
        private static string Expand(string template, NameValues values)
        {
            var sb = new StringBuilder(template.Length + 16);
            int i = 0;
            while (i < template.Length)
            {
                if (template[i] != '[')
                {
                    sb.Append(template[i++]);
                    continue;
                }
                int close = template.IndexOf(']', i + 1);
                string body = close > i ? template.Substring(i + 1, close - i - 1) : null;
                string name;
                int pad, offset;
                if (body == null || !TryParseToken(body, out name, out pad, out offset))
                {
                    sb.Append(template[i++]); // не токен — скобка идёт в имя как обычный символ
                    continue;
                }
                sb.Append(Value(name, pad, offset, values));
                i = close + 1;
            }
            return sb.ToString();
        }

        /// <summary>
        /// Разобрать тело скобок: ИМЯ[решётки][смещение]. Неизвестное имя — не токен.
        /// Чистая — под тест.
        /// </summary>
        internal static bool TryParseToken(string body, out string name, out int pad, out int offset)
        {
            name = null;
            pad = 0;
            offset = 0;
            if (string.IsNullOrEmpty(body))
                return false;
            int i = 0;
            while (i < body.Length && (char.IsLetter(body[i]) || body[i] == '_'))
                i++;
            string candidate = body.Substring(0, i);
            if (!KnownToken(candidate))
                return false;
            while (i < body.Length && body[i] == '#')
            {
                pad++;
                i++;
            }
            int digitsFrom = i;
            while (i < body.Length && body[i] >= '0' && body[i] <= '9')
                i++;
            if (i != body.Length)
                return false; // хвост из посторонних символов — тело не токен целиком
            if (i > digitsFrom && !int.TryParse(body.Substring(digitsFrom, i - digitsFrom),
                    NumberStyles.None, CultureInfo.InvariantCulture, out offset))
                return false;
            name = candidate;
            return true;
        }

        private static bool KnownToken(string candidate)
        {
            foreach (string t in Tokens)
                if (string.Equals(t, candidate, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        private static string Value(string name, int pad, int offset, NameValues v)
        {
            switch (name.ToUpperInvariant())
            {
                case "BASENAME": return v.BaseName ?? "";
                case "BOOKMARK": return v.BookmarkName ?? "";
                case "TIMESTAMP":
                    return v.Timestamp == DateTime.MinValue
                        ? "" : v.Timestamp.ToString("yyyy-MM-dd_HH-mm-ss", CultureInfo.InvariantCulture);
                case "FILENUMBER": return Number(v.FileNumber + offset, pad);
                case "CURRENTPAGE": return Number(v.CurrentPage + offset, pad);
                case "TOTAL_FILES": return Number(v.TotalFiles + offset, pad);
            }
            return "";
        }

        private static string Number(int value, int pad)
        {
            string s = value.ToString(CultureInfo.InvariantCulture);
            return pad > 1 && s.Length < pad ? s.PadLeft(pad, '0') : s;
        }

        /// <summary>
        /// Привести к безопасному имени файла: запрещённые символы и точки/пробелы по краям
        /// убираются, длина ограничивается. Слэши тоже запрещены — шаблон не должен уметь
        /// раскладывать файлы по подпапкам, иначе операция запишет за пределы выбранной папки.
        /// Чистая — под тест.
        /// </summary>
        internal static string Sanitize(string name)
        {
            if (string.IsNullOrEmpty(name))
                return "";
            var sb = new StringBuilder(name.Length);
            foreach (char c in name)
                sb.Append(IsForbidden(c) ? '_' : c);
            string s = sb.ToString().Trim().TrimEnd('.', ' ');
            const int maxLength = 120; // с запасом под путь и «_2» при разведении совпадений
            if (s.Length > maxLength)
                s = s.Substring(0, maxLength).TrimEnd('.', ' ');
            return s;
        }

        private static bool IsForbidden(char c)
        {
            if (c < ' ')
                return true;
            foreach (char bad in InvalidChars)
                if (c == bad)
                    return true;
            return false;
        }

        // Path.GetInvalidFileNameChars() зависит от платформы; фиксируем набор Windows явно,
        // чтобы имя, собранное шаблоном, вело себя одинаково и в тестах, и в поле.
        private static readonly char[] InvalidChars = { '\\', '/', ':', '*', '?', '"', '<', '>', '|' };
    }
}
