using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Формирует допустимые и уникальные имена листов Excel из имён файлов:
    /// запрещённые символы заменяются на «_», длина ограничивается 31 символом,
    /// дубликаты получают суффиксы _2, _3 и т.д.
    /// </summary>
    public class SheetNamer
    {
        private const int MaxLength = 31;
        private static readonly char[] ForbiddenChars = { ':', '\\', '/', '?', '*', '[', ']' };

        private readonly HashSet<string> _used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Резервирует имя, чтобы оно не досталось ни одному листу-источнику.</summary>
        public void Reserve(string name)
        {
            _used.Add(name);
        }

        public string Next(string fileNameWithoutExtension)
        {
            string baseName = Sanitize(fileNameWithoutExtension);
            if (_used.Add(baseName))
                return baseName;

            for (int i = 2; ; i++)
            {
                string suffix = "_" + i;
                string candidate = Truncate(baseName, MaxLength - suffix.Length) + suffix;
                if (_used.Add(candidate))
                    return candidate;
            }
        }

        private static string Sanitize(string raw)
        {
            char[] chars = (raw ?? "").Trim().ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (Array.IndexOf(ForbiddenChars, chars[i]) >= 0)
                    chars[i] = '_';
            }

            // Имя листа не может начинаться или заканчиваться апострофом.
            string name = new string(chars).Trim().Trim('\'').Trim();
            name = Truncate(name, MaxLength);
            if (name.Length == 0)
                name = "Лист";
            if (string.Equals(name, "History", StringComparison.OrdinalIgnoreCase))
                name = "History_"; // зарезервированное Excel имя

            return name;
        }

        private static string Truncate(string value, int maxLength)
        {
            if (value.Length <= maxLength)
                return value;
            string cut = value.Substring(0, maxLength).Trim().Trim('\'').Trim();
            return cut.Length > 0 ? cut : value.Substring(0, maxLength);
        }
    }
}
