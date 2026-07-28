using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

namespace ExcelMerger
{
    /// <summary>Одна запись истории: когда, что сделали и что получилось.</summary>
    public sealed class HistoryEntry
    {
        /// <summary>Момент завершения операции (UTC — местное время считаем при показе).</summary>
        public DateTime WhenUtc;
        /// <summary>Ключ операции для <see cref="Loc"/>: подпись зависит от языка, а файл — нет.</summary>
        public string Operation;
        /// <summary>Путь результата. Папка, если результатов несколько.</summary>
        public string Path;
    }

    /// <summary>
    /// История операций: только ПУТИ и имена, никаких копий самих файлов. Живёт рядом со
    /// статистикой (<see cref="UsageStats"/>) и по тем же правилам: файл в профиле
    /// пользователя, межпроцессная блокировка на read-modify-write, автоочистка по периоду.
    ///
    /// Хранится ключ операции, а не готовая подпись: подпись зависит от языка интерфейса, и
    /// записанная по-русски строка осталась бы русской после переключения на английский.
    ///
    /// Приватность. Путь к файлу — это сведения о человеке и его работе, поэтому историю
    /// можно выключить целиком и стереть одной кнопкой, а глубина ограничена кольцом: без
    /// него файл рос бы без предела и хранил бы то, о чём давно забыли.
    /// </summary>
    public static class OperationHistory
    {
        /// <summary>Сколько последних операций храним. Старые вытесняются молча.</summary>
        public const int MaxEntries = 200;

        private const char Sep = '\t';
        private const string EnabledKey = "enabled";
        private const string AutoKey = "autoClearDays";
        private const string EntryKey = "e";

        private static string FilePath { get { return AppPaths.HistoryFile; } }

        // ---------- разбор и сборка строки (чистые — под тест) ----------

        /// <summary>
        /// Строка файла из записи. Разделитель — ТАБУЛЯЦИЯ, а не запятая и не точка с запятой:
        /// в путях Windows встречается и то и другое, а табуляция в имени файла невозможна.
        /// Всё же экранируем — путь может прийти не из проводника, а из чужой программы.
        /// Чистая — под тест.
        /// </summary>
        internal static string FormatEntry(HistoryEntry entry)
        {
            return EntryKey + "=" + entry.WhenUtc.Ticks.ToString(CultureInfo.InvariantCulture) +
                   Sep + Escape(entry.Operation) + Sep + Escape(entry.Path);
        }

        /// <summary>Разбор строки файла; null — строка испорчена и её пропускают. Чистая — под тест.</summary>
        internal static HistoryEntry ParseEntry(string line)
        {
            if (line == null || !line.StartsWith(EntryKey + "=", StringComparison.Ordinal))
                return null;
            string[] parts = line.Substring(EntryKey.Length + 1).Split(Sep);
            if (parts.Length != 3)
                return null;
            long ticks;
            if (!long.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out ticks) ||
                ticks < DateTime.MinValue.Ticks || ticks > DateTime.MaxValue.Ticks)
                return null;
            string path = Unescape(parts[2]);
            if (string.IsNullOrEmpty(path))
                return null;
            return new HistoryEntry
            {
                WhenUtc = new DateTime(ticks, DateTimeKind.Utc),
                Operation = Unescape(parts[1]),
                Path = path
            };
        }

        /// <summary>
        /// Экранирование для однострочного поля: обратная косая, табуляция и перевод строки.
        /// Без него путь с переводом строки (такое имя Windows не даст, но файл может прийти
        /// из архива или с сетевого диска) разорвал бы запись на две и испортил бы соседнюю.
        /// Чистая — под тест.
        /// </summary>
        internal static string Escape(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";
            var sb = new StringBuilder(value.Length + 8);
            foreach (char c in value)
            {
                if (c == '\\') sb.Append("\\\\");
                else if (c == '\t') sb.Append("\\t");
                else if (c == '\r') sb.Append("\\r");
                else if (c == '\n') sb.Append("\\n");
                else sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>Обратное к <see cref="Escape"/>. Чистая — под тест.</summary>
        internal static string Unescape(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";
            var sb = new StringBuilder(value.Length);
            for (int i = 0; i < value.Length; i++)
            {
                if (value[i] != '\\' || i + 1 >= value.Length)
                {
                    sb.Append(value[i]);
                    continue;
                }
                char next = value[++i];
                if (next == '\\') sb.Append('\\');
                else if (next == 't') sb.Append('\t');
                else if (next == 'r') sb.Append('\r');
                else if (next == 'n') sb.Append('\n');
                else sb.Append(next); // неизвестная последовательность — берём как есть
            }
            return sb.ToString();
        }

        /// <summary>
        /// Оставить последние <see cref="MaxEntries"/> записей. Кольцо считается ЗДЕСЬ, а не
        /// при показе: иначе файл рос бы без предела, а вместе с ним и перечень путей, о
        /// которых человек давно забыл. Чистая — под тест.
        /// </summary>
        internal static List<HistoryEntry> Trim(List<HistoryEntry> entries)
        {
            if (entries == null)
                return new List<HistoryEntry>();
            if (entries.Count <= MaxEntries)
                return entries;
            return entries.GetRange(entries.Count - MaxEntries, MaxEntries);
        }

        /// <summary>
        /// Убрать записи старше указанного числа дней (0 — не убирать ничего).
        ///
        /// Для СПИСКА автоочистка — это скользящая давность, а не «раз в N дней стереть всё»,
        /// как у счётчиков. Счётчик копится от метки сброса, и обнулить его целиком осмысленно.
        /// Список же состоит из разновозрастных записей: правило «период прошёл — чистим всё»,
        /// применённое к нему, стёрло бы и сегодняшние операции из-за одной позавчерашней.
        /// Заодно так честнее к приватности: старое уходит само, недавнее остаётся.
        /// Чистая — под тест.
        /// </summary>
        internal static List<HistoryEntry> KeepRecent(List<HistoryEntry> entries, DateTime nowUtc, int days)
        {
            var kept = new List<HistoryEntry>();
            if (entries == null)
                return kept;
            if (days <= 0)
            {
                kept.AddRange(entries);
                return kept;
            }
            foreach (HistoryEntry e in entries)
                if ((nowUtc - e.WhenUtc).TotalDays < days)
                    kept.Add(e);
            return kept;
        }

        // ---------- состояние ----------

        /// <summary>Прочитанная история: настройки и сами записи.</summary>
        public sealed class Data
        {
            /// <summary>Записывать ли историю. По умолчанию включена, выключается в настройках.</summary>
            public bool Enabled = true;
            /// <summary>Через сколько дней чистить (0 — никогда). Те же периоды, что у статистики.</summary>
            public int AutoClearDays;
            /// <summary>Записи от старых к новым.</summary>
            public readonly List<HistoryEntry> Entries = new List<HistoryEntry>();
        }

        public static Data Load()
        {
            var d = new Data();
            try
            {
                if (!File.Exists(FilePath))
                    return d;
                foreach (string line in File.ReadAllLines(FilePath))
                {
                    if (line.StartsWith(EnabledKey + "=", StringComparison.Ordinal))
                    {
                        bool flag;
                        if (bool.TryParse(line.Substring(EnabledKey.Length + 1), out flag))
                            d.Enabled = flag;
                        continue;
                    }
                    if (line.StartsWith(AutoKey + "=", StringComparison.Ordinal))
                    {
                        int days;
                        if (int.TryParse(line.Substring(AutoKey.Length + 1), out days))
                            d.AutoClearDays = days;
                        continue;
                    }
                    HistoryEntry entry = ParseEntry(line);
                    if (entry != null) // испорченная строка пропускается молча, как в настройках
                        d.Entries.Add(entry);
                }
                // Устаревшие записи отсеиваются ПРИ ЧТЕНИИ и в память — писать отсюда нельзя:
                // Load зовут и снаружи блокировки (чтобы показать список), и запись без неё
                // столкнулась бы с записью второй копии приложения. На диск отсев попадает
                // первой же мутацией: Mutate сохраняет ровно то, что вернул Load.
                List<HistoryEntry> fresh = KeepRecent(d.Entries, DateTime.UtcNow, d.AutoClearDays);
                if (fresh.Count != d.Entries.Count)
                {
                    d.Entries.Clear();
                    d.Entries.AddRange(fresh);
                }
            }
            catch { } // повреждённая история не повод мешать работе
            return d;
        }

        private static void Save(Data d)
        {
            try
            {
                Directory.CreateDirectory(AppPaths.Root);
                var lines = new List<string>
                {
                    EnabledKey + "=" + d.Enabled,
                    AutoKey + "=" + d.AutoClearDays
                };
                foreach (HistoryEntry e in Trim(d.Entries))
                    lines.Add(FormatEntry(e));
                File.WriteAllLines(FilePath, lines);
            }
            catch { }
        }

        // ---------- атомарные мутации ----------

        private static void Mutate(Action<Data> change)
        {
            // Та же межпроцессная блокировка, что у статистики, но СВОЯ: общий мьютекс на два
            // файла заставлял бы запись истории ждать запись счётчиков без всякой причины.
            using (var mutex = new Mutex(false, @"Local\iwoHelperDesktop.history"))
            {
                bool held = false;
                try
                {
                    try { held = mutex.WaitOne(2000); }
                    catch (AbandonedMutexException) { held = true; } // прежний держатель умер
                    Data d = Load();
                    change(d);
                    Save(d);
                }
                finally
                {
                    if (held)
                        mutex.ReleaseMutex();
                }
            }
        }

        /// <summary>
        /// Записать завершённую операцию. Выключенная история молчит — и НЕ проверяет путь:
        /// человек отказался от записи, значит ни читать файл, ни трогать диск не за чем.
        /// </summary>
        public static void Record(string operationKey, string path)
        {
            if (string.IsNullOrEmpty(path))
                return;
            // Быстрый выход БЕЗ блокировки и без записи. Иначе выключенная история всё равно
            // перечитывала бы и переписывала файл на каждой операции — и воссоздавала бы его
            // после того, как человек его удалил. Решает всё равно проверка внутри мутации:
            // здесь мы лишь не трогаем диск понапрасну.
            if (!Load().Enabled)
                return;
            Mutate(delegate(Data d)
            {
                if (!d.Enabled)
                    return;
                d.Entries.Add(new HistoryEntry
                {
                    WhenUtc = DateTime.UtcNow,
                    Operation = operationKey,
                    Path = path
                });
            });
        }

        public static void SetEnabled(bool enabled)
        {
            // Выключение СРАЗУ стирает накопленное: оставить перечень путей у того, кто
            // только что отказался от их хранения, — это не то, о чём он просил.
            Mutate(delegate(Data d)
            {
                d.Enabled = enabled;
                if (!enabled)
                    d.Entries.Clear();
            });
        }

        public static void SetAutoClear(int days) { Mutate(delegate(Data d) { d.AutoClearDays = days; }); }

        public static void Clear() { Mutate(delegate(Data d) { d.Entries.Clear(); }); }
    }
}
