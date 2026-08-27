using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

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
            internal readonly List<string> UnknownLines = new List<string>();
        }

        public static Data Load()
        {
            Data data;
            return TryLoad(out data) ? data : new Data();
        }

        /// <summary>
        /// Существующий, но временно непрочитанный файл — не пустая история: мутация должна
        /// отказаться, а не перезаписать opt-out/defaults поверх неизвестного снимка.
        /// </summary>
        private static bool TryLoad(out Data data)
        {
            data = new Data();
            string[] lines;
            if (!AppStateFile.TryReadLines(FilePath, out lines))
                return false;

            foreach (string line in lines)
            {
                if (line.StartsWith(EnabledKey + "=", StringComparison.Ordinal))
                {
                    bool flag;
                    if (bool.TryParse(line.Substring(EnabledKey.Length + 1), out flag))
                        data.Enabled = flag;
                    continue;
                }
                if (line.StartsWith(AutoKey + "=", StringComparison.Ordinal))
                {
                    int days;
                    if (int.TryParse(line.Substring(AutoKey.Length + 1), out days) && days >= 0)
                        data.AutoClearDays = days;
                    continue;
                }
                HistoryEntry entry = ParseEntry(line);
                if (entry != null)
                    data.Entries.Add(entry);
                else
                    data.UnknownLines.Add(line);
            }
            List<HistoryEntry> fresh = KeepRecent(data.Entries, DateTime.UtcNow,
                data.AutoClearDays);
            if (fresh.Count != data.Entries.Count)
            {
                data.Entries.Clear();
                data.Entries.AddRange(fresh);
            }
            return true;
        }

        private static void Save(Data data)
        {
            var lines = new List<string>
            {
                EnabledKey + "=" + data.Enabled,
                AutoKey + "=" + data.AutoClearDays
            };
            foreach (HistoryEntry entry in Trim(data.Entries))
                lines.Add(FormatEntry(entry));
            lines.AddRange(data.UnknownLines);
            AppStateFile.WriteLines(FilePath, lines);
        }

        // ---------- атомарные мутации ----------

        private static bool Mutate(Func<Data, bool> change)
        {
            if (change == null)
                return false;
            return AppDataLock.TryRun(FilePath, delegate
            {
                Data data;
                if (!TryLoad(out data) || !change(data))
                    return false;
                Save(data);
                return true;
            });
        }

        private static bool MutateAndNotify(Func<Data, bool> change)
        {
            bool success = Mutate(change);
            if (success)
                NotifyChanged();
            return success;
        }

        /// <summary>
        /// Записать завершённую операцию. Проверка Enabled и добавление выполняются в одной
        /// блокировке: выключенная история не переписывается, а непрочитанный файл не теряется.
        /// </summary>
        public static void Record(string operationKey, string path)
        {
            if (string.IsNullOrEmpty(path))
                return;
            MutateAndNotify(delegate(Data data)
            {
                if (!data.Enabled)
                    return false;
                data.Entries.Add(new HistoryEntry
                {
                    WhenUtc = DateTime.UtcNow,
                    Operation = operationKey,
                    Path = path
                });
                return true;
            });
        }

        /// <summary>
        /// История пополнилась или очищена. Нужно стартовому экрану: он живёт всё время работы
        /// программы, а операции идут в других окнах — без уведомления список недавних оставался
        /// бы вчерашним до перезапуска.
        ///
        /// Событие приходит С ТОГО ПОТОКА, где закончилась операция (обычно фонового), поэтому
        /// подписчик обязан сам вернуться на поток интерфейса.
        /// </summary>
        public static event Action Changed;

        /// <summary>Сообщить об изменении, сделанном не через Record (очистка, выключение).</summary>
        private static void NotifyChanged()
        {
            Action changed = Changed;
            if (changed == null)
                return;
            foreach (Action handler in changed.GetInvocationList())
                try { handler(); } catch { }
        }

        public static bool SetEnabled(bool enabled)
        {
            // Выключение СРАЗУ стирает накопленное: оставить перечень путей у того, кто
            // только что отказался от их хранения, — это не то, о чём он просил.
            return MutateAndNotify(delegate(Data data)
            {
                data.Enabled = enabled;
                if (!enabled)
                    data.Entries.Clear();
                return true;
            });
        }

        public static bool SetAutoClear(int days)
        {
            int period = Math.Max(0, days);
            return MutateAndNotify(delegate(Data data)
            {
                data.AutoClearDays = period;
                List<HistoryEntry> fresh = KeepRecent(data.Entries, DateTime.UtcNow, period);
                data.Entries.Clear();
                data.Entries.AddRange(fresh);
                return true;
            });
        }

        public static bool Clear()
        {
            return MutateAndNotify(delegate(Data data)
            {
                data.Entries.Clear();
                return true;
            });
        }

        /// <summary>
        /// Последние результаты, к которым имеет смысл вернуться: от новых к старым, без
        /// повторов и без того, чего на диске уже нет. Проверка существования приходит
        /// параметром — так правило проверяется тестом, не заводя настоящих файлов.
        ///
        /// Показывать путь к исчезнувшему файлу нельзя: список быстрого доступа, половина
        /// которого не открывается, хуже отсутствия списка. Чистая — под тест.
        /// </summary>
        public static List<HistoryEntry> RecentFiles(Data data, int max, Func<string, bool> exists)
        {
            var result = new List<HistoryEntry>();
            if (data == null || max <= 0 || exists == null)
                return result;
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = data.Entries.Count - 1; i >= 0 && result.Count < max; i--)
            {
                HistoryEntry entry = data.Entries[i];
                if (entry == null || string.IsNullOrEmpty(entry.Path) || !seen.Add(entry.Path))
                    continue;
                bool alive;
                try { alive = exists(entry.Path); }
                catch { alive = false; }
                if (alive)
                    result.Add(entry);
            }
            return result;
        }
    }
}
