using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>Пользовательские настройки; хранятся в %APPDATA%\iwo Helper Desktop\settings.txt.</summary>
    public class UserSettings
    {
        public string LastInputFolder;
        public string LastOutputFolder;
        // «Заменить формулы значениями» сознательно НЕ запоминается: режим меняет
        // содержимое свода, включать его нужно осознанно на каждый запуск.
        public bool AddToc = true;              // «Содержание» по умолчанию включено
        public bool AllSheets;                  // все листы (по умолчанию — только первый)
        public string OutputExtension = ".xlsx";
        // Язык интерфейса «ru»/«en» (см. Loc). null = не задан (первый запуск без настроек):
        // Program берёт язык по системной локали. Установленную версию сидит инсталлер.
        public string Language;
        // Вид PDF-инструментов между запусками: ширина плитки (0 — не задана, брать умолчание)
        // и уровень сжатия (индекс CompressionLevel). Применяет/сохраняет PdfToolFormBase.
        public int ZoomWidth;
        public int CompressionLevel;
        // Размер/положение окон между запусками: ключ — имя типа формы, значение «x,y,w,h,m»
        // (см. WindowPlacement). Кросс-настройка (не «своя» ни у одного окна) — сохраняется из
        // свежей загрузки, поэтому долгоживущее окно не затрёт границы, записанные другим окном.
        public readonly Dictionary<string, string> WindowBounds =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // Проверять обновления при запуске. По умолчанию включено, выключается в настройках.
        public bool UpdateCheckOnStart = true;
        // Версия, о которой пользователь просил больше не напоминать («v1.18.0» → «1.18.0»).
        // Хранится ИМЕННО версия, а не флаг «никогда»: иначе один флажок отключил бы
        // уведомления навсегда, и о следующих выпусках человек не узнал бы вовсе.
        public string SkippedVersion;

        private static string FilePath
        {
            get { return AppPaths.SettingsFile; }
        }

        public static UserSettings Load()
        {
            var s = new UserSettings();
            try
            {
                if (!File.Exists(FilePath))
                    return s;
                foreach (string line in File.ReadAllLines(FilePath))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0)
                        continue;
                    string key = line.Substring(0, eq).Trim();
                    string value = line.Substring(eq + 1).Trim();
                    bool flag;
                    if (key == "lastInputFolder") s.LastInputFolder = value;
                    else if (key == "lastOutputFolder") s.LastOutputFolder = value;
                    else if (key == "addToc" && bool.TryParse(value, out flag)) s.AddToc = flag;
                    else if (key == "allSheets" && bool.TryParse(value, out flag)) s.AllSheets = flag;
                    else if (key == "outputExtension" && OutputFormats.FileFormatFor("x" + value) != 0) s.OutputExtension = value;
                    else if (key == "language" && (value == "ru" || value == "en")) s.Language = value;
                    else if (key == "zoomWidth") { int z; if (int.TryParse(value, out z)) s.ZoomWidth = z; }
                    else if (key == "compression") { int c; if (int.TryParse(value, out c)) s.CompressionLevel = c; }
                    else if (key == "updateCheckOnStart" && bool.TryParse(value, out flag)) s.UpdateCheckOnStart = flag;
                    // Пустая строка — «ничего не пропущено»: иначе пустое значение стало бы
                    // версией, которую не разобрать, и проверка молча перестала бы срабатывать.
                    else if (key == "skippedVersion") s.SkippedVersion = value.Length == 0 ? null : value;
                    else if (key.StartsWith("wnd.", StringComparison.Ordinal) && key.Length > 4)
                        s.WindowBounds[key.Substring(4)] = value;
                }
            }
            catch { } // повреждённые настройки не должны мешать запуску
            return s;
        }

        /// <summary>
        /// Сохранить настройки. Масштаб и уровень сжатия PDF-инструментов НЕ берутся из
        /// полей этого экземпляра, а сохраняются из свежайшего значения на диске: их
        /// владелец — PDF-окна, а долгоживущие окна (MainForm держит копию с запуска)
        /// иначе затёрли бы устаревшим значением чужой правкой (та же ловушка, что с
        /// языком). Осознанно эти поля пишет только <see cref="SaveView"/>.
        /// </summary>
        public void Save()
        {
            // Свежий снимок с диска, на него переносим ТОЛЬКО свои поля. Так список
            // «чужих» полей не приходится держать в голове: всё, что здесь не названо,
            // остаётся с диска и достаётся своему владельцу в целости. Каждое новое поле
            // по умолчанию оказывается чужим — это безопасная сторона ошибки.
            UserSettings disk = Load();
            disk.LastInputFolder = LastInputFolder;
            disk.LastOutputFolder = LastOutputFolder;
            disk.AddToc = AddToc;
            disk.AllSheets = AllSheets;
            disk.OutputExtension = OutputExtension;
            disk.WriteAll();
        }

        /// <summary>
        /// Сохранить настройки вместе с видом PDF-инструмента: масштаб и уровень сжатия
        /// пишутся ЯВНО (вызывает <see cref="PdfToolFormBase"/> при закрытии окна поверх
        /// свежей загрузки, поэтому прочие поля тоже актуальны).
        /// </summary>
        public void SaveView(int zoomWidth, int compressionLevel)
        {
            ZoomWidth = zoomWidth;
            CompressionLevel = compressionLevel;
            WriteAll();
        }

        /// <summary>
        /// Сохранить размер/положение одного окна (ключ — имя типа формы), не тронув прочие
        /// настройки и границы других окон: read-modify-write свежей загрузки (как SaveView для
        /// масштаба). Долгоживущее окно так не затрёт границы, записанные другим окном.
        /// </summary>
        public static void SaveWindowBounds(string formKey, string bounds)
        {
            Change(delegate(UserSettings s) { s.WindowBounds[formKey] = bounds; });
        }

        /// <summary>
        /// Запомнить версию, о которой просили не напоминать. Отдельным узким методом, а не
        /// парой «флажок + версия» в одном вызове: совмещённый метод требовал бы от каждого
        /// вызывающего передавать И текущее состояние проверки, а забытый параметр молча
        /// включил бы обратно то, что пользователь выключил.
        /// </summary>
        public static void SaveSkippedVersion(string version)
        {
            Change(delegate(UserSettings s) { s.SkippedVersion = version; });
        }

        /// <summary>Включить или выключить проверку обновлений при запуске.</summary>
        public static void SaveUpdateCheckOnStart(bool checkOnStart)
        {
            Change(delegate(UserSettings s) { s.UpdateCheckOnStart = checkOnStart; });
        }

        /// <summary>
        /// Поменять одно поле поверх СВЕЖЕЙ загрузки. Общий приём для настроек, у которых нет
        /// окна-владельца: правится ровно названное поле, остальные попадают на диск такими,
        /// какими их оставил их собственный владелец.
        /// </summary>
        private static void Change(Action<UserSettings> change)
        {
            UserSettings disk = Load();
            change(disk);
            disk.WriteAll();
        }

        private void WriteAll()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                var lines = new List<string>
                {
                    "lastInputFolder=" + (LastInputFolder ?? ""),
                    "lastOutputFolder=" + (LastOutputFolder ?? ""),
                    "addToc=" + AddToc,
                    "allSheets=" + AllSheets,
                    "outputExtension=" + (OutputExtension ?? ".xlsx"),
                    "zoomWidth=" + ZoomWidth,
                    "compression=" + CompressionLevel,
                    "updateCheckOnStart=" + UpdateCheckOnStart,
                    "skippedVersion=" + (SkippedVersion ?? ""),
                    // Язык — из живого Loc (единый источник истины), а НЕ из поля этого
                    // экземпляра: другие формы держат устаревшую копию настроек и иначе
                    // затёрли бы язык обратно при своём Save.
                    "language=" + Loc.Code(Loc.Current)
                };
                // Границы окон (wnd.<Форма>=x,y,w,h,m) — отсортированно: детерминированный файл.
                var keys = new List<string>(WindowBounds.Keys);
                keys.Sort(StringComparer.Ordinal);
                foreach (string k in keys)
                    lines.Add("wnd." + k + "=" + WindowBounds[k]);
                File.WriteAllLines(FilePath, lines);
            }
            catch { }
        }
    }
}
