using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>Пользовательские настройки; хранятся в %APPDATA%\iwo Helper Desktop\settings.txt.</summary>
    public class UserSettings
    {
        public string LastInputFolder;
        public string LastOutputFolder;
        // «Заменить формулы значениями» сознательно НЕ запоминается: режим меняет
        // содержимое свода, включать его нужно осознанно на каждый запуск.
        public bool AddToc = true;
        public bool AllSheets;
        public string OutputExtension = ".xlsx";
        public string Language;
        public int ZoomWidth;
        public int CompressionLevel;
        public readonly Dictionary<string, string> WindowBounds =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public bool UpdateCheckOnStart = true;
        public bool ShowWhatsNewOnStart = true;
        public string LastWhatsNewVersion;
        public string SkippedVersion;

        // Незнакомые строки сохраняются при downgrade: старая версия не должна стирать
        // настройки, добавленные более новой, только потому что пользователь закрыл окно.
        private readonly List<string> _unknownLines = new List<string>();

        private static string FilePath { get { return AppPaths.SettingsFile; } }

        public static UserSettings Load()
        {
            UserSettings settings;
            return TryLoad(out settings) ? settings : new UserSettings();
        }

        /// <summary>
        /// true означает достоверный снимок: файл прочитан либо действительно отсутствует.
        /// Существующий, но временно непрочитанный файл НЕЛЬЗЯ подменять defaults при мутации.
        /// </summary>
        private static bool TryLoad(out UserSettings settings)
        {
            settings = new UserSettings();
            string[] lines;
            if (!AppStateFile.TryReadLines(FilePath, out lines))
                return false;

            foreach (string line in lines)
            {
                int eq = line.IndexOf('=');
                if (eq <= 0)
                {
                    settings._unknownLines.Add(line);
                    continue;
                }
                string key = line.Substring(0, eq).Trim();
                string value = line.Substring(eq + 1).Trim();
                bool known = true;
                bool flag;
                if (key == "lastInputFolder") settings.LastInputFolder = value;
                else if (key == "lastOutputFolder") settings.LastOutputFolder = value;
                else if (key == "addToc" && bool.TryParse(value, out flag)) settings.AddToc = flag;
                else if (key == "allSheets" && bool.TryParse(value, out flag)) settings.AllSheets = flag;
                else if (key == "outputExtension" && OutputFormats.FileFormatFor("x" + value) != 0) settings.OutputExtension = value;
                else if (key == "language" && (value == "ru" || value == "en")) settings.Language = value;
                else if (key == "zoomWidth") { int z; if (int.TryParse(value, out z)) settings.ZoomWidth = z; }
                else if (key == "compression") { int c; if (int.TryParse(value, out c)) settings.CompressionLevel = c; }
                else if (key == "updateCheckOnStart" && bool.TryParse(value, out flag)) settings.UpdateCheckOnStart = flag;
                else if (key == "showWhatsNewOnStart" && bool.TryParse(value, out flag)) settings.ShowWhatsNewOnStart = flag;
                else if (key == "lastWhatsNewVersion") settings.LastWhatsNewVersion = value.Length == 0 ? null : value;
                else if (key == "skippedVersion") settings.SkippedVersion = value.Length == 0 ? null : value;
                else if (key.StartsWith("wnd.", StringComparison.Ordinal) && key.Length > 4)
                    settings.WindowBounds[key.Substring(4)] = value;
                else
                    known = false;
                if (!known)
                    settings._unknownLines.Add(line);
            }
            return true;
        }

        /// <summary>
        /// Сохранить настройки владельца Excel поверх свежего снимка. Возвращает false,
        /// если блокировку/чтение/атомарную публикацию выполнить не удалось.
        /// </summary>
        public bool Save()
        {
            return Change(delegate(UserSettings disk)
            {
                disk.LastInputFolder = LastInputFolder;
                disk.LastOutputFolder = LastOutputFolder;
                disk.AddToc = AddToc;
                disk.AllSheets = AllSheets;
                disk.OutputExtension = OutputExtension;
            });
        }

        /// <summary>
        /// Сохранить только вид PDF-инструмента поверх свежей загрузки: устаревший экземпляр
        /// окна не затирает пути, границы, язык и настройки обновлений другого окна.
        /// </summary>
        public bool SaveView(int zoomWidth, int compressionLevel)
        {
            bool saved = Change(delegate(UserSettings disk)
            {
                disk.ZoomWidth = zoomWidth;
                disk.CompressionLevel = compressionLevel;
            });
            if (saved)
            {
                ZoomWidth = zoomWidth;
                CompressionLevel = compressionLevel;
            }
            return saved;
        }

        public static void SaveWindowBounds(string formKey, string bounds)
        {
            Change(delegate(UserSettings s) { s.WindowBounds[formKey] = bounds; });
        }

        public static bool SaveSkippedVersion(string version)
        {
            return Change(delegate(UserSettings s) { s.SkippedVersion = version; });
        }

        public static bool SaveUpdateCheckOnStart(bool checkOnStart)
        {
            return Change(delegate(UserSettings s) { s.UpdateCheckOnStart = checkOnStart; });
        }

        public static bool SaveWhatsNew(bool showOnStart, string seenVersion)
        {
            return Change(delegate(UserSettings settings)
            {
                settings.ShowWhatsNewOnStart = showOnStart;
                settings.LastWhatsNewVersion = seenVersion;
            });
        }

        public static bool SaveLanguage(string language)
        {
            if (language != "ru" && language != "en")
                return false;
            return Change(delegate(UserSettings s) { s.Language = language; });
        }

        private static bool Change(Action<UserSettings> change)
        {
            if (change == null)
                return false;
            return AppDataLock.TryRun(FilePath, delegate
            {
                UserSettings disk;
                if (!TryLoad(out disk))
                    return false;
                change(disk);
                disk.WriteAll();
                return true;
            });
        }

        private void WriteAll()
        {
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
                "showWhatsNewOnStart=" + ShowWhatsNewOnStart,
                "lastWhatsNewVersion=" + (LastWhatsNewVersion ?? ""),
                "skippedVersion=" + (SkippedVersion ?? ""),
                "language=" + (Language == "en" ? "en" :
                    Language == "ru" ? "ru" : Loc.Code(Loc.Current))
            };
            var keys = new List<string>(WindowBounds.Keys);
            keys.Sort(StringComparer.Ordinal);
            foreach (string key in keys)
                lines.Add("wnd." + key + "=" + WindowBounds[key]);
            lines.AddRange(_unknownLines);
            AppStateFile.WriteLines(FilePath, lines);
        }
    }
}
