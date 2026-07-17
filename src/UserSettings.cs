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
        public string OutputExtension = ".xlsx";

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
                    else if (key == "outputExtension" && OutputFormats.FileFormatFor("x" + value) != 0) s.OutputExtension = value;
                }
            }
            catch { } // повреждённые настройки не должны мешать запуску
            return s;
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                File.WriteAllLines(FilePath, new List<string>
                {
                    "lastInputFolder=" + (LastInputFolder ?? ""),
                    "lastOutputFolder=" + (LastOutputFolder ?? ""),
                    "addToc=" + AddToc,
                    "outputExtension=" + (OutputExtension ?? ".xlsx")
                });
            }
            catch { }
        }
    }
}
