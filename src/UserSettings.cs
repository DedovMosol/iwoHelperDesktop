using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>Пользовательские настройки; хранятся в %APPDATA%\ExcelMerger\settings.txt.</summary>
    public class UserSettings
    {
        public string LastInputFolder;
        public string LastOutputFolder;
        public bool AddToc = true;     // «Содержание» по умолчанию включено
        public bool ValuesOnly;        // формулы по умолчанию сохраняются

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
                    else if (key == "valuesOnly" && bool.TryParse(value, out flag)) s.ValuesOnly = flag;
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
                    "valuesOnly=" + ValuesOnly
                });
            }
            catch { }
        }
    }
}
