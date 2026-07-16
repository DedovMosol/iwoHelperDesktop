using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>Последние выбранные пути; хранятся в %APPDATA%\ExcelMerger\settings.txt.</summary>
    public class UserSettings
    {
        public string LastInputFolder;
        public string LastOutputFolder;

        private static string FilePath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ExcelMerger", "settings.txt");
            }
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
                    if (key == "lastInputFolder") s.LastInputFolder = value;
                    else if (key == "lastOutputFolder") s.LastOutputFolder = value;
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
                    "lastOutputFolder=" + (LastOutputFolder ?? "")
                });
            }
            catch { }
        }
    }
}
