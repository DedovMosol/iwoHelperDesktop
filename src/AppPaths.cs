using System;
using System.IO;

namespace ExcelMerger
{
    /// <summary>Пути данных приложения в профиле пользователя.</summary>
    public static class AppPaths
    {
        public static string Root
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "iwo Helper Desktop");
            }
        }

        public static string SettingsFile
        {
            get { return Path.Combine(Root, "settings.txt"); }
        }

        public static string ReportsDir
        {
            get { return Path.Combine(Root, "reports"); }
        }
    }
}
