using System;
using System.IO;

namespace ExcelMerger
{
    /// <summary>Пути данных приложения в профиле пользователя.</summary>
    public static class AppPaths
    {
        // Переопределение корня для юнит-тестов: тесты настроек/статистики не должны
        // трогать живой %APPDATA% (падение процесса между записью и восстановлением
        // портило бы настройки пользователя, а параллельный запуск приложения — гонку).
        private static string _rootOverride;

        /// <summary>Направить все пути в указанный каталог (null — вернуть %APPDATA%). Только для тестов.</summary>
        internal static void SetRootForTests(string root)
        {
            _rootOverride = root;
        }

        public static string Root
        {
            get
            {
                return _rootOverride ?? Path.Combine(
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

        public static string StatsFile
        {
            get { return Path.Combine(Root, "stats.txt"); }
        }
    }
}
