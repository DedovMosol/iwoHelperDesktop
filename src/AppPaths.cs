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

        /// <summary>История операций — рядом со статистикой и по тем же правилам.</summary>
        public static string HistoryFile
        {
            get { return Path.Combine(Root, "history.txt"); }
        }

        /// <summary>
        /// Язык, выбранный в установщике (см. <see cref="SetupLanguage"/>). Отдельный файл, а не
        /// строка в settings.txt: установщик пишет в системной кодировке, а приложение читает
        /// UTF-8, и правка общего файла из установщика испортила бы кириллические пути в нём.
        /// </summary>
        public static string SetupLanguageFile
        {
            get { return Path.Combine(Root, "setup-language.txt"); }
        }
    }
}
