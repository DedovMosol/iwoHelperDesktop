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
        public string Language = "ru";          // язык интерфейса: «ru»/«en» (см. Loc)
        // Вид PDF-инструментов между запусками: ширина плитки (0 — не задана, брать умолчание)
        // и уровень сжатия (индекс CompressionLevel). Применяет/сохраняет PdfToolFormBase.
        public int ZoomWidth;
        public int CompressionLevel;

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
            UserSettings disk = Load(); // свежие значения полей, которыми этот экземпляр не владеет
            WriteAll(disk.ZoomWidth, disk.CompressionLevel);
        }

        /// <summary>
        /// Сохранить настройки вместе с видом PDF-инструмента: масштаб и уровень сжатия
        /// пишутся ЯВНО (вызывает <see cref="PdfToolFormBase"/> при закрытии окна поверх
        /// свежей загрузки, поэтому прочие поля тоже актуальны).
        /// </summary>
        public void SaveView(int zoomWidth, int compressionLevel)
        {
            WriteAll(zoomWidth, compressionLevel);
        }

        private void WriteAll(int zoomWidth, int compressionLevel)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                File.WriteAllLines(FilePath, new List<string>
                {
                    "lastInputFolder=" + (LastInputFolder ?? ""),
                    "lastOutputFolder=" + (LastOutputFolder ?? ""),
                    "addToc=" + AddToc,
                    "allSheets=" + AllSheets,
                    "outputExtension=" + (OutputExtension ?? ".xlsx"),
                    "zoomWidth=" + zoomWidth,
                    "compression=" + compressionLevel,
                    // Язык — из живого Loc (единый источник истины), а НЕ из поля этого
                    // экземпляра: другие формы держат устаревшую копию настроек и иначе
                    // затёрли бы язык обратно при своём Save.
                    "language=" + Loc.Code(Loc.Current)
                });
            }
            catch { }
        }
    }
}
