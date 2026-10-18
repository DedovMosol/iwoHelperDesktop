using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Подстановка шрифта и предельные кегли для писателей документов. Общий слой: и Word,
    /// и PowerPoint молча подменяют НЕустановленное семейство своим, и эта подмена меняет
    /// метрики — значит решать, чем писать ран, надо ДО вывода и одинаково. Чистая логика
    /// (<see cref="ResolveFontName"/>, <see cref="HasCyrillic"/>, <see cref="ClampSizePt"/>)
    /// покрыта юнит-тестами; список установленных шрифтов читается один раз на процесс.
    /// </summary>
    internal static class FontResolver
    {
        public const string DefaultFontName = "Times New Roman";
        public const double DefaultFontSize = 12;
        public const double MinFontSize = 5;   // защита от мусорного кегля из PDF
        public const double MaxFontSize = 72;  // потоковый документ (Word): выше — почти всегда мусор
        // Слайд: заголовки в 80–200 pt там обычное дело. Предел — сам формат PowerPoint
        // (4000 pt, см. PptxGeometry.Hundredths); Word такого кегля не примет (максимум 1638),
        // поэтому этот предел только для слайдов.
        public const double MaxSlideFontSize = 4000;

        /// <summary>
        /// Имя шрифта для рана: если шрифт источника установлен в системе — оставляем его, иначе
        /// подставляем установленный по умолчанию. КЛЮЧЕВОЕ: когда Word получает НЕустановленный
        /// шрифт, он уводит кириллицу в восточноазиатский фолбэк-слот (rFonts hint="eastAsia"),
        /// и при выключке по ширине раздвигает буквы по правилам CJK — получается «р а з р я д к а»
        /// (а латиница остаётся слитной). Установленный шрифт держит кириллицу в hAnsi — обычная
        /// выключка. Но и УСТАНОВЛЕННОГО мало: «не родные» для Word семейства (Liberation Serif
        /// и т.п.) получают тот же hint="eastAsia" при вводе кириллицы — поэтому кириллический
        /// текст пишется только шрифтами из сейф-листа Word-родных семейств, прочие уводятся в
        /// fallback (Liberation Serif — метрический клон Times New Roman, подмена нейтральна).
        /// Список шрифтов читается один раз.
        /// </summary>
        public static string Resolve(string requested, string text)
        {
            return ResolveFontName(requested, text, InstalledFonts, DefaultFontName);
        }

        // Word-родные семейства с полной кириллицей, за которыми Word не замечен в eastAsia-хинте.
        private static readonly HashSet<string> CyrillicSafeFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Times New Roman", "Arial", "Calibri", "Calibri Light", "Courier New",
            "Cambria", "Georgia", "Verdana", "Tahoma", "Segoe UI", "Consolas"
        };

        /// <summary>Чистая логика подстановки (под тест): установленный — оставить; кириллице — только сейф-лист.</summary>
        internal static string ResolveFontName(string requested, string text, ICollection<string> installed, string fallback)
        {
            if (string.IsNullOrEmpty(requested))
                return fallback;
            if (installed == null || !installed.Contains(requested))
                return fallback;
            return HasCyrillic(text) && !CyrillicSafeFonts.Contains(requested) ? fallback : requested;
        }

        /// <summary>Есть ли в тексте кириллица (U+0400–U+04FF).</summary>
        internal static bool HasCyrillic(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;
            for (int i = 0; i < text.Length; i++)
                if (text[i] >= 'Ѐ' && text[i] <= 'ӿ')
                    return true;
            return false;
        }

        private static readonly HashSet<string> InstalledFonts = LoadInstalledFonts();

        /// <summary>Семейства установленных шрифтов (без учёта регистра). Сбой чтения — пустой набор (всё уйдёт в fallback).</summary>
        private static HashSet<string> LoadInstalledFonts()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var col = new System.Drawing.Text.InstalledFontCollection())
                    foreach (System.Drawing.FontFamily fam in col.Families)
                        set.Add(fam.Name);
            }
            catch { }
            return set;
        }

        /// <summary>Кегль рана для документа: <see cref="ClampSizePt(double, double)"/> с пределом Word.</summary>
        internal static double ClampSizePt(double sizePt)
        {
            return ClampSizePt(sizePt, MaxFontSize);
        }

        /// <summary>
        /// Кегль рана, ПРИЖАТЫЙ к пределам [MinFontSize, maxPt]. Раньше всё вне пределов
        /// становилось кеглем по умолчанию — заголовок в 80 pt и подпись в 4,5 pt выходили
        /// одинаковыми 12 pt, хуже любого округления к границе. Умолчание — только когда
        /// кегля нет вовсе (NaN, ноль, отрицательный). Чистая — под тест.
        /// </summary>
        internal static double ClampSizePt(double sizePt, double maxPt)
        {
            if (double.IsNaN(sizePt) || sizePt <= 0)
                return DefaultFontSize;
            if (sizePt < MinFontSize)
                return MinFontSize;
            return sizePt > maxPt ? maxPt : sizePt;
        }
    }
}
