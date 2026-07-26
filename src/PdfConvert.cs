using System.IO;
using System.Text;

namespace ExcelMerger
{
    /// <summary>Преобразование готового PDF на месте — тем же движком, что и сжатие.</summary>
    public enum PdfConvertMode
    {
        /// <summary>В оттенки серого: цветные страницы дешевле печатать, файл обычно меньше.</summary>
        Grayscale = 0,
        /// <summary>
        /// Восстановление: перезапись файла движком чинит битую таблицу ссылок и мусор в
        /// конце — типичный «файл повреждён» у документов, докачанных или обрезанных на
        /// передаче. Содержимое при этом пересобирается, а не правится по месту.
        /// </summary>
        Repair = 1
    }

    /// <summary>
    /// Преобразования PDF через Ghostscript, отличные от сжатия. Делят с ним весь конвейер
    /// (<see cref="GsRewrite"/>): выполнить, проверить вывод, заменить оригинал только при
    /// успехе, вернуть его при любом сбое. Отличаются аргументами и политикой замены —
    /// здесь она мягче, чем у сжатия: результат должен примениться, даже если стал больше
    /// (серый вариант или починенный файл почти всегда другого размера).
    ///
    /// PDF/A сюда НЕ входит намеренно. Проба показала, что без внешнего цветового профиля
    /// движок молча выдаёт обычный PDF версии 1.7 вместо PDF/A, а такой файл наши же
    /// инструменты потом не открывают (нужен 1.4). Обещать соответствие, не умея его
    /// проверить, — хуже, чем не делать вовсе.
    /// </summary>
    public static class PdfConvert
    {
        private const int TimeoutMs = 300000; // как у сжатия — верхняя граница на большой скан

        /// <summary>
        /// Аргументы Ghostscript для режима. Версия вывода — 1.4, как и у сжатия: классическая
        /// таблица ссылок остаётся читаемой нашим же PdfSharp, иначе преобразованный файл
        /// нельзя было бы снова объединить или разделить. Чистая — под тест.
        /// </summary>
        public static string BuildArguments(string input, string output, PdfConvertMode mode, string bundledGsRoot)
        {
            var sb = new StringBuilder();
            sb.Append("-sDEVICE=pdfwrite -dCompatibilityLevel=1.4");
            if (mode == PdfConvertMode.Grayscale)
            {
                // Обе части обязательны: стратегия переводит цвета, модель устройства не даёт
                // движку вернуть цветовое пространство обратно на отдельных объектах.
                sb.Append(" -sColorConversionStrategy=Gray -dProcessColorModel=/DeviceGray");
            }
            sb.Append(" -dNOPAUSE -dBATCH -dQUIET -dSAFER");
            if (!string.IsNullOrEmpty(bundledGsRoot))
            {
                sb.Append(" -I ").Append(Quote(Path.Combine(bundledGsRoot, "lib")));
                sb.Append(" -I ").Append(Quote(Path.Combine(bundledGsRoot, "Resource", "Init")));
            }
            sb.Append(" -sOutputFile=").Append(Quote(output));
            sb.Append(' ').Append(Quote(input));
            return sb.ToString();
        }

        /// <summary>
        /// Заменять ли оригинал результатом преобразования: достаточно, чтобы вывод был
        /// годным PDF и непустым. Размер здесь не показатель — в отличие от сжатия, эти
        /// режимы меняют файл по существу, а не ради экономии места. Чистая — под тест.
        /// </summary>
        public static bool ShouldReplace(long origSize, long newSize, bool newValid)
        {
            return newValid && newSize > 0;
        }

        /// <summary>
        /// Преобразовать файл на месте. true — оригинал заменён. Без Ghostscript и при любом
        /// сбое возвращает false, оставляя файл нетронутым. Только с фонового потока и до
        /// открытия файла посторонними программами.
        /// </summary>
        public static bool Apply(string path, PdfConvertMode mode)
        {
            if (!Ghostscript.Available)
                return false;
            string args = BuildArguments(path, GsRewrite.TempOutput(path), mode, Ghostscript.BundledRoot);
            return GsRewrite.Run(path, args, TimeoutMs, ShouldReplace);
        }

        private static string Quote(string s)
        {
            return "\"" + s + "\"";
        }
    }
}
