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
                // Одной стратегии достаточно, и добавлять к ней модель устройства НЕ НАДО.
                // Документация pdfwrite говорит прямо: с версии 9.11 ColorConversionStrategy
                // сама выставляет ProcessColorModel, и задавать оба переключателя не следует;
                // на эту же связку заведён баг 693074 — «часть изображений остаётся цветной».
                // Раньше здесь стоял -dProcessColorModel=/DeviceGray с комментарием «обе части
                // обязательны»: это знание осталось от времён до 9.11. Проверено на восьми
                // пробах (вектор, цветной JPEG, CMYK, индексированная палитра, Separation,
                // группа прозрачности, SMask, реальная прозрачность): без модели устройства
                // все они дают чистое серое, и внутри файла остаётся только DeviceGray.
                sb.Append(" -sColorConversionStrategy=Gray");
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
        /// Заменять ли оригинал результатом преобразования: любой непустой вывод. Размер
        /// здесь не показатель — в отличие от сжатия, эти режимы меняют файл по существу, а
        /// не ради экономии места, и результат обязан примениться, даже если стал больше.
        /// Что вывод уцелел как документ, к этому моменту уже проверил конвейер
        /// (<see cref="GsRewrite.Run"/>). Чистая — под тест.
        /// </summary>
        public static bool ShouldReplace(long origSize, long newSize)
        {
            return newSize > 0;
        }

        /// <summary>
        /// Годен ли вывод режима по существу. Для серого это единственная честная проверка:
        /// движок сообщает об успехе, даже когда часть содержимого осталась цветной (цвет
        /// приходит из ICC-профилей, палитр, групп прозрачности — заранее их не перечислить),
        /// и без взгляда на результат человек получал бы зелёное «Готово» на цветном файле.
        /// Починке проверять нечего сверх того, что документ открылся, — это уже сделано.
        /// </summary>
        internal static bool Verify(string produced, PdfConvertMode mode)
        {
            return mode != PdfConvertMode.Grayscale || !PdfColorProbe.HasColorPages(produced);
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
            string temp = GsRewrite.TempOutput(path);
            string args = BuildArguments(path, temp, mode, Ghostscript.BundledRoot);
            return GsRewrite.Run(path, temp, args, TimeoutMs, ShouldReplace,
                delegate(string produced) { return Verify(produced, mode); });
        }

        private static string Quote(string s)
        {
            return "\"" + s + "\"";
        }
    }
}
