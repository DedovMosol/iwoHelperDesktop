using System;
using System.IO;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Уровень сжатия PDF. ЧИСЛА — идентификаторы, а не позиции: они уходят в settings.txt и
    /// возвращаются оттуда в следующей версии программы. Поэтому новый уровень добавляется
    /// В КОНЕЦ, даже если по смыслу стоит в середине, а порядок показа живёт отдельно
    /// (<see cref="PdfCompression.DisplayOrder"/>). Иначе у всех, кто выбрал «Хорошо»,
    /// после обновления молча оказался бы выбран другой уровень.
    /// </summary>
    public enum CompressionLevel
    {
        None = 0,    // «Отлично» — без сжатия (файл байт в байт, подписи целы)
        Good = 1,    // «Хорошо» — Ghostscript /ebook (~150 dpi)
        Small = 2,   // «Нормально» — Ghostscript /screen (~72 dpi)
        VeryGood = 3 // «Очень хорошо» — пересборка без пересчёта изображений (чёткость цела)
    }

    /// <summary>
    /// Пост-обработка готового PDF: сжатие «как в Acrobat» через Ghostscript
    /// (downsampling изображений, текст/вектор сохраняются). Дефолт — без сжатия.
    /// Чистые функции (Preset/BuildArguments/ShouldReplace) покрыты тестами;
    /// Compress вызывается ТОЛЬКО с фонового потока и до открытия файла.
    /// </summary>
    public static class PdfCompression
    {
        /// <summary>Максимальное время работы GS на один файл.</summary>
        private const int TimeoutMs = 300000; // 5 минут — верхняя граница на большой скан

        /// <summary>
        /// Уровни в порядке показа — от «ничего не трогаем» к «жмём сильнее всего».
        /// Единственное место, где этот порядок задан: и список в окне, и все переводы
        /// «индекс ↔ уровень» берут его отсюда, поэтому разойтись им негде.
        /// </summary>
        private static readonly CompressionLevel[] Order =
        {
            CompressionLevel.None,
            CompressionLevel.VeryGood,
            CompressionLevel.Good,
            CompressionLevel.Small
        };

        /// <summary>Копия порядка показа (наружу отдаём копию — массив изменяем только здесь).</summary>
        public static CompressionLevel[] DisplayOrder()
        {
            return (CompressionLevel[])Order.Clone();
        }

        /// <summary>Уровень по позиции в списке. Вне диапазона — «без сжатия». Чистая — под тест.</summary>
        public static CompressionLevel LevelAt(int index)
        {
            return index < 0 || index >= Order.Length ? CompressionLevel.None : Order[index];
        }

        /// <summary>Позиция уровня в списке. Неизвестный — позиция «без сжатия». Чистая — под тест.</summary>
        public static int IndexOf(CompressionLevel level)
        {
            for (int i = 0; i < Order.Length; i++)
                if (Order[i] == level)
                    return i;
            return 0;
        }

        /// <summary>
        /// Подпись уровня на текущем языке. Разрешение подставляется из <see cref="ImageDpi"/> —
        /// единственного места, где эти числа объявлены, поэтому подпись не может разойтись
        /// с тем, что делает движок.
        /// </summary>
        public static string Label(CompressionLevel level)
        {
            switch (level)
            {
                case CompressionLevel.VeryGood: return Loc.T("compress.level.veryGood");
                case CompressionLevel.Good:
                case CompressionLevel.Small:
                    return string.Format(Loc.T(level == CompressionLevel.Good
                        ? "compress.level.good" : "compress.level.small"), ImageDpi(level));
                default: return Loc.T("compress.level.none");
            }
        }

        /// <summary>Подписи для выпадающего списка в порядке показа. Метод, а не поле:
        /// читаются на текущем языке при каждом вызове (переключение языка пересоздаёт список).</summary>
        public static string[] LevelLabels()
        {
            var labels = new string[Order.Length];
            for (int i = 0; i < Order.Length; i++)
                labels[i] = Label(Order[i]);
            return labels;
        }

        /// <summary>Пресет Ghostscript для уровня. None сюда не передаётся (guard в Compress).</summary>
        public static string Preset(CompressionLevel level)
        {
            switch (level)
            {
                case CompressionLevel.Good: return "/ebook";
                case CompressionLevel.Small: return "/screen";
                // Пресет по умолчанию изображения не пересчитывает — вся экономия идёт от
                // пересборки документа: страницы пережимаются заново, одинаковые картинки
                // хранятся один раз, мусор не переносится. Замер на четырёх документах:
                // от −25% до −48% при нетронутой чёткости, а на файле из одинаковых
                // картинок — в 136 раз (это уже дедупликация, она включена по умолчанию).
                case CompressionLevel.VeryGood: return "/default";
                default: return null;
            }
        }

        /// <summary>Пересчитывает ли уровень изображения (даунсэмплинг). Чистая — под тест.</summary>
        public static bool Downsamples(CompressionLevel level)
        {
            return ImageDpi(level) > 0;
        }

        /// <summary>
        /// Часть статуса о выполненном сжатии: с разрешением там, где изображения
        /// пересчитаны, и без него там, где они не тронуты. Одно место на все три
        /// инструмента — иначе новый уровень отчитывался бы «до 0 dpi».
        /// </summary>
        public static string CompressedSuffix(CompressionLevel level)
        {
            return Downsamples(level)
                ? string.Format(Loc.T("common.suffix.compressed"), ImageDpi(level))
                : Loc.T("common.suffix.compressedKeep");
        }

        /// <summary>
        /// До какого разрешения пресет уменьшает цветные и серые изображения, dpi.
        /// Значения заданы самим Ghostscript в Resource\Init\gs_pdfwr.ps: /ebook — 150,
        /// /screen — 72. Штриховые (1-битные) изображения оба пресета оставляют на 300 dpi,
        /// а текст и вектор не растрируются вовсе. Единственное место, где живут эти числа,
        /// поэтому статусы инструментов берут их отсюда. 0 — уровень без сжатия.
        /// Чистая — под тест.
        /// </summary>
        public static int ImageDpi(CompressionLevel level)
        {
            switch (level)
            {
                case CompressionLevel.Good: return 150;
                case CompressionLevel.Small: return 72;
                default: return 0;
            }
        }

        /// <summary>
        /// Аргументы командной строки Ghostscript. Все пути в кавычках (пробелы!).
        /// bundledGsRoot != null (вшитый GS) добавляет -I на его lib/Resource\Init.
        /// Чистая — под тест. -dSAFER безопасен: он не блокирует чтение входа/записи
        /// выхода, переданных в командной строке (GS 10.x).
        /// </summary>
        public static string BuildArguments(string input, string output, CompressionLevel level, string bundledGsRoot)
        {
            var sb = new StringBuilder();
            // 1.4 (а не 1.5): классическая таблица xref без object streams — сжатый файл
            // остаётся читаемым нашим PdfSharp 1.50 (повторное объединение/разделение) и
            // всеми вьюерами. Downsampling изображений от уровня совместимости не зависит.
            sb.Append("-sDEVICE=pdfwrite -dCompatibilityLevel=1.4");
            sb.Append(" -dPDFSETTINGS=").Append(Preset(level));
            if (Preset(level) != null && !Downsamples(level))
            {
                // Обещание уровня — «чёткость цела», и держать его должен наш код, а не
                // чужие умолчания: пресет /default сегодня не пересчитывает изображения, но
                // портативная сборка работает с ТЕМ Ghostscript, который стоит у человека,
                // и ручаться за умолчания всех его версий мы не можем. Три ключа стоят
                // дёшево и делают обещание независимым от версии движка.
                sb.Append(" -dDownsampleColorImages=false -dDownsampleGrayImages=false");
                sb.Append(" -dDownsampleMonoImages=false");
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
        /// Заменять ли оригинал сжатым: только если тот строго меньше исходного (уже
        /// оптимизированный PDF движок может раздуть — тогда оставляем оригинал). Что вывод
        /// вообще годен, к этому моменту уже проверил конвейер (<see cref="GsRewrite.Run"/>),
        /// поэтому здесь остался ровно один вопрос — про размер. Чистая — под тест.
        /// </summary>
        public static bool ShouldReplace(long origSize, long newSize)
        {
            return newSize > 0 && newSize < origSize;
        }

        /// <summary>
        /// Сжимает PDF на месте. Возвращает true, если оригинал заменён меньшей копией.
        /// None или отсутствие Ghostscript → false (без изменений). Любые ошибки
        /// глушатся: сжатие опционально и не должно ронять объединение/разделение.
        /// ВНИМАНИЕ: только с фонового потока и до открытия/показа файла.
        /// </summary>
        public static bool Compress(string path, CompressionLevel level)
        {
            if (level == CompressionLevel.None || !Ghostscript.Available)
                return false;
            string args = BuildArguments(path, GsRewrite.TempOutput(path), level, Ghostscript.BundledRoot);
            return GsRewrite.Run(path, args, TimeoutMs, ShouldReplace);
        }

        /// <summary>
        /// Первые байты — заголовок «%PDF-». Дешёвая проба для самопроверки сборки
        /// (<c>--gscheck</c>): там нужно лишь убедиться, что движок вообще выдал PDF.
        /// Конвейер преобразований ею не пользуется — он проверяет вывод по существу,
        /// пересчитывая страницы (<see cref="PdfPageProbe"/>).
        /// </summary>
        internal static bool LooksLikePdf(string path)
        {
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var head = new byte[5];
                    if (fs.Read(head, 0, 5) != 5)
                        return false;
                    return head[0] == 0x25 && head[1] == 0x50 && head[2] == 0x44 &&
                           head[3] == 0x46 && head[4] == 0x2D; // %PDF-
                }
            }
            catch { return false; }
        }

        private static string Quote(string s)
        {
            return "\"" + s + "\"";
        }
    }
}
