using System;
using System.IO;
using System.Text;

namespace ExcelMerger
{
    /// <summary>Уровень сжатия PDF. Порядок = индекс в выпадающем списке.</summary>
    public enum CompressionLevel
    {
        None = 0, // «Отлично» — без сжатия (верность/подписи целы)
        Good = 1, // «Хорошо» — Ghostscript /ebook (~150 DPI)
        Small = 2 // «Нормально» — Ghostscript /screen (~72 DPI)
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

        /// <summary>Подписи для выпадающего списка (индекс = CompressionLevel).</summary>
        public static readonly string[] LevelLabels =
        {
            "Отлично — без сжатия",
            "Хорошо — меньше размер",
            "Нормально — минимальный размер"
        };

        /// <summary>Пресет Ghostscript для уровня. None сюда не передаётся (guard в Compress).</summary>
        public static string Preset(CompressionLevel level)
        {
            switch (level)
            {
                case CompressionLevel.Good: return "/ebook";
                case CompressionLevel.Small: return "/screen";
                default: return null;
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
        /// Заменять ли оригинал сжатым: только если он валиден, непуст и строго
        /// меньше исходного (уже оптимизированный PDF GS может раздуть — оставляем оригинал).
        /// Чистая — под тест.
        /// </summary>
        public static bool ShouldReplace(long origSize, long newSize, bool newValid)
        {
            return newValid && newSize > 0 && newSize < origSize;
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
            string tmp = path + ".gstmp";
            string bak = path + ".gsbak";
            try
            {
                long origSize = new FileInfo(path).Length;
                string args = BuildArguments(path, tmp, level, Ghostscript.BundledRoot);
                string stderr;
                int exit = Ghostscript.Run(args, TimeoutMs, out stderr);
                bool valid = exit == 0 && LooksLikePdf(tmp);
                long newSize = valid ? new FileInfo(tmp).Length : 0L;
                if (ShouldReplace(origSize, newSize, valid))
                {
                    ReplaceInPlace(path, tmp, bak);
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                TryDelete(tmp);
                TryDelete(bak);
            }
        }

        /// <summary>
        /// Безопасная замена: оригинал уводится в .gsbak, сжатый встаёт на его место,
        /// бэкап затем удаляется (в finally). При сбое оригинал восстанавливается —
        /// файл не теряется ни при каком исходе. Через File.Move (переименование в той
        /// же папке): работает и на сетевых дисках, где File.Replace может отказать.
        /// </summary>
        private static void ReplaceInPlace(string path, string tmp, string bak)
        {
            TryDelete(bak);
            File.Move(path, bak);       // оригинал — в сторону
            try
            {
                File.Move(tmp, path);   // сжатый — на место
            }
            catch
            {
                if (!File.Exists(path)) // откат: вернуть оригинал
                    File.Move(bak, path);
                throw;
            }
        }

        private static void TryDelete(string p)
        {
            try { if (File.Exists(p)) File.Delete(p); }
            catch { }
        }

        /// <summary>Первые байты — заголовок «%PDF-». Дешёвая проверка валидности вывода GS.</summary>
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
