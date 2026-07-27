using System;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Общий конвейер «переписать PDF через Ghostscript»: выполнить, проверить результат и
    /// заменить оригинал ТОЛЬКО если политика вызывающего это разрешила. Оригинал не
    /// теряется ни при каком исходе — он уходит в резервную копию и возвращается при сбое.
    ///
    /// Раньше эти шаги были встроены в сжатие. Их пришлось выделить, когда появились другие
    /// преобразования тем же движком (оттенки серого, восстановление): у них другая
    /// политика замены — сжатие заменяет файл, только если он стал МЕНЬШЕ, а перевод в
    /// серое или починка почти всегда дают другой размер и всё равно должны примениться.
    /// Различаются только аргументы и эта политика, всё остальное общее.
    /// </summary>
    internal static class GsRewrite
    {
        /// <summary>
        /// Прогнать файл через Ghostscript и заменить его результатом, если replace скажет
        /// «да». Возвращает true, если замена произошла. Ошибки, таймаут и негодный вывод
        /// глушатся: преобразование не должно ронять операцию, ради которой его запустили.
        /// Только с фонового потока и до открытия файла посторонними программами.
        /// </summary>
        public static bool Run(string path, string args, int timeoutMs,
            Func<long, long, bool, bool> replace)
        {
            string tmp = path + ".gstmp";
            string bak = path + ".gsbak";
            try
            {
                long origSize = new FileInfo(path).Length;
                string stderr;
                int exit = Ghostscript.Run(args, timeoutMs, out stderr);
                bool valid = EngineSucceeded(exit, stderr) && PdfCompression.LooksLikePdf(tmp);
                long newSize = valid ? new FileInfo(tmp).Length : 0L;
                if (!replace(origSize, newSize, valid))
                    return false;
                ReplaceInPlace(path, tmp, bak);
                return true;
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

        /// <summary>Путь к временному выходу для переданного файла (его же и подставит <see cref="Run"/>).</summary>
        public static string TempOutput(string path)
        {
            return path + ".gstmp";
        }

        /// <summary>
        /// Отработал ли движок на самом деле. Коду возврата верить НЕЛЬЗЯ: на файле, который
        /// он не смог прочитать (например защищённом паролем на открытие), Ghostscript выходит
        /// с нулём и всё равно оставляет годный по заголовку PDF — пустую заглушку в пару
        /// килобайт. Замер: исходник 20810 байт, выход 2522 байта, ровно одна страница, exit=0.
        /// Такая заглушка проходит и «валиден и непуст» (починка, серое), и «валиден и меньше»
        /// (сжатие), поэтому пользователь получал зелёное «Готово» и пустой документ.
        ///
        /// Признак настоящего отказа — «****» в потоке ошибок: этой строкой движок помечает
        /// свои сообщения об ошибке. Ложных срабатываний нет именно потому, что в аргументах
        /// стоит -dQUIET — на файле, который движок штатно чинит (главный риск такой проверки),
        /// поток ошибок пуст, проверено. Чистая — под тест.
        /// </summary>
        internal static bool EngineSucceeded(int exitCode, string stderr)
        {
            return exitCode == 0 &&
                   (string.IsNullOrEmpty(stderr) || stderr.IndexOf("****", StringComparison.Ordinal) < 0);
        }

        /// <summary>
        /// Безопасная замена: оригинал уводится в резервную копию, новый файл встаёт на его
        /// место, копия затем удаляется. При сбое оригинал возвращается — файл не теряется
        /// ни при каком исходе. Через переименование в той же папке: работает и на сетевых
        /// дисках, где замена одним вызовом может отказать.
        /// </summary>
        private static void ReplaceInPlace(string path, string tmp, string bak)
        {
            TryDelete(bak);
            File.Move(path, bak);       // оригинал — в сторону
            try
            {
                File.Move(tmp, path);   // новый — на место
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
    }
}
