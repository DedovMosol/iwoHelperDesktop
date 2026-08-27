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
    internal enum GsRewriteResult
    {
        Applied,
        RejectedByPolicy,
        Failed
    }

    internal static class GsRewrite
    {
        /// <summary>
        /// Прогнать файл через Ghostscript и заменить его результатом, если тот уцелел как
        /// документ и replace скажет «да» про его размер. Возвращает true, если замена
        /// произошла. Ошибки, таймаут и негодный вывод глушатся: преобразование не должно
        /// ронять операцию, ради которой его запустили.
        /// Только с фонового потока и до открытия файла посторонними программами.
        ///
        /// Проверок четыре, и порядок в них не косметический — от дешёвой к дорогой, каждая
        /// следующая выполняется только для того, что прошло предыдущие:
        /// 1) движок отработал (<see cref="EngineSucceeded"/>);
        /// 2) вывод открывается и в нём столько же страниц (<see cref="PdfPageProbe.PagesKept"/>) —
        ///    инвариант конвейера, общий всем режимам: что бы мы ни делали с документом,
        ///    документом он остаться обязан;
        /// 3) replace — политика вызывающего по размеру;
        /// 4) verify — необязательная проверка вывода по существу (перевод в серое смотрит
        ///    ею, что цвет действительно ушёл). Она последняя, потому что самая дорогая, и
        ///    незачем платить за неё, если результат всё равно не применяется.
        /// </summary>
        public static bool Run(string path, string tempOutput, string args, int timeoutMs,
            Func<long, long, bool> replace, Func<string, bool> verify = null)
        {
            return RunDetailed(path, tempOutput, args, timeoutMs, replace, verify) ==
                GsRewriteResult.Applied;
        }

        internal static GsRewriteResult RunDetailed(string path, string tempOutput,
            string args, int timeoutMs, Func<long, long, bool> replace,
            Func<string, bool> verify = null)
        {
            return RunCore(path, tempOutput, args, timeoutMs, replace, verify, true);
        }

        internal static GsRewriteResult RunDetailedUnpublished(string path,
            string tempOutput, string args, int timeoutMs, Func<long, long, bool> replace,
            Func<string, bool> verify = null)
        {
            return RunCore(path, tempOutput, args, timeoutMs, replace, verify, false);
        }

        private static GsRewriteResult RunCore(string path, string tempOutput,
            string args, int timeoutMs, Func<long, long, bool> replace,
            Func<string, bool> verify, bool recoverableTarget)
        {
            string tmp = tempOutput;
            try
            {
                long origSize = new FileInfo(path).Length;
                // Страницы исходника считаем ДО запуска движка, а не перед самой подменой.
                // Дело не в скорости: чтение открывает файл, а следом идёт File.Move этого же
                // файла — на Windows такое соседство упирается в антивирус, который берётся
                // за файл сразу после закрытия хэндла, и подмена изредка отказывает. Между
                // этим чтением и подменой теперь стоит вся работа движка, и окно закрыто.
                int pagesBefore = PdfPageProbe.PageCount(path);
                string stderr;
                int exit = Ghostscript.Run(args, timeoutMs, out stderr);
                if (!EngineSucceeded(exit, stderr))
                    return GsRewriteResult.Failed;
                if (!PdfPageProbe.PagesKept(pagesBefore, PdfPageProbe.PageCount(tmp)))
                    return GsRewriteResult.Failed;
                if (!replace(origSize, new FileInfo(tmp).Length))
                    return GsRewriteResult.RejectedByPolicy;
                if (verify != null && !verify(tmp))
                    return GsRewriteResult.Failed;
                if (recoverableTarget)
                {
                    using (var output = new AtomicOutput(path))
                    {
                        File.Move(tmp, output.TempPath);
                        tmp = null;
                        output.Commit();
                    }
                }
                else
                {
                    ReplaceUnpublished(path, tmp);
                    tmp = null;
                }
                return GsRewriteResult.Applied;
            }
            catch
            {
                return GsRewriteResult.Failed;
            }
            finally
            {
                TryDelete(tmp);
            }
        }

        private static void ReplaceUnpublished(string path, string replacement)
        {
            try
            {
                File.Replace(replacement, path, null, true);
            }
            catch (PlatformNotSupportedException)
            {
                File.Delete(path);
                File.Move(replacement, path);
            }
            catch (IOException)
            {
                File.Delete(path);
                File.Move(replacement, path);
            }
        }

        /// <summary>Путь к временному выходу; уникален для параллельных headless-запусков.</summary>
        public static string TempOutput(string path)
        {
            return path + ".iwo-gs-" + Guid.NewGuid().ToString("N") + ".tmp";
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

        private static void TryDelete(string p)
        {
            try { if (File.Exists(p)) File.Delete(p); }
            catch { }
        }
    }
}
