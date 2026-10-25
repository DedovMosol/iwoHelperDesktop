using System;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelMerger
{
    /// <summary>
    /// Транзакция записи одного результата: вызывающий пишет только во временный файл
    /// в той же папке, затем Commit атомарно ставит его на место. Сбой или отмена
    /// удаляют лишь наши временные файлы и оставляют прежний target целым.
    /// </summary>
    internal sealed class AtomicOutput : IDisposable
    {
        private static readonly Regex MarkerName = new Regex(
            @"^(?<target>.+)\.iwo-(?<id>[0-9a-fA-F]{32})\.txn$",
            RegexOptions.Compiled);
        private static readonly Regex TempName = new Regex(
            @"^(?<stem>.+)\.iwo-(?<id>[0-9a-fA-F]{32})(?<ext>\.[^.]+)$",
            RegexOptions.Compiled);

        public readonly string TargetPath;
        public readonly string TempPath;
        private readonly string _backupPath;
        private readonly string _markerPath;
        private FileStream _marker;
        private bool _committed;
        private bool _disposed;

        public AtomicOutput(string targetPath)
        {
            if (string.IsNullOrWhiteSpace(targetPath))
                throw new ArgumentException("нет пути результата", "targetPath");
            TargetPath = Path.GetFullPath(targetPath);
            string dir = Path.GetDirectoryName(TargetPath);
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
                throw new DirectoryNotFoundException(dir);
            string id = Guid.NewGuid().ToString("N");
            string suffix = ".iwo-" + id;
            string name = Path.GetFileNameWithoutExtension(TargetPath);
            string extension = Path.GetExtension(TargetPath);
            TempPath = Path.Combine(dir, name + suffix + extension);
            // Полное имя target сохраняется перед суффиксом. После аварии StartupSweep
            // однозначно восстановит result.pdf из result.pdf.iwo-<guid>.bak.
            _backupPath = Path.Combine(dir, Path.GetFileName(TargetPath) + suffix + ".bak");
            _markerPath = Path.Combine(dir, Path.GetFileName(TargetPath) + suffix + ".txn");
            ThrowIfDerivedTooLong(TargetPath, Math.Max(TempPath.Length,
                Math.Max(_backupPath.Length, _markerPath.Length)));

            // Пустой marker одновременно задаёт схему путей своим именем и отличает живую
            // транзакцию открытым handle. Полные пути из имени выводит этот же класс.
            try
            {
                _marker = new FileStream(_markerPath, FileMode.CreateNew, FileAccess.ReadWrite,
                    FileShare.Read);
            }
            catch
            {
                try { if (_marker != null) _marker.Dispose(); } catch { }
                _marker = null;
                DeleteQuietly(_markerPath);
                throw;
            }
        }

        public void Commit()
        {
            if (_committed) return;
            if (!File.Exists(TempPath))
                throw new FileNotFoundException(TempPath);
            FlushToDisk(TempPath);
            if (!File.Exists(TargetPath))
            {
                File.Move(TempPath, TargetPath);
                _committed = true;
                return;
            }

            // File.Replace — атомарная операция NTFS. На носителе без её поддержки
            // остаётся тот же safe rename с rollback, всё в одной папке.
            try
            {
                File.Replace(TempPath, TargetPath, _backupPath, true);
                _committed = true;
                DeleteQuietly(_backupPath);
                return;
            }
            catch (PlatformNotSupportedException) { }
            catch (IOException) { }

            // Только fallback создаёт окно без target: перед ним marker обязан попасть на диск.
            _marker.Flush(true);
            File.Move(TargetPath, _backupPath);
            try
            {
                File.Move(TempPath, TargetPath);
                _committed = true;
                DeleteQuietly(_backupPath);
            }
            catch
            {
                if (!File.Exists(TargetPath) && File.Exists(_backupPath))
                    File.Move(_backupPath, TargetPath);
                throw;
            }
        }

        /// <summary>
        /// Данные результата — на диск ДО переименования. NTFS журналирует метаданные, но не
        /// содержимое: переименование может пережить сбой питания, а записанные байты остаться
        /// в кэше, и тогда на месте прежнего файла оказывался бы файл из нулей, а прежнюю версию
        /// сразу удаляет следующая строка. Это классическая последовательность «записать,
        /// FlushFileBuffers, переименовать»; флаг REPLACEFILE_WRITE_THROUGH у ReplaceFile
        /// документирован как неподдерживаемый, поэтому сброс делаем сами.
        ///
        /// Сброс — улучшение надёжности, а не условие записи: нет права записи (файл помечен
        /// «только чтение»), его держит другой процесс без общего доступа на запись, носитель не
        /// поддерживает операцию — публикуем, как прежде, а не теряем готовый результат.
        /// </summary>
        private static void FlushToDisk(string path)
        {
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Write,
                    FileShare.ReadWrite | FileShare.Delete))
                    stream.Flush(true); // FlushFileBuffers: сбрасывает кэш ФАЙЛА, кто бы его ни писал
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            catch (NotSupportedException) { }
        }

        public void Dispose()
        {
            if (_disposed)
                return;
            _disposed = true;
            DeleteQuietly(TempPath);
            bool unresolved = false;
            if (_committed)
            {
                // После успешной публикации backup — только мусор. При неуспешной он
                // может быть единственной копией исходника и ниже НИКОГДА не удаляется.
                if (File.Exists(TargetPath))
                    DeleteQuietly(_backupPath);
                else if (File.Exists(_backupPath))
                    unresolved = true;
            }
            else if (File.Exists(_backupPath) && !File.Exists(TargetPath))
            {
                try { File.Move(_backupPath, TargetPath); }
                catch
                {
                    // Последняя попытка вернуть файл не удалась. Backup и journal остаются
                    // рядом: StartupSweep повторит восстановление, не уничтожая данные.
                    unresolved = true;
                }
            }

            try { if (_marker != null) _marker.Dispose(); } catch { }
            _marker = null;
            if (!unresolved)
                DeleteQuietly(_markerPath);
        }

        /// <summary>
        /// MAX_PATH вместе с завершающим нулём: без longPathAware в манифесте Win32 не создаёт
        /// файл с путём длиннее 259 символов.
        /// </summary>
        internal const int MaxPath = 260;

        /// <summary>
        /// Служебные файлы длиннее цели на суффикс «.iwo-&lt;guid&gt;.txn» (41 символ), поэтому
        /// цель в 219–259 символов допустима сама, а журнал рядом с ней — нет. Без этой проверки
        /// человек получал «не найдена часть пути» про файл, которого никогда не видел. Отказ —
        /// ДО создания чего-либо на диске и с длиной, которую можно сократить. Путь «\\?\» под
        /// лимит не попадает — его не проверяем.
        /// </summary>
        private static void ThrowIfDerivedTooLong(string targetPath, int longestDerived)
        {
            if (longestDerived < MaxPath || targetPath.StartsWith(@"\\?\", StringComparison.Ordinal))
                return;
            int allowed = MaxPath - 1 - (longestDerived - targetPath.Length);
            throw new MergeException(string.Format(Loc.T("err.output.pathTooLong"),
                targetPath.Length, allowed));
        }

        internal static bool TryGetRecoveryPaths(string markerPath, out string target,
            out string temp, out string backup)
        {
            target = temp = backup = null;
            try
            {
                string directory = Path.GetDirectoryName(Path.GetFullPath(markerPath));
                Match match = MarkerName.Match(Path.GetFileName(markerPath));
                if (!match.Success || string.IsNullOrEmpty(directory))
                    return false;
                string targetName = match.Groups["target"].Value;
                string id = match.Groups["id"].Value;
                target = Path.Combine(directory, targetName);
                temp = Path.Combine(directory, Path.GetFileNameWithoutExtension(targetName) +
                    ".iwo-" + id + Path.GetExtension(targetName));
                backup = Path.Combine(directory, targetName + ".iwo-" + id + ".bak");
                return true;
            }
            catch { return false; }
        }

        internal static bool TryGetMarkerForTemp(string tempPath, out string markerPath)
        {
            markerPath = null;
            try
            {
                string full = Path.GetFullPath(tempPath);
                Match match = TempName.Match(Path.GetFileName(full));
                if (!match.Success)
                    return false;
                markerPath = Path.Combine(Path.GetDirectoryName(full),
                    match.Groups["stem"].Value + match.Groups["ext"].Value +
                    ".iwo-" + match.Groups["id"].Value + ".txn");
                return true;
            }
            catch { return false; }
        }

        private static void DeleteQuietly(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch { }
        }
    }
}
