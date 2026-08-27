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
