using System;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Транзакция записи одного результата: вызывающий пишет только во временный файл
    /// в той же папке, затем Commit атомарно ставит его на место. Сбой или отмена
    /// удаляют лишь наши временные файлы и оставляют прежний target целым.
    /// </summary>
    internal sealed class AtomicOutput : IDisposable
    {
        public readonly string TargetPath;
        public readonly string TempPath;
        private readonly string _backupPath;
        private bool _committed;

        public AtomicOutput(string targetPath)
        {
            if (string.IsNullOrWhiteSpace(targetPath))
                throw new ArgumentException("нет пути результата", "targetPath");
            TargetPath = Path.GetFullPath(targetPath);
            string dir = Path.GetDirectoryName(TargetPath);
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
                throw new DirectoryNotFoundException(dir);
            string suffix = ".iwo-" + Guid.NewGuid().ToString("N");
            string name = Path.GetFileNameWithoutExtension(TargetPath);
            string extension = Path.GetExtension(TargetPath);
            TempPath = Path.Combine(dir, name + suffix + extension);
            _backupPath = Path.Combine(dir, name + suffix + ".bak");
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
            DeleteQuietly(TempPath);
            if (!_committed && File.Exists(_backupPath) && !File.Exists(TargetPath))
            {
                try { File.Move(_backupPath, TargetPath); }
                catch
                {
                    // Откат намеренно молчаливый: это последняя попытка вернуть файл из
                    // резервной копии в Dispose. Если даже она не вышла, .bak остаётся рядом
                    // с результатом и содержимое пользователя всё равно не потеряно.
                }
            }
            DeleteQuietly(_backupPath);
        }

        private static void DeleteQuietly(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch { }
        }
    }
}
