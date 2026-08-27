using System;
using System.IO;

namespace ExcelMerger
{
    /// <summary>Проверка идентичности путей и race-safe публикация под свободным именем.</summary>
    public static class OutputFile
    {
        public static bool IsSameFile(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
                return false;
            try
            {
                return string.Equals(Path.GetFullPath(a), Path.GetFullPath(b),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        internal static UniqueOutput CreateUnique(string dir, string safeName,
            string extension)
        {
            return new UniqueOutput(dir, safeName, extension);
        }

        internal static string Candidate(string dir, string safeName, string extension,
            int number)
        {
            return Path.Combine(dir, safeName + (number <= 1 ? "" : "_" + number) +
                extension);
        }
    }

    /// <summary>
    /// Новый результат без placeholder/check-then-create: запись идёт в уникальный temp,
    /// Commit публикует rename-ом без overwrite и при коллизии пробует следующий суффикс.
    /// </summary>
    internal sealed class UniqueOutput : IDisposable
    {
        private readonly string _directory;
        private readonly string _safeName;
        private readonly string _extension;
        private bool _committed;

        internal readonly string TempPath;
        internal string TargetPath { get; private set; }

        internal UniqueOutput(string directory, string safeName, string extension)
        {
            _directory = Path.GetFullPath(directory);
            _safeName = safeName;
            _extension = extension;
            Directory.CreateDirectory(_directory);
            TempPath = Path.Combine(_directory, safeName + ".iwo-" +
                Guid.NewGuid().ToString("N") + extension);
        }

        internal string Commit()
        {
            if (_committed)
                return TargetPath;
            if (!File.Exists(TempPath))
                throw new FileNotFoundException(TempPath);
            for (int number = 1; ; number++)
            {
                string candidate = OutputFile.Candidate(_directory, _safeName,
                    _extension, number);
                try
                {
                    File.Move(TempPath, candidate); // never overwrites on .NET Framework
                    TargetPath = candidate;
                    _committed = true;
                    return candidate;
                }
                catch (IOException)
                {
                    if (!File.Exists(candidate))
                        throw;
                }
            }
        }

        public void Dispose()
        {
            if (_committed)
                return;
            try { if (File.Exists(TempPath)) File.Delete(TempPath); } catch { }
        }
    }
}
