using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace ExcelMerger
{
    /// <summary>Проверка идентичности путей и race-safe публикация под свободным именем.</summary>
    public static class OutputFile
    {
        // Коды Windows в HRESULT (FACILITY_WIN32): файл держит другой процесс.
        private const int SharingViolation = unchecked((int)0x80070020);
        private const int LockViolation = unchecked((int)0x80070021);

        /// <summary>
        /// Ведут ли два пути к одному файлу. Сначала дешёвое сравнение текста полных путей;
        /// если тексты разные, а оба файла существуют, — сравнение самих файлов (том + номер
        /// файла). Одного текста мало: короткое имя 8.3 («REPORT~1.PDF»), подключённый диск
        /// против UNC-пути и жёсткая ссылка выглядят разными путями, и объединение или
        /// --compress записали бы результат поверх исходника.
        /// </summary>
        public static bool IsSameFile(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
                return false;
            string fullA, fullB;
            try
            {
                fullA = Path.GetFullPath(a);
                fullB = Path.GetFullPath(b);
            }
            catch { return false; }
            if (string.Equals(fullA, fullB, StringComparison.OrdinalIgnoreCase))
                return true;
            ulong volumeA, indexA, volumeB, indexB;
            return TryGetIdentity(fullA, out volumeA, out indexA) &&
                   TryGetIdentity(fullB, out volumeB, out indexB) &&
                   volumeA == volumeB && indexA == indexB;
        }

        /// <summary>
        /// Том и номер файла на нём. Открытие — только на чтение атрибутов и с полным
        /// разделением доступа: проверка не должна ни мешать программе, держащей файл, ни
        /// спотыкаться о неё. false — файла нет или узнать не удалось; тогда вызывающий
        /// остаётся при сравнении текста. Номер 64-битный: на ReFS он не гарантированно
        /// уникален, но ошибка тогда только в безопасную сторону — «тот же файл», то есть
        /// отказ записать, а не перезапись исходника.
        /// </summary>
        private static bool TryGetIdentity(string path, out ulong volume, out ulong index)
        {
            volume = index = 0;
            try
            {
                using (SafeFileHandle handle = CreateFileW(path, FileReadAttributes,
                    FileShare.ReadWrite | FileShare.Delete, IntPtr.Zero, FileMode.Open, 0,
                    IntPtr.Zero))
                {
                    if (handle.IsInvalid)
                        return false;
                    ByHandleFileInformation info;
                    if (!GetFileInformationByHandle(handle, out info))
                        return false;
                    volume = info.VolumeSerialNumber;
                    index = ((ulong)info.FileIndexHigh << 32) | info.FileIndexLow;
                    return true;
                }
            }
            catch { return false; }
        }

        /// <summary>
        /// Можно ли записать PDF-результат в outputPath: null — можно, иначе сообщение для
        /// человека. «Файл занят» говорится только тогда, когда его действительно держит
        /// другая программа; отсутствующая папка, нет прав, сетевой сбой называются своими
        /// словами — раньше всё это превращалось в «возможно, открыт в другой программе», и
        /// человек закрывал программы, которые были ни при чём.
        /// Проба та же, что у <see cref="MergeService.CheckOutputWritable"/>, но причина
        /// определяется по коду ошибки, а не по типу исключения: там любая IOException —
        /// «занят», а это и «нет сети», и «слишком длинный путь».
        /// </summary>
        internal static string PdfWriteProblem(string outputPath)
        {
            try
            {
                string full = Path.GetFullPath(outputPath);
                string dir = Path.GetDirectoryName(full);
                if (dir == null || !Directory.Exists(dir))
                    return string.Format(Loc.T("err.merge.folderMissing"), dir);
                if (File.Exists(full))
                {
                    // Эксклюзивное открытие: файл, открытый во вьюере, даст нарушение совместного доступа.
                    using (new FileStream(full, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
                }
                else
                {
                    using (new FileStream(full, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { }
                    File.Delete(full);
                }
                return null;
            }
            catch (IOException ex) when (IsBusy(ex))
            {
                return Loc.T("err.pdf.fileBusy");
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                return string.Format(Loc.T("err.pdf.saveFailed"), DiskSpace.Describe(ex, outputPath));
            }
        }

        /// <summary>Файл держит другой процесс (а не «нет папки/прав/места»). Чистая — под тест.</summary>
        internal static bool IsBusy(Exception ex)
        {
            return ex is IOException &&
                   (ex.HResult == SharingViolation || ex.HResult == LockViolation);
        }

        private const uint FileReadAttributes = 0x80;

        [StructLayout(LayoutKind.Sequential)]
        private struct ByHandleFileInformation
        {
            public uint FileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME CreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWriteTime;
            public uint VolumeSerialNumber;
            public uint FileSizeHigh;
            public uint FileSizeLow;
            public uint NumberOfLinks;
            public uint FileIndexHigh;
            public uint FileIndexLow;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFileW(string fileName, uint desiredAccess,
            FileShare shareMode, IntPtr securityAttributes, FileMode creationDisposition,
            uint flagsAndAttributes, IntPtr templateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandle(SafeFileHandle file,
            out ByHandleFileInformation information);

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
