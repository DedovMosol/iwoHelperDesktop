using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace ExcelMerger
{
    /// <summary>
    /// Межпроцессная блокировка read-modify-write для одного файла состояния. Файловый
    /// lock лежит рядом с данными и действует между Windows-сеансами одного профиля.
    /// При таймауте/неизвестном снимке действие не публикуется.
    /// </summary>
    internal static class AppDataLock
    {
        private const int WaitMilliseconds = 2000;
        private const int RetryMilliseconds = 25;

        public static bool TryRun(string stateFile, Func<bool> action)
        {
            if (string.IsNullOrWhiteSpace(stateFile) || action == null)
                return false;
            string lockPath;
            try
            {
                lockPath = Path.GetFullPath(stateFile) + ".lock";
                string directory = Path.GetDirectoryName(lockPath);
                if (!string.IsNullOrEmpty(directory))
                    Directory.CreateDirectory(directory);
            }
            catch { return false; }

            var elapsed = Stopwatch.StartNew();
            while (true)
            {
                try
                {
                    using (new FileStream(lockPath, FileMode.OpenOrCreate,
                        FileAccess.ReadWrite, FileShare.None, 1, FileOptions.None))
                        return action();
                }
                catch (IOException ex)
                {
                    if (!IsSharingViolation(ex) ||
                        elapsed.ElapsedMilliseconds >= WaitMilliseconds)
                        return false;
                }
                catch
                {
                    // Optional state must not abort the document operation or crash the UI.
                    return false;
                }
                Thread.Sleep(RetryMilliseconds);
            }
        }

        private static bool IsSharingViolation(IOException error)
        {
            int code = error.HResult & 0xFFFF;
            return code == 32 || code == 33; // ERROR_SHARING_VIOLATION / ERROR_LOCK_VIOLATION
        }
    }
}
