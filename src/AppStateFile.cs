using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>Общий trustworthy read / atomic write для небольших optional state files.</summary>
    internal static class AppStateFile
    {
        /// <summary>
        /// Missing — достоверный пустой снимок; существующий, но непрочитанный файл — false.
        /// </summary>
        internal static bool TryReadLines(string path, out string[] lines)
        {
            lines = new string[0];
            try
            {
                lines = File.ReadAllLines(path);
                return true;
            }
            catch (FileNotFoundException) { return true; }
            catch (DirectoryNotFoundException) { return true; }
            catch { return false; }
        }

        internal static void WriteLines(string path, IEnumerable<string> lines)
        {
            string directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);
            using (var output = new AtomicOutput(path))
            {
                File.WriteAllLines(output.TempPath, lines);
                output.Commit();
            }
        }
    }
}
