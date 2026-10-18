using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelMerger
{
    /// <summary>
    /// Восстановление и уборка осиротевших файлов «.iwo-*». Живую транзакцию отличает
    /// открытый exclusive-journal; после аварии backup сначала возвращается на место и
    /// только заведомо лишние temp/backup удаляются. Единственная копия документа мусором
    /// не считается никогда.
    /// </summary>
    internal static class StartupSweep
    {
        internal static readonly TimeSpan MinAge = TimeSpan.FromHours(1);

        private static readonly Regex OrphanName = new Regex(
            @"\.iwo-(?:gs-)?[0-9a-fA-F]{32}(\.[^.]*)?$",
            RegexOptions.Compiled);
        private static readonly Regex GsBackup = new Regex(
            @"^(?<target>.+)\.iwo-gs-(?<id>[0-9a-fA-F]{32})\.bak$",
            RegexOptions.Compiled);

        // Копии PDF во %TEMP%, которые могут быть РАСШИФРОВАННЫМИ: страница защищённого файла
        // для Ghostscript (PageRasterizer.RenderPage), собранный документ PdfOpsForm и папка
        // подложек PageBackgrounds (внутри — decrypted.pdf из RenderPagesWithoutText). Свои
        // finally их удаляют, но аварийный выход оставлял бы открытый текст защищённого
        // документа в общей папке навсегда. Имена — ровно те, что создаёт код, чужое не трогаем.
        private static readonly Regex TempCopyFile = new Regex(
            @"^iwo_(?:pg_dec|ops)_[0-9a-fA-F]{32}\.pdf$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex TempCopyDir = new Regex(
            @"^iwo_bg_[0-9a-fA-F]{32}$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Восстановить прерванные транзакции и удалить старые временные файлы. Возвращает
        /// число удалённых файлов (восстановленный backup удалением не считается).
        /// </summary>
        internal static int Sweep(IEnumerable<string> directories, DateTime nowUtc)
        {
            if (directories == null)
                return 0;
            int removed = 0;
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string value in directories)
            {
                string dir;
                try { dir = string.IsNullOrWhiteSpace(value) ? null : Path.GetFullPath(value); }
                catch { continue; }
                if (dir == null || !seen.Add(dir))
                    continue;
                try
                {
                    if (!Directory.Exists(dir))
                        continue;
                    var files = new List<string>(Directory.EnumerateFiles(dir));
                    removed += RecoverJournals(files);
                    removed += RecoverLooseBackups(files, nowUtc);
                    removed += DeleteOldTemps(files, nowUtc);
                }
                catch
                {
                    // Недоступная папка/сеть/ACL не должны задержать или сорвать запуск.
                }
            }
            return removed;
        }

        private static int RecoverJournals(IEnumerable<string> files)
        {
            int removed = 0;
            foreach (string marker in files)
            {
                if (!marker.EndsWith(".txn", StringComparison.OrdinalIgnoreCase))
                    continue;
                string target, temp, backup;
                if (!AtomicOutput.TryGetRecoveryPaths(marker, out target, out temp, out backup))
                    continue;
                FileStream held = null;
                bool resolved = false;
                try
                {
                    // Живой AtomicOutput держит marker с несовместимым sharing mode.
                    held = new FileStream(marker, FileMode.Open, FileAccess.ReadWrite,
                        FileShare.None);
                    if (!File.Exists(target) && File.Exists(backup))
                        File.Move(backup, target);
                    if (File.Exists(target))
                    {
                        removed += DeleteOne(temp);
                        removed += DeleteOne(backup);
                    }
                    else
                    {
                        // Target до операции не существовал, backup нет: незавершённый temp
                        // не публикуем — вызывающий не успел подтвердить Commit.
                        removed += DeleteOne(temp);
                    }
                    resolved = true;
                }
                // Живой журнал (sharing violation), нет прав, странное имя — любой сбой
                // касается только этой транзакции; остальные разбираем дальше.
                catch { }
                finally
                {
                    if (held != null)
                        try { held.Dispose(); } catch { }
                }
                // Сбой удаления ОДНОГО журнала (занят антивирусом, нет прав) раньше вылетал из
                // цикла и обрывал уборку всей папки: остальные журналы не восстанавливались.
                // Журнал уже разобран, его повторный разбор при следующем запуске безопасен.
                if (resolved)
                    removed += TryDeleteOne(marker);
            }
            return removed;
        }

        private static int TryDeleteOne(string path)
        {
            try { return DeleteOne(path); }
            catch { return 0; }
        }

        private static int RecoverLooseBackups(IEnumerable<string> files, DateTime nowUtc)
        {
            int removed = 0;
            foreach (string file in files)
            {
                if (!file.EndsWith(".bak", StringComparison.OrdinalIgnoreCase))
                    continue;
                Match match = GsBackup.Match(Path.GetFileName(file));
                if (!match.Success)
                {
                    // AtomicOutput до recovery-marker терял расширение target в имени
                    // backup; без marker старый и новый форматы неотличимы. Не угадываем.
                    continue;
                }
                string target = Path.Combine(Path.GetDirectoryName(file),
                    match.Groups["target"].Value);
                try
                {
                    if (!File.Exists(target))
                    {
                        File.Move(file, target);
                        continue;
                    }
                    if (IsOld(file, nowUtc))
                        removed += DeleteOne(file);
                }
                catch { }
            }
            return removed;
        }

        private static int DeleteOldTemps(IEnumerable<string> files, DateTime nowUtc)
        {
            int removed = 0;
            foreach (string file in files)
            {
                string name = Path.GetFileName(file);
                if (!IsOrphanName(name) || name.EndsWith(".bak", StringComparison.OrdinalIgnoreCase) ||
                    name.EndsWith(".txn", StringComparison.OrdinalIgnoreCase))
                    continue;
                string marker;
                if (AtomicOutput.TryGetMarkerForTemp(file, out marker) && File.Exists(marker))
                    continue;
                try
                {
                    if (IsOld(file, nowUtc))
                        removed += DeleteOne(file);
                }
                catch { }
            }
            return removed;
        }

        /// <summary>
        /// Удалить старые (не моложе <see cref="MinAge"/>) временные копии PDF из tempRoot —
        /// см. <see cref="TempCopyFile"/>. Корень параметром, чтобы тест не трогал настоящий
        /// %TEMP%. Возвращает число удалённых файлов и папок; сбои молчат, как и вся уборка.
        /// </summary>
        internal static int SweepTempCopies(string tempRoot, DateTime nowUtc)
        {
            int removed = 0;
            try
            {
                if (string.IsNullOrWhiteSpace(tempRoot) || !Directory.Exists(tempRoot))
                    return 0;
                foreach (string file in Directory.EnumerateFiles(tempRoot, "iwo_*.pdf"))
                {
                    if (!TempCopyFile.IsMatch(Path.GetFileName(file)))
                        continue;
                    try
                    {
                        if (IsOld(file, nowUtc))
                            removed += DeleteOne(file);
                    }
                    catch { }
                }
                foreach (string dir in Directory.EnumerateDirectories(tempRoot, "iwo_bg_*"))
                {
                    if (!TempCopyDir.IsMatch(Path.GetFileName(dir)))
                        continue;
                    try
                    {
                        // Возраст папки — по САМОМУ свежему, что в ней есть: время самой папки
                        // не меняется, пока долгий рендер пишет во вложенные src*\run*.
                        if (nowUtc - NewestWriteUtc(dir) < MinAge)
                            continue;
                        Directory.Delete(dir, true);
                        removed++;
                    }
                    catch { }
                }
            }
            catch
            {
                // Недоступный %TEMP% не должен задержать или сорвать запуск.
            }
            return removed;
        }

        private static DateTime NewestWriteUtc(string dir)
        {
            DateTime newest = Directory.GetLastWriteTimeUtc(dir);
            foreach (string entry in Directory.EnumerateFileSystemEntries(dir, "*",
                SearchOption.AllDirectories))
            {
                DateTime written = File.GetLastWriteTimeUtc(entry);
                if (written > newest)
                    newest = written;
            }
            return newest;
        }

        private static bool IsOld(string file, DateTime nowUtc)
        {
            return nowUtc - File.GetLastWriteTimeUtc(file) >= MinAge;
        }

        private static int DeleteOne(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                return 0;
            File.Delete(path);
            return 1;
        }

        internal static bool IsOrphanName(string fileName)
        {
            return !string.IsNullOrEmpty(fileName) && OrphanName.IsMatch(fileName);
        }

        /// <summary>Папки известных результатов; выполняется только на фоновом потоке.</summary>
        internal static IEnumerable<string> DirectoriesToSweep(UserSettings settings)
        {
            var dirs = new List<string> { AppPaths.Root };
            if (settings != null)
            {
                dirs.Add(settings.LastInputFolder);
                dirs.Add(settings.LastOutputFolder);
            }
            try
            {
                foreach (HistoryEntry entry in OperationHistory.Load().Entries)
                {
                    if (entry == null || string.IsNullOrEmpty(entry.Path))
                        continue;
                    dirs.Add(entry.Path);
                    try { dirs.Add(Path.GetDirectoryName(entry.Path)); } catch { }
                }
            }
            catch { }
            return dirs;
        }
    }
}
