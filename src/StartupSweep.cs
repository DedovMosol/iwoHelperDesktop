using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelMerger
{
    /// <summary>
    /// Уборка осиротевших временных файлов «.iwo-*» при старте (CWE-459: неполная очистка).
    /// Аварийное завершение (краш, выключение питания, снятие процесса) оставляет рядом с
    /// результатом файлы, которые штатный путь удаляет в finally: промежуточные
    /// «имя.iwo-&lt;guid&gt;.расширение» / «.bak» (AtomicOutput) и «имя…​.iwo-gs-&lt;guid&gt;.tmp/.bak»
    /// (GsRewrite). Они накапливались бы в папках пользователя навсегда.
    ///
    /// Безопасность:
    ///  • вызывается ТОЛЬКО после <see cref="SingleInstance.TryAcquire"/> — наши собственные
    ///    живые файлы не существуют, а второй экземпляр не запущен;
    ///  • возраст старше <see cref="MinAge"/> отпугивает случайное удаление файла, который
    ///    прямо сейчас пишет чужой процесс (в т.ч. консольный запуск из прошлого сеанса);
    ///  • имя обязано совпасть с точным рисунком наших суффиксов (32 шестнадцатеричных знака
    ///    после «.iwo-» или «.iwo-gs-»): пользовательский файл с похожим названием не страдает;
    ///  • любой сбой уборки молча глотается — очистка мусора не должна мешать запуску.
    /// </summary>
    internal static class StartupSweep
    {
        /// <summary>Возраст, после которого временный файл считается осиротевшим.</summary>
        internal static readonly TimeSpan MinAge = TimeSpan.FromHours(1);

        // Оба производителя пишут «.iwo-»/«.iwo-gs-» + 32 hex-знака (Guid «N»); после них —
        // расширение или конец имени. Требование полной тройки «.»-разделителей не даёт
        // зацепить пользовательские файлы вроде «отчёт.iwo-заметки.txt».
        private static readonly Regex OrphanName = new Regex(
            @"\.iwo-(?:gs-)?[0-9a-fA-F]{32}(\.[^.]*)?$",
            RegexOptions.Compiled);

        /// <summary>
        /// Удалить осиротевшие файлы из указанных папок. Возвращает число удалённых.
        /// Чистая функция от (папки, «сейчас») — под юнит-тест; ошибок не бросает никогда.
        /// </summary>
        internal static int Sweep(IEnumerable<string> directories, DateTime nowUtc)
        {
            int removed = 0;
            if (directories == null)
                return 0;
            foreach (string dir in directories)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir))
                        continue;
                    foreach (string file in Directory.GetFiles(dir))
                    {
                        string name = Path.GetFileName(file);
                        if (!IsOrphanName(name))
                            continue;
                        if (nowUtc - File.GetLastWriteTimeUtc(file) < MinAge)
                            continue;
                        File.Delete(file);
                        removed++;
                    }
                }
                catch
                {
                    // Уборка мусора не должна срывать запуск: папка без прав,
                    // сетевой диск без связи, файл, который кто-то держит, — пропускаем.
                }
            }
            return removed;
        }

        /// <summary>Похоже ли имя на наш временный файл (без проверки возраста).</summary>
        internal static bool IsOrphanName(string fileName)
        {
            return !string.IsNullOrEmpty(fileName) && OrphanName.IsMatch(fileName);
        }

        /// <summary>
        /// Папки, где уборка имеет смысл: последние папки ввода/вывода из настроек плюс
        /// папки результатов из истории операций. Других мест программа не использует.
        /// </summary>
        internal static IEnumerable<string> DirectoriesToSweep(UserSettings settings)
        {
            var dirs = new List<string>();
            if (settings != null)
            {
                dirs.Add(settings.LastInputFolder);
                dirs.Add(settings.LastOutputFolder);
            }
            try
            {
                foreach (HistoryEntry entry in OperationHistory.Load().Entries)
                {
                    if (string.IsNullOrEmpty(entry.Path))
                        continue;
                    dirs.Add(File.Exists(entry.Path)
                        ? Path.GetDirectoryName(entry.Path)
                        : (Directory.Exists(entry.Path) ? entry.Path : null));
                }
            }
            catch
            {
                // История — необязательное дополнение к папкам из настроек.
            }
            return dirs;
        }
    }
}
