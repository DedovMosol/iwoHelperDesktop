using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// План «Повторить пропущенные»: что повторять и как собрать итог. Чистая логика без Excel —
    /// отдельно от <see cref="MergeService"/>, который владеет экземпляром Excel и копированием.
    /// Окно (<c>MainForm</c>) и запуск (<see cref="MergeRun"/>) берут план отсюда же, поэтому
    /// строки списка, набор повтора и итог не расходятся.
    ///
    /// Единица повтора — лист. Каждый <see cref="FileResult"/> помнит лист источника
    /// (<see cref="FileResult.SourceSheet"/>): файл повторяется целиком, только если в нём не
    /// перенеслось ничего; иначе — только его упавшие листы. Раньше повтор в режиме «все листы»
    /// копировал заново и уже перенесённые листы: в своде появлялись «Янв_2», «Фев_2», а
    /// оглавление и счётчики их дублировали.
    /// </summary>
    internal static class MergeRetry
    {
        /// <summary>Ключ файла: полный путь (недопустимый — как есть). Один на весь план повтора.</summary>
        internal static string Key(string path)
        {
            if (path == null)
                return null;
            try { return Path.GetFullPath(path); }
            catch { return path; }
        }

        /// <summary>
        /// Какие листы каждого файла повторять: имена упавших листов, или null — файл целиком.
        /// Пустой набор — повторять нечего: упал лист без имени (не прочиталось) в файле, где
        /// другие листы перенеслись; файл целиком дал бы дубли перенесённых. Чистая — под тест.
        /// </summary>
        internal static Dictionary<string, HashSet<string>> Sheets(MergeResult previous)
        {
            var failed = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            var anyOk = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FileResult fr in previous.Files)
            {
                string key = Key(fr.FullPath);
                if (key == null)
                    continue;
                if (fr.Ok)
                {
                    anyOk.Add(key);
                    continue;
                }
                HashSet<string> sheets;
                if (!failed.TryGetValue(key, out sheets))
                    failed[key] = sheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // имена листов Excel не различают регистр
                if (fr.SourceSheet != null)
                    sheets.Add(fr.SourceSheet);
            }
            var plan = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, HashSet<string>> item in failed)
                plan[item.Key] = anyOk.Contains(item.Key) ? item.Value : null;
            return plan;
        }

        /// <summary>
        /// Файлы повтора — каждый ОДИН раз, в порядке прошлого итога, и только те, где по плану
        /// есть что повторять. Раньше файл с двумя упавшими листами попадал сюда дважды и
        /// переносился дважды, а файл, где повторять нечего, открывался впустую. Пустой список —
        /// повторять нечего (кнопки повтора нет). Чистая — под тест.
        /// </summary>
        internal static List<string> Paths(MergeResult previous)
        {
            Dictionary<string, HashSet<string>> plan = Sheets(previous);
            var paths = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FileResult fr in previous.Files)
            {
                string key = Key(fr.FullPath);
                HashSet<string> sheets;
                if (fr.Ok || key == null || !plan.TryGetValue(key, out sheets) ||
                    (sheets != null && sheets.Count == 0) || !seen.Add(key))
                    continue;
                paths.Add(fr.FullPath);
            }
            return paths;
        }

        /// <summary>
        /// Записи прошлого итога, которые повтор не заменяет: успешные листы и пропуски, которые
        /// не повторяются. С них окно начинает строки повторяемых файлов — иначе строка файла с
        /// перенесённым «Янв» и упавшим «Фев» после повтора показывала бы один «Фев».
        /// Чистая — под тест.
        /// </summary>
        internal static List<FileResult> Kept(MergeResult previous)
        {
            Dictionary<string, HashSet<string>> plan = Sheets(previous);
            var kept = new List<FileResult>();
            foreach (FileResult fr in previous.Files)
                if (!IsRetried(fr, plan))
                    kept.Add(fr);
            return kept;
        }

        /// <summary>
        /// Итог после повтора: повторённые записи прошлого прогона заменяются свежими результатами
        /// своего файла (на месте первой из них), остальные записи и порядок сохраняются.
        /// Чистая — под тест.
        /// </summary>
        internal static MergeResult Combine(MergeResult previous, IList<FileResult> attempts)
        {
            // Свежие результаты, сгруппированные по файлу (в режиме «все листы» их несколько).
            var freshByFile = new Dictionary<string, List<FileResult>>(StringComparer.OrdinalIgnoreCase);
            foreach (FileResult a in attempts)
            {
                string key = Key(a.FullPath) ?? a.FileName;
                List<FileResult> bucket;
                if (!freshByFile.TryGetValue(key, out bucket))
                    freshByFile[key] = bucket = new List<FileResult>();
                bucket.Add(a);
            }

            var combined = new MergeResult();
            combined.OutputPath = previous.OutputPath;
            combined.MacrosCarried = previous.MacrosCarried; // свод тот же; при сохранении проверится заново
            Dictionary<string, HashSet<string>> plan = Sheets(previous);
            var expanded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FileResult old in previous.Files)
            {
                string key = Key(old.FullPath) ?? old.FileName;
                List<FileResult> fresh;
                if (key != null && IsRetried(old, plan) && freshByFile.TryGetValue(key, out fresh))
                {
                    if (expanded.Add(key)) // файл разворачивается один раз
                        combined.Files.AddRange(fresh);
                }
                else
                {
                    combined.Files.Add(old); // успешная или неповторённая запись остаётся
                }
            }
            foreach (FileResult f in combined.Files)
                if (f.Ok) combined.OkCount++; else combined.SkipCount++;
            return combined;
        }

        /// <summary>Заменяет ли повтор эту запись прошлого итога (по плану <see cref="Sheets"/>).</summary>
        private static bool IsRetried(FileResult fr, Dictionary<string, HashSet<string>> plan)
        {
            if (fr.Ok)
                return false;
            HashSet<string> sheets;
            string key = Key(fr.FullPath);
            if (key == null || !plan.TryGetValue(key, out sheets))
                return false;
            return sheets == null || (fr.SourceSheet != null && sheets.Contains(fr.SourceSheet));
        }
    }
}
