using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Локальные счётчики операций (без телеметрии): %APPDATA%\iwo Helper Desktop\stats.txt.
    /// Все мутации и автоматический сброс — один read-modify-write под межпроцессной
    /// файловой блокировкой, поэтому параллельный процесс не теряет свежий инкремент.
    /// </summary>
    public class UsageStats
    {
        public int ExcelDigests;
        public int PdfMerges;
        public int PdfExtracts;
        public int PdfSplitRanges;
        public int PdfSplitEveryN;
        public int PdfSplitBookmarks;
        public int PdfComparisons;
        public int PdfToWord;
        public int PdfToPptx;
        public int PdfCompressions;
        public int AutoClearDays;
        public DateTime SinceUtc = DateTime.UtcNow;

        private readonly List<string> _unknownLines = new List<string>();

        public int Total
        {
            get
            {
                long total = (long)ExcelDigests + PdfMerges + PdfExtracts + PdfSplitRanges +
                    PdfSplitEveryN + PdfSplitBookmarks + PdfComparisons + PdfToWord + PdfToPptx;
                return total > int.MaxValue ? int.MaxValue : (int)Math.Max(0, total);
            }
        }

        public static bool ShouldAutoClear(DateTime sinceUtc, DateTime nowUtc, int periodDays)
        {
            return periodDays > 0 && (nowUtc - sinceUtc).TotalDays >= periodDays;
        }

        public static UsageStats Load()
        {
            UsageStats result = null;
            bool locked = AppDataLock.TryRun(AppPaths.StatsFile, delegate
            {
                UsageStats snapshot;
                if (!TryLoadRaw(out snapshot))
                    return false;
                if (ApplyAutoClear(snapshot, DateTime.UtcNow))
                    snapshot.SaveCore();
                result = snapshot;
                return true;
            });
            if (locked)
                return result;

            // Таймаут блокировки не мешает показать статистику. Это только чтение: никаких
            // defaults поверх чужого файла и никакого сброса без lock здесь нет.
            if (TryLoadRaw(out result))
                return result;
            return new UsageStats();
        }

        private static bool TryLoadRaw(out UsageStats stats)
        {
            stats = new UsageStats();
            string[] lines;
            if (!AppStateFile.TryReadLines(AppPaths.StatsFile, out lines))
                return false;

            foreach (string line in lines)
            {
                int eq = line.IndexOf('=');
                if (eq <= 0)
                {
                    stats._unknownLines.Add(line);
                    continue;
                }
                string key = line.Substring(0, eq).Trim();
                string value = line.Substring(eq + 1).Trim();
                int number;
                bool known = true;
                if (key == "excelDigests" && TryCounter(value, out number)) stats.ExcelDigests = number;
                else if (key == "pdfMerges" && TryCounter(value, out number)) stats.PdfMerges = number;
                else if (key == "pdfExtracts" && TryCounter(value, out number)) stats.PdfExtracts = number;
                else if (key == "pdfSplitRanges" && TryCounter(value, out number)) stats.PdfSplitRanges = number;
                else if (key == "pdfSplitEveryN" && TryCounter(value, out number)) stats.PdfSplitEveryN = number;
                else if (key == "pdfSplitBookmarks" && TryCounter(value, out number)) stats.PdfSplitBookmarks = number;
                else if (key == "pdfComparisons" && TryCounter(value, out number)) stats.PdfComparisons = number;
                else if (key == "pdfToWord" && TryCounter(value, out number)) stats.PdfToWord = number;
                else if (key == "pdfToPptx" && TryCounter(value, out number)) stats.PdfToPptx = number;
                else if (key == "pdfCompressions" && TryCounter(value, out number)) stats.PdfCompressions = number;
                else if (key == "autoClearDays" && int.TryParse(value, out number) && number >= 0) stats.AutoClearDays = number;
                else if (key == "sinceUtc")
                {
                    long ticks;
                    if (long.TryParse(value, out ticks) && ticks >= DateTime.MinValue.Ticks &&
                        ticks <= DateTime.MaxValue.Ticks)
                        stats.SinceUtc = new DateTime(ticks, DateTimeKind.Utc);
                }
                else
                    known = false;
                if (!known)
                    stats._unknownLines.Add(line);
            }
            return true;
        }

        private static bool TryCounter(string value, out int number)
        {
            return int.TryParse(value, out number) && number >= 0;
        }

        private static bool ApplyAutoClear(UsageStats stats, DateTime nowUtc)
        {
            if (!ShouldAutoClear(stats.SinceUtc, nowUtc, stats.AutoClearDays))
                return false;
            stats.ResetCounters(nowUtc);
            return true;
        }

        private void SaveCore()
        {
            var lines = new List<string>
            {
                "excelDigests=" + ExcelDigests,
                "pdfMerges=" + PdfMerges,
                "pdfExtracts=" + PdfExtracts,
                "pdfSplitRanges=" + PdfSplitRanges,
                "pdfSplitEveryN=" + PdfSplitEveryN,
                "pdfSplitBookmarks=" + PdfSplitBookmarks,
                "pdfComparisons=" + PdfComparisons,
                "pdfToWord=" + PdfToWord,
                "pdfToPptx=" + PdfToPptx,
                "pdfCompressions=" + PdfCompressions,
                "autoClearDays=" + AutoClearDays,
                "sinceUtc=" + SinceUtc.Ticks
            };
            lines.AddRange(_unknownLines);
            AppStateFile.WriteLines(AppPaths.StatsFile, lines);
        }

        private void ResetCounters(DateTime nowUtc)
        {
            ExcelDigests = 0;
            PdfMerges = 0;
            PdfExtracts = 0;
            PdfSplitRanges = 0;
            PdfSplitEveryN = 0;
            PdfSplitBookmarks = 0;
            PdfComparisons = 0;
            PdfToWord = 0;
            PdfToPptx = 0;
            PdfCompressions = 0;
            SinceUtc = nowUtc;
        }

        private static bool Mutate(Action<UsageStats> change)
        {
            if (change == null)
                return false;
            return AppDataLock.TryRun(AppPaths.StatsFile, delegate
            {
                UsageStats stats;
                if (!TryLoadRaw(out stats))
                    return false;
                ApplyAutoClear(stats, DateTime.UtcNow);
                change(stats);
                stats.SaveCore();
                return true;
            });
        }

        private static void Increment(ref int value, int amount = 1)
        {
            if (amount <= 0 || value == int.MaxValue)
                return;
            value = amount > int.MaxValue - value ? int.MaxValue : value + amount;
        }

        public static void RecordExcelDigest() { Mutate(delegate(UsageStats s) { Increment(ref s.ExcelDigests); }); }
        public static void RecordPdfMerge() { Mutate(delegate(UsageStats s) { Increment(ref s.PdfMerges); }); }
        public static void RecordPdfExtract() { Mutate(delegate(UsageStats s) { Increment(ref s.PdfExtracts); }); }
        public static void RecordPdfSplitRanges() { Mutate(delegate(UsageStats s) { Increment(ref s.PdfSplitRanges); }); }
        public static void RecordPdfSplitEveryN() { Mutate(delegate(UsageStats s) { Increment(ref s.PdfSplitEveryN); }); }
        public static void RecordPdfSplitBookmarks() { Mutate(delegate(UsageStats s) { Increment(ref s.PdfSplitBookmarks); }); }
        public static void RecordPdfCompare() { Mutate(delegate(UsageStats s) { Increment(ref s.PdfComparisons); }); }
        public static void RecordPdfToWord() { Mutate(delegate(UsageStats s) { Increment(ref s.PdfToWord); }); }
        public static void RecordPdfToPptx() { Mutate(delegate(UsageStats s) { Increment(ref s.PdfToPptx); }); }
        public static void RecordPdfCompress(int count = 1) { if (count > 0) Mutate(delegate(UsageStats s) { Increment(ref s.PdfCompressions, count); }); }

        public static bool SetAutoClear(int days) { return Mutate(delegate(UsageStats s) { s.AutoClearDays = Math.Max(0, days); }); }
        public static bool ClearCounters() { return Mutate(delegate(UsageStats s) { s.ResetCounters(DateTime.UtcNow); }); }
    }
}
