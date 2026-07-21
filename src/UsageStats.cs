using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Локальные счётчики операций (без телеметрии): %APPDATA%\iwo Helper Desktop\stats.txt.
    /// Мутации — через read-modify-write (весь UI в одном потоке, поэтому без гонок и
    /// потери инкрементов между окнами). Поддержана ручная очистка и опциональная
    /// авто-очистка раз в N дней (0 — выключена).
    /// </summary>
    public class UsageStats
    {
        public int ExcelDigests;
        public int PdfMerges;
        public int PdfExtracts;
        public int PdfSplitRanges;
        public int PdfSplitEveryN;
        public int PdfSplitBookmarks;
        public int PdfToWord;                  // конвертаций «цифровой PDF → Word (.docx)»
        public int PdfCompressions;            // сжатых файлов; НЕ входит в Total (это параметр, не операция)
        public int AutoClearDays;              // 0 — выкл; 1 / 7 / 30
        public DateTime SinceUtc = DateTime.UtcNow;

        public int Total
        {
            get { return ExcelDigests + PdfMerges + PdfExtracts + PdfSplitRanges + PdfSplitEveryN + PdfSplitBookmarks + PdfToWord; }
        }

        /// <summary>Пора ли авто-очистка (период прошёл). Чистая — под тест.</summary>
        public static bool ShouldAutoClear(DateTime sinceUtc, DateTime nowUtc, int periodDays)
        {
            return periodDays > 0 && (nowUtc - sinceUtc).TotalDays >= periodDays;
        }

        public static UsageStats Load()
        {
            var s = new UsageStats();
            try
            {
                if (File.Exists(AppPaths.StatsFile))
                {
                    foreach (string line in File.ReadAllLines(AppPaths.StatsFile))
                    {
                        int eq = line.IndexOf('=');
                        if (eq <= 0)
                            continue;
                        string k = line.Substring(0, eq).Trim();
                        string v = line.Substring(eq + 1).Trim();
                        int n;
                        long ticks;
                        if (k == "excelDigests" && int.TryParse(v, out n)) s.ExcelDigests = n;
                        else if (k == "pdfMerges" && int.TryParse(v, out n)) s.PdfMerges = n;
                        else if (k == "pdfExtracts" && int.TryParse(v, out n)) s.PdfExtracts = n;
                        else if (k == "pdfSplitRanges" && int.TryParse(v, out n)) s.PdfSplitRanges = n;
                        else if (k == "pdfSplitEveryN" && int.TryParse(v, out n)) s.PdfSplitEveryN = n;
                        else if (k == "pdfSplitBookmarks" && int.TryParse(v, out n)) s.PdfSplitBookmarks = n;
                        else if (k == "pdfToWord" && int.TryParse(v, out n)) s.PdfToWord = n;
                        else if (k == "pdfCompressions" && int.TryParse(v, out n)) s.PdfCompressions = n;
                        else if (k == "autoClearDays" && int.TryParse(v, out n)) s.AutoClearDays = n;
                        else if (k == "sinceUtc" && long.TryParse(v, out ticks)) s.SinceUtc = new DateTime(ticks, DateTimeKind.Utc);
                    }
                }
            }
            catch { } // повреждённая статистика не должна мешать работе

            if (ShouldAutoClear(s.SinceUtc, DateTime.UtcNow, s.AutoClearDays))
            {
                int keepPeriod = s.AutoClearDays;
                s.ResetCounters();
                s.AutoClearDays = keepPeriod; // период очистки сохраняется
                s.Save();
            }
            return s;
        }

        private void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(AppPaths.StatsFile));
                File.WriteAllLines(AppPaths.StatsFile, new List<string>
                {
                    "excelDigests=" + ExcelDigests,
                    "pdfMerges=" + PdfMerges,
                    "pdfExtracts=" + PdfExtracts,
                    "pdfSplitRanges=" + PdfSplitRanges,
                    "pdfSplitEveryN=" + PdfSplitEveryN,
                    "pdfSplitBookmarks=" + PdfSplitBookmarks,
                    "pdfToWord=" + PdfToWord,
                    "pdfCompressions=" + PdfCompressions,
                    "autoClearDays=" + AutoClearDays,
                    "sinceUtc=" + SinceUtc.Ticks
                });
            }
            catch { }
        }

        private void ResetCounters()
        {
            ExcelDigests = 0;
            PdfMerges = 0;
            PdfExtracts = 0;
            PdfSplitRanges = 0;
            PdfSplitEveryN = 0;
            PdfSplitBookmarks = 0;
            PdfToWord = 0;
            PdfCompressions = 0;
            SinceUtc = DateTime.UtcNow;
        }

        // ---------- атомарные мутации (read-modify-write) ----------

        private static void Mutate(Action<UsageStats> change)
        {
            UsageStats s = Load();
            change(s);
            s.Save();
        }

        public static void RecordExcelDigest() { Mutate(delegate(UsageStats s) { s.ExcelDigests++; }); }
        public static void RecordPdfMerge() { Mutate(delegate(UsageStats s) { s.PdfMerges++; }); }
        public static void RecordPdfExtract() { Mutate(delegate(UsageStats s) { s.PdfExtracts++; }); }
        public static void RecordPdfSplitRanges() { Mutate(delegate(UsageStats s) { s.PdfSplitRanges++; }); }
        public static void RecordPdfSplitEveryN() { Mutate(delegate(UsageStats s) { s.PdfSplitEveryN++; }); }
        public static void RecordPdfSplitBookmarks() { Mutate(delegate(UsageStats s) { s.PdfSplitBookmarks++; }); }
        public static void RecordPdfToWord() { Mutate(delegate(UsageStats s) { s.PdfToWord++; }); }
        public static void RecordPdfCompress(int count = 1) { if (count > 0) Mutate(delegate(UsageStats s) { s.PdfCompressions += count; }); }

        public static void SetAutoClear(int days) { Mutate(delegate(UsageStats s) { s.AutoClearDays = days; }); }
        public static void ClearCounters() { Mutate(delegate(UsageStats s) { s.ResetCounters(); }); }
    }
}
