using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>Что просили сделать в командной строке.</summary>
    internal enum PdfCliKind
    {
        None = 0, Merge, Extract, Split, Compress, Grayscale, Repair, ToImages, ToText
    }

    /// <summary>Как разбивать документ (только для <see cref="PdfCliKind.Split"/>).</summary>
    internal enum PdfCliSplitMode { Ranges = 0, Every, Bookmarks }

    /// <summary>
    /// Разобранная команда. Отдельный тип, а не набор out-параметров: разбор проверяется
    /// тестами целиком, одним сравнением ожидаемого с полученным.
    /// </summary>
    internal sealed class PdfCliCommand
    {
        public PdfCliKind Kind;
        public List<string> Inputs = new List<string>();
        public string Output;                    // файл или папка — смотря по команде
        public string Error;                     // разбор не удался: что именно не так
        public CompressionLevel Level = CompressionLevel.None;
        public PdfCliSplitMode SplitMode = PdfCliSplitMode.Ranges;
        public string Ranges;                    // «1-3,5» для диапазонов и извлечения
        public int Every;                        // страниц в части
        public int Dpi = 150;
        public ImageExportFormat Format = ImageExportFormat.Png;
    }

    /// <summary>
    /// Те же операции, что и в окнах, но из командной строки — для пакетной обработки и
    /// сценариев. Ничего нового здесь не считается: разбор превращает строку в описание
    /// команды, а выполняет её тот же сервис, что и кнопка в окне. Поэтому CLI не может
    /// разойтись с интерфейсом в поведении — расходиться нечему.
    ///
    /// Коды возврата, как и у режима Excel: 0 — сделано, 1 — не получилось,
    /// 2 — неверно вызвано (эти три различает любой сценарий-обёртка).
    /// </summary>
    internal static class PdfCli
    {
        public const int Ok = 0;
        public const int Failed = 1;
        public const int BadUsage = 2;

        private static readonly string[] Commands =
        {
            "--merge", "--extract", "--split", "--compress", "--grayscale", "--repair",
            "--to-image", "--to-text"
        };

        /// <summary>Начинается ли строка запуска с нашей команды. Чистая — под тест.</summary>
        public static bool IsCommand(string[] args)
        {
            if (args == null || args.Length == 0)
                return false;
            foreach (string name in Commands)
                if (string.Equals(args[0], name, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        /// <summary>
        /// Разобрать строку запуска. Ошибка возвращается в поле, а не исключением: вызывающему
        /// нужно напечатать её вместе со справкой, а не ловить. Чистая — под тест.
        /// </summary>
        public static PdfCliCommand Parse(string[] args)
        {
            var cmd = new PdfCliCommand();
            if (args == null || args.Length == 0)
                return Bad(cmd, "no command");
            string head = args[0].ToLowerInvariant();
            var rest = new List<string>();
            for (int i = 1; i < args.Length; i++)
                rest.Add(args[i]);

            switch (head)
            {
                case "--merge":      return ParseMerge(cmd, rest);
                case "--extract":    return ParseExtract(cmd, rest);
                case "--split":      return ParseSplit(cmd, rest);
                case "--compress":   return ParseTwoPaths(cmd, PdfCliKind.Compress, rest, true);
                case "--grayscale":  return ParseTwoPaths(cmd, PdfCliKind.Grayscale, rest, false);
                case "--repair":     return ParseTwoPaths(cmd, PdfCliKind.Repair, rest, false);
                case "--to-image":   return ParseToImages(cmd, rest);
                case "--to-text":    return ParseTwoPaths(cmd, PdfCliKind.ToText, rest, false);
                default:             return Bad(cmd, "unknown command: " + args[0]);
            }
        }

        /// <summary>Уровень сжатия по имени. Чистая — под тест.</summary>
        public static bool TryParseLevel(string text, out CompressionLevel level)
        {
            level = CompressionLevel.None;
            if (string.IsNullOrEmpty(text))
                return false;
            switch (text.ToLowerInvariant())
            {
                case "none":     level = CompressionLevel.None; return true;
                case "verygood": level = CompressionLevel.VeryGood; return true;
                case "good":     level = CompressionLevel.Good; return true;
                case "normal":   level = CompressionLevel.Small; return true;
                default:         return false;
            }
        }

        /// <summary>Справка — единственный текст CLI, который живёт не в каталоге переводов:
        /// командная строка отвечает на языке своих ключей, а они английские.</summary>
        public static string Usage()
        {
            return
                "iwoHelperDesktop PDF commands:\r\n" +
                "  --merge <out.pdf> <in.pdf> [in.pdf ...] [--level none|verygood|good|normal]\r\n" +
                "  --extract <in.pdf> <pages> <out.pdf>            pages: 1-3,5\r\n" +
                "  --split <in.pdf> <out_dir> [--ranges 1-3,5 | --every N | --bookmarks]\r\n" +
                "  --compress <in.pdf> <out.pdf> [--level verygood|good|normal]\r\n" +
                "  --grayscale <in.pdf> <out.pdf>\r\n" +
                "  --repair <in.pdf> <out.pdf>\r\n" +
                "  --to-image <in.pdf> <out_dir> [--dpi 150] [--format png|jpg]\r\n" +
                "  --to-text <in.pdf> <out.txt>\r\n" +
                "Exit codes: 0 done, 1 failed, 2 bad usage.";
        }

        /// <summary>
        /// Выполнить разобранную команду. Печатает результат через log и возвращает код.
        /// Исключения сюда не выходят: сценарию нужен код возврата и строка, а не трасса.
        /// </summary>
        public static int Execute(PdfCliCommand cmd, Action<string> log)
        {
            if (cmd == null || cmd.Error != null)
            {
                Say(log, "ERROR: " + (cmd == null ? "no command" : cmd.Error));
                Say(log, Usage());
                return BadUsage;
            }
            try
            {
                switch (cmd.Kind)
                {
                    case PdfCliKind.Merge:     return RunMerge(cmd, log);
                    case PdfCliKind.Extract:   return RunExtract(cmd, log);
                    case PdfCliKind.Split:     return RunSplit(cmd, log);
                    case PdfCliKind.Compress:  return RunCompress(cmd, log);
                    case PdfCliKind.Grayscale: return RunConvert(cmd, PdfConvertMode.Grayscale, log);
                    case PdfCliKind.Repair:    return RunConvert(cmd, PdfConvertMode.Repair, log);
                    case PdfCliKind.ToImages:  return RunToImages(cmd, log);
                    case PdfCliKind.ToText:    return RunToText(cmd, log);
                    default:
                        Say(log, Usage());
                        return BadUsage;
                }
            }
            catch (Exception ex)
            {
                Say(log, "ERROR: " + ex.Message);
                return Failed;
            }
        }

        // ---------- разбор ----------

        private static PdfCliCommand ParseMerge(PdfCliCommand cmd, List<string> rest)
        {
            cmd.Kind = PdfCliKind.Merge;
            List<string> plain = TakeOptions(cmd, rest);
            if (cmd.Error != null)
                return cmd;
            if (plain.Count < 2)
                return Bad(cmd, "--merge needs an output file and at least one input");
            cmd.Output = plain[0];
            for (int i = 1; i < plain.Count; i++)
                cmd.Inputs.Add(plain[i]);
            return cmd;
        }

        private static PdfCliCommand ParseExtract(PdfCliCommand cmd, List<string> rest)
        {
            cmd.Kind = PdfCliKind.Extract;
            List<string> plain = TakeOptions(cmd, rest);
            if (cmd.Error != null)
                return cmd;
            if (plain.Count != 3)
                return Bad(cmd, "--extract needs <in.pdf> <pages> <out.pdf>");
            cmd.Inputs.Add(plain[0]);
            cmd.Ranges = plain[1];
            cmd.Output = plain[2];
            return cmd;
        }

        private static PdfCliCommand ParseSplit(PdfCliCommand cmd, List<string> rest)
        {
            cmd.Kind = PdfCliKind.Split;
            List<string> plain = TakeOptions(cmd, rest);
            if (cmd.Error != null)
                return cmd;
            if (plain.Count != 2)
                return Bad(cmd, "--split needs <in.pdf> <out_dir>");
            cmd.Inputs.Add(plain[0]);
            cmd.Output = plain[1];
            if (cmd.SplitMode == PdfCliSplitMode.Ranges && string.IsNullOrEmpty(cmd.Ranges))
                return Bad(cmd, "--split needs --ranges, --every or --bookmarks");
            return cmd;
        }

        private static PdfCliCommand ParseTwoPaths(PdfCliCommand cmd, PdfCliKind kind, List<string> rest, bool needLevel)
        {
            cmd.Kind = kind;
            List<string> plain = TakeOptions(cmd, rest);
            if (cmd.Error != null)
                return cmd;
            if (plain.Count != 2)
                return Bad(cmd, Name(kind) + " needs <in.pdf> <out>");
            cmd.Inputs.Add(plain[0]);
            cmd.Output = plain[1];
            if (needLevel && cmd.Level == CompressionLevel.None)
                cmd.Level = CompressionLevel.Good; // сжатие без уровня бессмысленно — берём средний
            return cmd;
        }

        private static PdfCliCommand ParseToImages(PdfCliCommand cmd, List<string> rest)
        {
            cmd.Kind = PdfCliKind.ToImages;
            List<string> plain = TakeOptions(cmd, rest);
            if (cmd.Error != null)
                return cmd;
            if (plain.Count != 2)
                return Bad(cmd, "--to-image needs <in.pdf> <out_dir>");
            cmd.Inputs.Add(plain[0]);
            cmd.Output = plain[1];
            return cmd;
        }

        /// <summary>
        /// Вынуть из хвоста именованные ключи, вернув остальное позиционным списком. Один
        /// разборщик на все команды: ключи везде значат одно и то же, и повторять их разбор
        /// в каждой ветке значило бы однажды разойтись.
        /// </summary>
        private static List<string> TakeOptions(PdfCliCommand cmd, List<string> rest)
        {
            var plain = new List<string>();
            for (int i = 0; i < rest.Count; i++)
            {
                string arg = rest[i];
                if (!arg.StartsWith("--", StringComparison.Ordinal))
                {
                    plain.Add(arg);
                    continue;
                }
                switch (arg.ToLowerInvariant())
                {
                    case "--bookmarks":
                        cmd.SplitMode = PdfCliSplitMode.Bookmarks;
                        break;
                    case "--level":
                        if (++i >= rest.Count || !TryParseLevel(rest[i], out cmd.Level))
                            return Fail(cmd, plain, "--level: none|verygood|good|normal");
                        break;
                    case "--ranges":
                        if (++i >= rest.Count)
                            return Fail(cmd, plain, "--ranges needs a value like 1-3,5");
                        cmd.Ranges = rest[i];
                        cmd.SplitMode = PdfCliSplitMode.Ranges;
                        break;
                    case "--every":
                        int every;
                        if (++i >= rest.Count || !int.TryParse(rest[i], out every) || every < 1)
                            return Fail(cmd, plain, "--every needs a number of pages (1 or more)");
                        cmd.Every = every;
                        cmd.SplitMode = PdfCliSplitMode.Every;
                        break;
                    case "--dpi":
                        int dpi;
                        if (++i >= rest.Count || !int.TryParse(rest[i], out dpi) || dpi < 1)
                            return Fail(cmd, plain, "--dpi needs a positive number");
                        cmd.Dpi = dpi;
                        break;
                    case "--format":
                        if (++i >= rest.Count)
                            return Fail(cmd, plain, "--format: png|jpg");
                        string format = rest[i].ToLowerInvariant();
                        if (format == "png") cmd.Format = ImageExportFormat.Png;
                        else if (format == "jpg" || format == "jpeg") cmd.Format = ImageExportFormat.Jpeg;
                        else return Fail(cmd, plain, "--format: png|jpg");
                        break;
                    default:
                        return Fail(cmd, plain, "unknown option: " + arg);
                }
            }
            return plain;
        }

        private static List<string> Fail(PdfCliCommand cmd, List<string> plain, string message)
        {
            cmd.Error = message;
            return plain;
        }

        private static PdfCliCommand Bad(PdfCliCommand cmd, string message)
        {
            cmd.Error = message;
            return cmd;
        }

        private static string Name(PdfCliKind kind)
        {
            switch (kind)
            {
                case PdfCliKind.Compress:  return "--compress";
                case PdfCliKind.Grayscale: return "--grayscale";
                case PdfCliKind.Repair:    return "--repair";
                case PdfCliKind.ToText:    return "--to-text";
                default:                   return "command";
            }
        }

        // ---------- выполнение (те же сервисы, что и у кнопок) ----------

        private static int RunMerge(PdfCliCommand cmd, Action<string> log)
        {
            var order = new List<PdfPageRef>();
            foreach (string path in cmd.Inputs)
            {
                int pages = PdfMergeService.LoadPages(path).Count;
                for (int i = 0; i < pages; i++)
                    order.Add(new PdfPageRef { SourcePath = path, PageIndex = i });
            }
            PdfMergeService.Merge(order, cmd.Output);
            Compress(cmd.Output, cmd.Level, log);
            Say(log, string.Format("OK: {0} ({1} pages from {2} files)", cmd.Output, order.Count, cmd.Inputs.Count));
            return Ok;
        }

        private static int RunExtract(PdfCliCommand cmd, Action<string> log)
        {
            int pageCount = PdfMergeService.LoadPages(cmd.Inputs[0]).Count;
            var indexes = new List<int>();
            foreach (PageRange r in PageRanges.Parse(cmd.Ranges, pageCount))
                for (int i = r.Start; i <= r.End; i++)
                    indexes.Add(i);
            PdfSplitService.Extract(cmd.Inputs[0], indexes, cmd.Output);
            Compress(cmd.Output, cmd.Level, log);
            Say(log, string.Format("OK: {0} ({1} pages)", cmd.Output, indexes.Count));
            return Ok;
        }

        private static int RunSplit(PdfCliCommand cmd, Action<string> log)
        {
            string source = cmd.Inputs[0];
            string baseName = Path.GetFileNameWithoutExtension(source);
            // Папку назначения создаём здесь: в окне её выбирают в диалоге, и она заведомо
            // есть, а в командной строке её пишут как строку — и обычно ещё не существующую.
            Directory.CreateDirectory(cmd.Output);
            List<string> parts;
            switch (cmd.SplitMode)
            {
                case PdfCliSplitMode.Every:
                    parts = PdfSplitService.SplitEveryN(source, cmd.Every, cmd.Output, baseName);
                    break;
                case PdfCliSplitMode.Bookmarks:
                    parts = PdfSplitService.SplitByBookmarks(source, cmd.Output, baseName);
                    break;
                default:
                    int pageCount = PdfMergeService.LoadPages(source).Count;
                    parts = PdfSplitService.SplitByRanges(source,
                        PageRanges.Parse(cmd.Ranges, pageCount), cmd.Output, baseName);
                    break;
            }
            foreach (string part in parts)
                Compress(part, cmd.Level, log);
            Say(log, string.Format("OK: {0} files in {1}", parts.Count, cmd.Output));
            return Ok;
        }

        private static int RunCompress(PdfCliCommand cmd, Action<string> log)
        {
            if (!Ghostscript.Available)
            {
                Say(log, "ERROR: Ghostscript not found (needed for compression)");
                return Failed;
            }
            CopyToOutput(cmd.Inputs[0], cmd.Output);
            long before = new FileInfo(cmd.Output).Length;
            bool shrank = PdfCompression.Compress(cmd.Output, cmd.Level);
            long after = new FileInfo(cmd.Output).Length;
            Say(log, shrank
                ? string.Format("OK: {0} ({1} -> {2} bytes)", cmd.Output, before, after)
                : string.Format("OK: {0} (not smaller, left as is)", cmd.Output));
            return Ok;
        }

        private static int RunConvert(PdfCliCommand cmd, PdfConvertMode mode, Action<string> log)
        {
            if (!Ghostscript.Available)
            {
                Say(log, "ERROR: Ghostscript not found (needed for this command)");
                return Failed;
            }
            CopyToOutput(cmd.Inputs[0], cmd.Output);
            if (!PdfConvert.Apply(cmd.Output, mode))
            {
                TryDelete(cmd.Output);            // огрызок не оставляем — как и в окне
                Say(log, "ERROR: the engine could not perform this conversion");
                return Failed;
            }
            Say(log, "OK: " + cmd.Output);
            return Ok;
        }

        private static int RunToImages(PdfCliCommand cmd, Action<string> log)
        {
            int pageCount = PdfMergeService.LoadPages(cmd.Inputs[0]).Count;
            var indexes = new List<int>();
            for (int i = 0; i < pageCount; i++)
                indexes.Add(i);
            List<string> files = PdfExportService.ToImages(cmd.Inputs[0], indexes, cmd.Output,
                null, cmd.Format, cmd.Dpi);
            Say(log, string.Format("OK: {0} images in {1}", files.Count, cmd.Output));
            return Ok;
        }

        private static int RunToText(PdfCliCommand cmd, Action<string> log)
        {
            PdfExportService.ToText(cmd.Inputs[0], cmd.Output);
            Say(log, "OK: " + cmd.Output);
            return Ok;
        }

        /// <summary>Сжать готовый файл, если уровень задан. Молча: сжатие необязательно.</summary>
        private static void Compress(string path, CompressionLevel level, Action<string> log)
        {
            if (level == CompressionLevel.None || !Ghostscript.Available)
                return;
            if (!PdfCompression.Compress(path, level))
                Say(log, "note: compression did not make " + Path.GetFileName(path) + " smaller");
        }

        /// <summary>
        /// Команды, меняющие файл, работают с КОПИЕЙ: исходник приложение не трогает никогда,
        /// и командная строка тут не исключение.
        /// </summary>
        private static void CopyToOutput(string source, string output)
        {
            if (string.Equals(Path.GetFullPath(source), Path.GetFullPath(output), StringComparison.OrdinalIgnoreCase))
                throw new MergeException("input and output must be different files");
            string dir = Path.GetDirectoryName(Path.GetFullPath(output));
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);
            File.Copy(source, output, true);
        }

        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch { }
        }

        private static void Say(Action<string> log, string line)
        {
            if (log != null)
                log(line);
        }
    }
}
