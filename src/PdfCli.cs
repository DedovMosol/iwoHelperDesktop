using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>Что просили сделать в командной строке.</summary>
    internal enum PdfCliKind
    {
        None = 0, Help, Merge, Extract, Split, Compress, Grayscale, Repair, ToImages, ToText
    }

    /// <summary>Как разбивать документ (только для <see cref="PdfCliKind.Split"/>).</summary>
    internal enum PdfCliSplitMode { Ranges = 0, Every, Bookmarks }

    [Flags]
    internal enum PdfCliOption
    {
        None = 0,
        Level = 1,
        Ranges = 2,
        Every = 4,
        Bookmarks = 8,
        Dpi = 16,
        Format = 32,
        SplitSelectors = Ranges | Every | Bookmarks
    }

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
        public PdfCliOption SeenOptions;
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
            "--to-image", "--to-text", "--help", "-h", "/?"
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
                case "--help":
                case "-h":
                case "/?":
                    cmd.Kind = PdfCliKind.Help;
                    return cmd;
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
                "  --extract <in.pdf> <pages> <out.pdf> [--level none|verygood|good|normal]  pages: 1-3,5\r\n" +
                "  --split <in.pdf> <out_dir> [--ranges 1-3,5 | --every N | --bookmarks] [--level none|verygood|good|normal]\r\n" +
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
                    case PdfCliKind.Help:
                        Say(log, Usage());
                        return Ok;
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
            if (cmd.Error != null || !ValidateOptions(cmd, PdfCliOption.Level))
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
            if (cmd.Error != null || !ValidateOptions(cmd, PdfCliOption.Level))
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
            if (cmd.Error != null || !ValidateOptions(cmd,
                PdfCliOption.Level | PdfCliOption.SplitSelectors))
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
            PdfCliOption allowed = kind == PdfCliKind.Compress
                ? PdfCliOption.Level : PdfCliOption.None;
            if (cmd.Error != null || !ValidateOptions(cmd, allowed))
                return cmd;
            if (plain.Count != 2)
                return Bad(cmd, Name(kind) + " needs <in.pdf> <out>");
            cmd.Inputs.Add(plain[0]);
            cmd.Output = plain[1];
            if (needLevel && (cmd.SeenOptions & PdfCliOption.Level) == 0)
                cmd.Level = CompressionLevel.Good;
            else if (needLevel && cmd.Level == CompressionLevel.None)
                return Bad(cmd, "--compress does not accept --level none");
            return cmd;
        }

        private static PdfCliCommand ParseToImages(PdfCliCommand cmd, List<string> rest)
        {
            cmd.Kind = PdfCliKind.ToImages;
            List<string> plain = TakeOptions(cmd, rest);
            if (cmd.Error != null || !ValidateOptions(cmd,
                PdfCliOption.Dpi | PdfCliOption.Format))
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
            bool positionalOnly = false;
            for (int i = 0; i < rest.Count; i++)
            {
                string arg = rest[i];
                if (positionalOnly || !arg.StartsWith("--", StringComparison.Ordinal))
                {
                    plain.Add(arg);
                    continue;
                }
                if (arg == "--")
                {
                    positionalOnly = true;
                    continue;
                }
                switch (arg.ToLowerInvariant())
                {
                    case "--bookmarks":
                        if ((cmd.SeenOptions & PdfCliOption.SplitSelectors) != 0)
                            return Fail(cmd, plain, "split selectors are mutually exclusive");
                        cmd.SeenOptions |= PdfCliOption.Bookmarks;
                        cmd.SplitMode = PdfCliSplitMode.Bookmarks;
                        break;
                    case "--level":
                        if ((cmd.SeenOptions & PdfCliOption.Level) != 0)
                            return Fail(cmd, plain, "--level specified more than once");
                        cmd.SeenOptions |= PdfCliOption.Level;
                        if (++i >= rest.Count || !TryParseLevel(rest[i], out cmd.Level))
                            return Fail(cmd, plain, "--level: none|verygood|good|normal");
                        break;
                    case "--ranges":
                        if ((cmd.SeenOptions & PdfCliOption.SplitSelectors) != 0)
                            return Fail(cmd, plain, "split selectors are mutually exclusive");
                        cmd.SeenOptions |= PdfCliOption.Ranges;
                        if (++i >= rest.Count)
                            return Fail(cmd, plain, "--ranges needs a value like 1-3,5");
                        cmd.Ranges = rest[i];
                        cmd.SplitMode = PdfCliSplitMode.Ranges;
                        break;
                    case "--every":
                        if ((cmd.SeenOptions & PdfCliOption.SplitSelectors) != 0)
                            return Fail(cmd, plain, "split selectors are mutually exclusive");
                        cmd.SeenOptions |= PdfCliOption.Every;
                        int every;
                        if (++i >= rest.Count || !int.TryParse(rest[i], out every) || every < 1)
                            return Fail(cmd, plain, "--every needs a number of pages (1 or more)");
                        cmd.Every = every;
                        cmd.SplitMode = PdfCliSplitMode.Every;
                        break;
                    case "--dpi":
                        if ((cmd.SeenOptions & PdfCliOption.Dpi) != 0)
                            return Fail(cmd, plain, "--dpi specified more than once");
                        cmd.SeenOptions |= PdfCliOption.Dpi;
                        int dpi;
                        if (++i >= rest.Count || !int.TryParse(rest[i], out dpi) ||
                            !RasterBudget.IsValidExportDpi(dpi))
                            return Fail(cmd, plain, "--dpi needs a number from 1 to " +
                                RasterBudget.MaxExportDpi);
                        cmd.Dpi = dpi;
                        break;
                    case "--format":
                        if ((cmd.SeenOptions & PdfCliOption.Format) != 0)
                            return Fail(cmd, plain, "--format specified more than once");
                        cmd.SeenOptions |= PdfCliOption.Format;
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

        private static bool ValidateOptions(PdfCliCommand cmd, PdfCliOption allowed)
        {
            PdfCliOption invalid = cmd.SeenOptions & ~allowed;
            if (invalid == PdfCliOption.None)
                return true;
            cmd.Error = OptionName(invalid) + " is not valid for " + Name(cmd.Kind);
            return false;
        }

        private static string OptionName(PdfCliOption options)
        {
            if ((options & PdfCliOption.Level) != 0) return "--level";
            if ((options & PdfCliOption.Ranges) != 0) return "--ranges";
            if ((options & PdfCliOption.Every) != 0) return "--every";
            if ((options & PdfCliOption.Bookmarks) != 0) return "--bookmarks";
            if ((options & PdfCliOption.Dpi) != 0) return "--dpi";
            return "--format";
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
            using (AtomicOutput output = CopyToOutput(cmd.Inputs[0], cmd.Output))
            {
                long before = new FileInfo(output.TempPath).Length;
                GsRewriteResult result = PdfCompression.CompressDetailedUnpublished(
                    output.TempPath, cmd.Level);
                if (result == GsRewriteResult.Failed)
                {
                    Say(log, "ERROR: Ghostscript failed or produced an invalid result");
                    return Failed;
                }
                long after = new FileInfo(output.TempPath).Length;
                output.Commit();
                Say(log, result == GsRewriteResult.Applied
                    ? string.Format("OK: {0} ({1} -> {2} bytes)", cmd.Output, before, after)
                    : string.Format("OK: {0} (not smaller, left as is)", cmd.Output));
            }
            return Ok;
        }

        private static int RunConvert(PdfCliCommand cmd, PdfConvertMode mode, Action<string> log)
        {
            if (!Ghostscript.Available)
            {
                Say(log, "ERROR: Ghostscript not found (needed for this command)");
                return Failed;
            }
            using (AtomicOutput output = CopyToOutput(cmd.Inputs[0], cmd.Output))
            {
                if (!PdfConvert.ApplyUnpublished(output.TempPath, mode))
                {
                    Say(log, "ERROR: the engine could not perform this conversion");
                    return Failed;
                }
                output.Commit();
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

        /// <summary>Сжать готовый файл, если уровень задан; отсутствие движка явно сообщается.</summary>
        private static void Compress(string path, CompressionLevel level, Action<string> log)
        {
            if (level == CompressionLevel.None)
                return;
            if (!Ghostscript.Available)
            {
                Say(log, "WARNING: Ghostscript not found; output was left uncompressed: " +
                    Path.GetFileName(path));
                return;
            }
            GsRewriteResult result = PdfCompression.CompressDetailed(path, level);
            if (result == GsRewriteResult.Failed)
                Say(log, "WARNING: compression failed; output was left uncompressed: " +
                    Path.GetFileName(path));
            else if (result == GsRewriteResult.RejectedByPolicy)
                Say(log, "NOTE: compression did not make " + Path.GetFileName(path) + " smaller");
        }

        /// <summary>
        /// Подготовить копию во временном файле транзакции. Команды, меняющие файл,
        /// публикуют его только после полного завершения движка — исходный target и
        /// предыдущий результат остаются целыми при сбое или отмене.
        /// </summary>
        private static AtomicOutput CopyToOutput(string source, string output)
        {
            if (string.Equals(Path.GetFullPath(source), Path.GetFullPath(output), StringComparison.OrdinalIgnoreCase))
                throw new MergeException("input and output must be different files");
            string dir = Path.GetDirectoryName(Path.GetFullPath(output));
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);
            AtomicOutput transaction = new AtomicOutput(output);
            try
            {
                File.Copy(source, transaction.TempPath, true);
                return transaction;
            }
            catch
            {
                transaction.Dispose();
                throw;
            }
        }

        private static void Say(Action<string> log, string line)
        {
            if (log != null)
                log(line);
        }
    }
}
