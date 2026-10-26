using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelMerger
{
    /// <summary>Результат обработки одного исходного файла.</summary>
    public class FileResult
    {
        public string FileName;
        public string FullPath;  // полный путь источника (для повторной попытки)
        public string SheetName; // имя листа в итоговой книге (null, если пропущен)
        public bool Ok;
        public string Note;      // причина пропуска или предупреждение
        public bool Linkable;    // на лист можно сослаться гиперссылкой (есть ячейка A1)
        // Имя листа В ИСТОЧНИКЕ; null — пропущен файл целиком (не открылся, нет листов).
        // По нему повтор переносит только упавшие листы, а не весь файл заново.
        public string SourceSheet;
    }

    public class MergeResult
    {
        public readonly List<FileResult> Files = new List<FileResult>();
        public int OkCount;
        public int SkipCount;
        public bool Cancelled;
        public string OutputPath;
        public string TocError; // не null, если лист «Содержание» создать не удалось
        public bool MacrosCarried; // итог в формате с макросами получил VBA-код скопированных листов

        /// <summary>Число исходных файлов (в режиме «все листы» листов больше).</summary>
        public int FileCount
        {
            get
            {
                var files = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
                foreach (FileResult f in Files)
                    files.Add(FileKey(f));
                return files.Count;
            }
        }

        /// <summary>
        /// Пропуски по файлам, а не по записям: в режиме «все листы» SkipCount считает листы.
        /// skippedFiles — файлы, не давшие своду ни одного листа; lostSheets — непереносённые
        /// листы файлов, вошедших в свод частично. Чистая — под тест.
        /// </summary>
        public void CountSkips(out int skippedFiles, out int lostSheets)
        {
            var included = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            foreach (FileResult f in Files)
                if (f.Ok)
                    included.Add(FileKey(f));
            var skipped = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            lostSheets = 0;
            foreach (FileResult f in Files)
            {
                if (f.Ok)
                    continue;
                if (included.Contains(FileKey(f)))
                    lostSheets++;
                else
                    skipped.Add(FileKey(f));
            }
            skippedFiles = skipped.Count;
        }

        // Все записи одного исходного файла (в режиме «все листы» их несколько) — по его пути.
        private static string FileKey(FileResult f)
        {
            return f.FullPath ?? f.FileName ?? "";
        }
    }

    /// <summary>Ошибка объединения с понятным пользователю сообщением.</summary>
    public class MergeException : Exception
    {
        public MergeException(string message) : base(message) { }

        /// <summary>
        /// Можно ли завернуть исключение в понятное «файл повреждён/недоступен»: критические
        /// сбои процесса (OutOfMemory) маскировать нельзя — причина не в файле, и такое
        /// сообщение направит пользователя по ложному следу. Для when-фильтров catch.
        /// </summary>
        public static bool ShouldWrap(Exception ex)
        {
            // Отмену (OperationCanceledException) НЕ оборачиваем — это не ошибка файла, её
            // ловят формы отдельно; OOM тоже пробрасываем как есть (нельзя маскировать).
            return !(ex is OutOfMemoryException) && !(ex is OperationCanceledException);
        }
    }

    /// <summary>
    /// Файл подвесил Excel (Workbooks перестали отвечать) — внутренний сигнал
    /// к перезапуску экземпляра без этого файла. Наружу не выходит.
    /// </summary>
    internal class ExcelWedgedException : Exception
    {
        public readonly string FilePath;
        public readonly string Note; // примечание к файлу в следующем проходе (почему пропущен)
        public ExcelWedgedException(string filePath, string note = MergeService.WedgedNote)
        {
            FilePath = filePath;
            Note = note;
        }
    }

    /// <summary>
    /// Объединяет первый видимый лист каждого Excel-файла папки в один итоговый
    /// файл через COM-автоматизацию установленного Excel (перенос без потерь:
    /// форматирование, диаграммы, сводные таблицы, картинки). Поддерживает
    /// повторное слияние пропущенных файлов в существующий свод.
    /// Методы Merge/RetrySkipped должны вызываться в STA-потоке.
    /// </summary>
    public class MergeService
    {
        public event Action<int, int, string> Progress; // номер файла, всего, имя файла
        public event Action<string> Trace;              // пошаговая диагностика (используется в --cli)
        public event Action<FileResult> FileDone;
        public event Action Restarting;                 // экземпляр Excel перезапускается — прежние FileDone отменяются

        // Служебное имя листа-заглушки новой книги; резервируется в SheetNamer.
        private const string PlaceholderName = "zz_tmp_5f2a9c";

        private const string TocSheetName = "Содержание";
        private const int XlCalculationManual = -4135;
        private const int XlCalculationAutomatic = -4105;

        // Заведомо неверный пароль: защищённый файл даёт исключение,
        // а не модальный диалог, который повесил бы скрытый Excel.
        private const string WrongPassword = "\u0001\u0002";

        private static readonly string[] Extensions = { ".xlsx", ".xls", ".xlsm", ".xlsb" };

        /// <summary>Имя-основа листа в своде: в режиме «все листы» — «файл · лист».</summary>
        public static string SheetBaseName(string fileBaseName, string sheetName, bool allSheets)
        {
            if (!allSheets)
                return fileBaseName;
            return fileBaseName + " · " + sheetName;
        }

        private volatile bool _cancel;

        /// <summary>
        /// Отмена: останов после текущего файла, без сохранения результата. Файл, который Excel не
        /// закончил за отсрочку сторожа (<see cref="ComWatchdog.Limits.CancelMs"/>), прерывается
        /// остановкой процесса Excel. Запись свода, если уже началась, доводится до конца.
        /// </summary>
        public void Cancel()
        {
            _cancel = true;
        }

        /// <summary>Предел одного шага (файл, оглавление, запись): Excel, занятый дольше, считается зависшим.</summary>
        internal const int StepLimitMs = 10 * 60 * 1000;

        /// <summary>Пределы сторожа зависшего вызова Excel; тесты подставляют короткие.</summary>
        internal ComWatchdog.Limits WatchdogLimits = new ComWatchdog.Limits { StepMs = StepLimitMs };

        /// <summary>
        /// Шов для живых тестов: вызывается в шаге сторожа с экземпляром Excel и путём источника —
        /// перед его открытием, а с path = null — перед записью свода. Тест вешает здесь Excel
        /// модальным окном или долгим расчётом, как это делает «трудный» файл.
        /// </summary>
        internal Action<object, string> StepHookForTests { get; set; }

        /// <summary>
        /// Файлы Excel папки в естественном порядке; временные файлы «~$*»
        /// и сам итоговый файл исключаются.
        /// </summary>
        public static List<string> FindSourceFiles(string folder, string outputPath)
        {
            var files = new List<string>();
            string outputFull = outputPath != null ? Path.GetFullPath(outputPath) : null;
            foreach (string path in Directory.GetFiles(folder))
            {
                string name = Path.GetFileName(path);
                if (name.StartsWith("~$", StringComparison.Ordinal))
                    continue;
                string ext = Path.GetExtension(path).ToLowerInvariant();
                if (Array.IndexOf(Extensions, ext) < 0)
                    continue;
                if (outputFull != null &&
                    string.Equals(Path.GetFullPath(path), outputFull, StringComparison.OrdinalIgnoreCase))
                    continue;
                files.Add(path);
            }
            // Порядок как в Проводнике: «Отчет 2» раньше «Отчет 10».
            files.Sort(delegate(string a, string b)
            {
                return NaturalStringComparer.Instance.Compare(Path.GetFileName(a), Path.GetFileName(b));
            });
            return files;
        }

        /// <summary>
        /// Быстрая проверка, что итоговый файл можно записать. null — можно,
        /// иначе понятное пользователю сообщение. Вызывается до запуска Excel,
        /// чтобы о занятом файле стало известно сразу, а не после обработки
        /// всех исходных файлов. Окончательная защита — обработчик сохранения.
        /// </summary>
        public static string CheckOutputWritable(string outputPath)
        {
            try
            {
                string full = Path.GetFullPath(outputPath);
                string dir = Path.GetDirectoryName(full);
                if (dir == null || !Directory.Exists(dir))
                    return string.Format(Loc.T("err.merge.folderMissing"), dir);
                if (File.Exists(full))
                {
                    // Эксклюзивное открытие: открытый в Excel файл даст IOException.
                    using (new FileStream(full, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
                }
                else
                {
                    // Файла нет — пробное создание с немедленным удалением.
                    using (new FileStream(full, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { }
                    File.Delete(full);
                }
                return null;
            }
            catch (IOException)
            {
                return Loc.T("err.merge.outputBusy");
            }
            catch (UnauthorizedAccessException)
            {
                return Loc.T("err.merge.noWritePerm");
            }
            catch (Exception ex)
            {
                return string.Format(Loc.T("err.merge.outputNotWritable"), ShortMessage(ex));
            }
        }

        /// <summary>
        /// Объединяет заданные файлы (в указанном порядке) в новый итоговый файл.
        /// Итоговый файл сам исключается из источников, дубликаты убираются.
        /// </summary>
        public MergeResult Merge(IList<string> files, string outputPath, MergeOptions options)
        {
            if (options == null)
                throw new ArgumentNullException("options");

            List<string> sources = PrepareSourceList(files, outputPath);
            if (sources.Count == 0)
                throw new MergeException(Loc.T("err.merge.noFiles"));

            ValidateOutput(outputPath);
            return Run(outputPath, sources, options, null);
        }

        /// <summary>
        /// Готовит список источников: убирает сам итоговый файл (по полному пути)
        /// и дубликаты, сохраняя порядок. Чистая функция — покрыта тестами.
        /// </summary>
        public static List<string> PrepareSourceList(IList<string> files, string outputPath)
        {
            var result = new List<string>();
            if (files == null)
                return result;
            string outputFull = outputPath != null ? Path.GetFullPath(outputPath) : null;
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string path in files)
            {
                if (string.IsNullOrEmpty(path))
                    continue;
                string full = Path.GetFullPath(path);
                if (outputFull != null && string.Equals(full, outputFull, StringComparison.OrdinalIgnoreCase))
                    continue; // не сливать свод сам в себя
                if (seen.Add(full))
                    result.Add(path);
            }
            return result;
        }

        /// <summary>
        /// Повторяет только пропущенные файлы предыдущего прогона, дослияя их
        /// листы в уже существующий итоговый файл. Оглавление перестраивается.
        /// </summary>
        public MergeResult RetrySkipped(string outputPath, MergeOptions options, MergeResult previous)
        {
            if (options == null)
                throw new ArgumentNullException("options");
            if (previous == null)
                throw new ArgumentNullException("previous");
            if (!File.Exists(outputPath))
                throw new MergeException(Loc.T("err.merge.noOutput"));

            List<string> retryPaths = MergeRetry.Paths(previous);
            if (retryPaths.Count == 0)
                throw new MergeException(Loc.T("err.merge.nothingToRetry"));

            ValidateOutput(outputPath);
            return Run(outputPath, retryPaths, options, previous, MergeRetry.Sheets(previous));
        }

        private static void ValidateOutput(string outputPath)
        {
            if (OutputFormats.FileFormatFor(outputPath) == 0)
                throw new MergeException(string.Format(Loc.T("err.merge.badExtension"),
                    string.Join(", ", OutputFormats.Extensions) + "."));
            string lockError = CheckOutputWritable(outputPath);
            if (lockError != null)
                throw new MergeException(lockError);
        }

        private const int MaxWedgeRestarts = 3;
        private const long MinFreeBytes = 200L * 1024 * 1024; // 200 МБ на рабочем диске

        /// <summary>
        /// Сообщение о нехватке места (или null). Excel создаёт временные файлы на
        /// системном/temp-диске; при почти полном диске любой Workbooks.Open падает
        /// с невнятным «свойство Open» — понятнее сказать это заранее. Чистая, тестируемая.
        /// </summary>
        internal static string LowSpaceMessage(string root, long freeBytes)
        {
            if (freeBytes >= MinFreeBytes)
                return null;
            return string.Format(Loc.T("err.merge.lowSpace"), root, freeBytes / (1024 * 1024));
        }

        /// <summary>Предполётная проверка рабочих дисков (системный, temp, диск свода).</summary>
        private static string CheckWorkingSpace(string outputPath)
        {
            var roots = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string[] paths =
            {
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                Path.GetTempPath(),
                Path.GetDirectoryName(Path.GetFullPath(outputPath))
            };
            foreach (string p in paths)
            {
                if (string.IsNullOrEmpty(p))
                    continue;
                string root = Path.GetPathRoot(p);
                if (string.IsNullOrEmpty(root) || !roots.Add(root))
                    continue;
                try
                {
                    var drive = new DriveInfo(root);
                    if (drive.IsReady)
                    {
                        string msg = LowSpaceMessage(root, drive.AvailableFreeSpace);
                        if (msg != null)
                            return msg;
                    }
                }
                catch { } // диск недоступен для оценки — не мешаем работе
            }
            return null;
        }

        /// <summary>
        /// Движок слияния с самовосстановлением: если файл подвешивает Excel
        /// (Workbooks перестают отвечать), экземпляр перезапускается, а виновный
        /// файл исключается — без перезагрузки машины и без потери остальных.
        /// previous == null — новый свод; иначе дослияние в существующий файл.
        /// </summary>
        private MergeResult Run(string outputPath, List<string> files, MergeOptions options, MergeResult previous,
            Dictionary<string, HashSet<string>> onlySheets = null)
        {
            string spaceError = CheckWorkingSpace(outputPath);
            if (spaceError != null)
                throw new MergeException(spaceError);

            // Файлы, подвесившие Excel, — с примечанием, почему пропущены.
            var excluded = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int restart = 0; ; restart++)
            {
                try
                {
                    return RunOnce(outputPath, files, options, previous, excluded, onlySheets);
                }
                catch (ExcelWedgedException wedge)
                {
                    // Виновный файл в чёрный список и перезапуск свежего Excel.
                    // Уже исключённый файл снова виноват быть не может — защита от зацикливания.
                    if (restart >= MaxWedgeRestarts || excluded.ContainsKey(wedge.FilePath))
                        throw new MergeException(string.Format(Loc.T("err.merge.excelUnstable"),
                            Path.GetFileName(wedge.FilePath)));
                    excluded.Add(wedge.FilePath, wedge.Note);
                    RaiseTrace("перезапуск Excel после зависания на " + Path.GetFileName(wedge.FilePath));
                    var restarting = Restarting;
                    if (restarting != null)
                        restarting(); // GUI очистит строки: прошлый проход отменён, будет переотправлен
                }
            }
        }

        /// <summary>
        /// Один проход слияния в отдельном экземпляре Excel. Файлы из excluded
        /// пропускаются без открытия (ранее подвесили Excel) с их примечанием. Бросает
        /// ExcelWedgedException, если очередной файл подвесил Workbooks или на нём сработал
        /// сторож вызова (<see cref="ComWatchdog"/>): модальное окно или предел времени шага.
        /// Все вызовы Excel идут внутри шагов сторожа; запись свода — шаг без отмены.
        /// </summary>
        private MergeResult RunOnce(string outputPath, List<string> files, MergeOptions options, MergeResult previous,
            Dictionary<string, string> excluded, Dictionary<string, HashSet<string>> onlySheets)
        {
            bool intoExisting = previous != null;
            bool legacy = OutputFormats.FileFormatFor(outputPath) == OutputFormats.Xls;
            var copyTarget = new CopyTarget
            {
                Legacy = legacy,
                SmallGrid = intoExisting && legacy, // существующий .xls открывается в режиме совместимости — сетка старая
                OnlySheets = onlySheets
            };
            var attempt = new MergeResult();
            attempt.OutputPath = outputPath;

            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new MergeException(Loc.T("err.merge.excelMissing"));

            dynamic excel = null;
            dynamic target = null;
            OwnedOfficeProcess excelProcess = null; // страховка от осиротевшего EXCEL.EXE (см. finally)
            ComWatchdog watchdog = null;            // сторож зависшего вызова этого экземпляра
            bool askToUpdateLinks = true; // настройка экземпляра до свода — вернуть, если он достанется человеку
            ComMessageFilter.Register(); // авто-повтор вызовов, отклонённых занятым Excel
            try
            {
                excel = Activator.CreateInstance(excelType);
                excelProcess = OwnExcelProcess((object)excel);
                watchdog = new ComWatchdog(excelProcess, delegate { return _cancel; }, WatchdogLimits);
                watchdog.Begin(true); // подготовка: на диске ничего не меняется — отмена безопасна
                excel.Visible = false;
                excel.DisplayAlerts = false;
                excel.ScreenUpdating = false;
                excel.EnableEvents = false;
                try
                {
                    askToUpdateLinks = (bool)excel.AskToUpdateLinks;
                    excel.AskToUpdateLinks = false;
                }
                catch { }
                try { excel.AutomationSecurity = 3; } catch { } // msoAutomationSecurityForceDisable: макросы источников не выполняются

                WaitExcelReady((object)excel); // под нагрузкой Excel готов к Workbooks не сразу

                if (intoExisting)
                {
                    target = OpenExistingDigest((object)excel, outputPath);
                }
                else
                {
                    target = excel.Workbooks.Add(-4167); // xlWBATWorksheet: новая книга с одним листом
                    dynamic placeholder = target.Sheets[1];
                    placeholder.Name = PlaceholderName;
                }

                // Без пересчёта после вставки каждого листа — заметно быстрее на формулах.
                try { excel.Calculation = XlCalculationManual; } catch { }

                var namer = new SheetNamer();
                if (intoExisting)
                {
                    // Старое «Содержание» удаляется — будет перестроено по общему итогу.
                    if (options.AddToc)
                        try { target.Sheets[TocSheetName].Delete(); } catch { }
                    ReserveExistingSheetNames((object)target, namer);
                }
                else
                {
                    namer.Reserve(PlaceholderName);
                }
                if (options.AddToc)
                    namer.Reserve(TocSheetName); // файл с таким именем получит суффикс _2
                watchdog.EndStep();

                int index = 0;
                foreach (string path in files)
                {
                    if (_cancel)
                    {
                        attempt.Cancelled = true;
                        break;
                    }
                    index++;
                    RaiseProgress(index, files.Count, Path.GetFileName(path));

                    string full = Path.GetFullPath(path);
                    string excludedNote;
                    if (excluded.TryGetValue(full, out excludedNote))
                    {
                        // Файл ранее подвесил Excel — не открываем повторно.
                        FileResult skip = Skipped(Path.GetFileName(path), path, excludedNote);
                        attempt.Files.Add(skip);
                        attempt.SkipCount++;
                        RaiseFileDone(skip);
                        continue;
                    }

                    watchdog.Begin(true);
                    List<FileResult> frs;
                    try
                    {
                        frs = CopySheets((object)excel, (object)target, path, namer, options, copyTarget);
                    }
                    catch (Exception ex)
                    {
                        // Защита от ошибок вне per-file обработчика (например, сбой COM-привязки).
                        var fr = new FileResult();
                        fr.FileName = Path.GetFileName(path);
                        fr.FullPath = path;
                        fr.Note = "не удалось обработать (" + ShortMessage(ex) + ")";
                        frs = new List<FileResult> { fr };
                    }
                    bool responsive = IsExcelResponsive((object)excel); // тоже вызов Excel — внутри шага
                    WatchdogStop stop = watchdog.End();
                    if (stop == WatchdogStop.Cancel)
                    {
                        // Файл прерван остановкой Excel: его ошибки — следствие отмены, не показываем.
                        RaiseTrace("отмена: Excel не закончил «" + Path.GetFileName(path) + "» вовремя — процесс остановлен");
                        attempt.Cancelled = true;
                        break;
                    }
                    if (stop != WatchdogStop.None)
                    {
                        string note = StalledNote(stop, watchdog.DialogTitle, watchdog.StepLimitMs);
                        RaiseTrace("сторож: " + note + " на «" + Path.GetFileName(path) + "» — процесс Excel остановлен");
                        throw new ExcelWedgedException(full, note); // перезапуск без этого файла
                    }
                    foreach (FileResult fr in frs)
                    {
                        attempt.Files.Add(fr);
                        if (fr.Ok) attempt.OkCount++; else attempt.SkipCount++;
                        RaiseFileDone(fr);
                    }

                    // Отказоустойчивость: если файл подвесил Excel (Workbooks не
                    // отвечают), уходим на перезапуск экземпляра без него — иначе
                    // все следующие файлы были бы ложно «пропущены».
                    if (!responsive)
                        throw new ExcelWedgedException(full);
                }

                if (attempt.Cancelled)
                    return attempt; // без сохранения: существующий свод не тронут

                if (!intoExisting && attempt.OkCount == 0)
                    throw new MergeException(Loc.T("err.merge.noSheets"));

                MergeResult final = intoExisting
                    ? MergeRetry.Combine(previous, attempt.Files)
                    : attempt;

                if (intoExisting && attempt.OkCount == 0)
                    return final; // ни одна повторная попытка не удалась — файл не трогаем

                watchdog.Begin(true); // оглавление и подготовка к записи: свод ещё в памяти
                if (!intoExisting)
                    try { target.Sheets[PlaceholderName].Delete(); } catch { }

                // Итоговый файл сохраняется со стандартным авторасчётом (один пересчёт здесь).
                try { excel.Calculation = XlCalculationAutomatic; } catch { }

                // Ws.Copy переносит и модуль кода листа. В .xlsx он отбрасывается при сохранении,
                // а форматы с макросами донесут чужой VBA-код получателю свода — предупреждаем.
                // HasVBProject не требует доверия к объектной модели VBA (на живом Excel
                // VBProject отказал «не является доверенным», HasVBProject прочитался); try —
                // на случай политики, закрывающей и его.
                if (KeepsMacros(outputPath))
                {
                    try { final.MacrosCarried = (bool)target.HasVBProject; } catch { }
                }

                if (options.AddToc)
                {
                    try
                    {
                        TocBuilder.Build((object)target, TocSheetName, final.Files);
                        TocBuilder.AddReturnButtons((object)target, TocSheetName); // ссылка на оглавление на каждом листе
                    }
                    catch (Exception ex)
                    {
                        final.TocError = string.Format(Loc.T("err.merge.tocFailed"), ShortMessage(ex));
                    }
                }

                watchdog.EndStep();

                // Точка невозврата: дальше запись. Отмена, пришедшая до неё, — отмена: файл не создаётся
                // (существующий свод не меняется), как и при отмене между файлами.
                if (_cancel)
                {
                    attempt.Cancelled = true;
                    return attempt;
                }
                watchdog.Begin(false); // запись не прерываем отменой: остановить можно только зависшую
                try
                {
                    if (StepHookForTests != null)
                        StepHookForTests((object)excel, null);
                    if (intoExisting)
                        target.Save(); // формат существующего файла сохраняется
                    else
                        target.SaveAs(outputPath, OutputFormats.FileFormatFor(outputPath));
                }
                catch (Exception ex)
                {
                    throw new MergeException(string.Format(Loc.T("err.merge.saveFailed"), DiskSpace.Describe(ex, outputPath)));
                }
                // Запись вернулась без ошибки — файл записан, даже если сторож остановил Excel в тот
                // же миг: успех не перечёркиваем (закрытие ниже переживёт остановленный процесс).
                watchdog.End();

                return final;
            }
            catch (Exception ex) when (watchdog != null && watchdog.Stop != WatchdogStop.None && !(ex is ExcelWedgedException))
            {
                // Сторож остановил Excel вне файла (запуск, свод, оглавление, запись): любая ошибка
                // после этого — следствие остановки; человеку — её причина, а не «сбой RPC».
                if (watchdog.Stop == WatchdogStop.Cancel)
                {
                    RaiseTrace("отмена: Excel не ответил вовремя — процесс остановлен");
                    attempt.Cancelled = true;
                    return attempt;
                }
                string message = watchdog.Describe("Excel");
                RaiseTrace("сторож: " + message);
                throw new MergeException(message);
            }
            finally
            {
                // Закрытие тоже под сторожем: Quit, упёршийся в модальное окно, иначе висел бы вечно.
                if (watchdog != null)
                    watchdog.Begin(true);
                // Статические ссылки: после Close/Quit никаких динамических операций
                // над объектами (см. комментарий у ComSafe.Release).
                object targetObj = target;
                if (targetObj != null)
                {
                    try { target.Close(false); } catch { }
                    ComSafe.Release(targetObj);
                }
                object excelObj = excel;
                int foreignBooks = 0;
                if (excelObj != null)
                {
                    foreignBooks = QuitOrHandOver(excelObj, askToUpdateLinks);
                    ComSafe.Release(excelObj);
                }
                if (watchdog != null)
                    watchdog.Dispose(); // дальше вызовов Excel нет
                ComSafe.Collect();
                ComMessageFilter.Revoke();
                if (foreignBooks > 0)
                    RaiseTrace("в экземпляре Excel открыты книги пользователя (" + foreignBooks +
                        ") — Excel оставлен открытым и показан");
                // Последний рубеж: заклинивший экземпляр не выходит и после Quit/Release —
                // ждём штатного выхода и завершаем только СВОЙ процесс. Отданный человеку —
                // уже не свой: его не трогаем.
                else if (excelProcess != null && excelProcess.KillIfLingering(ExcelExitWaitMs))
                    RaiseTrace("Excel не завершился сам — процесс " + excelProcess.Id + " остановлен");
            }
        }

        private const int ExcelExitWaitMs = 5000;

        /// <summary>
        /// Закрыть свой экземпляр Excel — или, если в нём остались ЧУЖИЕ книги, отдать его
        /// человеку. Вызывается, когда свод и источники уже закрыты: всё, что ещё открыто, —
        /// книги, которые Windows передала в наш скрытый экземпляр (двойной щелчок по файлу во
        /// время свода). Quit при DisplayAlerts = false закрыл бы их без сохранения, а
        /// страховка KillIfLingering затем завершила бы процесс. Вместо этого экземпляр
        /// становится обычным окном Excel с прежними настройками (askToUpdateLinks — значение
        /// до свода). Возвращает число чужих книг (0 — экземпляр закрыт). После Quit
        /// динамических вызовов нет (правила ComSafe).
        /// </summary>
        internal static int QuitOrHandOver(object excelObj, bool askToUpdateLinks = true)
        {
            dynamic excel = excelObj;
            int open = 0;
            try { open = (int)excel.Workbooks.Count; }
            catch { open = 0; } // экземпляр не отвечает — закрываем, как раньше
            if (open == 0)
            {
                try { excel.Quit(); } catch { }
                return 0;
            }
            try { excel.Calculation = XlCalculationAutomatic; } catch { }
            try { excel.ScreenUpdating = true; } catch { }
            try { excel.EnableEvents = true; } catch { }
            try { excel.DisplayAlerts = true; } catch { }
            try { excel.AskToUpdateLinks = askToUpdateLinks; } catch { }
            try { excel.Visible = true; } catch { }
            try { excel.UserControl = true; } catch { } // без этого освобождение ссылок закрыло бы Excel
            return open;
        }

        /// <summary>
        /// Страховка от осиротевшего EXCEL.EXE — процесс по окну приложения. null — определить
        /// не удалось (сбой Hwnd, недействительное окно): свод идёт без страховки, но не молча —
        /// иначе зависший после прогона Excel в трассе ничем не объяснялся бы.
        /// </summary>
        internal OwnedOfficeProcess OwnExcelProcess(object excelObj)
        {
            dynamic excel = excelObj;
            string reason = "окно приложения недействительно";
            try
            {
                OwnedOfficeProcess owned = OwnedOfficeProcess.FromWindow((long)excel.Hwnd);
                if (owned != null)
                    return owned;
            }
            catch (Exception ex)
            {
                reason = ShortMessage(ex);
            }
            RaiseTrace("процесс Excel не определён (" + reason + ") — страховки от зависшего EXCEL.EXE нет");
            return null;
        }

        /// <summary>Формат итогового файла хранит VBA (.xlsm/.xlsb/.xls); .xlsx код отбрасывает.</summary>
        internal static bool KeepsMacros(string outputPath)
        {
            int format = OutputFormats.FileFormatFor(outputPath);
            return format != 0 && format != 51; // 51 — xlOpenXMLWorkbook (.xlsx)
        }

        /// <summary>
        /// Предупреждение о макросах для отчёта и записки — по-русски, как все генерируемые
        /// документы. Русская строка каталога (err.merge.macrosCarried) — та же фраза литералом
        /// (выгрузка подписей читает только литералы); расхождение ловит тест.
        /// </summary>
        internal const string MacrosWarningText =
            "итоговый файл содержит макросы VBA из скопированных листов — для передачи свода безопаснее формат .xlsx";

        /// <summary>
        /// Открывает существующий свод для дослияния. Заведомо неверные пароли — как у
        /// источников: свод, который успели защитить паролем (открытия или записи), даёт
        /// исключение, а не модальный запрос в невидимом Excel, ждущий вечно. Проверено на
        /// живом Excel: без паролей Open такого свода висел; с ними оба вида защиты дают
        /// исключение, а свод без пароля открывается для записи как прежде.
        /// </summary>
        private static object OpenExistingDigest(object excelObj, string outputPath)
        {
            dynamic excel = excelObj;
            try
            {
                return excel.Workbooks.Open(
                    Filename: outputPath,
                    UpdateLinks: 0,
                    ReadOnly: false,
                    Password: WrongPassword,
                    WriteResPassword: WrongPassword,
                    IgnoreReadOnlyRecommended: true,
                    Notify: false,
                    AddToMru: false);
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                throw new MergeException(DigestOpenError(Unwrap(ex).Message));
            }
        }

        /// <summary>Понятное сообщение о неудачном открытии свода для дослияния. Чистая — под тест.</summary>
        internal static string DigestOpenError(string excelMessage)
        {
            if (IsPasswordError(excelMessage))
                return Loc.T("err.merge.outputPassword");
            return string.Format(Loc.T("err.merge.outputOpenFailed"), ShortText(excelMessage));
        }

        /// <summary>
        /// Ждёт готовности только что созданного Excel: свежий экземпляр под
        /// нагрузкой отвечает на Workbooks не сразу, и первый вызов Open иначе
        /// падает с «Невозможно получить свойство Open класса Workbooks».
        /// </summary>
        private static void WaitExcelReady(object excelObj)
        {
            dynamic excel = excelObj;
            for (int i = 0; i < 60; i++) // до ~6 секунд
            {
                try
                {
                    int probe = (int)excel.Workbooks.Count;
                    System.Threading.Thread.Sleep(300); // дать Excel окончательно инициализироваться
                    return;
                }
                catch
                {
                    System.Threading.Thread.Sleep(100);
                }
            }
        }

        private static void ReserveExistingSheetNames(object targetObj, SheetNamer namer)
        {
            dynamic target = targetObj;
            dynamic sheets = target.Sheets;
            int count = (int)sheets.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic s = sheets[i];
                namer.Reserve((string)s.Name);
            }
        }

        internal const long LegacyRows = 65536; // сетка формата .xls (OutputFormats.Xls)
        internal const long LegacyColumns = 256;

        /// <summary>Куда переносятся листы: от этого зависит, что проверять у источника.</summary>
        private sealed class CopyTarget
        {
            public bool Legacy;     // свод сохраняется в .xls
            public bool SmallGrid;  // сама книга свода уже в старой сетке (дослияние в .xls)
            public Dictionary<string, HashSet<string>> OnlySheets; // повтор: только эти листы файла
        }

        /// <summary>
        /// Примечание о данных за пределами сетки .xls; null — лист в неё помещается. Без него
        /// Excel при сохранении в .xls молча отбрасывал всё, что дальше 65 536-й строки и
        /// столбца IV (DisplayAlerts выключен, вопроса проверки совместимости нет). Чистая — под тест.
        /// </summary>
        internal static string LegacyGridNote(long lastRow, long lastColumn)
        {
            return lastRow > LegacyRows || lastColumn > LegacyColumns
                ? "данные за пределами 65 536 строк × 256 столбцов не попадут в свод .xls"
                : null;
        }

        internal const string BigGridIntoLegacyNote =
            "лист книги нового формата нельзя добавить в свод .xls — соберите свод в .xlsx";

        /// <summary>
        /// Последняя заполненная строка и столбец листа (1-based). UsedRange дёшев, но завышает
        /// из-за одного только оформления, поэтому, если он выходит за сетку .xls, границу
        /// уточняет поиск последнего значения или формулы.
        /// </summary>
        private static void LastUsedCell(dynamic ws, out long lastRow, out long lastColumn)
        {
            dynamic used = ws.UsedRange;
            lastRow = (long)(int)used.Row + (long)(int)used.Rows.Count - 1;
            lastColumn = (long)(int)used.Column + (long)(int)used.Columns.Count - 1;
            if (LegacyGridNote(lastRow, lastColumn) == null)
                return;
            const int xlFormulas = -4123, xlPart = 2, xlByRows = 1, xlByColumns = 2, xlPrevious = 2;
            dynamic byRow = ws.Cells.Find("*", Type.Missing, xlFormulas, xlPart, xlByRows, xlPrevious);
            dynamic byColumn = ws.Cells.Find("*", Type.Missing, xlFormulas, xlPart, xlByColumns, xlPrevious);
            lastRow = byRow == null ? 0 : (long)(int)byRow.Row;
            lastColumn = byColumn == null ? 0 : (long)(int)byColumn.Column;
        }

        private List<FileResult> CopySheets(object excelObj, object targetObj, string path, SheetNamer namer,
            MergeOptions options, CopyTarget into)
        {
            dynamic excel = excelObj;
            dynamic target = targetObj;
            string fileName = Path.GetFileName(path);
            string fileBase = Path.GetFileNameWithoutExtension(path);
            var results = new List<FileResult>();

            // Отсекаем до Excel то, что подвешивает Workbooks.Open, а с ним и все
            // следующие файлы: битый файл (не ZIP и не OLE2) и зашифрованную
            // парольную книгу. Незашифрованный OOXML — всегда ZIP; контейнер OLE2
            // у .xlsx/.xlsm/.xlsb — шифрованная книга или чужой OLE2 под этим расширением.
            ExcelContainer container = FileSignature.Detect(path);
            if (container == ExcelContainer.NotExcel)
            {
                results.Add(Skipped(fileName, path, "файл повреждён или не является книгой Excel"));
                return results;
            }
            if (container == ExcelContainer.Ole2 &&
                !string.Equals(Path.GetExtension(path), ".xls", StringComparison.OrdinalIgnoreCase))
            {
                results.Add(Skipped(fileName, path, Ole2MismatchNote(FileSignature.HasEncryptedPackage(path))));
                return results;
            }

            dynamic source = null;
            try
            {
                RaiseTrace("open: " + fileName);
                try
                {
                    if (StepHookForTests != null)
                        StepHookForTests(excelObj, path);
                    source = OpenSource(excel, path);
                }
                catch (Exception ex)
                {
                    results.Add(Skipped(fileName, path, ClassifyError(ex)));
                    return results;
                }

                RaiseTrace("opened");
                List<object> sheets = VisibleSheets(source, options.AllSheets);
                if (sheets.Count == 0)
                {
                    results.Add(Skipped(fileName, path, "в файле нет видимых листов"));
                    return results;
                }

                HashSet<string> wanted = null;
                if (options.AllSheets && into.OnlySheets != null)
                    into.OnlySheets.TryGetValue(MergeRetry.Key(path), out wanted); // null — файл целиком
                if (wanted != null)
                {
                    var pending = new HashSet<string>(wanted, StringComparer.OrdinalIgnoreCase);
                    sheets = sheets.FindAll(delegate(object s)
                    {
                        string name;
                        try { name = (string)((dynamic)s).Name; }
                        catch { return false; }
                        return pending.Remove(name);
                    });
                    foreach (string missing in pending)
                    {
                        FileResult gone = Skipped(fileName, path, "лист «" + missing + "» больше не найден в файле");
                        gone.SourceSheet = missing;
                        results.Add(gone);
                    }
                }

                bool fileNotesPending = true;
                foreach (object wsObj in sheets)
                {
                    dynamic ws = wsObj;
                    var fr = new FileResult();
                    fr.FileName = fileName;
                    fr.FullPath = path;
                    try { fr.SourceSheet = (string)ws.Name; }
                    catch { } // имя не прочиталось — повтор возьмёт файл целиком
                    try
                    {
                        string baseName = SheetBaseName(fileBase, fr.SourceSheet ?? (string)ws.Name, options.AllSheets);
                        string gridNote = null;
                        if (into.Legacy)
                        {
                            bool bigGrid = false;
                            try { bigGrid = (long)(int)ws.Rows.Count > LegacyRows; }
                            catch { } // лист-диаграмма: ячеек нет, сетке нечего превышать
                            if (bigGrid && into.SmallGrid)
                            {
                                // Excel отказал бы в Copy (1004) своим длинным текстом; причина известна заранее.
                                fr.Note = BigGridIntoLegacyNote;
                                results.Add(fr);
                                continue;
                            }
                            if (bigGrid)
                            {
                                long lastRow, lastColumn;
                                try
                                {
                                    LastUsedCell(ws, out lastRow, out lastColumn);
                                    gridNote = LegacyGridNote(lastRow, lastColumn);
                                }
                                catch { } // размер не выяснился — сохраняем как раньше, без примечания
                            }
                        }
                        dynamic last = target.Sheets[target.Sheets.Count];
                        ws.Copy(Type.Missing, last); // After: в конец итоговой книги
                        dynamic copied = target.Sheets[target.Sheets.Count];

                        string name = namer.Next(baseName);
                        try
                        {
                            copied.Name = name;
                        }
                        catch (Exception)
                        {
                            try { copied.Delete(); } catch { }
                            fr.Note = "не удалось назначить имя листа «" + name + "»";
                            results.Add(fr);
                            continue;
                        }

                        RaiseTrace("renamed: " + name);
                        fr.SheetName = name;
                        fr.Ok = true;
                        if (gridNote != null)
                            AppendNote(fr, gridNote);

                        // Лист-диаграмма не имеет ячеек: гиперссылка на A1 невозможна.
                        try { dynamic probe = copied.Range("A1"); fr.Linkable = true; }
                        catch { fr.Linkable = false; }

                        if (options.ValuesOnly && fr.Linkable)
                        {
                            bool replaced;
                            try { replaced = ReplaceFormulasWithValues((object)copied); }
                            catch (Exception) { replaced = false; }
                            if (!replaced)
                                AppendNote(fr, "формулы заменены не полностью");
                        }

                        if (fileNotesPending)
                        {
                            // Предупреждения уровня файла — на первом его листе.
                            try
                            {
                                if ((bool)source.HasVBProject)
                                    AppendNote(fr, "файл содержит макросы (не выполнялись)");
                            }
                            catch { }
                            try
                            {
                                if ((bool)source.Date1904)
                                    AppendNote(fr, "источник использует систему дат 1904 — проверьте даты");
                            }
                            catch { }
                            fileNotesPending = false;
                        }
                    }
                    catch (Exception ex)
                    {
                        fr.Ok = false;
                        fr.Note = "лист не перенесён (" + ShortMessage(ex) + ")";
                    }
                    results.Add(fr);
                }
                return results;
            }
            finally
            {
                // Статическая ссылка: после Close никаких динамических операций над source.
                object sourceObj = source;
                if (sourceObj != null)
                {
                    try { source.Close(false); } catch { }
                    ComSafe.Release(sourceObj);
                }
                RaiseTrace("closed source");
            }
        }

        /// <summary>
        /// Открывает книгу-источник с повтором транзиентных сбоев COM (Excel занят
        /// под нагрузкой). Реальные проблемы файла (формат, пароль, повреждение)
        /// пробрасываются сразу, без бесполезных повторов.
        /// </summary>
        private static dynamic OpenSource(dynamic excel, string path)
        {
            Exception last = null;
            for (int attempt = 0; attempt < 4; attempt++)
            {
                try
                {
                    return excel.Workbooks.Open(
                        Filename: path,
                        UpdateLinks: 0,
                        ReadOnly: true,
                        Password: WrongPassword,
                        IgnoreReadOnlyRecommended: true,
                        Notify: false,
                        AddToMru: false);
                }
                catch (Exception ex)
                {
                    last = ex;
                    // Проблема самого файла или процесс Excel завершён (сторож, сбой) —
                    // повторять бессмысленно.
                    if (IsPermanentOpenError(Unwrap(ex).Message) || IsServerGone(Unwrap(ex).HResult))
                        throw;
                    System.Threading.Thread.Sleep(500); // транзиентный сбой — подождать и повторить
                }
            }
            throw last;
        }

        /// <summary>
        /// Ошибка открытия — свойство самого файла (пароль/формат/повреждение), ретраи не помогут.
        /// Excel не даёт различимых HRESULT (всё DISP_E_EXCEPTION), поэтому по тексту сообщения;
        /// паттерны и русского, и английского Office — иначе на английском 4 пустых ретрая по
        /// 500 мс и общая формулировка вместо причины. Чистая — под тест.
        /// </summary>
        internal static bool IsPermanentOpenError(string message)
        {
            string m = (message ?? "").ToLowerInvariant();
            return IsPasswordError(m)
                || m.Contains("недопустим") || m.Contains("поврежд")
                || m.Contains("format") || m.Contains("invalid") || m.Contains("not valid")
                || m.Contains("damag") || m.Contains("corrupt");
        }

        /// <summary>Сообщение Excel о пароле (русский и английский Office) — для источников и свода.</summary>
        private static bool IsPasswordError(string message)
        {
            string m = (message ?? "").ToLowerInvariant();
            return m.Contains("парол") || m.Contains("password");
        }

        // Примечания к файлам (Note) НАРОЧНО не через Loc: те же строки попадают в отчёт
        // (ReportWriter) и записку Word (NoteText), которые по политике приложения всегда
        // русские, как все генерируемые документы; список в UI показывает их без перевода,
        // чтобы экран, отчёт и записка совпадали дословно.
        private static FileResult Skipped(string fileName, string path, string note)
        {
            var fr = new FileResult();
            fr.FileName = fileName;
            fr.FullPath = path;
            fr.Note = note;
            return fr;
        }

        /// <summary>
        /// Примечание к OLE2-файлу под расширением OOXML. Поток EncryptedPackage — книга
        /// зашифрована паролем; без него это другой формат (обычно переименованный .xls),
        /// и «защищён паролем» было бы неправдой. Чистая — под тест.
        /// </summary>
        internal static string Ole2MismatchNote(bool encrypted)
        {
            return encrypted
                ? PasswordNote
                : "формат файла не соответствует расширению (возможно, это переименованная книга .xls)";
        }

        private const string PasswordNote = "файл защищён паролем";

        /// <summary>
        /// Файл ранее подвесил Excel. Без «пропущен:» — статус пропуска добавляют сами строка
        /// списка и оглавление (иначе там было бы «пропущен: пропущен: …»).
        /// </summary>
        internal const string WedgedNote = "вызывал зависание Excel";

        /// <summary>
        /// Примечание к файлу, на котором сторож остановил Excel (Dialog или Timeout; без «пропущен:»,
        /// как <see cref="WedgedNote"/>): человек видит, что именно случилось, — окно с вопросом,
        /// которое в скрытом Excel некому закрыть, или бесконечную работу. Чистая — под тест.
        /// </summary>
        internal static string StalledNote(WatchdogStop stop, string dialogTitle, int stepLimitMs)
        {
            if (stop == WatchdogStop.Timeout)
                return "Excel не ответил за " + ComWatchdog.LimitMinutes(stepLimitMs) + " мин";
            if (stop == WatchdogStop.Dialog && !string.IsNullOrEmpty(dialogTitle))
                return "Excel ждал ответа в окне «" + dialogTitle + "»";
            if (stop == WatchdogStop.Dialog)
                return "Excel ждал ответа в диалоговом окне";
            return WedgedNote;
        }

        /// <summary>
        /// Процесс-сервер COM завершён или недоступен (его остановил сторож или он упал): вызовы в
        /// него не пройдут, повторять бессмысленно. RPC_S_CALL_FAILED, RPC_S_SERVER_UNAVAILABLE,
        /// RPC_E_DISCONNECTED, RPC_E_SERVER_DIED, RPC_E_SERVER_DIED_DNE. Чистая — под тест.
        /// </summary>
        internal static bool IsServerGone(int hresult)
        {
            switch ((uint)hresult)
            {
                case 0x800706BE:
                case 0x800706BA:
                case 0x80010108:
                case 0x80010007:
                case 0x80010012:
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Есть ли формулы по значению UsedRange.HasFormula: true — все ячейки, false — ни
        /// одной, null (DBNull) — часть. В отличие от SpecialCells, читается и на защищённом
        /// листе (проверено на живом Excel). Чистая — под тест.
        /// </summary>
        internal static bool AnyFormula(object hasFormula)
        {
            return !(hasFormula is bool) || (bool)hasFormula;
        }

        /// <summary>
        /// Заменяет формулы листа их вычисленными (кэшированными) значениями.
        /// Обрабатываются только области с формулами (SpecialCells) — массовое
        /// присваивание по областям, а не по ячейкам. Область с объединёнными
        /// ячейками не принимает массив — для неё поячеечный фолбэк через
        /// верхнюю левую ячейку MergeArea. false — формулы остались: иначе отчёт обещал
        /// бы значения, а в своде были бы формулы. Проверено на живом Excel: защищённый
        /// лист отказывает в SpecialCells, а формулу массива, рассечённую областями
        /// SpecialCells, не принимают ни область, ни отдельная ячейка — раньше оба случая
        /// молча сходили за «заменено».
        /// </summary>
        private static bool ReplaceFormulasWithValues(object sheetObj)
        {
            dynamic sheet = sheetObj;
            if (!AnyFormula((object)sheet.UsedRange.HasFormula))
                return true; // формул на листе нет

            dynamic formulaCells = null;
            try
            {
                formulaCells = sheet.UsedRange.SpecialCells(-4123); // xlCellTypeFormulas
            }
            catch
            {
                // Формулы есть, а выбрать их нельзя (защищённый лист). Защиту не снимаем:
                // пароль неизвестен, и это решение автора книги.
                return false;
            }
            dynamic areas = formulaCells.Areas;
            int count = (int)areas.Count;
            // Неудавшийся разлив возвращён формулой: общий проход идёт дальше (остальные формулы,
            // в том числе ссылки на источник, обязаны стать значениями), но области с уцелевшим
            // разливом заменяются по ячейкам в обход якоря — иначе его значение стёрло бы разлив.
            bool spillsKept = !ReplaceSpills(sheet, areas, count);
            for (int i = 1; i <= count; i++)
            {
                dynamic area = areas[i];
                if (spillsKept && HasAnySpill(area))
                {
                    ReplaceCellByCell(area, true);
                    continue;
                }
                try
                {
                    object values = area.Value2;
                    if (values != null)
                        area.Value2 = CellText.EscapeValues(values); // строковые результаты — буквально, без повторного разбора
                }
                catch (Exception)
                {
                    ReplaceCellByCell(area);
                }
            }

            // Контроль: поячеечный фолбэк глотает отказы (часть массива, особая ячейка) —
            // оставшиеся формулы должны попасть в примечание, а не в «заменено».
            return !AnyFormula((object)sheet.UsedRange.HasFormula);
        }

        /// <summary>
        /// Динамические массивы (Excel 2021 / Microsoft 365): формула живёт ТОЛЬКО в ячейке-якоре,
        /// а разлив — значения без формул. SpecialCells(формулы) даёт один якорь, запись его
        /// значения стирает разлив, и данные пропадали, причём проверка «формул не осталось»
        /// честно проходила. Поэтому разливы заменяются ДО общего прохода и целиком:
        /// читаем значения всего разлива, пишем значение якоря (разлив исчезает) и затем весь
        /// прямоугольник по прежнему адресу. Не записалось — формула якоря возвращается, и
        /// лист получает примечание «заменены не полностью», а не тихую потерю.
        ///
        /// Excel без динамических массивов (2019 и старше) свойства HasSpill не знает: чтение
        /// бросает, и лист идёт прежним путём. Свойство агрегатное (true/false/null для
        /// диапазона), поэтому по ячейкам идём только в областях, где разлив вообще есть.
        /// </summary>
        private static bool ReplaceSpills(dynamic sheet, dynamic areas, int count)
        {
            object any;
            try { any = sheet.UsedRange.HasSpill; }
            catch (Exception ex) when (!(ex is OutOfMemoryException)) { return true; } // нет динамических массивов
            if (!AnyFormula(any))
                return true;
            var anchors = new List<object>();
            try
            {
                for (int i = 1; i <= count; i++)
                {
                    dynamic area = areas[i];
                    if (!HasAnySpill(area))
                        continue;
                    int rows = (int)area.Rows.Count, cols = (int)area.Columns.Count;
                    for (int r = 1; r <= rows; r++)
                        for (int c = 1; c <= cols; c++)
                        {
                            dynamic cell = area.Cells[r, c];
                            if (IsSpill(cell))
                                anchors.Add((object)cell); // ячейка с формулой внутри разлива — его якорь
                        }
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return false; // ничего ещё не записано; разливы остаются формулами
            }
            bool all = true;
            foreach (object anchorObj in anchors)
            {
                dynamic anchor = anchorObj;
                string address, formula;
                object values;
                try
                {
                    dynamic spill = anchor.SpillingToRange;
                    address = (string)spill.Address;
                    formula = (string)anchor.Formula2;
                    values = spill.Value2;
                }
                catch (Exception ex) when (!(ex is OutOfMemoryException))
                {
                    all = false; // не прочитался — не трогаем
                    continue;
                }
                var grid = values as object[,];
                object first = grid != null ? grid.GetValue(grid.GetLowerBound(0), grid.GetLowerBound(1)) : values;
                try
                {
                    anchor.Value2 = CellText.EscapeValues(first);
                    sheet.Range(address).Value2 = CellText.EscapeValues(values);
                }
                catch (Exception ex) when (!(ex is OutOfMemoryException))
                {
                    try { anchor.Formula2 = formula; } catch { } // вернуть разлив, а не потерять его
                    all = false;
                }
            }
            return all;
        }

        /// <summary>Есть ли в диапазоне разлив (HasSpill агрегатный: true, false или null). Без динамических массивов — нет.</summary>
        private static bool HasAnySpill(dynamic range)
        {
            try { return AnyFormula((object)range.HasSpill); }
            catch (Exception ex) when (!(ex is OutOfMemoryException)) { return false; }
        }

        private static bool IsSpill(dynamic cell)
        {
            object spills = cell.HasSpill;
            return spills is bool && (bool)spills;
        }

        private static void ReplaceCellByCell(dynamic area, bool skipSpills = false)
        {
            int rows = (int)area.Rows.Count;
            int cols = (int)area.Columns.Count;
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    try
                    {
                        // Значение объединённой ячейки живёт в её верхней левой ячейке.
                        if (skipSpills && IsSpill(area.Cells[r, c]))
                            continue; // уцелевший разлив: значение якоря стёрло бы его
                        dynamic topLeft = area.Cells[r, c].MergeArea.Cells[1, 1];
                        object v = topLeft.Value2;
                        if (v != null)
                            topLeft.Value2 = CellText.EscapeValues(v);
                    }
                    catch { } // не top-left объединённой области — значение уже заменено
                }
            }
        }

        private static void AppendNote(FileResult fr, string note)
        {
            fr.Note = string.IsNullOrEmpty(fr.Note) ? note : fr.Note + "; " + note;
        }

        /// <summary>
        /// Отвечает ли Excel: обращение к Workbooks — именно оно виснет при
        /// заклинивании после битого файла (и падает, если Excel аварийно вышел).
        /// </summary>
        private static bool IsExcelResponsive(object excelObj)
        {
            dynamic excel = excelObj;
            // Несколько попыток: транзиентный COM-блип (занят) не должен ложно
            // трактоваться как заклинивание и вызывать лишний перезапуск.
            for (int i = 0; i < 3; i++)
            {
                try
                {
                    int probe = (int)excel.Workbooks.Count;
                    return true;
                }
                catch (Exception)
                {
                    System.Threading.Thread.Sleep(150);
                }
            }
            return false;
        }

        /// <summary>Видимые листы книги: все или только первый (в порядке книги).</summary>
        private static List<object> VisibleSheets(dynamic workbook, bool all)
        {
            var list = new List<object>();
            dynamic sheets = workbook.Sheets; // включая листы-диаграммы
            int count = (int)sheets.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic s = sheets[i];
                if ((int)s.Visible == -1) // xlSheetVisible
                {
                    list.Add(s);
                    if (!all)
                        break;
                }
            }
            return list;
        }

        private static string ClassifyError(Exception ex)
        {
            Exception e = Unwrap(ex);
            if (IsPasswordError(e.Message))
                return PasswordNote;
            return "не удалось открыть или скопировать (" + ShortMessage(e) + ")";
        }

        private static Exception Unwrap(Exception ex)
        {
            var t = ex as TargetInvocationException;
            if (t != null && t.InnerException != null)
                return t.InnerException;
            return ex;
        }

        private static string ShortMessage(Exception ex)
        {
            return ShortText(Unwrap(ex).Message);
        }

        private static string ShortText(string message)
        {
            string m = (message ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            while (m.Contains("  "))
                m = m.Replace("  ", " ");
            if (m.Length > 200)
                m = m.Substring(0, 200) + "…";
            return m;
        }

        private void RaiseTrace(string msg)
        {
            var h = Trace;
            if (h != null) h(msg);
        }

        private void RaiseProgress(int current, int total, string fileName)
        {
            var h = Progress;
            if (h != null) h(current, total, fileName);
        }

        private void RaiseFileDone(FileResult fr)
        {
            var h = FileDone;
            if (h != null) h(fr);
        }
    }
}
