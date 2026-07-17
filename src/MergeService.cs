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
    }

    public class MergeResult
    {
        public readonly List<FileResult> Files = new List<FileResult>();
        public int OkCount;
        public int SkipCount;
        public bool Cancelled;
        public string OutputPath;
        public string TocError; // не null, если лист «Содержание» создать не удалось

        /// <summary>Число исходных файлов (в режиме «все листы» листов больше).</summary>
        public int FileCount
        {
            get
            {
                var files = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
                foreach (FileResult f in Files)
                    files.Add(f.FullPath ?? f.FileName);
                return files.Count;
            }
        }
    }

    /// <summary>Ошибка объединения с понятным пользователю сообщением.</summary>
    public class MergeException : Exception
    {
        public MergeException(string message) : base(message) { }
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

        /// <summary>Мягкая отмена: останов после текущего файла, без сохранения результата.</summary>
        public void Cancel()
        {
            _cancel = true;
        }

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
                    return "Папка сохранения не существует: " + dir;
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
                return "Итоговый файл занят другой программой — закройте его (обычно он открыт в Excel) и повторите.";
            }
            catch (UnauthorizedAccessException)
            {
                return "Нет прав на запись в папку сохранения.";
            }
            catch (Exception ex)
            {
                return "Итоговый файл недоступен для записи (" + ShortMessage(ex) + ").";
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
                throw new MergeException("Не выбрано ни одного файла Excel для объединения.");

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
                throw new MergeException("Итоговый файл не найден — сначала выполните обычное объединение.");

            var retryPaths = new List<string>();
            foreach (FileResult fr in previous.Files)
            {
                if (!fr.Ok && fr.FullPath != null)
                    retryPaths.Add(fr.FullPath);
            }
            if (retryPaths.Count == 0)
                throw new MergeException("Пропущенных файлов нет — повторять нечего.");

            ValidateOutput(outputPath);
            return Run(outputPath, retryPaths, options, previous);
        }

        /// <summary>
        /// Итог после повторной попытки: пропущенные записи предыдущего прогона
        /// заменяются свежими результатами (по полному пути), успешные и порядок
        /// сохраняются. Чистая функция — покрыта юнит-тестами.
        /// </summary>
        public static MergeResult CombineRetryResults(MergeResult previous, IList<FileResult> attempts)
        {
            // Свежие результаты повтора, сгруппированные по исходному файлу
            // (в режиме «все листы» один файл даёт несколько листов).
            var freshByPath = new Dictionary<string, List<FileResult>>(StringComparer.OrdinalIgnoreCase);
            foreach (FileResult a in attempts)
            {
                string key = a.FullPath ?? a.FileName;
                List<FileResult> bucket;
                if (!freshByPath.TryGetValue(key, out bucket))
                {
                    bucket = new List<FileResult>();
                    freshByPath[key] = bucket;
                }
                bucket.Add(a);
            }

            var combined = new MergeResult();
            combined.OutputPath = previous.OutputPath;
            var expanded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FileResult old in previous.Files)
            {
                if (old.Ok)
                {
                    combined.Files.Add(old); // успешные листы сохраняются
                    continue;
                }
                string key = old.FullPath ?? old.FileName;
                List<FileResult> fresh;
                if (key != null && freshByPath.TryGetValue(key, out fresh))
                {
                    if (expanded.Add(key)) // разворачиваем файл один раз
                        combined.Files.AddRange(fresh);
                }
                else
                {
                    combined.Files.Add(old); // файл не повторялся
                }
            }
            foreach (FileResult f in combined.Files)
                if (f.Ok) combined.OkCount++; else combined.SkipCount++;
            return combined;
        }

        private static void ValidateOutput(string outputPath)
        {
            if (OutputFormats.FileFormatFor(outputPath) == 0)
                throw new MergeException("Неподдерживаемое расширение итогового файла. Допустимы: " +
                    string.Join(", ", OutputFormats.Extensions) + ".");
            string lockError = CheckOutputWritable(outputPath);
            if (lockError != null)
                throw new MergeException(lockError);
        }

        /// <summary>
        /// Движок слияния. previous == null — новый свод; иначе дослияние
        /// в существующий файл с объединением результатов.
        /// </summary>
        private MergeResult Run(string outputPath, List<string> files, MergeOptions options, MergeResult previous)
        {
            bool intoExisting = previous != null;
            var attempt = new MergeResult();
            attempt.OutputPath = outputPath;

            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new MergeException("Microsoft Excel не установлен: COM-компонент Excel.Application не найден.");

            dynamic excel = null;
            dynamic target = null;
            ComMessageFilter.Register(); // авто-повтор вызовов, отклонённых занятым Excel
            try
            {
                excel = Activator.CreateInstance(excelType);
                excel.Visible = false;
                excel.DisplayAlerts = false;
                excel.ScreenUpdating = false;
                excel.EnableEvents = false;
                try { excel.AskToUpdateLinks = false; } catch { }
                try { excel.AutomationSecurity = 3; } catch { } // msoAutomationSecurityForceDisable: макросы источников не выполняются

                WaitExcelReady((object)excel); // под нагрузкой Excel готов к Workbooks не сразу

                if (intoExisting)
                {
                    target = excel.Workbooks.Open(
                        Filename: outputPath,
                        UpdateLinks: 0,
                        ReadOnly: false,
                        IgnoreReadOnlyRecommended: true,
                        Notify: false,
                        AddToMru: false);
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

                    List<FileResult> frs;
                    try
                    {
                        frs = CopySheets((object)excel, (object)target, path, namer, options);
                    }
                    catch (Exception ex)
                    {
                        // Защита от ошибок вне per-file обработчика (например, сбой COM-привязки).
                        var fr = new FileResult();
                        fr.FileName = Path.GetFileName(path);
                        fr.FullPath = path;
                        fr.Note = "не удалось обработать (" + ShortMessage(ex) + ")";
                        EnsureExcelAlive((object)excel, fr.FileName);
                        frs = new List<FileResult> { fr };
                    }
                    foreach (FileResult fr in frs)
                    {
                        attempt.Files.Add(fr);
                        if (fr.Ok) attempt.OkCount++; else attempt.SkipCount++;
                        RaiseFileDone(fr);
                    }
                }

                if (attempt.Cancelled)
                    return attempt; // без сохранения: существующий свод не тронут

                if (!intoExisting && attempt.OkCount == 0)
                    throw new MergeException("Не удалось перенести ни один лист — итоговый файл не создан. Причины указаны в списке файлов.");

                MergeResult final = intoExisting
                    ? CombineRetryResults(previous, attempt.Files)
                    : attempt;

                if (intoExisting && attempt.OkCount == 0)
                    return final; // ни одна повторная попытка не удалась — файл не трогаем

                if (!intoExisting)
                    try { target.Sheets[PlaceholderName].Delete(); } catch { }

                // Итоговый файл сохраняется со стандартным авторасчётом (один пересчёт здесь).
                try { excel.Calculation = XlCalculationAutomatic; } catch { }

                if (options.AddToc)
                {
                    try
                    {
                        TocBuilder.Build((object)target, TocSheetName, final.Files);
                    }
                    catch (Exception ex)
                    {
                        final.TocError = "лист «Содержание» создать не удалось (" + ShortMessage(ex) + ")";
                    }
                }

                try
                {
                    if (intoExisting)
                        target.Save(); // формат существующего файла сохраняется
                    else
                        target.SaveAs(outputPath, OutputFormats.FileFormatFor(outputPath));
                }
                catch (Exception ex)
                {
                    throw new MergeException(
                        "Не удалось сохранить итоговый файл. Возможно, он открыт в Excel или нет прав на запись в папку.\n(" +
                        ShortMessage(ex) + ")");
                }

                return final;
            }
            finally
            {
                // Статические ссылки: после Close/Quit никаких динамических операций
                // над объектами (см. комментарий у ComSafe.Release).
                object targetObj = target;
                if (targetObj != null)
                {
                    try { target.Close(false); } catch { }
                    ComSafe.Release(targetObj);
                }
                object excelObj = excel;
                if (excelObj != null)
                {
                    try { excel.Quit(); } catch { }
                    ComSafe.Release(excelObj);
                }
                ComSafe.Collect();
                ComMessageFilter.Revoke();
            }
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

        private List<FileResult> CopySheets(object excelObj, object targetObj, string path, SheetNamer namer, MergeOptions options)
        {
            dynamic excel = excelObj;
            dynamic target = targetObj;
            string fileName = Path.GetFileName(path);
            string fileBase = Path.GetFileNameWithoutExtension(path);
            var results = new List<FileResult>();

            dynamic source = null;
            try
            {
                RaiseTrace("open: " + fileName);
                try
                {
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

                bool fileNotesPending = true;
                foreach (object wsObj in sheets)
                {
                    dynamic ws = wsObj;
                    var fr = new FileResult();
                    fr.FileName = fileName;
                    fr.FullPath = path;
                    try
                    {
                        string baseName = SheetBaseName(fileBase, (string)ws.Name, options.AllSheets);
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

                        // Лист-диаграмма не имеет ячеек: гиперссылка на A1 невозможна.
                        try { dynamic probe = copied.Range("A1"); fr.Linkable = true; }
                        catch { fr.Linkable = false; }

                        if (options.ValuesOnly && fr.Linkable)
                        {
                            try { ReplaceFormulasWithValues((object)copied); }
                            catch (Exception) { AppendNote(fr, "формулы заменены не полностью"); }
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
                    string m = (Unwrap(ex).Message ?? "").ToLowerInvariant();
                    // Проблема самого файла — повторять бессмысленно.
                    if (m.Contains("парол") || m.Contains("password") ||
                        m.Contains("недопустим") || m.Contains("format") || m.Contains("поврежд"))
                        throw;
                    System.Threading.Thread.Sleep(500); // транзиентный сбой — подождать и повторить
                }
            }
            throw last;
        }

        private static FileResult Skipped(string fileName, string path, string note)
        {
            var fr = new FileResult();
            fr.FileName = fileName;
            fr.FullPath = path;
            fr.Note = note;
            return fr;
        }

        /// <summary>
        /// Заменяет формулы листа их вычисленными (кэшированными) значениями.
        /// Обрабатываются только области с формулами (SpecialCells) — массовое
        /// присваивание по областям, а не по ячейкам. Область с объединёнными
        /// ячейками не принимает массив — для неё поячеечный фолбэк через
        /// верхнюю левую ячейку MergeArea.
        /// </summary>
        private static void ReplaceFormulasWithValues(object sheetObj)
        {
            dynamic sheet = sheetObj;
            dynamic formulaCells = null;
            try
            {
                formulaCells = sheet.UsedRange.SpecialCells(-4123); // xlCellTypeFormulas
            }
            catch
            {
                return; // формул на листе нет
            }
            dynamic areas = formulaCells.Areas;
            int count = (int)areas.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic area = areas[i];
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
        }

        private static void ReplaceCellByCell(dynamic area)
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

        /// <summary>Если Excel аварийно завершился, продолжать бессмысленно — падаем с понятной ошибкой.</summary>
        private static void EnsureExcelAlive(object excelObj, string fileName)
        {
            dynamic excel = excelObj;
            try
            {
                string probe = (string)excel.Name;
            }
            catch (Exception)
            {
                throw new MergeException("Excel аварийно завершился при обработке файла «" + fileName +
                    "». Итоговый файл не создан. Запустите объединение повторно, при повторении — исключите этот файл.");
            }
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
            string m = (e.Message ?? "").ToLowerInvariant();
            if (m.Contains("парол") || m.Contains("password"))
                return "файл защищён паролем";
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
            string m = (Unwrap(ex).Message ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
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
