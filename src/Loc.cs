using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>Язык интерфейса.</summary>
    public enum Lang { Ru, En }

    /// <summary>
    /// Локализация интерфейса и сообщений. Ресурсных .resx в проекте нет (UI строится кодом),
    /// поэтому — централизованный строковый каталог: ключ → (русский, английский). Текущий язык
    /// хранится в настройках (settings.txt), переключается через <see cref="Set"/>; окна
    /// пересобираются подписчиками на <see cref="Changed"/> (см. ShellContext). Промах по ключу
    /// возвращает сам ключ — чтобы пропущенная строка сразу бросалась в глаза.
    ///
    /// СОДЕРЖИМОЕ генерируемых документов (записка ГОСТ 7.0.97, отчёты, оглавление свода) НЕ
    /// локализуется — по решению остаётся русским независимо от языка UI.
    /// </summary>
    internal static class Loc
    {
        private const int Ru = 0, En = 1;
        private static Lang _current = Lang.Ru;

        /// <summary>Текущий язык интерфейса.</summary>
        public static Lang Current { get { return _current; } }

        /// <summary>Событие смены языка — подписчики (окна) пересобирают себя.</summary>
        public static event Action Changed;

        /// <summary>Задать язык при старте БЕЗ события (до создания окон).</summary>
        public static void Init(Lang lang) { _current = lang; }

        /// <summary>Сменить язык: сохранить в настройки и уведомить подписчиков. No-op, если не изменился.</summary>
        public static void Set(Lang lang)
        {
            if (lang == _current)
                return;
            _current = lang;
            // Сохранить, сохранив прочие настройки: Load читает файл, Save пишет их обратно и
            // проставляет язык из Loc.Current (см. UserSettings.Save).
            try { UserSettings.Load().Save(); }
            catch { } // не удалось сохранить — язык всё равно применится в этой сессии
            Action h = Changed;
            if (h != null)
                h();
        }

        /// <summary>Код языка для настроек: «ru»/«en».</summary>
        public static string Code(Lang lang) { return lang == Lang.En ? "en" : "ru"; }

        /// <summary>Разобрать код языка из настроек; неизвестный → русский.</summary>
        public static Lang Parse(string code)
        {
            return string.Equals(code, "en", StringComparison.OrdinalIgnoreCase) ? Lang.En : Lang.Ru;
        }

        /// <summary>Строка по ключу на текущем языке; нет ключа → сам ключ (видимый промах).</summary>
        public static string T(string key)
        {
            string[] pair;
            if (key != null && Catalog.TryGetValue(key, out pair))
                return pair[_current == Lang.En ? En : Ru];
            return key;
        }

        // ---- Каталог: ключ → [ru, en]. Группы по префиксам. ----
        private static readonly Dictionary<string, string[]> Catalog = Build();

        /// <summary>Все ключи каталога (для тестов полноты).</summary>
        internal static IEnumerable<string> Keys { get { return Catalog.Keys; } }

        /// <summary>Пара [ru, en] по ключу или null (для тестов).</summary>
        internal static string[] Pair(string key)
        {
            string[] p;
            return Catalog.TryGetValue(key, out p) ? p : null;
        }

        private static Dictionary<string, string[]> Build()
        {
            var d = new Dictionary<string, string[]>(StringComparer.Ordinal);
            void A(string key, string ru, string en) { d[key] = new[] { ru, en }; }

            // menu.* — меню окон (бывшая «Справка»)
            A("menu.root", "☰ Меню", "☰ Menu");
            A("menu.howTo", "Как пользоваться", "How to use");
            A("menu.stats", "Статистика", "Statistics");
            A("menu.language", "Язык / Language", "Язык / Language");
            A("menu.lang.ru", "RU", "RU");
            A("menu.lang.en", "EN", "EN");

            // lang.* — выбор языка на главной (глобус)
            A("lang.tooltip", "Язык интерфейса / Interface language", "Язык интерфейса / Interface language");

            // shell.* — оболочка окон
            A("shell.toolOpen.title", "Инструмент уже открыт", "Tool already open");
            A("shell.toolOpen.body", "«{0}» уже запущен — открыто его окно.", "“{0}” is already running — its window is open.");

            // hub.* — стартовый экран (выбор инструмента)
            A("hub.subtitle", "Выберите инструмент", "Choose a tool");
            A("hub.excel.name", "Свод Excel", "Excel Digest");
            A("hub.excel.desc",
                "Объединить листы из нескольких файлов Excel в один свод: оглавление, замена формул значениями, сопроводительная записка Word.",
                "Merge sheets from several Excel files into one digest: table of contents, replace formulas with values, a Word cover note.");
            A("hub.pdf.name", "Объединение PDF", "Merge PDF");
            A("hub.pdf.desc",
                "Собрать один PDF из нескольких файлов: выбрать нужные страницы и задать их порядок. Страницы копируются без искажений.",
                "Build one PDF from several files: pick the pages you need and set their order. Pages are copied without distortion.");
            A("hub.split.name", "Разделение PDF", "Split PDF");
            A("hub.split.desc",
                "Извлечь выбранные страницы в один PDF или разбить документ на несколько: по диапазонам, каждые N страниц или по закладкам.",
                "Extract selected pages into one PDF, or split the document into several: by ranges, every N pages, or by bookmarks.");
            A("hub.ocr.name", "PDF → Word", "PDF → Word");
            A("hub.ocr.desc",
                "Извлечь текст цифрового PDF (сохранённого из Word и т.п.) в редактируемый Word (.docx). Поддержка отсканированных документов в настоящее время недоступна.",
                "Extract the text of a born‑digital PDF (saved from Word, etc.) into an editable Word (.docx). Scanned documents are not supported yet.");
            A("hub.update", "⟳ Проверить обновления", "⟳ Check for updates");
            A("hub.about", "О программе", "About");

            // common.* — общие элементы нескольких окон
            A("common.browse", "Обзор…", "Browse…");
            A("common.copy", "Копировать", "Copy");
            A("common.home", "⌂ Главная", "⌂ Home");
            A("common.homeTip", "Открыть экран выбора инструмента", "Open the tool chooser");
            A("common.zoom", "Масштаб:", "Zoom:");
            A("common.tip.zoom", "Масштаб миниатюр (также Ctrl+колесо мыши)", "Thumbnail zoom (also Ctrl+mouse wheel)");
            A("common.busy", "Дождитесь завершения…", "Wait for it to finish…");
            A("common.err.openFailed", "Не удалось открыть", "Could not open");
            // Общие для PDF-инструментов: перестановка/удаление страниц, диалоги выбора файлов
            A("common.earlier", "◀ Раньше", "◀ Earlier");
            A("common.later", "Позже ▶", "Later ▶");
            A("common.remove", "Удалить", "Remove");
            A("common.tip.earlier", "Переместить страницу раньше (Alt+←)", "Move the page earlier (Alt+←)");
            A("common.tip.later", "Переместить страницу позже (Alt+→)", "Move the page later (Alt+→)");
            A("common.tip.remove", "Убрать выбранные страницы из вывода (Delete)", "Remove the selected pages from the output (Delete)");
            A("common.pdfFilter", "Документы PDF (*.pdf)|*.pdf", "PDF documents (*.pdf)|*.pdf");
            A("common.pdfSaveFilter", "Документ PDF (*.pdf)|*.pdf", "PDF document (*.pdf)|*.pdf");
            A("common.pickPdf", "Выберите PDF-файлы", "Choose PDF files");
            A("common.fileNotAdded", "Файл не добавлен", "File not added");
            A("common.addPdf", "Добавить PDF…", "Add PDF…");
            A("common.tip.addPdf", "Файлы также можно перетащить в окно", "You can also drag files onto the window");
            A("common.tip.removePages", "Удалить выбранные страницы (Delete)", "Remove the selected pages (Delete)");
            A("common.status.pageCountList", "Страниц в списке: {0}.", "Pages in the list: {0}.");
            A("common.status.saving", "Сохранение…", "Saving…");
            A("common.busySaving", "Дождитесь завершения сохранения…", "Wait for saving to finish…");
            A("common.status.notDone", "Не выполнено.", "Failed.");
            A("common.close", "Закрыть", "Close");
            A("common.yes", "Да", "Yes");
            A("common.no", "Нет", "No");
            A("common.compression", "Сжатие:", "Compression:");
            A("grid.pageTip", "{0} — стр. {1}", "{0} — p. {1}");
            A("grid.pageLabel", "{0}\nстр. {1}", "{0}\np. {1}");
            A("common.tip.compression",
                "«Хорошо»/«Нормально» уменьшают размер, снижая разрешение изображений (как в Acrobat).\n" +
                "Текст и вектор сохраняются. У подписанных PDF подпись станет недействительной.",
                "“Good”/“Normal” reduce the size by downsampling images (as in Acrobat).\n" +
                "Text and vectors are preserved. A signed PDF’s signature becomes invalid.");
            // err.pdf.* — сообщения PDF-сервисов (объединение/разделение/загрузка), показываются в диалогах
            A("err.pdf.noPages", "Нет страниц для объединения.", "No pages to merge.");
            A("err.pdf.fileBusy", "Файл PDF недоступен для записи — возможно, открыт в другой программе.",
                "The PDF file is not writable — it may be open in another program.");
            A("err.pdf.noPagesIn", "В файле «{0}» нет страниц.", "The file “{0}” has no pages.");
            A("err.pdf.cantOpen", "Не удалось открыть «{0}»: файл повреждён, защищён паролем или использует неподдерживаемые возможности PDF. ({1})",
                "Could not open “{0}”: the file is corrupt, password‑protected, or uses unsupported PDF features. ({1})");
            A("err.pdf.cantOpenShort", "Не удалось открыть «{0}»: {1}", "Could not open “{0}”: {1}");
            A("err.pdf.pageGone", "В «{0}» нет страницы {1} — файл изменился после добавления в список.",
                "“{0}” has no page {1} — the file changed after it was added to the list.");
            A("err.pdf.saveFailed", "Не удалось сохранить PDF: {0}", "Could not save the PDF: {0}");
            A("err.split.noPages", "Не выбрано ни одной страницы.", "No pages selected.");
            A("err.split.noRanges", "Не задано ни одного диапазона.", "No ranges specified.");
            A("err.split.badN", "Число страниц в части должно быть не меньше 1.", "Pages per part must be at least 1.");
            A("err.split.rangeOutside", "Диапазон {0} вне файла (страниц: {1}).", "Range {0} is outside the file ({1} pages).");
            A("err.split.noBookmarks", "В файле нет закладок верхнего уровня — этот режим не применим.",
                "The file has no top‑level bookmarks — this mode does not apply.");
            A("err.split.saveFailed", "Не удалось сохранить «{0}»: {1}", "Could not save “{0}”: {1}");
            A("err.ranges.empty", "Укажите диапазоны страниц, например: 1-3, 5, 8-", "Enter page ranges, for example: 1-3, 5, 8-");
            A("err.ranges.outside", "Диапазон «{0}» вне 1–{1}.", "Range “{0}” is outside 1–{1}.");
            A("err.ranges.badPage", "Не понял номер страницы в «{0}».", "Could not read a page number in “{0}”.");
            A("err.word.notInstalled", "Microsoft Word не установлен: COM-компонент Word.Application не найден.",
                "Microsoft Word is not installed: the Word.Application COM component was not found.");
            A("err.word.fileBusy", "«{0}» недоступен для записи — возможно, открыт в Word или другой программе.",
                "“{0}” is not writable — it may be open in Word or another program.");
            A("word.label.docx", "Файл Word", "Word file");
            A("word.label.note", "Записка Word", "Word note");
            // cli.* — консольный режим
            A("cli.usage", "Использование: iwoHelperDesktop.exe --cli <папка> <итоговый> [--toc] [--values] [--allsheets]",
                "Usage: iwoHelperDesktop.exe --cli <folder> <output> [--toc] [--values] [--allsheets]");
            A("cli.unknownFlag", "неизвестный параметр «{0}»", "unknown parameter “{0}”");
            A("split.partInfix", "_часть_", "_part_");
            A("split.unnamed", "без_имени", "unnamed");
            A("err.ocr.noPages", "Не выбрано ни одной страницы для конвертации.", "No pages selected to convert.");
            A("err.ocr.scanned",
                "В выбранных PDF нет извлекаемого текста — похоже, это отсканированные документы (изображения). " +
                "Поддержка отсканированных документов в настоящее время недоступна.",
                "The selected PDFs have no extractable text — they look like scanned documents (images). " +
                "Scanned documents are not supported yet.");
            A("err.ocr.extractFailed",
                "Не удалось извлечь текст из «{0}»: файл повреждён, зашифрован или без прав на извлечение. ({1})",
                "Could not extract text from “{0}”: the file is corrupt, encrypted, or extraction is not allowed. ({1})");

            // crash.* — последний рубеж обработки ошибок (CrashReport)
            A("crash.title", "Непредвиденная ошибка", "Unexpected error");
            A("crash.body",
                "{0}\n\nПриложение продолжит работу. Техническая информация сохранена в файл:\n{1}",
                "{0}\n\nThe app will keep running. Technical details were saved to:\n{1}");

            // update.* — проверка обновлений
            A("update.err.title", "Не удалось проверить обновления", "Could not check for updates");
            A("update.err.network", "Проверьте подключение к интернету. ({0})", "Check your internet connection. ({0})");
            A("update.err.badResponse", "Непонятный ответ сервера.", "Unclear server response.");
            A("update.err.parseVersion", "не удалось прочитать версию из ответа GitHub",
                "could not read the version from the GitHub response");
            A("update.available.title", "Доступна новая версия {0}", "A new version {0} is available");
            A("update.available.body", "У вас {0}. Открыть страницу загрузки в браузере?",
                "You have {0}. Open the download page in your browser?");
            A("update.none.title", "Обновлений нет", "No updates");
            A("update.none.body", "У вас последняя версия ({0}).", "You have the latest version ({0}).");

            // gs.* — предупреждение об отсутствии Ghostscript
            A("gs.title", "Сжатие недоступно", "Compression unavailable");
            A("gs.heading", "Нужен Ghostscript", "Ghostscript required");
            A("gs.body",
                "Сжатие PDF использует Ghostscript — он не найден в системе. " +
                "Установите его (бесплатно), затем перезапустите приложение — либо " +
                "используйте установщик приложения, в него Ghostscript уже входит.",
                "PDF compression uses Ghostscript, which was not found on this system. " +
                "Install it (free) and restart the app — or use the app installer, " +
                "which already bundles Ghostscript.");
            A("gs.download", "Скачать Ghostscript", "Download Ghostscript");
            A("compress.level.none", "Отлично — без сжатия", "Excellent — no compression");
            A("compress.level.good", "Хорошо — меньше размер", "Good — smaller size");
            A("compress.level.small", "Нормально — минимальный размер", "Normal — minimal size");

            // err.merge.* — сообщения сервиса «Свод Excel», показываемые в диалогах/статусе
            A("err.merge.folderMissing", "Папка сохранения не существует: {0}", "The save folder does not exist: {0}");
            A("err.merge.outputBusy", "Итоговый файл занят другой программой — закройте его (обычно он открыт в Excel) и повторите.",
                "The output file is in use by another program — close it (usually open in Excel) and try again.");
            A("err.merge.noWritePerm", "Нет прав на запись в папку сохранения.", "No permission to write to the save folder.");
            A("err.merge.outputNotWritable", "Итоговый файл недоступен для записи ({0}).", "The output file is not writable ({0}).");
            A("err.merge.noFiles", "Не выбрано ни одного файла Excel для объединения.", "No Excel files selected to merge.");
            A("err.merge.noOutput", "Итоговый файл не найден — сначала выполните обычное объединение.",
                "The output file was not found — run a normal merge first.");
            A("err.merge.nothingToRetry", "Пропущенных файлов нет — повторять нечего.", "No skipped files — nothing to retry.");
            A("err.merge.badExtension", "Неподдерживаемое расширение итогового файла. Допустимы: {0}",
                "Unsupported output file extension. Allowed: {0}");
            A("err.merge.lowSpace", "На диске {0} почти нет свободного места ({1} МБ). Excel не сможет открыть файлы — освободите место и повторите.",
                "Drive {0} is almost out of free space ({1} MB). Excel won’t be able to open files — free up space and try again.");
            A("err.merge.excelUnstable", "Excel не удалось стабилизировать после файла «{0}». Исключите этот файл из списка (снимите галочку) и повторите.",
                "Excel could not be stabilized after the file “{0}”. Exclude this file from the list (untick it) and try again.");
            A("err.merge.excelMissing", "Microsoft Excel не установлен: COM-компонент Excel.Application не найден.",
                "Microsoft Excel is not installed: the Excel.Application COM component was not found.");
            A("err.merge.tocFailed", "лист «Содержание» создать не удалось ({0})",
                "the “Contents” sheet could not be created ({0})");
            A("err.merge.noSheets", "Не удалось перенести ни один лист — итоговый файл не создан. Причины указаны в списке файлов.",
                "No sheet could be transferred — the output file was not created. See the file list for reasons.");
            A("err.merge.saveFailed", "Не удалось сохранить итоговый файл. Возможно, он открыт в Excel или нет прав на запись в папку.\n({0})",
                "Could not save the output file. It may be open in Excel, or there is no write permission for the folder.\n({0})");

            // about.* — окно «О программе» (AboutForm)
            A("about.version", "Версия {0}", "Version {0}");
            A("about.desc",
                "Офисные инструменты: свод листов Excel, объединение, разделение и сжатие PDF, конвертация PDF → Word.",
                "Office tools: Excel sheet digest, merge, split and compress PDFs, convert PDF → Word.");
            A("about.author", "Автор: Dodonov Andrey (DedovMosol)", "Author: Dodonov Andrey (DedovMosol)");
            A("about.license", "© 2026 · Лицензия MIT", "© 2026 · MIT License");
            A("about.privacy", "Политика конфиденциальности", "Privacy Policy");
            A("about.privacyNote", "(данные не покидают ваш ПК)", "(your data never leaves your PC)");
            A("about.donate", "Поддержать проект (донаты):", "Support the project (donations):");
            A("about.account", "Счёт:", "Account:");
            A("about.bank", "Банк:", "Bank:");

            // stats.* — окно «Статистика» (StatsForm)
            A("stats.since", "Считается с {0}.", "Counting since {0}.");
            A("stats.row.excel", "Своды Excel", "Excel digests");
            A("stats.row.merge", "Объединения PDF", "PDF merges");
            A("stats.row.extract", "Извлечения страниц (PDF)", "Page extractions (PDF)");
            A("stats.row.ranges", "Разбиение по диапазонам", "Split by ranges");
            A("stats.row.everyN", "Разбиение: каждые N страниц", "Split: every N pages");
            A("stats.row.bookmarks", "Разбиение по закладкам", "Split by bookmarks");
            A("stats.row.pdftoword", "Конвертации PDF → Word", "PDF → Word conversions");
            A("stats.row.compress", "Сжатия PDF (файлов)", "PDF compressions (files)");
            A("stats.total", "Всего операций: {0}", "Total operations: {0}");
            A("stats.autoClear", "Автоочистка:", "Auto‑clear:");
            A("stats.auto.off", "Выключена", "Off");
            A("stats.auto.daily", "Раз в день", "Once a day");
            A("stats.auto.7days", "Раз в 7 дней", "Every 7 days");
            A("stats.auto.30days", "Раз в 30 дней", "Every 30 days");
            A("stats.tip.auto", "Счётчики будут автоматически обнуляться с выбранной периодичностью",
                "Counters will be reset automatically at the chosen interval");
            A("stats.btn.clear", "Очистить", "Clear");
            A("stats.confirm.clear.title", "Очистить счётчики?", "Clear the counters?");
            A("stats.confirm.clear.body", "Все накопленные числа обнулятся. Действие необратимо.",
                "All accumulated numbers will be reset. This cannot be undone.");

            // split.* — инструмент «Разделение PDF» (PdfSplitForm)
            A("split.header.subtitle", "Извлечение страниц из документа формата *.pdf со сжатием.",
                "Extract pages from a *.pdf document, with compression.");
            A("split.btn.open", "Открыть PDF…", "Open PDF…");
            A("split.tip.open", "Файл также можно перетащить в окно", "You can also drag the file onto the window");
            A("split.lbl.mode", "Режим:", "Mode:");
            A("split.mode.extract", "Извлечь выбранные", "Extract selected");
            A("split.mode.ranges", "По диапазонам", "By ranges");
            A("split.mode.everyN", "Каждые N страниц", "Every N pages");
            A("split.mode.bookmarks", "По закладкам", "By bookmarks");
            A("split.lbl.ranges", "Диапазоны (напр. 1-3, 5, 8-):", "Ranges (e.g. 1-3, 5, 8-):");
            A("split.lbl.n", "Страниц в части:", "Pages per part:");
            A("split.chk.combine", "Объединить в один файл", "Combine into one file");
            A("split.tip.combine", "Все указанные страницы — в один PDF, а не по файлу на диапазон",
                "All listed pages into one PDF, not one file per range");
            A("split.status.openPdf", "Откройте PDF — кнопкой «Открыть PDF…» или перетащив его в окно.",
                "Open a PDF — with “Open PDF…” or by dragging it onto the window.");
            A("split.pickPdf", "Выберите PDF для разделения", "Choose a PDF to split");
            A("split.err.fileNotOpened", "Файл не открыт", "File not opened");
            A("split.status.opened", "Открыт «{0}»: страниц {1}.", "Opened “{0}”: {1} pages.");
            A("split.hint.extract", "Выделите нужные страницы в сетке (Ctrl+A — все).",
                "Select the pages in the grid (Ctrl+A — all).");
            A("split.hint.bookmarks", "По одному файлу на закладку верхнего уровня.",
                "One file per top‑level bookmark.");
            A("split.btn.extract", "Извлечь…", "Extract…");
            A("split.btn.split", "Разделить…", "Split…");
            A("split.err.noPages.title", "Не выбраны страницы", "No pages selected");
            A("split.err.noPages.body", "Выделите страницы в сетке (Ctrl+A — все).", "Select pages in the grid (Ctrl+A — all).");
            A("split.err.badRanges", "Диапазоны заданы неверно", "Ranges are invalid");
            A("split.suffix.selected", "_выбранные.pdf", "_selected.pdf");
            A("split.suffix.combined", "_объединённые.pdf", "_combined.pdf");
            A("split.pickBase", "Базовое имя и папка для частей (к имени добавятся номера)",
                "Base name and folder for the parts (numbers are appended)");
            A("split.status.splitting", "Разделение…", "Splitting…");
            A("split.status.extracting", "Извлечение…", "Extracting…");
            A("split.err.splitFailed", "Разделение не выполнено", "Split failed");
            A("split.err.extractFailed", "Извлечение не выполнено", "Extraction failed");
            A("split.status.created", "✓ Создано файлов: {0}.", "✓ Files created: {0}.");
            A("split.status.done", "✓ Готово.", "✓ Done.");
            A("split.suffix.compressed", " · сжато: {0}", " · compressed: {0}");
            A("split.status.largeHint", " · файл крупный — включите «Сжатие», чтобы уменьшить размер.",
                " · the file is large — turn on “Compression” to reduce its size.");
            A("split.help.body",
                "1. Откройте PDF — кнопкой «Открыть PDF…» или перетащив его в окно; появится сетка страниц.\n" +
                "2. Выберите режим:\n" +
                "   • «Извлечь выбранные» — выделите страницы в сетке (Ctrl+A — все) → сохранит их в один PDF;\n" +
                "   • «По диапазонам» — «1-3, 5, 8-»: каждый диапазон → отдельный файл;\n" +
                "   • «Каждые N страниц» — равные части (1 — каждая страница отдельно);\n" +
                "   • «По закладкам» — по одному файлу на закладку верхнего уровня, имена из заголовков.\n" +
                "3. При необходимости выберите «Сжатие» (по умолчанию «Отлично» — без сжатия): " +
                "«Хорошо»/«Нормально» уменьшают размер за счёт понижения разрешения изображений " +
                "(как в Acrobat), текст сохраняется. Требуется Ghostscript.\n" +
                "4. Нажмите «Извлечь…»/«Разделить…» и укажите имя и папку для результата " +
                "(при разбиении к имени добавятся номера или метки).\n\n" +
                "Страницы копируются как есть, без переконвертации. Исходный файл не изменяется; " +
                "имена не перезаписываются (при совпадении добавляется номер).\n" +
                "Сжатие меняет содержимое файла, поэтому у подписанных PDF подпись станет " +
                "недействительной (как и при сжатии в Acrobat) — сжимайте до подписания.",
                "1. Open a PDF — with “Open PDF…” or by dragging it onto the window; a page grid appears.\n" +
                "2. Choose a mode:\n" +
                "   • “Extract selected” — select pages in the grid (Ctrl+A — all) → saves them into one PDF;\n" +
                "   • “By ranges” — “1-3, 5, 8-”: each range → a separate file;\n" +
                "   • “Every N pages” — equal parts (1 — each page separately);\n" +
                "   • “By bookmarks” — one file per top‑level bookmark, names from the headings.\n" +
                "3. Optionally choose “Compression” (default “Excellent” — no compression): " +
                "“Good”/“Normal” shrink the size by downsampling images " +
                "(as in Acrobat), text is preserved. Ghostscript required.\n" +
                "4. Click “Extract…”/“Split…” and choose a name and folder for the result " +
                "(when splitting, numbers or labels are appended to the name).\n\n" +
                "Pages are copied as‑is, without re‑conversion. The source file is not changed; " +
                "names are not overwritten (a number is added on a clash).\n" +
                "Compression changes the file bytes, so a signed PDF’s signature becomes " +
                "invalid (as with Acrobat) — compress before signing.");

            // pdf.* — инструмент «Объединение PDF» (PdfMergeForm)
            A("pdf.header.subtitle",
                "Объединение документов формата *.pdf с возможностью изменения порядка страниц и сжатием.",
                "Merge *.pdf documents with page reordering and compression.");
            A("pdf.status.addPdf", "Добавьте PDF-файлы — кнопкой или перетащив их в окно.",
                "Add PDF files — with the button or by dragging them onto the window.");
            A("pdf.btn.save", "Сохранить PDF…", "Save PDF…");
            A("pdf.defaultName", "Объединённый.pdf", "Merged.pdf");
            A("pdf.status.saveFailed", "PDF не сохранён.", "PDF was not saved.");
            A("pdf.err.saveFailed", "PDF не сохранён", "PDF was not saved");
            A("pdf.status.saved", "✓ Сохранено страниц: {0}.", "✓ Pages saved: {0}.");
            A("pdf.status.savedCompressed", "✓ Сохранено страниц: {0} · сжато.", "✓ Pages saved: {0} · compressed.");
            A("pdf.help.body",
                "1. Добавьте PDF-файлы — кнопкой «Добавить PDF…» или перетащив их в окно.\n" +
                "2. Появится сетка миниатюр страниц. Масштаб — ползунком внизу или Ctrl+колесо мыши.\n" +
                "3. Задайте порядок: перетаскивайте миниатюры или используйте «◀ Раньше» / «Позже ▶».\n" +
                "   Лишние страницы удаляйте кнопкой «Удалить».\n" +
                "4. При необходимости выберите «Сжатие» (по умолчанию «Отлично» — без сжатия). " +
                "«Хорошо»/«Нормально» уменьшают размер за счёт понижения разрешения изображений " +
                "(как в Acrobat); текст сохраняется. Требуется Ghostscript.\n" +
                "5. «Сохранить PDF…» соберёт один документ в выбранном порядке.\n\n" +
                "Горячие клавиши: Delete — удалить выбранные, Alt+←/→ — порядок, " +
                "Ctrl+A — выделить всё, Ctrl+колесо — масштаб.\n" +
                "Страницы копируются как есть, без переконвертации — сканы, печати и подписи " +
                "не искажаются. Битые и защищённые паролем файлы пропускаются с причиной.\n" +
                "Сжатие меняет содержимое файла, поэтому у подписанных PDF подпись станет " +
                "недействительной (как и при сжатии в Acrobat) — сжимайте до подписания.",
                "1. Add PDF files — with “Add PDF…” or by dragging them onto the window.\n" +
                "2. A grid of page thumbnails appears. Zoom with the slider below or Ctrl+mouse wheel.\n" +
                "3. Set the order: drag thumbnails or use “◀ Earlier” / “Later ▶”.\n" +
                "   Remove pages you don’t need with “Remove”.\n" +
                "4. Optionally choose “Compression” (default “Excellent” — no compression). " +
                "“Good”/“Normal” shrink the size by downsampling images " +
                "(as in Acrobat); text is preserved. Ghostscript required.\n" +
                "5. “Save PDF…” assembles one document in the chosen order.\n\n" +
                "Shortcuts: Delete — remove selected, Alt+←/→ — order, " +
                "Ctrl+A — select all, Ctrl+wheel — zoom.\n" +
                "Pages are copied as‑is, without re‑conversion — scans, stamps and signatures " +
                "are not distorted. Broken and password‑protected files are skipped with a reason.\n" +
                "Compression changes the file bytes, so a signed PDF’s signature becomes " +
                "invalid (as with Acrobat) — compress before signing.");

            // ocr.* — инструмент «PDF → Word» (OcrForm)
            A("ocr.header.subtitle",
                "Извлечение текста и таблиц из документов формата *.pdf с возможностью изменения порядка страниц.",
                "Extract text and tables from *.pdf documents, with page reordering.");
            A("ocr.btn.open", "Добавить PDF…", "Add PDF…");
            A("ocr.tip.open", "Можно выбрать несколько файлов или перетащить их в окно",
                "Pick several files, or drag them onto the window");
            A("ocr.btn.convert", "Конвертировать в Word…", "Convert to Word…");
            A("ocr.tip.convert", "Извлечь текст в редактируемый .docx", "Extract the text into an editable .docx");
            A("ocr.status.addPdf", "Добавьте цифровые PDF — кнопкой или перетащив их в окно.",
                "Add born‑digital PDFs — with the button or by dragging them onto the window.");
            A("ocr.status.pageCount", "Страниц к переводу: {0}.", "Pages to convert: {0}.");
            A("ocr.status.converting", "Конвертация в Word…", "Converting to Word…");
            A("ocr.status.failed", "Не выполнено.", "Failed.");
            A("ocr.status.done", "✓ Готово: страниц {0} → Word (.docx).", "✓ Done: {0} pages → Word (.docx).");
            A("ocr.err.convertFailed", "Конвертация не выполнена", "Conversion failed");
            A("ocr.docxFilter", "Документ Word (*.docx)|*.docx", "Word document (*.docx)|*.docx");
            A("ocr.defaultMerged", "Объединённый.docx", "Merged.docx");
            A("ocr.help.body",
                "1. Добавьте один или несколько PDF — кнопкой «Добавить PDF…» (можно выбрать сразу " +
                "несколько) или перетащив их в окно. Страницы всех файлов показываются одной сеткой.\n" +
                "2. При необходимости измените порядок страниц: перетащите миниатюру или выделите " +
                "её и нажмите «◀ Раньше»/«Позже ▶» (Alt+←/→). Лишние страницы уберите из вывода " +
                "кнопкой «Удалить» (Delete). В Word попадут страницы в показанном порядке.\n" +
                "3. Нажмите «Конвертировать в Word…» и укажите имя .docx — все выбранные страницы " +
                "соберутся в один документ.\n\n" +
                "Извлекается ТЕКСТОВЫЙ СЛОЙ цифровых PDF (например, сохранённых из Word, " +
                "«Microsoft Print to PDF», экспортированных из браузера). Переносятся: текст " +
                "абзацами в порядке чтения — с шрифтом, размером, начертанием, цветом, " +
                "подчёркиванием, выравниванием и красной строкой; таблицы с линиями (границами) " +
                "восстанавливаются ячейками, включая объединённые; книжная и альбомная " +
                "ориентация страниц сохраняется постранично; изображения и гиперссылки.\n\n" +
                "Текущие ограничения перевода в Word:\n" +
                "• Отсканированные документы (страницы-изображения без текстового слоя) не " +
                "поддерживаются — появится сообщение, файл не пострадает.\n" +
                "• Если шрифт из PDF не установлен в системе, текст оформляется шрифтом " +
                "Times New Roman — начертание может немного отличаться от оригинала.\n" +
                "• Таблицы БЕЗ линий (границ), врезки, несколько колонок переносятся " +
                "простыми абзацами в одну колонку — их, возможно, придётся поправить вручную.\n" +
                "• Если PDF сохранён с испорченной кодировкой текста (без корректного ToUnicode), " +
                "извлечённый текст будет нечитаемым — это дефект самого файла, а не конвертации; " +
                "проверить можно, скопировав текст в самом PDF (Ctrl+C).",
                "1. Add one or several PDFs — with “Add PDF…” (you can pick several at once) or by " +
                "dragging them onto the window. Pages of all files are shown in a single grid.\n" +
                "2. Reorder pages if needed: drag a thumbnail, or select it and click " +
                "“◀ Earlier”/“Later ▶” (Alt+←/→). Drop pages you don’t need with " +
                "“Remove” (Delete). Word gets the pages in the order shown.\n" +
                "3. Click “Convert to Word…” and choose a .docx name — all selected pages " +
                "are assembled into one document.\n\n" +
                "The TEXT LAYER of born‑digital PDFs is extracted (e.g. saved from Word, " +
                "“Microsoft Print to PDF”, exported from a browser). Transferred: text as " +
                "paragraphs in reading order — with font, size, weight, colour, underline, " +
                "alignment and first‑line indent; bordered tables are rebuilt as cells, " +
                "including merged ones; portrait/landscape orientation is kept per page; " +
                "images and hyperlinks.\n\n" +
                "Current limitations:\n" +
                "• Scanned documents (image pages with no text layer) are not supported — " +
                "a message is shown, the file is untouched.\n" +
                "• If a PDF font is not installed, the text is set in Times New Roman — the " +
                "look may differ slightly from the original.\n" +
                "• Tables WITHOUT ruled borders, text boxes and multi‑column layouts are " +
                "flattened to single‑column paragraphs — you may need to fix them by hand.\n" +
                "• If the PDF was saved with broken text encoding (no proper ToUnicode), the " +
                "extracted text will be unreadable — a defect of the file, not the conversion; " +
                "check by copying text inside the PDF itself (Ctrl+C).");

            // excel.* — инструмент «Свод Excel» (MainForm)
            A("excel.defaultName", "Свод_", "Digest_");
            A("excel.noteFileSuffix", " — записка.docx", " — note.docx");
            A("excel.header.subtitle", "Объедините листы Excel-файлов из папки в один свод.",
                "Merge sheets of the Excel files in a folder into one digest.");
            A("excel.sec.inputFolder", "ПАПКА С ИСХОДНЫМИ ФАЙЛАМИ", "SOURCE FILES FOLDER");
            A("excel.sec.output", "ИТОГОВЫЙ ФАЙЛ", "OUTPUT FILE");
            A("excel.sec.params", "ПАРАМЕТРЫ", "OPTIONS");
            A("excel.sec.files", "ФАЙЛЫ ДЛЯ ОБЪЕДИНЕНИЯ", "FILES TO MERGE");
            A("excel.lbl.name", "Имя:", "Name:");
            A("excel.lbl.folder", "Папка:", "Folder:");
            A("excel.lbl.sheets", "Листы:", "Sheets:");
            A("excel.tip.format", "Формат итогового файла; .xls — старый формат Excel 97–2003",
                "Output file format; .xls is the old Excel 97–2003 format");
            A("excel.scope.first", "Только первый лист", "First sheet only");
            A("excel.scope.all", "Все листы", "All sheets");
            A("excel.tip.scope", "Из каждого файла брать только первый видимый лист или все видимые",
                "Take only the first visible sheet of each file, or all visible sheets");
            A("excel.chk.toc", "Добавить лист «Содержание» с оглавлением и ссылками",
                "Add a “Contents” sheet with a table of contents and links");
            A("excel.tip.toc", "Первым листом свода будет оглавление: гиперссылки на листы и статусы всех файлов",
                "The first sheet becomes a table of contents: hyperlinks to sheets and each file’s status");
            A("excel.chk.values", "Заменить формулы значениями", "Replace formulas with values");
            A("excel.tip.values", "Свод не будет зависеть от исходных файлов: вместо формул — вычисленные значения",
                "The digest won’t depend on the sources: computed values instead of formulas");
            A("excel.btn.merge", "Объединить", "Merge");
            A("excel.tip.merge", "Собрать свод из файлов выбранной папки (Enter)",
                "Build the digest from the files in the chosen folder (Enter)");
            A("excel.btn.cancel", "Отменить", "Cancel");
            A("excel.tip.cancel", "Остановить после текущего файла (Esc)", "Stop after the current file (Esc)");
            A("excel.btn.up", "▲ Выше", "▲ Up");
            A("excel.tip.up", "Переместить выбранный файл выше (Alt+↑)", "Move the selected file up (Alt+↑)");
            A("excel.btn.down", "▼ Ниже", "▼ Down");
            A("excel.tip.down", "Переместить выбранный файл ниже (Alt+↓)", "Move the selected file down (Alt+↓)");
            A("excel.btn.sortName", "По имени", "By name");
            A("excel.tip.sortName", "Вернуть естественный порядок по имени файла", "Restore natural order by file name");
            A("excel.btn.checkAll", "Отметить все", "Check all");
            A("excel.btn.uncheckAll", "Снять все", "Uncheck all");
            A("excel.btn.retry", "Повторить пропущенные", "Retry skipped");
            A("excel.tip.retry", "Дослить исправленные файлы в существующий свод без полного пересбора",
                "Append fixed files to the existing digest without a full rebuild");
            A("excel.link.openFile", "Открыть файл", "Open file");
            A("excel.link.openFolder", "Открыть папку", "Open folder");
            A("excel.link.openReport", "Открыть отчёт", "Open report");
            A("excel.tip.openReport", "Отчёт о слиянии; в истории хранятся три последних",
                "The merge report; the three latest are kept in history");
            A("excel.link.note", "Записка Word", "Word note");
            A("excel.tip.note", "Сопроводительная записка к своду (.docx): итоги, пропущенные файлы, оформление по ГОСТ",
                "A cover note for the digest (.docx): totals, skipped files, GOST formatting");
            A("excel.tip.input", "Папку можно перетащить мышью в окно программы", "You can drag a folder onto the window");
            A("excel.tip.name", "Расширение .xlsx добавится автоматически", "The .xlsx extension is added automatically");
            A("excel.tip.outDir", "Пусто — итоговый файл сохранится в папку с исходными",
                "Empty — the output is saved next to the sources");
            A("excel.menu.reports", "Папка отчётов", "Reports folder");
            A("excel.pick.input", "Папка с исходными файлами Excel", "Folder with the source Excel files");
            A("excel.pick.output", "Папка для сохранения итогового файла", "Folder to save the output file");
            A("excel.status.chooseFolder", "Выберите папку с исходными файлами.", "Choose the folder with the source files.");
            A("excel.status.startingExcel", "Запуск Excel…", "Starting Excel…");
            A("excel.status.fileProgress", "Файл {0} из {1}: {2}", "File {0} of {1}: {2}");
            A("excel.status.failed", "Объединение не выполнено.", "Merge failed.");
            A("excel.status.cancelled", "Отменено — итоговый файл не создан.", "Cancelled — no output file was created.");
            A("excel.status.doneWithSkips", "Готово: перенесено {0}, пропущено {1} — причины в списке.",
                "Done: {0} transferred, {1} skipped — reasons are in the list.");
            A("excel.status.doneClean", "✓ Готово: перенесено листов — {0}.", "✓ Done: sheets transferred — {0}.");
            A("excel.status.tocWarn", " Внимание: {0}.", " Note: {0}.");
            A("excel.status.cancelling", "Отмена после текущего файла…", "Cancelling after the current file…");
            A("excel.status.finishing", "Завершение…", "Finishing…");
            A("excel.status.noteBusy", "Готовится записка Word…", "Preparing the Word note…");
            A("excel.status.noteFailed", "Записка не создана.", "The note was not created.");
            A("excel.status.noteSaved", "Записка сохранена рядом со сводом.", "The note was saved next to the digest.");
            A("excel.status.waitNote", "Дождитесь завершения записки Word…", "Wait for the Word note to finish…");
            A("excel.found.chooseFolder", "Укажите папку или перетащите её в окно.", "Choose a folder or drag it onto the window.");
            A("excel.found.notFound", "Папка не найдена.", "Folder not found.");
            A("excel.found.readError", "Не удалось прочитать папку: {0}", "Could not read the folder: {0}");
            A("excel.found.noExcel", "Файлы Excel (.xlsx, .xls, .xlsm, .xlsb) не найдены.",
                "No Excel files (.xlsx, .xls, .xlsm, .xlsb) found.");
            A("excel.found.count", "Найдено файлов: {0}, выбрано: {1}", "Files found: {0}, selected: {1}");
            A("excel.col.file", "Файл", "File");
            A("excel.col.result", "Результат", "Result");
            A("excel.col.note", "Примечание", "Note");
            A("excel.row.skipped", "✗ пропущен", "✗ skipped");
            A("excel.row.moved", "✓ перенесён", "✓ transferred");
            A("excel.row.sheets", "✓ листов: {0}", "✓ sheets: {0}");
            A("excel.row.sheetsPartial", "⚠ листов: {0} из {1}", "⚠ sheets: {0} of {1}");
            A("excel.err.openReports", "Не удалось открыть папку отчётов", "Could not open the reports folder");
            A("excel.err.folderNotFound.title", "Папка с исходными файлами не найдена", "Source files folder not found");
            A("excel.err.folderNotFound.body", "Проверьте путь: {0}", "Check the path: {0}");
            A("excel.err.noName.title", "Укажите имя итогового файла", "Enter the output file name");
            A("excel.err.noName.body", "Поле «Имя» не заполнено.", "The “Name” field is empty.");
            A("excel.err.badName.title", "Недопустимое имя файла", "Invalid file name");
            A("excel.err.badName.body", "Имя не должно содержать символы  \\ / : * ? \" < > |",
                "The name must not contain  \\ / : * ? \" < > |");
            A("excel.confirm.createFolder.title", "Папка сохранения не существует", "Save folder does not exist");
            A("excel.confirm.createFolder.body", "Создать папку?\n{0}", "Create the folder?\n{0}");
            A("excel.err.createFolder", "Не удалось создать папку", "Could not create the folder");
            A("excel.err.noFiles.title", "Не выбрано ни одного файла", "No files selected");
            A("excel.err.noFiles.body", "Отметьте галочками файлы для объединения.", "Tick the files to merge.");
            A("excel.confirm.overwrite.title", "Файл уже существует", "File already exists");
            A("excel.confirm.overwrite.body", "Файл «{0}» уже есть в папке сохранения.\nПерезаписать его?",
                "“{0}” already exists in the save folder.\nOverwrite it?");
            A("excel.err.outputLocked", "Итоговый файл недоступен для записи", "The output file is not writable");
            A("excel.err.mergeFailed.title", "Объединение не выполнено", "Merge failed");
            A("excel.err.noteFailed.title", "Записка не создана", "The note was not created");
            A("excel.confirm.closeBusy.title", "Идёт объединение", "Merge in progress");
            A("excel.confirm.closeBusy.body", "Прервать объединение и закрыть программу?", "Stop the merge and close the app?");
            A("excel.help.body",
                "1. Укажите папку с исходными файлами — «Обзор…» или перетащите папку в окно.\n" +
                "2. Задайте имя свода; папку сохранения можно сменить (пустая — папка с исходными).\n" +
                "3. В списке «Файлы для объединения» задайте порядок и состав: перетаскиванием " +
                "строк или кнопками «▲ Выше»/«▼ Ниже»; снимите галочку у ненужного файла. " +
                "«По имени» вернёт естественный порядок, «Отметить все»/«Снять все» — быстрый выбор.\n" +
                "4. Нажмите «Объединить»: из каждого выбранного файла переносится первый видимый " +
                "лист со всем оформлением, формулами и диаграммами.\n\n" +
                "Параметры:\n" +
                "• «Листы» — только первый видимый лист каждого файла или все видимые,\n" +
                "• лист «Содержание» — оглавление свода с гиперссылками и статусами файлов,\n" +
                "• «Заменить формулы значениями» — свод не зависит от исходных файлов.\n\n" +
                "После слияния результат по каждому файлу виден в тех же строках. Битые " +
                "и запароленные файлы пропускаются, причина видна в списке и в отчёте.\n\n" +
                "Горячие клавиши в списке: Alt+↑/↓ — порядок, Delete — исключить, " +
                "Ctrl+A — выделить всё, Ctrl+C — копировать.\n" +
                "Отчёты (три последних): ☰ Меню → «Папка отчётов».",
                "1. Choose the source files folder — “Browse…” or drag a folder onto the window.\n" +
                "2. Set the digest name; you can change the save folder (empty — the sources folder).\n" +
                "3. In the “Files to merge” list set the order and selection: drag rows or use " +
                "“▲ Up”/“▼ Down”; untick a file you don’t need. " +
                "“By name” restores natural order, “Check all”/“Uncheck all” select quickly.\n" +
                "4. Click “Merge”: the first visible sheet of each selected file is transferred " +
                "with all its formatting, formulas and charts.\n\n" +
                "Options:\n" +
                "• “Sheets” — the first visible sheet of each file, or all visible sheets,\n" +
                "• the “Contents” sheet — a table of contents with hyperlinks and file statuses,\n" +
                "• “Replace formulas with values” — the digest won’t depend on the sources.\n\n" +
                "After the merge each file’s result appears in the same rows. Broken and " +
                "password‑protected files are skipped, with the reason shown in the list and report.\n\n" +
                "List shortcuts: Alt+↑/↓ — order, Delete — exclude, " +
                "Ctrl+A — select all, Ctrl+C — copy.\n" +
                "Reports (three latest): ☰ Menu → “Reports folder”.");

            return d;
        }
    }
}
