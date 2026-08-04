using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using ExcelMerger;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Юнит-тесты без внешних фреймворков: компилируются вместе с src\*.cs
    /// в отдельный консольный exe (tests\build_tests.cmd) и не попадают
    /// в производственную сборку. Код выхода 0 — все тесты прошли.
    /// </summary>
    internal static class UnitTests
    {
        private static int _passed;
        private static int _failed;

        private static int Main()
        {
            Run("SheetNamer: запрещённые символы заменяются на «_»", TestNamerForbiddenChars);
            Run("SheetNamer: обрезка до 31 символа", TestNamerTruncation);
            Run("SheetNamer: дедупликация суффиксами _2, _3", TestNamerDedupe);
            Run("SheetNamer: дедупликация длинного имени не превышает 31", TestNamerDedupeLong);
            Run("SheetNamer: Reserve исключает имя", TestNamerReserve);
            Run("SheetNamer: пустое имя становится «Лист»", TestNamerEmpty);
            Run("SheetNamer: History зарезервировано Excel", TestNamerHistory);
            Run("SheetNamer: апострофы по краям срезаются", TestNamerApostrophes);
            Run("Natural: «Отчет 2» раньше «Отчет 10»", TestNaturalNumbers);
            Run("Natural: регистр не учитывается", TestNaturalCase);
            Run("Natural: null меньше любой строки", TestNaturalNulls);
            Run("Natural: сортировка списка как в Проводнике", TestNaturalSortList);
            Run("FindSourceFiles: фильтры, исключения, порядок", TestFindSourceFiles);
            Run("CellText: строки экранируются апострофом", TestCellTextEscape);
            Run("CellText: не-строки проходят без изменений", TestCellTextPassthrough);
            Run("CellText: массив области (включая 1-базный COM)", TestCellTextArrays);
            Run("CLI: разбор флагов --toc/--values", TestCliOptions);
            Run("CLI: неизвестный флаг — ошибка", TestCliUnknownFlag);
            Run("ReportWriter: формат строки файла", TestReportLine);
            Run("ReportWriter: содержимое полного отчёта", TestReportBuild);
            Run("ReportWriter: ротация хранит не более 3 отчётов", TestReportRotation);
            Run("ReportWriter: коллизия имён в одну секунду", TestReportNameCollision);
            Run("ToolRegistry: открыт/закрыт, авто-удаление Disposed", TestToolRegistry);
            Run("HelpMenu: структура «Справка» и вставка доп. пунктов", TestHelpMenu);
            Run("ThumbZoom: кламп ширины плитки", TestThumbZoomClamp);
            Run("ThumbZoom: колесо меняет масштаб и упирается в границы", TestThumbZoomWheel);
            Run("ThumbZoom: высота плитки пропорциональна", TestThumbZoomTile);
            Run("ListReorder: MoveUp/Down границы и обмен", TestListReorderMoves);
            Run("ListReorder: Move (перенос) и RemoveAt (набор)", TestListReorderMoveRemove);
            Run("SourceFileList: порядок, включение, IncludedInOrder", TestSourceFileList);
            Run("SourceFileList: сортировка по имени как в Проводнике", TestSourceFileSort);
            Run("PrepareSourceList: исключение свода и дубликатов", TestPrepareSourceList);
            Run("FileSignature: ZIP/OLE2 — книга, текст/пусто — нет", TestFileSignature);
            Run("LowSpaceMessage: мало места — понятная ошибка, иначе null", TestLowSpaceMessage);
            Run("PageRanges.Parse: диапазоны, пробелы, открытый конец", TestPageRangesParse);
            Run("PageRanges.Parse: неверный ввод — ошибка", TestPageRangesParseErrors);
            Run("PageRanges.EveryN: нарезка на равные части", TestPageRangesEveryN);
            Run("PageRanges.ToIndices: диапазоны -> индексы (порядок, повторы)", TestPageRangesToIndices);
            Run("UpdateChecker: разбор тега и сравнение версий", TestUpdateChecker);
            Run("UsageStats.ShouldAutoClear: период очистки", TestShouldAutoClear);
            Run("UsageStats.Total: включает PdfToWord, исключает сжатия", TestUsageTotal);
            Run("MessageForm.ButtonX: одна по центру, две по краям", TestMessageButtonX);
            Run("PdfSplitService.Sanitize: недопустимые символы", TestSanitize);
            Run("PdfSplit (живой): извлечение, диапазоны, каждые N, закладки", TestPdfSplitLive);
            Run("PdfPageGrid.ClampWindow: окно видимых с буфером", TestClampWindow);
            Run("Theme.ToBgr: упаковка цвета 0x00BBGGRR", TestThemeToBgr);
            Run("TocBuilder.SheetRef: ссылка на A1, апострофы удвоены", TestSheetRef);
            Run("WindowChrome: COLORREF упакован как 0x00BBGGRR", TestWindowChromeColorRef);
            Run("HeaderBand: строится с заголовком, двойная буферизация", TestHeaderBand);
            Run("HeaderBand.TextRightBound: текст не заходит под кнопку", TestHeaderTextBound);
            Run("MainForm.ClassifyListKey: Alt+↑/↓, Ctrl+C/A, Delete, Enter", TestClassifyListKey);
            Run("PdfToolFormBase.ClassifyPageKey: Delete, Alt+←/→, Ctrl+A, Enter", TestClassifyPageKey);
            Run("PdfPageOrder: добавление и границы MoveUp/MoveDown", TestPdfOrderMoves);
            Run("PdfPageOrder: перенос drag&drop в обе стороны", TestPdfOrderDragMove);
            Run("PdfPageOrder: удаление набора строк + Clear", TestPdfOrderRemove);
            Run("PdfPageExtraction.Assemble: сборка страниц из нескольких PDF, границы", TestAssemble);
            Run("NoteText: период, счётчики, файл свода", TestNoteBasics);
            Run("NoteText: таблица пропущенных", TestNoteSkippedTable);
            Run("NoteText: без пропусков — «замечания отсутствуют»", TestNoteClean);
            Run("SheetBaseName: первый лист vs все листы", TestSheetBaseName);
            Run("MergeResult.FileCount: файлы, а не листы", TestFileCount);
            Run("Retry: один пропущенный файл -> несколько листов", TestCombineRetryMultiSheet);
            Run("Retry: пропущенные заменяются свежими результатами", TestCombineRetryReplaces);
            Run("Retry: неудачный повтор обновляет причину", TestCombineRetryKeepsFailed);
            Run("Retry: порядок и успешные записи не меняются", TestCombineRetryOrder);
            Run("OutputFormats: код формата по расширению", TestOutputFormatCodes);
            Run("OutputFormats: срез введённого расширения", TestStripExtension);
            Run("CrashReport.Format: метка времени, версия, текст исключения", TestCrashReportFormat);
            Run("MergeService.IsPermanentOpenError: ru/en паттерны, транзиент — нет", TestPermanentOpenError);
            Run("MergeException.ShouldWrap: OOM не маскируется под «файл повреждён»", TestShouldWrap);
            Run("CheckOutputWritable: занятый файл распознан", TestOutputLocked);
            Run("CheckOutputWritable: свободный и новый файлы", TestOutputWritable);
            Run("CheckOutputWritable: несуществующая папка", TestOutputBadFolder);
            Run("PdfCompression.Preset: уровень -> пресет Ghostscript", TestCompressionPreset);
            Run("Сжатие: число уровня — идентификатор, порядок показа — отдельно", TestCompressionLevelOrder);
            Run("Сжатие: «Очень хорошо» не пересчитывает изображения (обещание в аргументах)", TestCompressionKeepsResolution);
            Run("Сжатие: уровень переживает запись в настройки и совместим с прежними", TestCompressionLevelRoundTrip);
            Run("PdfCompression.BuildArguments: кавычки, пресет, -I для бандла", TestCompressionArgs);
            Run("PdfCompression.ShouldReplace: только валидный и строго меньше", TestCompressionShouldReplace);
            Run("Ghostscript.PickFirstExisting: первый существующий из кандидатов", TestGhostscriptPick);
            Run("PdfCompression (живой): крупный PDF сжимается, страницы целы", TestCompressLive);
            Run("AboutForm: пункты доната валидны (20 цифр, банк не пуст)", TestDonationRequisites);
            Run("AboutForm: описание выделяемое и говорит про цифровой PDF", TestAboutDescription);
            Run("SingleInstance: имя занимается первым и освобождается после выхода", TestSingleInstanceName);
            Run("LruCache: вытеснение самого несвежего, Count, порядок", TestLruEviction);
            Run("LruCache: touch через TryGet переносит вытеснение", TestLruTouchOnGet);
            Run("LruCache: замена ключа не растит размер и не вытесняет", TestLruReplace);
            Run("LruCache: ключи регистронезависимы (пути файлов)", TestLruCaseInsensitive);
            Run("LruCache: Clear освобождает все элементы", TestLruClear);
            Run("LruCache: ёмкость < 1 запрещена", TestLruCapacityGuard);
            Run("LruCache: TryPeek не освежает, Remove без onEvict, KeysSnapshot", TestLruPeekRemove);
            Run("PdfPageRef.ComposeRotation: нормализация в {0,90,180,270}", TestRotationCompose);
            Run("PdfPageRef.Clone: независимый снимок страницы", TestPageRefClone);
            Run("ListReorder.MoveRange: перенос набора и позиция вставки", TestMoveRange);
            Run("ListReorder: AdjustedInsertIndex и NormalizeIndices", TestMoveRangeHelpers);
            Run("PdfPageOrder: InsertDocument/InsertAt в позицию с клампом", TestPdfOrderInsertAt);
            Run("ThumbZoom.RenderWidthFor: DPI, база и крупный зум, границы", TestRenderWidthFor);
            Run("ThumbZoom.CellSize: ячейка с полосой номера, максимум 400", TestThumbZoomCell);
            Run("PdfPageGrid.TileRectFromIcon: плитка по центру иконной зоны", TestTileRectFromIcon);
            Run("PdfPageGrid.HoverRotateButtons: два чипа у нижней кромки, без пересечений", TestHoverRotateButtons);
            Run("PdfPageGrid.CellRectFromTile: ячейка накрывает плитку и чипы на максимуме зума", TestCellRectCoversChips);
            Run("PdfPageOrder: Checkpoint/Undo, лимит, повороты не откатываются", TestPdfOrderUndo);
            Run("ThumbZoom.PageCacheCapacity: бюджет → ёмкость с границами", TestPageCacheCapacity);
            Run("PdfPageGrid: PageLabel (позиция/исходная), TileKey с поворотом, FlipFor", TestGridLabelTileKey);
            Run("PdfPageGrid.PasteIndex: каретка, выделение, конец", TestPasteIndex);
            Run("PdfToolFormBase.ClassifyPageKey: буфер, поворот, Ctrl+G, Esc", TestClassifyClipboardKeys);
            Run("PdfToolFormBase.IsResetZoomKey: Ctrl+0 (оба ряда)", TestIsResetZoomKey);
            Run("PdfSplitService.RotationAt: карта поворотов и границы", TestRotationAt);
            Run("Merge (живой): /Rotate пишется и складывается с исходным", TestMergeRotationLive);
            Run("SplitByRanges (живой): карта поворотов по индексу страницы", TestSplitRotationLive);
            Run("PageRotation: MapPoint/MapBox по углам, Inverse, At, свап размеров", TestPageRotationMap);
            Run("PageRotation.RotatePage: слова, линии H<->V, картинки (рамки+пиксели)", TestPageRotationRotatePage);
            Run("PdfTextExtract.ApplyUnderscoreHeights: прочерк растёт, слово не трогается", TestUnderscoreHeights);
            Run("PdfPageExtraction.BuildRotations: первый экземпляр решает, рост карты, null", TestBuildRotations);
            Run("PDF->Word (живой): боковой текст выправляется поворотом страницы", TestExtractRotationLive);
            Run("PdfDrop.ExtractPaths: фильтр .pdf, несуществующие, пустой дроп", TestPdfDropExtract);
            Run("PdfPageGrid.DropInsertIndex: позиция дропа по метке/в конец", TestDropInsertIndex);
            Run("PdfPageGrid.HintFor: подсказка пусто/дроп, скрыта при непустом", TestHintFor);
            Run("PdfToolFormBase.RestingStatus: счётчик выделения или idle", TestRestingStatus);
            Run("PdfToolFormBase.SuccessStatus: галочка, разделители, точка", TestSuccessStatus);
            Run("PdfCompression.ImageDpi: 150/72 совпадают с пресетами Ghostscript", TestCompressionDpi);
            Run("Итоговые строки объединения и разделения на обоих языках", TestDoneStatusLines);
            Run("Предпросмотр: доворот картинки до поворота страницы", TestPreviewRotationDelta);
            Run("PdfToolFormBase.ProgressItem: страница из процента, границы", TestProgressItem);
            Run("PdfToolFormBase.BuildShortcuts: набор клавиш по возможностям", TestBuildShortcuts);
            Run("ChoiceCard.FilterByExtension: фильтр по расширению и существованию", TestCardFilter);
            Run("ThumbZoom.Percent: масштаб в процентах от умолчания", TestZoomPercent);
            Run("ThumbZoom.WidthFromPercent: обратная к Percent, round-trip и кламп", TestWidthFromPercent);
            Run("PdfPageGrid.IsOnLabel: точка в полосе номера, не в плитке", TestIsOnLabel);
            Run("PdfToolFormBase.ShouldOfferCancel: отмена от порога страниц", TestShouldOfferCancel);
            Run("UserSettings: общий Save не затирает масштаб/сжатие устаревшим экземпляром", TestSettingsViewNotClobbered);
            Run("UserSettings: границы окон сохраняются и не затираются устаревшим Save", TestWindowBoundsPersistence);
            Run("WindowPlacement.Attach (живой): окно запоминает и восстанавливает место", TestWindowPlacementAttachLive);
            Run("Merge (живой): отмена бросает и не создаёт файл", TestCancelMergeLive);
            Run("SplitEveryN (живой): отмена бросает и удаляет частичные файлы", TestCancelSplitLive);
            Run("ToImages (живой): отмена бросает и удаляет уже сохранённые картинки", TestCancelExportImagesLive);
            Run("Прочие операции (живое окно): «Отмена» показывается и поднимает токен", TestOpsCancelWiringLive);
            Run("PdfPageGrid.BuildKeySet: ключи набора без дублей, null -> пусто", TestGridBuildKeySet);
            Run("PdfPageGrid.StaleKeys: вытесняются только отсутствующие в keep", TestGridStaleKeys);
            Run("PdfPageGrid.LowerBound: бинарный поиск по монотонному предикату", TestLowerBound);
            Run("PdfPageGrid.VisibleRange: видимый диапазон по Top/Bottom", TestVisibleRange);
            Run("PdfSplitForm.ShouldSuggestCompression: без сжатия и ≥90% исходника", TestSuggestCompression);
            Run("OcrLayout: порядок чтения (сверху вниз, слева направо)", TestOcrReadingOrder);
            Run("OcrLayout: разрыв абзаца по вертикальному зазору", TestOcrParagraphs);
            Run("OcrLayout: разрыв абзаца по красной строке (justified)", TestOcrParagraphsIndent);
            Run("OcrLayout: измерен отступ красной строки", TestOcrIndentDetected);
            Run("OcrLayout: без отступов -> красная строка не навязывается", TestOcrNoIndentReported);
            Run("FontNames.Clean: нормализация имени шрифта", TestFontNames);
            Run("WordDocxWriter: неустановленный шрифт -> fallback (против eastAsia-разрядки)", TestResolveFontName);
            Run("PdfToolFormBase: проценты прогресса (сделано/всего, клампы)", TestProgressPercent);
            Run("OcrLayout: смена шрифта -> раны", TestOcrRunsFontFamily);
            Run("OcrLayout: стиль рана (курсив, кегль)", TestOcrParagraphStyle);
            Run("OcrLayout: смешанный формат -> раны", TestOcrRunsMixedFormat);
            Run("OcrLayout: надстрочный ран", TestOcrSuperscript);
            Run("OcrLayout: ран-гиперссылка", TestOcrHyperlinkRun);
            Run("OcrLayout: цвет рана сохранён", TestOcrColorRun);
            Run("OcrLayout: рваный абзац по левому краю", TestOcrLeftAligned);
            Run("OcrLayout: центрированная строка", TestOcrCentered);
            Run("OcrLayout: IsCentered — узкая/широкая/красная строка/рваная/правая", TestIsCenteredPredicate);
            Run("OcrLayout: центрированный многострочный титул -> один центрированный абзац", TestOcrCenteredBlock);
            Run("OcrLayout: фрагменты слова склеиваются по зазору", TestOcrGlueFragments);
            Run("OcrLayout: узкий настоящий пробел сохранён (не склейка)", TestOcrNarrowSpaceKept);
            Run("OcrLayout: тонкое тире остаётся в строке", TestOcrThinDashStaysOnLine);
            Run("OcrLayout: перенос с дефисом склеивает слово", TestOcrHyphenation);
            Run("OcrLayout: кириллический дефис-перенос сохраняется (составное слово)", TestOcrHyphenCyrillicKept);
            Run("OcrLayout: умышленный перевод строки рвёт абзац (подпись)", TestOcrHardLineBreak);
            Run("OcrLayout: подпись в левой колонке рвётся по строкам (доступное место)", TestOcrSignatureHardBreak);
            Run("OcrLayout: мелкая цифра у базовой линии — сносочный маркер (Super)", TestOcrFootnoteDigitSuper);
            Run("OcrLayout: строка с красной строкой не центрируется случайной симметрией", TestOcrRedIndentNotCentered);
            Run("WordDocxWriter: межблочные интервалы — типичный зазор, лишек, кап", TestDocxGapMath);
            Run("ShellContext.MoveActiveLast: активное окно пересоздаётся последним", TestMoveActiveLast);
            Run("WindowPlacement: сериализация границ окна и отбраковка мусора", TestWindowBoundsRoundTrip);
            Run("WindowPlacement.ClampToWorkingArea: край/вне экрана/мин/макс/мультимонитор", TestClampToWorkingArea);
            Run("OcrLayout: гигантский зазор внутри строки рвёт её на зоны (в ячейке — нет)", TestOcrWideGapSplit);
            Run("PageBlocks.CoalesceRowBands: блоки одной строки — в полосу", TestCoalesceRowBands);
            Run("WordDocxWriter.AnchorIndents: красная строка по факту / позиция колонки", TestAnchorIndents);
            Run("OcrLayout: дефис лат+кириллица на переносе сохраняется", TestOcrHyphenMixedKept);
            Run("PdfTextExtract: слово под низкой наложенной картинкой скрывается", TestCoveredByLowImage);
            Run("OcrLayout: пустой ввод -> нет абзацев", TestOcrEmpty);
            Run("ListMarker: нумерованный «1.»/«12)»", TestListMarkerNumbered);
            Run("ListMarker: маркированный «•»/«—»", TestListMarkerBulleted);
            Run("ListMarker: не список (год, проценты, без пробела, обычный текст)", TestListMarkerNegatives);
            Run("OcrLayout: плотный нумерованный список -> отдельные пункты с ListKind", TestOcrNumberedList);
            Run("OcrLayout: маркированный список -> ListKind=Bulleted, содержимое без маркера", TestOcrBulletedList);
            Run("StampDetector: штамп -> область со всеми словами, вне полосы не берётся", TestStampDetected);
            Run("StampDetector: нет одного опорного слова -> не штамп", TestStampMissingAnchor);
            Run("StampDetector: опорные слова разбросаны по странице -> не штамп", TestStampScatteredRejected);
            Run("Loc: каталог полон — у каждого ключа непустые ru и en", TestLocCatalogComplete);
            Run("Loc: каждый запрашиваемый кодом ключ есть в каталоге", TestLocKeysUsedInCodeExist);
            Run("Loc: плейсхолдеры {N} у ru и en совпадают", TestLocPlaceholders);
            Run("Loc: перетаскивание — повелительное и «окно программы»", TestLocDragHints);
            Run("Loc: Init/Current/Parse/Code", TestLocInit);
            Run("Loc: EN — в построенных формах нет кириллицы (кроме двуязычных меток)", TestNoCyrillicInEnglishForms);

            Run("TableDetector: сетка 2x2 -> строки/колонки, текст ячеек", TestTable2x2);
            Run("TableDetector: пропущенная гориз. граница -> rowspan", TestTableRowSpan);
            Run("TableDetector: пропущенная верт. граница -> colspan", TestTableColSpan);
            Run("TableDetector: одиночные линии (подчёркивания) -> не таблица", TestTableStrayLines);
            Run("TableDetector: рамка 1x1 без внутренних линий -> не таблица", TestTableSingleBox);
            Run("TableDetector: слова вне таблицы остаются в потоке", TestTableWordsOutside);
            Run("TableDetector: нет линий -> нет таблиц, все слова в потоке", TestTableNoLines);
            Run("PdfPageExtraction: страница-таблица считается текстовой (не «скан»)", TestHasExtractableContent);
            Run("UnderlineDetector: линия под словом -> подчёркнуто", TestUnderlineMarks);
            Run("UnderlineDetector: далёкая/короткая линия -> не подчёркнуто", TestUnderlineIgnores);
            Run("UnderlineDetector: линия во всю ширину (разделитель) -> не подчёркнуто", TestUnderlineWideRule);
            Run("OcrLayout: левый сайдбар отделяется от тела (не перемешиваются)", TestSidebarSeparation);
            Run("OcrLayout: одноколоночный текст не делится (сайдбар не срабатывает)", TestNoSidebarSingleColumn);
            Run("SetMargins: поля по словам И картинкам; нижнее — с капом", TestMarginsWithImages);
            Run("ColumnConfineIndents: правая колонка -> левый отступ; левая -> правый; полная ширина -> нет", TestColumnConfine);
            Run("FontResolver.HasCyrillic: кириллица / латиница / пусто", TestHasCyrillic);
            Run("GridDetector: широкий одиночный ряд -> ColSpan во всю ширину", TestGridColSpan);
            Run("GridDetector: увеличенный зазор между группами -> интервал после строки", TestGridRowSpacing);
            Run("TableDetector: прочерк кусками (коллинеарные) -> LoneLines, не таблица", TestLoneCollinearRule);
            Run("TableDetector: линия внутри рамки таблицы -> не прочерк (LoneLines пуст)", TestRuleInsideTableExcluded);
            Run("GlyphDedup: сдвоенные глифы схлопываются, слово помечается жирным", TestGlyphDedupDoubled);
            Run("GlyphDedup: настоящие соседние одинаковые символы («77») не склеиваются", TestGlyphDedupRealPair);
            Run("GlyphDedup: единичный дубль чистится, но слово не жирное", TestGlyphDedupSparse);
            Run("XyCut: существенный блок + одинокая пометка справа — режется", TestXyCutOneSubstantialColumn);
            Run("GridDetector: чек «метка … значение» -> безграничная таблица", TestGridReceipt);
            Run("GridDetector: justified-текст без широких зазоров — не сетка", TestGridJustifiedNegative);
            Run("GridDetector: две строки пар — мало для сетки", TestGridTwoRowsNegative);
            Run("AddRuleWords: одиночная линия -> прочерк «____», подчёркивание/толстая — нет", TestRuleWords);
            Run("PdfTextExtract: «_»-слово и PUA-мусор распознаются", TestUnderscoreAndPua);
            Run("XyCut: колонки с общими базовыми линиями — левая целиком раньше правой", TestXyCutColumns);
            Run("XyCut: этаж под колонками выводится после обеих колонок", TestXyCutFloorsThenColumns);
            Run("XyCut: широкий пробел одной строки («подпись … дата») — не колонки", TestXyCutGuardSingleLine);
            Run("XyCut: крошка рядом с существенной колонкой — режется (пометка)", TestXyCutGuardThinColumn);
            Run("OcrLayout: двухколоночная шапка — абзацы колонок не смешаны, левая раньше", TestOcrTwoColumnsSeparated);
            Run("OcrLayout: ячейка таблицы (splitColumns=false) — «метка … число» одной строкой", TestOcrCellNoColumns);
            Run("OcrLayout: шапка не размывает красную строку тела (гейт по justified)", TestOcrIndentWithHeader);
            Run("PageBlocks.OrderedItems: колонки -> side-by-side полоса (левая|правая), нижний одиночный", TestOrderedItemsColumns);
            Run("PageBlocks.OrderedItems: внутри листа — строки сверху вниз, бок о бок слева направо", TestOrderedItemsWithinLeaf);
            Run("WordDocxWriter.BandColumnWidths: границы ячеек по середине зазора колонок", TestBandColumnWidths);
            Run("WordDocxWriter: центрированное изображение (логотип) -> по центру, врезка/штамп -> нет", TestImageCentered);
            Run("PageRasterizer: рамка PDF (Y-вверх) -> пиксельный кроп, кламп по краю", TestCropRect);

            // ---------- 1.17.4 ----------
            Run("PdfPageGrid.ClassifyDoubleClick: чип поворота > номер > предпросмотр", TestClassifyDoubleClick);
            Run("NumberPromptDialog.DialogWidth: база или подпись с полями", TestNumberDialogWidth);
            Run("Ui.AppIcon: один экземпляр на процесс (HICON не плодится)", TestAppIconCached);
            Run("Ui.Font: кэш по размеру/стилю/семейству, один экземпляр", TestFontCached);
            Run("PdfPageGrid.WheelBasis: неприменённая цель колеса старше текущей ширины", TestWheelBasis);

            // ---------- 1.17.8 ----------
            Run("Окна: при минимальном размере ничего не обрезано и кнопки не наложены", TestWindowsSurviveMinimumSize);
            Run("Диалоги: ничего не свисает из окна и контролы не наложены", TestDialogsLayoutIsSound);
            Run("Кнопка: радиус и поле подписи считаются от размера", TestRoundedButtonMetrics);
            Run("Карточка: значок, заголовок и описание стоят по центру высоты", TestCardContentCentered);
            Run("Кнопки: ни одна подпись не обрезана многоточием", TestButtonCaptionsFit);
            Run("Хаб: разделы, «Назад» и Esc показывают свой набор карточек", TestHubNavigation);
            Run("Хаб: придержанные файлы забываются при уходе из раздела", TestHubPendingFilesCleared);
            Run("Хаб: из карточек открывается каждый инструмент, и именно свой", TestHubOpensEveryTool);
            Run("«Разделение» передаёт открытый документ в «Прочие операции»", TestSplitHandsDocumentToOps);
            Run("«Объединение» передаёт собранный файл в «Прочие операции»", TestMergeHandsResultToOps);
            Run("Просмотр: окно сворачивается и имеет кнопку на панели задач", TestPreviewCanMinimize);
            Run("Окна: «Главная» последняя в обходе Tab (фокус не на ней)", TestHomeHeaderLastInTabOrder);
            Run("HeaderBand.TextRows: заголовок и подпись помещаются на любом масштабе экрана", TestHeaderRowsFitAnyDpi);
            Run("MathUtil.Median: нижняя медиана, пустой вход, вход не переставлен", TestMedian);
            Run("ListMarker: не-ASCII цифры («１.», «١.») номером списка не считаются", TestListMarkerNonAsciiDigits);
            Run("Loc.T: отсутствующий ключ возвращается как есть, null не роняет", TestLocMissingKey);
            Run("PdfCompression.LevelLabels: подпись на уровень, переводы, а не ключи (оба языка)", TestCompressionLevelLabels);
            Run("Cancellation.ThrowIf: null молчит, поднятый флаг бросает", TestCancellationThrowIf);
            Run("Cancellation.NoPartialOutput: и отмена, и сбой удаляют созданное", TestNoPartialOutput);
            Run("WindowPlacement.BestWorkArea: экран по наибольшему пересечению, фолбэк без него", TestBestWorkArea);
            Run("XyCut.OrderTree: этажи и колонки деревом, AvailRight по соседней колонке", TestXyCutOrderTreeShape);
            Run("Ui.OnUi: null/без хэндла/разрушенное окно — false без исключения", TestOnUiGuard);
            Run("UserSettings: язык не затирается записью устаревшего экземпляра", TestSettingsLanguageNotClobbered);
            Run("Установщик: язык приезжает маркером и применяется один раз", TestSetupLanguageMarker);

            // ---------- 1.17.8 ----------
            Run("Чередование: порядок разбирается на документы по непрерывным отрезкам", TestInterleaveRuns);
            Run("Чередование: пачки одностороннего сканера (обратная сторона с конца)", TestInterleaveScannerCase);
            Run("Чередование: шаг, хвост длинной пачки, перестановка тех же ссылок", TestInterleavePaceAndTails);
            Run("Шаблон имени: подстановка токенов", TestNameTemplateBasics);
            Run("Шаблон имени: дополнение нулями и смещение нумерации", TestNameTemplatePadAndOffset);
            Run("Шаблон имени: неизвестные токены текстом, путь обезврежен", TestNameTemplateUnknownAndUnsafe);
            Run("Шаблон имени: различает ли шаблон файлы между собой", TestNameTemplateUniqueness);
            Run("Текст: таблицы не теряются, ячейки через табуляцию", TestPlainTextKeepsTables);
            Run("Текст: порядок чтения сверху вниз и слева направо", TestPlainTextReadingOrder);
            Run("Текст: накрытые ячейки не сдвигают колонки", TestPlainTextMergedCells);
            Run("Текст: страницы через перевод страницы, пустые на месте", TestPlainTextDocument);
            Run("Экспорт: ширина в пикселях по разрешению, расширения файлов", TestExportPixelWidth);
            Run("Ghostscript: аргументы режимов серого и починки", TestConvertArguments);
            Run("Ghostscript: политика замены мягче, чем у сжатия", TestConvertShouldReplace);
            Run("Ghostscript: нулевой код возврата с «****» в потоке — это отказ", TestEngineSucceeded);
            Run("Ghostscript: документ обязан остаться документом (страницы целы)", TestPagesKept);
            Run("Проба цвета: выборка страниц по всей длине документа", TestColorSamplePages);
            Run("Проба цвета: насыщенность и доля, а не одинокий пиксель", TestColorHasColor);
            Run("Рендер: потолок высоты растра ужимает ширину, а не память", TestRenderFitWidth);
            Run("Проба цвета (живая): предельно длинная страница не рвёт память и время", TestColorProbeExtremePage);
            Run("Ghostscript (стресс): подмена файла не отказывает на повторах", TestGsReplaceStress);
            Run("Ghostscript (живой): страницы считает PdfPig — PdfSharp не читает 1.5", TestPageProbeLive);
            Run("Сжатие (живое): «Очень хорошо» уменьшает файл, не трогая пиксели картинок", TestCompressKeepsResolutionLive);
            Run("Ghostscript (живой): серое на JPEG, CMYK и прозрачности", TestGrayscaleHardLive);
            Run("Ghostscript (живой): серое действительно серое, битый файл чинится", TestConvertLive);
            Run("Просмотр: ступени масштаба и проценты", TestPreviewZoomSteps);
            Run("Просмотр: вписывание по окну без растягивания мелкой страницы", TestPreviewZoomFit);
            Run("Просмотр: Ctrl+колесо увеличивает к точке под курсором", TestPreviewZoomAnchor);
            Run("Просмотр: панорама и порог перетаскивания", TestPreviewPan);
            Run("Просмотр: положение страницы считается в координатах содержимого", TestPreviewCentered);
            Run("Сетка: выбор чётных и нечётных страниц (нумерация с единицы)", TestSelectEveryOther);
            Run("Добивка: позиции пустых страниц для двусторонней печати", TestBlankPagePositions);
            Run("Добивка (живая): пустые страницы в файле и нужного размера", TestBlankPageMergeLive);
            Run("Шаблон имени частей: пустой сохраняет прежние имена", TestPartNameOptional);
            Run("Шаблон имени частей (живой): имена файлов по шаблону", TestSplitTemplateLive);
            Run("Защита: запись поверх исходника распознаётся заранее", TestSameFileGuard);
            Run("Печать: страница вписывается в лист целиком и по центру", TestPrintFitToPage);
            Run("Сжатие: подписи называют разрешение из того же источника", TestCompressionLabelsNameDpi);
            Run("Смена языка: главный экран остаётся рабочим", TestLanguageRebuildKeepsHubUsable);
            Run("Смена языка: свёрнутый главный экран не уезжает за экран", TestLanguageRebuildKeepsMinimizedHubOnScreen);
            Run("Руководство: проверка вшитости умеет отвечать «нет»", TestUserManualPackedDetectsAbsence);
            Run("Руководство: распаковывается рядом с настройками, под своим именем", TestUserManualPath);
            Run("Обновления: решение показывать уведомление при запуске", TestShouldNotifyUpdate);
            Run("Обновления: версия из двух чисел не роняет показ", TestVersionDisplay);
            Run("История: строка переживает запись и чтение, испорченная пропускается", TestHistoryEntryRoundTrip);
            Run("История: кольцо хранит последние записи", TestHistoryTrim);
            Run("История: автоочистка убирает старое, не трогая свежее", TestHistoryKeepRecent);
            Run("История (живое): запись, счётчик, очистка и выключение", TestSettingsHistoryLive);
            Run("История (живое): окно списка — свежее сверху, подписи из каталога", TestHistoryWindowLive);
            Run("Обновления: окно показывается один раз, сторож отпускает", TestUpdateWindowShownOnce);
            Run("Обновления: настройки не затираются устаревшим Save", TestUpdatePrefsNotClobbered);
            Run("Обновления: подпись флажка помещается в диалог на обоих языках", TestUpdateSkipCaptionFits);
            Run("Настройки (живое): три действия истории — в ряд и без обрезки подписей", TestSettingsHistoryRowLive);
            Run("Хаб (живое): значки настроек и справки доступны с клавиатуры и диктору", TestHubIconButtonsLive);
            Run("Крупный системный шрифт: окна и меню живут своим шрифтом", TestSystemFontIgnoredLive);
            Run("Масштаб (живое): двойной клик по полю процентов возвращает 100%", TestZoomDoubleClickResetsLive);
            Run("Настройки (живое): флажок пишется, «снова напоминать» появляется по делу", TestSettingsUpdateControlsLive);

            Run("XmlText: экранируются & < > и кавычка", TestXmlEscape);
            Run("XmlText: очистка — управляющие, непарные суррогаты, перевод строки", TestXmlSanitize);
            Run("PptxGeometry: пункты → EMU, знак координаты и неотрицательный размер", TestPptxEmu);
            Run("PptxGeometry: кегль в сотых пункта и пределы разметки", TestPptxHundredths);
            Run("PptxGeometry: размер слайда — самый частый размер страницы", TestPptxSlideSize);
            Run("PptxGeometry: страница другого размера вписывается и центрируется", TestPptxFit);
            Run("PptxGeometry: ось Y переворачивается от высоты страницы", TestPptxFlipY);
            Run("OoxmlPackage: [Content_Types].xml первой записью, Override со слэшем", TestOoxmlContentTypes);
            Run("OoxmlPackage: имя части — ASCII, без ведущего слэша, без дублей", TestOoxmlPartNames);
            Run("PPTX: пакет самосогласован — части, связи, отсутствие BOM", TestPptxPackageConsistency);
            Run("PPTX: валидатор Open XML не находит ошибок в собранном файле", TestPptxValidator);
            Run("PPTX: валидатор — зависимость ТОЛЬКО тестов, в src его нет", TestPptxValidatorNotInApp);
            Run("PPTX: абзац — надпись с нулевыми полями, без автоподбора и заливки", TestPptxTextBox);
            Run("PPTX: формат рана — кегль, начертание, цвет, шрифт и порядок элементов", TestPptxRunProps);
            Run("PPTX: выравнивание и литеральный маркер списка", TestPptxAlignAndBullet);
            Run("PPTX: пустой абзац фигуры не создаёт", TestPptxEmptyParagraph);
            Run("PPTX: запас ширины надписи не выходит за край страницы", TestPptxWidthSlack);
            Run("PPTX: одна картинка на трёх слайдах — одна часть, три связи", TestPptxMediaDedup);
            Run("PPTX: гиперссылка — связь наружу, повтор адреса не плодит связей", TestPptxHyperlink);
            Run("PPTX: слайд с содержимым проходит валидатор", TestPptxValidatorWithContent);
            Run("PPTX: прочерк-заполнитель на слайд не переносится", TestPptxFillerRule);
            Run("Разбор для слайда: далеко разнесённые строки — разные блоки", TestSlideGapBreak);
            Run("PPTX: однострочная надпись не переносится, многострочная переносится", TestPptxSingleLine);
            Run("PPTX: многострочный абзац — ОДНА надпись со строками на своих местах", TestPptxParagraphKeptWhole);
            Run("PPTX: поправка на подъёмную часть строки", TestPptxAscentGap);
            Run("PPTX: подложка — фон слайда, а при вписывании — закреплённая картинка", TestPptxBackground);
            Run("PPTX: разрешение подложек падает с ростом числа страниц", TestPptxBackgroundDpi);
            Run("PPTX: объединение ячеек — пролёты у владельца, пометки у накрытых", TestPptxMergeGrid);
            Run("PPTX: таблица — сетка колонок, полные строки, границы по источнику", TestPptxTable);
            Run("PPTX: пробел между ранами не теряется", TestPptxKeepsSpaces);
            Run("PPTX: шаг строк — по базовым линиям, а не по верху чернил", TestPptxPitchFromBaselines);
            Run("Разбор: базовая линия слова — медиана по буквам, поворот её переносит", TestWordBaseline);
            Run("Палитра: значок каждого инструмента отличим от соседних", TestGlyphColoursDistinct);
            Run("Разметка PDF: строки одного блока не рвутся, разных — рвутся", TestStructureBlocksGroup);
            Run("PPTX: у каждой строки абзаца свой отступ от края рамки", TestPptxPerLineIndent);
            Run("Нехватка места названа диском и остатком", TestDiskFullMessage);
            Run("Страницы без текста названы в результате", TestConvertDoneStatus);
            Run("Разделение (живое): отмена на сжатии не оставляет частей", TestSplitCancelDuringCompressionLive);
            Run("Окна: работа с диском не на потоке интерфейса", TestNoDiskWorkOnUiThread);
            Run("Окна: сбой при опросе пароля не оставит окно заблокированным", TestPasswordPromptCannotHangWindow);
            Run("Недавние: сколько карточек влезает в полосу", TestRecentFit);
            Run("Стартовый экран (живой): недавние не наезжают на карточки инструментов", TestStartScreenWithRecentSurvivesMinimum);
            Run("Недавние файлы: значок по инструменту, затем по расширению", TestRecentCardGlyph);
            Run("Недавние файлы: свежие, без повторов и без исчезнувших", TestRecentFiles);
            Run("Пустая страница: формат берётся у соседа", TestBlankSheetSize);
            Run("Пустая страница (живая): обёртка нужного размера собирается в документ", TestBlankSheetLive);
            Run("Печать: предпросмотр рисует ограниченное число листов", TestPrintPreviewPageLimit);
            Run("Печать (живая): задание одно и то же для показа и для печати", TestPrintDocumentReusable);
            Run("Командная строка: разбор команд и отказов", TestPdfCliParse);
            Run("Командная строка: имена уровней сжатия", TestPdfCliLevels);
            Run("Командная строка (живая): команды делают то же, что кнопки", TestPdfCliLive);
            Run("Пароли: реестр помнит файл, а не написание пути", TestPasswordRegistry);
            Run("Пароли: «нужен пароль» отличается от «файл битый»", TestPasswordDetection);
            Run("Пароли: ширина окна по самой длинной строке подписи", TestPasswordDialogWidth);
            Run("Пароли (живое): защищённый файл открывается и работает", TestPasswordLive);
            Run("Состав документа: когда советовать уровень с пересчётом картинок", TestContentProfileAdvice);
            Run("Состав документа (живой): доля изображений у скана и у текста", TestContentProfileLive);
            Run("Оглавление: пересчёт под новый состав страниц", TestBookmarksRemap);
            Run("Оглавление: карта подряд идущего диапазона", TestBookmarksRangeMap);
            Run("Оглавление (живое): объединение и разделение его не теряют", TestBookmarksLive);
            Run("В исходниках нет управляющих символов", TestNoControlCharactersInSources);
            Run("В репозитории нет рабочих упоминаний и путей машины", TestNoWorkReferencesInRepo);
            Run("Переводы помечены как бета и объясняют, чего им не хватает", TestBetaNotice);
            Run("Руководство выбирается по языку интерфейса", TestManualBothLanguages);
            Run("Примечание «бета» переносится по словам, а не обрезается", TestBetaNoteWraps);
            Run("Полоска «бета» (живое): на минимальной ширине текст цел", TestBetaNoteFitsLive);
            Run("Собранное = исходник: чем правка отличается от файла", TestAssembledPristine);
            Run("Постранично можно читать прямо из исходника (подмножество по возрастанию)", TestPlainSubset);
            Run("Дроп: набор расширений (PDF, а у «Прочих операций» и картинки)", TestDropExtensions);
            Run("Картинка → лист A4: ориентация, поля, пропорции", TestImageLayoutOnPage);
            Run("Картинка → EXIF-поворот применяется, а не игнорируется", TestImageExifRotation);
            Run("Картинки (живое): JPEG переносится как есть, PNG и TIFF собираются", TestImageToPdfLive);
            Run("Прочие операции (живое): сетка правится, кнопки по составу набора", TestOpsGridEditableLive);

            Console.WriteLine();
            Console.WriteLine("Пройдено: " + _passed + ", провалено: " + _failed);
            // Нижняя граница числа тестов: без неё удалённая строка Run(...) проходит незаметно —
            // прогон остаётся зелёным, просто проверок становится меньше. Растёт вместе с набором.
            const int MinTests = 359;
            int total = _passed + _failed;
            int code = _failed == 0 ? 0 : 1;
            if (total < MinTests)
            {
                Console.WriteLine("ОШИБКА: тестов " + total + ", а должно быть не меньше " + MinTests +
                    " — из прогона пропала проверка.");
                code = 1;
            }
            // Выходим ЖЁСТКО, как и само приложение (см. FastExit): живые тесты открывают
            // настоящие окна, а те трогают WinRT, и штатная выгрузка процесса роняет access
            // violation уже ПОСЛЕ того, как все проверки прошли. Прогон при этом выглядит
            // зелёным, а код возврата — ненулевым, и полная пирамида встаёт на шаге, который
            // на самом деле прошёл.
            FastExit.Now(code);
            return code; // недостижимо, но компилятору нужен возврат
        }

        // ---------- SheetNamer ----------

        private static void TestNamerForbiddenChars()
        {
            var n = new SheetNamer();
            AssertEqual("От_чет _март_ 1_", n.Next("От:чет [март] 1?"), "санитизация");
        }

        private static void TestNamerTruncation()
        {
            var n = new SheetNamer();
            string name = n.Next("Очень длинное имя файла отчета за первый квартал");
            AssertTrue(name.Length <= 31, "длина " + name.Length + " превышает 31");
            AssertEqual("Очень длинное имя файла отчета", name, "обрезка с зачисткой хвоста");
        }

        private static void TestNamerDedupe()
        {
            var n = new SheetNamer();
            AssertEqual("Отчет", n.Next("Отчет"), "первое имя");
            AssertEqual("Отчет_2", n.Next("Отчет"), "второе имя");
            // Дубликаты ищутся без учёта регистра, но регистр исходного имени сохраняется.
            AssertEqual("отчет_3", n.Next("отчет"), "регистронезависимый дубль");
        }

        private static void TestNamerDedupeLong()
        {
            var n = new SheetNamer();
            string baseName = new string('а', 31);
            string first = n.Next(baseName);
            string second = n.Next(baseName);
            AssertEqual(31, first.Length, "длина первого");
            AssertTrue(second.Length <= 31, "длина дубля " + second.Length);
            AssertTrue(second.EndsWith("_2"), "суффикс дубля: " + second);
        }

        private static void TestNamerReserve()
        {
            var n = new SheetNamer();
            n.Reserve("Содержание");
            AssertEqual("Содержание_2", n.Next("Содержание"), "резерв обходится суффиксом");
        }

        private static void TestNamerEmpty()
        {
            var n = new SheetNamer();
            AssertEqual("Лист", n.Next("   "), "пустое имя");
        }

        private static void TestNamerHistory()
        {
            var n = new SheetNamer();
            AssertEqual("History_", n.Next("History"), "зарезервированное имя");
        }

        private static void TestNamerApostrophes()
        {
            var n = new SheetNamer();
            AssertEqual("абв", n.Next("'абв'"), "апострофы по краям");
        }

        // ---------- NaturalStringComparer ----------

        private static void TestNaturalNumbers()
        {
            AssertTrue(NaturalStringComparer.Instance.Compare("Отчет 2", "Отчет 10") < 0, "2 < 10");
            AssertTrue(NaturalStringComparer.Instance.Compare("Отчет 10", "Отчет 2") > 0, "10 > 2");
        }

        private static void TestNaturalCase()
        {
            AssertEqual(0, NaturalStringComparer.Instance.Compare("отчет 10", "ОТЧЕТ 10"), "регистр");
        }

        private static void TestNaturalNulls()
        {
            AssertTrue(NaturalStringComparer.Instance.Compare(null, "a") < 0, "null < строки");
            AssertTrue(NaturalStringComparer.Instance.Compare("a", null) > 0, "строка > null");
            AssertEqual(0, NaturalStringComparer.Instance.Compare(null, null), "null == null");
        }

        private static void TestNaturalSortList()
        {
            var items = new List<string> { "Файл 10", "Файл 2", "Файл 1", "Файл 20" };
            items.Sort(NaturalStringComparer.Instance);
            AssertEqual("Файл 1|Файл 2|Файл 10|Файл 20", string.Join("|", items.ToArray()), "порядок");
        }

        // ---------- FindSourceFiles ----------

        private static void TestFindSourceFiles()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                File.WriteAllText(Path.Combine(dir, "Отчет 10.xlsx"), "");
                File.WriteAllText(Path.Combine(dir, "Отчет 2.XLSX"), "");   // регистр расширения
                File.WriteAllText(Path.Combine(dir, "старый.xls"), "");
                File.WriteAllText(Path.Combine(dir, "макрос.xlsm"), "");
                File.WriteAllText(Path.Combine(dir, "бинарный.xlsb"), "");
                File.WriteAllText(Path.Combine(dir, "~" + "$Отчет 10.xlsx"), ""); // временный Excel
                File.WriteAllText(Path.Combine(dir, "прочее.txt"), "");
                File.WriteAllText(Path.Combine(dir, "Свод.xlsx"), "");      // итоговый файл

                List<string> files = MergeService.FindSourceFiles(dir, Path.Combine(dir, "Свод.xlsx"));

                var names = new List<string>();
                foreach (string f in files)
                    names.Add(Path.GetFileName(f));
                AssertEqual("бинарный.xlsb|макрос.xlsm|Отчет 2.XLSX|Отчет 10.xlsx|старый.xls",
                    string.Join("|", names.ToArray()), "состав и порядок");
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        // ---------- CellText ----------

        private static void TestCellTextEscape()
        {
            AssertEqual("'=SUM(A1)", CellText.EscapeForEntry("=SUM(A1)"), "формула-инъекция");
            AssertEqual("'12 345", CellText.EscapeForEntry("12 345"), "число-подобная строка");
            AssertEqual("'01.02.2026", CellText.EscapeForEntry("01.02.2026"), "дата-подобная строка");
            AssertEqual("''уже с апострофом", CellText.EscapeForEntry("'уже с апострофом"), "апостроф сохраняется");
            AssertEqual(null, CellText.EscapeForEntry(null), "null");
            AssertEqual("", CellText.EscapeForEntry(""), "пустая строка");
        }

        private static void TestCellTextPassthrough()
        {
            AssertEqual(42.5, CellText.EscapeValues(42.5), "число");
            AssertEqual(true, CellText.EscapeValues(true), "булево");
            AssertEqual(null, CellText.EscapeValues(null), "null");
        }

        private static void TestCellTextArrays()
        {
            // Обычный 0-базный массив
            var plain = new object[1, 2];
            plain[0, 0] = "=x";
            plain[0, 1] = 7.0;
            var outPlain = (object[,])CellText.EscapeValues(plain);
            AssertEqual("'=x", outPlain[0, 0], "строка в 0-базном массиве");
            AssertEqual(7.0, outPlain[0, 1], "число в 0-базном массиве");

            // 1-базный массив — именно такой возвращает Range.Value2 через COM
            var comStyle = (object[,])Array.CreateInstance(
                typeof(object), new[] { 2, 1 }, new[] { 1, 1 });
            comStyle[1, 1] = "12 345";
            comStyle[2, 1] = 30.0;
            var outCom = (object[,])CellText.EscapeValues(comStyle);
            AssertEqual("'12 345", outCom[1, 1], "строка в 1-базном массиве");
            AssertEqual(30.0, outCom[2, 1], "число в 1-базном массиве");
        }

        // ---------- CLI ----------

        private static void TestCliOptions()
        {
            MergeOptions o;
            string err;
            AssertTrue(Program.TryParseCliOptions(new[] { "--cli", "in", "out" }, 3, out o, out err), "без флагов");
            AssertTrue(!o.AddToc && !o.ValuesOnly, "умолчания CLI — выключено");

            AssertTrue(Program.TryParseCliOptions(new[] { "--cli", "in", "out", "--toc", "--values" }, 3, out o, out err), "оба флага");
            AssertTrue(o.AddToc && o.ValuesOnly, "флаги применены");

            AssertTrue(Program.TryParseCliOptions(new[] { "--cli", "in", "out", "--VALUES" }, 3, out o, out err), "регистр флага");
            AssertTrue(o.ValuesOnly, "флаг в верхнем регистре");
        }

        private static void TestCliUnknownFlag()
        {
            MergeOptions o;
            string err;
            AssertTrue(!Program.TryParseCliOptions(new[] { "--cli", "in", "out", "--wtf" }, 3, out o, out err), "неизвестный флаг отвергнут");
            AssertTrue(err != null && err.Contains("--wtf"), "текст ошибки называет флаг");
        }

        // ---------- ReportWriter ----------

        private static void TestReportLine()
        {
            var ok = new FileResult();
            ok.FileName = "а.xlsx";
            ok.Ok = true;
            ok.SheetName = "а";
            AssertEqual("OK      а.xlsx -> [а]", ReportWriter.FormatFileLine(ok), "перенесённый");

            var skip = new FileResult();
            skip.FileName = "б.xlsx";
            skip.Note = "битый";
            AssertEqual("SKIPPED б.xlsx | битый", ReportWriter.FormatFileLine(skip), "пропущенный");
        }

        private static void TestReportBuild()
        {
            var result = new MergeResult();
            result.OutputPath = @"C:\out\Свод.xlsx";
            result.OkCount = 1;
            result.SkipCount = 1;
            var ok = new FileResult(); ok.FileName = "а.xlsx"; ok.Ok = true; ok.SheetName = "а";
            var skip = new FileResult(); skip.FileName = "б.xlsx"; skip.Note = "битый";
            result.Files.Add(ok);
            result.Files.Add(skip);
            var options = new MergeOptions(); options.AddToc = true;

            string report = ReportWriter.BuildReport(result, @"C:\in", options, new DateTime(2026, 7, 16, 14, 0, 0));
            AssertTrue(report.Contains("2026-07-16 14:00:00"), "дата");
            AssertTrue(report.Contains(@"C:\in"), "входная папка");
            AssertTrue(report.Contains(@"C:\out\Свод.xlsx"), "итоговый файл");
            AssertTrue(report.Contains("лист «Содержание»: да"), "параметры");
            AssertTrue(report.Contains("перенесено 1, пропущено 1"), "итог");
            AssertTrue(report.Contains("OK      а.xlsx"), "строка файла");
            AssertTrue(report.Contains("SKIPPED б.xlsx"), "строка пропуска");
        }

        private static void TestReportRotation()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N"));
            try
            {
                var stamp = new DateTime(2026, 7, 16, 10, 0, 0);
                for (int i = 0; i < 5; i++)
                    ReportWriter.SaveWithRotation(dir, "отчёт " + i, stamp.AddMinutes(i), 3);

                string[] files = Directory.GetFiles(dir, "report_*.txt");
                AssertEqual(3, files.Length, "число отчётов после ротации");
                Array.Sort(files, StringComparer.OrdinalIgnoreCase);
                AssertTrue(files[0].Contains("10-02-00"), "старейший из оставшихся");
                AssertTrue(files[2].Contains("10-04-00"), "новейший");
                AssertEqual("отчёт 4", File.ReadAllText(files[2]).TrimStart('﻿'), "содержимое новейшего");
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        private static void TestReportNameCollision()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N"));
            try
            {
                var stamp = new DateTime(2026, 7, 16, 10, 0, 0);
                string p1 = ReportWriter.SaveWithRotation(dir, "первый", stamp, 3);
                string p2 = ReportWriter.SaveWithRotation(dir, "второй", stamp, 3);
                AssertTrue(!string.Equals(p1, p2, StringComparison.OrdinalIgnoreCase), "имена различаются");
                AssertEqual(2, Directory.GetFiles(dir, "report_*.txt").Length, "оба сохранены");
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        // ---------- ToolRegistry ----------

        private static void TestToolRegistry()
        {
            var reg = new ToolRegistry();
            System.Windows.Forms.Form f;
            AssertTrue(!reg.TryGetOpen("a", out f), "пустой реестр — не открыт");

            var form = new System.Windows.Forms.Form();
            reg.Add("a", form);
            AssertTrue(reg.TryGetOpen("a", out f) && ReferenceEquals(f, form), "добавленный — открыт");
            AssertEqual(1, reg.OpenForms().Count, "одно живое окно");

            form.Dispose();
            AssertTrue(!reg.TryGetOpen("a", out f), "после Dispose — не открыт");
            AssertEqual(0, reg.OpenForms().Count, "закрытые не считаются");

            var f2 = new System.Windows.Forms.Form();
            reg.Add("b", f2);
            reg.Remove("b");
            AssertTrue(!reg.TryGetOpen("b", out f), "после Remove — не открыт");
            f2.Dispose();
        }

        // ---------- HelpMenu (общее меню «Справка») ----------

        private static void TestHelpMenu()
        {
            var extra = new System.Windows.Forms.ToolStripMenuItem("Папка отчётов");
            using (System.Windows.Forms.MenuStrip menu = HelpMenu.Create(null, delegate { }, extra))
            {
                AssertEqual(1, menu.Items.Count, "один пункт строки меню");
                var help = (System.Windows.Forms.ToolStripMenuItem)menu.Items[0];
                AssertEqual(Loc.T("menu.root"), help.Text, "название пункта");

                var texts = new List<string>();
                foreach (System.Windows.Forms.ToolStripItem it in help.DropDownItems)
                    if (it is System.Windows.Forms.ToolStripMenuItem)
                        texts.Add(it.Text);
                AssertTrue(texts.Contains(Loc.T("menu.howTo")), "есть «Как пользоваться»");
                // Статистика переехала ВНУТРЬ «Настроек»: отдельного пункта в меню быть не
                // должно, иначе одно окно открывалось бы двумя путями и оба пришлось бы
                // помнить при каждой правке.
                AssertTrue(texts.Contains(Loc.T("settings.title")), "есть «Настройки»");
                AssertTrue(!texts.Contains(Loc.T("menu.stats")), "отдельного пункта «Статистика» больше нет");
                AssertTrue(texts.Contains(Loc.T("menu.language")), "есть выбор языка");
                AssertTrue(texts.Contains("Папка отчётов"), "доп. пункт вставлен");
                // «О программе» перенесена на стартовый экран — в меню её быть не должно.
                AssertTrue(!texts.Contains(Loc.T("hub.about")), "«О программе» убрана из меню");
            }

            // Без доп. пунктов: «Как пользоваться», «Статистика», «Язык / Language».
            using (System.Windows.Forms.MenuStrip menu = HelpMenu.Create(null, delegate { }))
            {
                var help = (System.Windows.Forms.ToolStripMenuItem)menu.Items[0];
                int menuItems = 0;
                foreach (System.Windows.Forms.ToolStripItem it in help.DropDownItems)
                    if (it is System.Windows.Forms.ToolStripMenuItem)
                        menuItems++;
                AssertEqual(3, menuItems, "без extras — три пункта (справка, статистика, язык)");
            }
        }

        // ---------- ThumbZoom ----------

        private static void TestThumbZoomClamp()
        {
            AssertEqual(ThumbZoom.MinWidth, ThumbZoom.Clamp(ThumbZoom.MinWidth - 50), "ниже минимума");
            AssertEqual(ThumbZoom.MaxWidth, ThumbZoom.Clamp(ThumbZoom.MaxWidth + 50), "выше максимума");
            AssertEqual(150, ThumbZoom.Clamp(150), "в диапазоне");
        }

        private static void TestThumbZoomWheel()
        {
            int up = ThumbZoom.StepFromWheel(132, 120);   // один щелчок вверх
            int down = ThumbZoom.StepFromWheel(132, -120); // один щелчок вниз
            AssertTrue(up > 132, "колесо вверх увеличивает: " + up);
            AssertTrue(down < 132, "колесо вниз уменьшает: " + down);
            AssertEqual(ThumbZoom.MaxWidth, ThumbZoom.StepFromWheel(ThumbZoom.MaxWidth, 1200), "не выше максимума");
            AssertEqual(ThumbZoom.MinWidth, ThumbZoom.StepFromWheel(ThumbZoom.MinWidth, -1200), "не ниже минимума");
        }

        private static void TestThumbZoomTile()
        {
            System.Drawing.Size s = ThumbZoom.TileSize(160);
            AssertEqual(160, s.Width, "ширина");
            AssertTrue(s.Height > s.Width, "высота больше ширины (портрет): " + s.Height);

            // Сетка owner-draw: лимита ImageList 256×256 больше нет, но ширина клампится
            // диапазоном масштаба [MinWidth..MaxWidth].
            System.Drawing.Size over = ThumbZoom.TileSize(10000);
            AssertEqual(ThumbZoom.MaxWidth, over.Width, "кламп к максимуму масштаба");
            System.Drawing.Size under = ThumbZoom.TileSize(1);
            AssertEqual(ThumbZoom.MinWidth, under.Width, "кламп к минимуму масштаба");
        }

        // ---------- ListReorder ----------

        private static void TestListReorderMoves()
        {
            var l = new List<string> { "a", "b", "c" };
            AssertEqual(0, ListReorder.MoveUp(l, 0), "верхний вверх — на месте");
            AssertEqual(2, ListReorder.MoveDown(l, 2), "нижний вниз — на месте");
            AssertEqual(0, ListReorder.MoveUp(l, 1), "b вверх -> индекс 0");
            AssertEqual("b|a|c", string.Join("|", l.ToArray()), "после MoveUp");
            AssertEqual(1, ListReorder.MoveDown(l, 0), "b вниз -> индекс 1");
            AssertEqual("a|b|c", string.Join("|", l.ToArray()), "MoveDown вернул порядок");
        }

        private static void TestListReorderMoveRemove()
        {
            var l = new List<string> { "a", "b", "c", "d" };
            ListReorder.Move(l, 0, 3);
            AssertEqual("b|c|a|d", string.Join("|", l.ToArray()), "перенос a перед позицией 3");
            ListReorder.Move(l, 2, 0);
            AssertEqual("a|b|c|d", string.Join("|", l.ToArray()), "перенос обратно");
            ListReorder.RemoveAt(l, new[] { 3, 1 });
            AssertEqual("a|c", string.Join("|", l.ToArray()), "удаление набора индексов");
        }

        // ---------- SourceFileList ----------

        private static string IncludedSig(SourceFileList list)
        {
            return string.Join("|", list.IncludedInOrder().ConvertAll(System.IO.Path.GetFileName).ToArray());
        }

        private static void TestSourceFileList()
        {
            var list = new SourceFileList();
            list.SetFiles(new[] { @"C:\in\А.xlsx", @"C:\in\Б.xlsx", @"C:\in\В.xlsx" });
            AssertEqual(3, list.Count, "три файла");
            AssertEqual(3, list.IncludedCount, "все включены");

            list[1].Include = false; // исключаем Б
            AssertEqual(2, list.IncludedCount, "два включённых");
            AssertEqual("А.xlsx|В.xlsx", IncludedSig(list), "исключённый не в списке");

            // Перестановка по позициям всего списка [А, Б, В] -> [В, А, Б]
            list.MoveUp(2); // [А, В, Б]
            list.MoveUp(1); // [В, А, Б]
            AssertEqual("В.xlsx|А.xlsx", IncludedSig(list), "порядок среди включённых изменился");

            list.SetAllIncluded(true);
            AssertEqual(3, list.IncludedCount, "все снова включены");
        }

        private static void TestSourceFileSort()
        {
            var list = new SourceFileList();
            list.SetFiles(new[] { @"C:\in\Отчет 10.xlsx", @"C:\in\Отчет 2.xlsx", @"C:\in\Отчет 1.xlsx" });
            list.SortByName();
            AssertEqual("Отчет 1.xlsx|Отчет 2.xlsx|Отчет 10.xlsx", IncludedSig(list), "естественный порядок");
        }

        // ---------- PrepareSourceList ----------

        private static void TestPrepareSourceList()
        {
            var files = new List<string>
            {
                @"C:\in\А.xlsx", @"C:\in\Свод.xlsx", @"C:\in\А.xlsx", @"C:\in\Б.xlsx"
            };
            List<string> prepared = MergeService.PrepareSourceList(files, @"C:\in\Свод.xlsx");
            var names = prepared.ConvertAll(System.IO.Path.GetFileName);
            AssertEqual("А.xlsx|Б.xlsx", string.Join("|", names.ToArray()),
                "свод исключён, дубликат убран, порядок сохранён");
            AssertEqual(0, MergeService.PrepareSourceList(null, "x").Count, "null -> пусто");
        }

        // ---------- FileSignature ----------

        private static void TestFileSignature()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string txt = Path.Combine(dir, "битый.xlsx");
                File.WriteAllText(txt, "this is not an excel file");
                AssertEqual(ExcelContainer.NotExcel, FileSignature.Detect(txt), "текст с расширением .xlsx");

                string empty = Path.Combine(dir, "пустой.xlsx");
                File.WriteAllBytes(empty, new byte[0]);
                AssertEqual(ExcelContainer.NotExcel, FileSignature.Detect(empty), "пустой файл");

                string zip = Path.Combine(dir, "ooxml.xlsx");
                File.WriteAllBytes(zip, new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00 });
                AssertEqual(ExcelContainer.Zip, FileSignature.Detect(zip), "ZIP-сигнатура PK — OOXML");

                string ole = Path.Combine(dir, "запарол.xlsx");
                File.WriteAllBytes(ole, new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 });
                AssertEqual(ExcelContainer.Ole2, FileSignature.Detect(ole), "OLE2/CFB — xls или шифр");

                AssertEqual(ExcelContainer.Unreadable, FileSignature.Detect(Path.Combine(dir, "нет.xlsx")),
                    "отсутствующий файл — решает Excel");
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        private static void TestLowSpaceMessage()
        {
            AssertEqual(null, MergeService.LowSpaceMessage(@"C:\", 5L * 1024 * 1024 * 1024), "5 ГБ — достаточно");
            AssertEqual(null, MergeService.LowSpaceMessage(@"C:\", 200L * 1024 * 1024), "ровно порог 200 МБ — достаточно");
            string low = MergeService.LowSpaceMessage(@"C:\", 10L * 1024 * 1024);
            AssertTrue(low != null && low.Contains(@"C:\") && low.Contains("10 МБ"), "мало места: " + low);
            string zero = MergeService.LowSpaceMessage(@"C:\", 0);
            AssertTrue(zero != null && zero.Contains("0 МБ"), "ноль байт: " + zero);
        }

        // ---------- WindowChrome / HeaderBand ----------

        private static void TestHeaderTextBound()
        {
            // Нет дочерних контролов — граница у правого края с отступом.
            AssertEqual(760, HeaderBand.TextRightBound(780, int.MaxValue), "без кнопки — правый край - 20");
            // Кнопка слева от края — текст обрезается до её левой границы минус зазор.
            AssertEqual(588, HeaderBand.TextRightBound(780, 600), "с кнопкой — левее её на 12");
            AssertEqual(488, HeaderBand.TextRightBound(700, 500), "узкое окно — левее кнопки");
            // Кнопка правее правого поля — не выходим за край панели.
            AssertEqual(760, HeaderBand.TextRightBound(780, 900), "кнопка за краем — ограничены полем");
        }

        private static void TestClassifyListKey()
        {
            var Alt = System.Windows.Forms.Keys.Alt;
            var Ctrl = System.Windows.Forms.Keys.Control;
            AssertEqual(MainForm.ListKeyAction.MoveUp, MainForm.ClassifyListKey(Alt | System.Windows.Forms.Keys.Up), "Alt+Up");
            AssertEqual(MainForm.ListKeyAction.MoveDown, MainForm.ClassifyListKey(Alt | System.Windows.Forms.Keys.Down), "Alt+Down");
            AssertEqual(MainForm.ListKeyAction.Copy, MainForm.ClassifyListKey(Ctrl | System.Windows.Forms.Keys.C), "Ctrl+C — копировать");
            AssertEqual(MainForm.ListKeyAction.SelectAll, MainForm.ClassifyListKey(Ctrl | System.Windows.Forms.Keys.A), "Ctrl+A — выделить всё");
            AssertEqual(MainForm.ListKeyAction.Exclude, MainForm.ClassifyListKey(System.Windows.Forms.Keys.Delete), "Delete — исключить");
            AssertEqual(MainForm.ListKeyAction.Swallow, MainForm.ClassifyListKey(System.Windows.Forms.Keys.Enter), "Enter — не сливать");
            AssertEqual(MainForm.ListKeyAction.None, MainForm.ClassifyListKey(System.Windows.Forms.Keys.Up), "просто ↑ — навигация");
        }

        private static void TestIsResetZoomKey()
        {
            var Ctrl = System.Windows.Forms.Keys.Control;
            AssertTrue(PdfToolFormBase.IsResetZoomKey(Ctrl | System.Windows.Forms.Keys.D0), "Ctrl+0 (основной ряд)");
            AssertTrue(PdfToolFormBase.IsResetZoomKey(Ctrl | System.Windows.Forms.Keys.NumPad0), "Ctrl+0 (цифровой блок)");
            AssertTrue(!PdfToolFormBase.IsResetZoomKey(System.Windows.Forms.Keys.D0), "0 без Ctrl — нет");
            AssertTrue(!PdfToolFormBase.IsResetZoomKey(Ctrl | System.Windows.Forms.Keys.D1), "Ctrl+1 — нет");
        }

        private static void TestClassifyPageKey()
        {
            var Alt = System.Windows.Forms.Keys.Alt;
            var Ctrl = System.Windows.Forms.Keys.Control;
            AssertEqual(PdfToolFormBase.PageKeyAction.Remove, PdfToolFormBase.ClassifyPageKey(System.Windows.Forms.Keys.Delete), "Delete — удалить");
            AssertEqual(PdfToolFormBase.PageKeyAction.MoveEarlier, PdfToolFormBase.ClassifyPageKey(Alt | System.Windows.Forms.Keys.Left), "Alt+← — раньше");
            AssertEqual(PdfToolFormBase.PageKeyAction.MoveLater, PdfToolFormBase.ClassifyPageKey(Alt | System.Windows.Forms.Keys.Right), "Alt+→ — позже");
            AssertEqual(PdfToolFormBase.PageKeyAction.SelectAll, PdfToolFormBase.ClassifyPageKey(Ctrl | System.Windows.Forms.Keys.A), "Ctrl+A — выделить всё");
            AssertEqual(PdfToolFormBase.PageKeyAction.Swallow, PdfToolFormBase.ClassifyPageKey(System.Windows.Forms.Keys.Enter), "Enter — не сохранять");
            AssertEqual(PdfToolFormBase.PageKeyAction.None, PdfToolFormBase.ClassifyPageKey(System.Windows.Forms.Keys.Left), "просто ← — навигация");
        }

        private static string RangeSig(System.Collections.Generic.List<PageRange> ranges)
        {
            var parts = new System.Collections.Generic.List<string>();
            foreach (PageRange r in ranges)
                parts.Add(r.Start + "-" + r.End);
            return string.Join("|", parts.ToArray());
        }

        private static void TestPageRangesParse()
        {
            AssertEqual("0-2|4-4|7-9", RangeSig(PageRanges.Parse("1-3, 5, 8-", 10)), "1-3,5,8- при 10");
            AssertEqual("1-1", RangeSig(PageRanges.Parse("2", 5)), "одна страница");
            AssertEqual("0-2", RangeSig(PageRanges.Parse("-3", 10)), "открытое начало");
            AssertEqual("0-1|3-3", RangeSig(PageRanges.Parse("  1 - 2 , 4 ", 5)), "пробелы");
            AssertEqual("5-8", PageRanges.Parse("5-", 8)[0].Label, "открытый конец 5.. -> 5-8");
            AssertEqual("8", PageRanges.Parse("8-", 8)[0].Label, "последняя страница -> метка 8");
            AssertEqual("1-3", PageRanges.Parse("1-3", 3)[0].Label, "метка диапазона");
        }

        private static void TestPageRangesParseErrors()
        {
            AssertThrows("пусто", delegate { PageRanges.Parse("", 10); });
            AssertThrows("0-3 (ниже 1)", delegate { PageRanges.Parse("0-3", 10); });
            AssertThrows("5-2 (начало>конец)", delegate { PageRanges.Parse("5-2", 10); });
            AssertThrows("3-99 (выше pageCount)", delegate { PageRanges.Parse("3-99", 10); });
            AssertThrows("нечисло", delegate { PageRanges.Parse("abc", 10); });
        }

        private static void TestPageRangesEveryN()
        {
            AssertEqual("0-2|3-5|6-8|9-9", RangeSig(PageRanges.EveryN(10, 3)), "10 по 3");
            AssertEqual("0-1|2-3|4-5", RangeSig(PageRanges.EveryN(6, 2)), "6 по 2");
            AssertEqual("0-0|1-1|2-2", RangeSig(PageRanges.EveryN(3, 1)), "по одной");
            AssertEqual("0-4", RangeSig(PageRanges.EveryN(5, 10)), "n больше всего");
            AssertThrows("n<1", delegate { PageRanges.EveryN(5, 0); });
        }

        private static void TestPageRangesToIndices()
        {
            var idx = PageRanges.ToIndices(PageRanges.Parse("1-3, 5, 8-", 10)); // 0-2,4,7-9
            AssertEqual("0,1,2,4,7,8,9", string.Join(",", idx.ConvertAll(i => i.ToString()).ToArray()), "порядок");
            // Повторы и обратный смысл ввода сохраняются: «1-4, 1» -> ...,0
            var dup = PageRanges.ToIndices(PageRanges.Parse("1-4, 1", 10));
            AssertEqual("0,1,2,3,0", string.Join(",", dup.ConvertAll(i => i.ToString()).ToArray()), "повторы сохраняются");
        }

        private static void TestUpdateChecker()
        {
            AssertEqual(new Version(1, 11, 2), UpdateChecker.ParseTag("v1.11.2"), "тег с v");
            AssertEqual(new Version(1, 12, 0), UpdateChecker.ParseTag("1.12.0"), "тег без v");
            AssertEqual(null, UpdateChecker.ParseTag("release"), "мусор -> null");
            AssertTrue(UpdateChecker.IsNewer(new Version(1, 12, 0), new Version(1, 11, 2)), "1.12.0 новее 1.11.2");
            AssertTrue(!UpdateChecker.IsNewer(new Version(1, 11, 2), new Version(1, 11, 2)), "равные — не новее");
            AssertTrue(!UpdateChecker.IsNewer(new Version(1, 11, 0), new Version(1, 11, 2)), "старее — не новее");
            AssertTrue(!UpdateChecker.IsNewer(null, new Version(1, 0, 0)), "null latest — не новее");
        }

        /// <summary>
        /// Решение «показывать ли уведомление при запуске». Ветка «пропущенная версия»
        /// сравнивается по СТАРШИНСТВУ, а не по равенству: один флажок не должен отключать
        /// уведомления навсегда, иначе о следующих выпусках человек не узнает.
        /// </summary>
        private static void TestShouldNotifyUpdate()
        {
            var current = new Version(1, 17, 9);
            var next = new Version(1, 18, 0);
            var later = new Version(1, 18, 1);

            AssertTrue(UpdateChecker.ShouldNotify(next, current, null), "новее и ничего не пропущено");
            AssertTrue(!UpdateChecker.ShouldNotify(current, current, null), "та же версия — молчим");
            AssertTrue(!UpdateChecker.ShouldNotify(new Version(1, 17, 8), current, null), "старее — молчим");
            AssertTrue(!UpdateChecker.ShouldNotify(null, current, null), "не разобрали ответ — молчим");

            AssertTrue(!UpdateChecker.ShouldNotify(next, current, "1.18.0"), "про эту версию просили не напоминать");
            AssertTrue(UpdateChecker.ShouldNotify(later, current, "1.18.0"), "СЛЕДУЮЩАЯ версия — напоминаем снова");
            AssertTrue(!UpdateChecker.ShouldNotify(next, current, "1.19.0"), "пропущенная выше найденной — молчим");
            AssertTrue(UpdateChecker.ShouldNotify(next, current, "мусор"), "непонятная пропущенная не глушит проверку");
            AssertTrue(UpdateChecker.ShouldNotify(next, current, ""), "пустая пропущенная не глушит проверку");
            AssertTrue(!UpdateChecker.ShouldNotify(next, current, "v1.18.0"), "пропущенная с «v» разбирается так же");
        }

        /// <summary>
        /// История: строка файла собирается и разбирается обратно без потерь, а испорченная
        /// строка пропускается, а не роняет разбор. Разделитель — табуляция, потому что в
        /// путях Windows встречаются и запятая, и точка с запятой; всё, что могло бы порвать
        /// строку, экранируется.
        /// </summary>
        private static void TestHistoryEntryRoundTrip()
        {
            var when = new DateTime(2026, 7, 28, 9, 15, 30, DateTimeKind.Utc);
            foreach (string path in new[]
            {
                @"C:\Папка\файл.pdf",
                @"C:\с запятой, и точкой с запятой; внутри\файл.pdf",
                "C:\\с\tтабуляцией.pdf",
                "C:\\с\nпереводом строки.pdf",
                "C:\\с\rвозвратом каретки.pdf",
                @"\\сервер\общая папка\файл.pdf",
                @"C:\обратная\\косая.pdf"
            })
            {
                var e = new HistoryEntry { WhenUtc = when, Operation = "hist.op.merge", Path = path };
                string line = OperationHistory.FormatEntry(e);
                AssertTrue(line.IndexOf('\n') < 0 && line.IndexOf('\r') < 0,
                    "запись не должна содержать переводов строки: " + path);
                HistoryEntry back = OperationHistory.ParseEntry(line);
                AssertTrue(back != null, "строка не разобралась: " + path);
                AssertEqual(path, back.Path, "путь пережил запись и чтение");
                AssertEqual("hist.op.merge", back.Operation, "ключ операции цел");
                AssertEqual(when, back.WhenUtc, "время цело");
            }

            // Испорченное — пропускается, а не роняет.
            AssertTrue(OperationHistory.ParseEntry(null) == null, "null");
            AssertTrue(OperationHistory.ParseEntry("") == null, "пустая строка");
            AssertTrue(OperationHistory.ParseEntry("enabled=True") == null, "не запись истории");
            AssertTrue(OperationHistory.ParseEntry("e=мусор	op	C:\f.pdf") == null, "время не число");
            AssertTrue(OperationHistory.ParseEntry("e=123	op") == null, "полей меньше трёх");
            AssertTrue(OperationHistory.ParseEntry("e=123	op	") == null, "пустой путь — не запись");
            AssertTrue(OperationHistory.ParseEntry("e=-99999999999999999999	op	C:\f.pdf") == null,
                "время вне диапазона DateTime не роняет разбор");
        }

        /// <summary>
        /// Кольцо истории: хранится не больше заданного числа записей, и остаются ПОСЛЕДНИЕ.
        /// Без кольца файл рос бы без предела, а с ним и перечень путей, о которых человек
        /// давно забыл, — то есть сведения, которых он не просил хранить.
        /// </summary>
        private static void TestHistoryTrim()
        {
            var few = new List<HistoryEntry>();
            for (int i = 0; i < 5; i++)
                few.Add(new HistoryEntry { Path = "f" + i });
            AssertEqual(5, OperationHistory.Trim(few).Count, "мало записей — не режем");

            var many = new List<HistoryEntry>();
            for (int i = 0; i < OperationHistory.MaxEntries + 25; i++)
                many.Add(new HistoryEntry { Path = "f" + i });
            List<HistoryEntry> cut = OperationHistory.Trim(many);
            AssertEqual(OperationHistory.MaxEntries, cut.Count, "обрезано до предела");
            AssertEqual("f25", cut[0].Path, "остались ПОСЛЕДНИЕ, старые вытеснены");
            AssertEqual("f" + (OperationHistory.MaxEntries + 24), cut[cut.Count - 1].Path, "самая новая на месте");

            AssertEqual(0, OperationHistory.Trim(null).Count, "null — пустой список, не исключение");
        }

        /// <summary>
        /// Автоочистка истории — СКОЛЬЗЯЩАЯ ДАВНОСТЬ, а не «раз в N дней стереть всё».
        /// Правило счётчиков сюда не переносится: счётчик копится от метки сброса, а список
        /// состоит из разновозрастных записей, и «период прошёл — чистим всё» стёрло бы
        /// сегодняшние операции из-за одной позавчерашней. Ровно эта ошибка и была допущена.
        /// </summary>
        private static void TestHistoryKeepRecent()
        {
            var now = new DateTime(2026, 7, 28, 12, 0, 0, DateTimeKind.Utc);
            var all = new List<HistoryEntry>
            {
                new HistoryEntry { WhenUtc = now.AddDays(-30), Path = "старая" },
                new HistoryEntry { WhenUtc = now.AddDays(-8), Path = "восьмидневная" },
                new HistoryEntry { WhenUtc = now.AddHours(-2), Path = "сегодняшняя" }
            };

            AssertEqual(3, OperationHistory.KeepRecent(all, now, 0).Count, "период 0 — не убираем ничего");

            List<HistoryEntry> week = OperationHistory.KeepRecent(all, now, 7);
            AssertEqual(1, week.Count, "при недельной давности остаётся только свежая");
            AssertEqual("сегодняшняя", week[0].Path, "СЕГОДНЯШНЯЯ ЗАПИСЬ НЕ СТИРАЕТСЯ из-за старой");

            AssertEqual(2, OperationHistory.KeepRecent(all, now, 30).Count, "при 30 днях уходит только тридцатидневная");
            AssertEqual(1, OperationHistory.KeepRecent(all, now, 1).Count, "при сутках остаётся только двухчасовая");
            AssertEqual(0, OperationHistory.KeepRecent(null, now, 7).Count, "null — пустой список, не исключение");

            // Ровно на границе периода запись уже старая: «хранить 7 дней» значит меньше семи.
            var edge = new List<HistoryEntry> { new HistoryEntry { WhenUtc = now.AddDays(-7), Path = "граница" } };
            AssertEqual(0, OperationHistory.KeepRecent(edge, now, 7).Count, "ровно семь дней — уже не хранится");
        }

        /// <summary>
        /// Версия для показа не должна ронять программу. Version.ToString(3) БРОСАЕТ на
        /// версии из двух чисел, а тег приходит с GitHub — то есть снаружи. На пути проверки
        /// при запуске это был бы отчёт о сбое сразу после открытия программы, у всех разом
        /// и без их участия: тег «v1.18» поставить никто не мешает.
        /// </summary>
        private static void TestVersionDisplay()
        {
            AssertEqual("1.18.0", UpdateChecker.Display(new Version(1, 18, 0)), "три числа — как есть");
            AssertEqual("1.18", UpdateChecker.Display(new Version(1, 18)), "два числа не роняют показ");
            AssertEqual("1.18.0", UpdateChecker.Display(new Version(1, 18, 0, 7)), "четвёртое число не показываем");
            AssertEqual("", UpdateChecker.Display(null), "null — пустая строка, не исключение");
            // Сквозной случай: тег с GitHub из двух чисел проходит весь путь без исключения.
            AssertEqual("1.18", UpdateChecker.Display(UpdateChecker.ParseTag("v1.18")), "тег «v1.18» показывается целиком");
        }

        /// <summary>
        /// Окно об обновлении в любой момент одно. Проверка при запуске и проверка по кнопке —
        /// два независимых воркера, и нажатие кнопки в первые десять секунд после запуска дало
        /// бы два одинаковых сообщения, одно поверх другого. Проверяем и обратное: сторож
        /// обязан отпускать, иначе первое же окно закрыло бы уведомления до перезапуска.
        /// </summary>
        private static void TestUpdateWindowShownOnce()
        {
            int shown = 0;
            UpdateUi.ShowOnce(delegate
            {
                shown++;
                UpdateUi.ShowOnce(delegate { shown++; }); // как второй воркер во время показа
            });
            AssertEqual(1, shown, "вложенный показ отсечён");

            UpdateUi.ShowOnce(delegate { shown++; });
            AssertEqual(2, shown, "после закрытия окна следующий показ проходит");

            // Сорвавшийся показ не должен запирать сторож навсегда.
            try { UpdateUi.ShowOnce(delegate { throw new InvalidOperationException("сбой показа"); }); }
            catch (InvalidOperationException) { }
            UpdateUi.ShowOnce(delegate { shown++; });
            AssertEqual(3, shown, "исключение внутри показа отпускает сторож");
        }

        private static void TestShouldAutoClear()
        {
            var now = new DateTime(2026, 7, 20, 12, 0, 0, DateTimeKind.Utc);
            AssertTrue(!UsageStats.ShouldAutoClear(now.AddDays(-3), now, 0), "период 0 — никогда");
            AssertTrue(!UsageStats.ShouldAutoClear(now.AddHours(-5), now, 1), "меньше суток — рано");
            AssertTrue(UsageStats.ShouldAutoClear(now.AddDays(-1), now, 1), "сутки прошли — пора");
            AssertTrue(!UsageStats.ShouldAutoClear(now.AddDays(-6), now, 7), "6 из 7 дней — рано");
            AssertTrue(UsageStats.ShouldAutoClear(now.AddDays(-8), now, 7), "8 дней при периоде 7 — пора");
            AssertTrue(UsageStats.ShouldAutoClear(now.AddDays(-31), now, 30), "31 день при 30 — пора");
        }

        private static void TestUsageTotal()
        {
            var s = new UsageStats
            {
                ExcelDigests = 1, PdfMerges = 2, PdfExtracts = 3, PdfSplitRanges = 4,
                PdfSplitEveryN = 5, PdfSplitBookmarks = 6, PdfToWord = 7, PdfCompressions = 99
            };
            // Total — сумма ОПЕРАЦИЙ (1+…+7); сжатие — параметр, в Total не входит.
            AssertEqual(28, s.Total, "Total включает PdfToWord и исключает PdfCompressions");
        }

        private static void TestMessageButtonX()
        {
            AssertEqual(164, MessageForm.ButtonX(0, 1, 440, 112, 20), "одна кнопка — по центру");
            AssertEqual(20, MessageForm.ButtonX(0, 2, 440, 112, 20), "две: первая слева");
            AssertEqual(308, MessageForm.ButtonX(1, 2, 440, 112, 20), "две: вторая справа");
        }

        private static void TestSanitize()
        {
            AssertEqual("Глава 1", PdfSplitService.Sanitize("Глава 1"), "обычный заголовок");
            AssertEqual("a_b_c", PdfSplitService.Sanitize("a/b:c"), "недопустимые символы");
            AssertEqual("без_имени", PdfSplitService.Sanitize(""), "пустое имя");
            AssertEqual("без_имени", PdfSplitService.Sanitize("   "), "только пробелы");
        }

        private static int PdfPageCount(string path)
        {
            using (PdfDocument d = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                return d.PageCount;
        }

        private static string PageCounts(List<string> files)
        {
            var parts = new List<string>();
            foreach (string f in files)
                parts.Add(PdfPageCount(f).ToString());
            return string.Join(",", parts.ToArray());
        }

        private static void TestPdfSplitLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = Path.Combine(dir, "исходник.pdf");
                using (var doc = new PdfDocument())
                {
                    for (int i = 0; i < 10; i++)
                        doc.AddPage();
                    doc.Outlines.Add("Раздел А", doc.Pages[0]);
                    doc.Outlines.Add("Раздел Б", doc.Pages[4]);
                    doc.Outlines.Add("Раздел В", doc.Pages[7]);
                    doc.Save(src);
                }

                // Извлечь выбранные [0,2,4] → один файл, 3 страницы
                string extract = Path.Combine(dir, "выбранные.pdf");
                PdfSplitService.Extract(src, new List<int> { 0, 2, 4 }, extract);
                AssertEqual(3, PdfPageCount(extract), "извлечено 3 страницы");

                // По диапазонам 1-3,5,8- → 3 файла (3,1,3)
                List<string> ranges = PdfSplitService.SplitByRanges(src, PageRanges.Parse("1-3,5,8-", 10), dir, "диап");
                AssertEqual(3, ranges.Count, "диапазонов — 3 файла");
                AssertEqual("3,1,3", PageCounts(ranges), "страниц по диапазонам");
                AssertEqual("диап_1-3.pdf", Path.GetFileName(ranges[0]), "имя первого диапазона");

                // Каждые 3 страницы → 4 файла (3,3,3,1)
                List<string> everyN = PdfSplitService.SplitEveryN(src, 3, dir, "часть");
                AssertEqual(4, everyN.Count, "каждые 3 — 4 файла");
                AssertEqual("3,3,3,1", PageCounts(everyN), "страниц по частям");

                // По закладкам → 3 файла (4,3,3), имена с заголовками
                List<string> byMark = PdfSplitService.SplitByBookmarks(src, dir, "закл");
                AssertEqual(3, byMark.Count, "закладок — 3 файла");
                AssertEqual("4,3,3", PageCounts(byMark), "страниц по закладкам");
                AssertTrue(Path.GetFileName(byMark[0]).Contains("Раздел А"), "имя по закладке: " + Path.GetFileName(byMark[0]));
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        private static void TestCompressLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = Path.Combine(dir, "крупный.pdf");
                MakeImagePdf(src, 2);
                long before = new FileInfo(src).Length;
                int pagesBefore = PdfPageCount(src);
                AssertTrue(before > 300 * 1024, "исходник с изображением должен быть крупным: " + before);

                if (Ghostscript.Available)
                {
                    // «Нормально» (/screen, 72 DPI) — понижает разрешение изображения → размер падает.
                    bool applied = PdfCompression.Compress(src, CompressionLevel.Small);
                    long after = new FileInfo(src).Length;
                    AssertTrue(applied, "сжатие применено (GS есть)");
                    AssertTrue(after < before, "размер уменьшился: " + before + " -> " + after);
                    AssertTrue(PdfCompression.LooksLikePdf(src), "результат — валидный PDF");
                    // Страницы сохранены (это НЕ растр всего документа, а downsampling — структура цела).
                    AssertEqual(pagesBefore, PdfPageCount(src), "число страниц сохранено");
                }
                else
                {
                    // Без Ghostscript сжатие — безопасный no-op: файл не тронут.
                    bool applied = PdfCompression.Compress(src, CompressionLevel.Small);
                    AssertTrue(!applied, "без Ghostscript — без изменений");
                    AssertEqual(before, new FileInfo(src).Length, "файл не тронут без GS");
                }
            }
            finally
            {
                Directory.Delete(dir, true);
            }
        }

        /// <summary>PDF с крупным «шумовым» изображением (плохо сжимается) и текстом на каждой странице.</summary>
        private static void MakeImagePdf(string path, int pages)
        {
            string jpg = Path.Combine(Path.GetDirectoryName(path), "noise.jpg");
            const int side = 1800; // ~245 DPI на A4 → /screen (72) заведомо понизит
            using (var bmp = new System.Drawing.Bitmap(side, side, System.Drawing.Imaging.PixelFormat.Format24bppRgb))
            {
                System.Drawing.Imaging.BitmapData bd = bmp.LockBits(
                    new System.Drawing.Rectangle(0, 0, side, side),
                    System.Drawing.Imaging.ImageLockMode.WriteOnly,
                    System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                int total = Math.Abs(bd.Stride) * side;
                var buf = new byte[total];
                new Random(12345).NextBytes(buf); // шум — JPEG почти не сжимает, исходник тяжёлый
                System.Runtime.InteropServices.Marshal.Copy(buf, 0, bd.Scan0, total);
                bmp.UnlockBits(bd);
                bmp.Save(jpg, JpegEncoder(), JpegQuality(90));
            }
            try
            {
                using (var doc = new PdfDocument())
                using (XImage img = XImage.FromFile(jpg))
                {
                    var font = new XFont("Arial", 14);
                    for (int i = 0; i < pages; i++)
                    {
                        PdfPage page = doc.AddPage();
                        page.Width = 595;
                        page.Height = 842;
                        using (XGraphics gfx = XGraphics.FromPdfPage(page))
                        {
                            gfx.DrawImage(img, 0, 0, page.Width, page.Height);
                            gfx.DrawString("Страница " + (i + 1) + " — текст должен сохраниться.",
                                font, XBrushes.White, new XPoint(30, 30));
                        }
                    }
                    doc.Save(path);
                }
            }
            finally
            {
                try { File.Delete(jpg); } catch { }
            }
        }

        private static System.Drawing.Imaging.ImageCodecInfo JpegEncoder()
        {
            foreach (System.Drawing.Imaging.ImageCodecInfo c in System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders())
                if (c.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid)
                    return c;
            return null;
        }

        private static System.Drawing.Imaging.EncoderParameters JpegQuality(long quality)
        {
            var p = new System.Drawing.Imaging.EncoderParameters(1);
            p.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                System.Drawing.Imaging.Encoder.Quality, quality);
            return p;
        }

        private static void TestDonationRequisites()
        {
            AssertEqual(20, AboutForm.DonationAccount.Length, "счёт — 20 цифр");
            foreach (char c in AboutForm.DonationAccount)
                AssertTrue(c >= '0' && c <= '9', "в счёте только цифры");
            AssertTrue(AboutForm.DonationBank.Length > 0, "банк не пуст");
        }

        /// <summary>
        /// Окно «О программе»: описание должно предупреждать, что в Word переводится только
        /// цифровой PDF (то же, что обещает карточка инструмента), быть выделяемым (значит
        /// read-only поле, а не подпись) и целиком помещаться в окно. Копирайт с лицензией
        /// стоят внизу слева, ниже всего прочего содержимого.
        /// </summary>
        private static void TestAboutDescription()
        {
            foreach (Lang lang in new[] { Lang.Ru, Lang.En })
            {
                string[] desc = Loc.Pair("about.desc");
                string text = desc[lang == Lang.En ? 1 : 0];
                string marker = lang == Lang.En ? "digital" : "цифров";
                AssertTrue(text.Contains(marker),
                    "описание не уточняет про цифровой PDF (" + lang + "): " + text);
                AssertTrue(text.Contains("scanned") || text.Contains("сканированны"),
                    "описание не предупреждает про сканы (" + lang + ")");
            }

            Lang saved = Loc.Current;
            string failure = null;
            var th = new System.Threading.Thread(delegate()
            {
                // Оба языка: тексты разной длины ложатся в разное число строк, и высота поля
                // считается замером, поэтому по-английски описание тоже может не влезть.
                foreach (Lang lang in new[] { Lang.Ru, Lang.En })
                {
                    if (failure != null)
                        return;
                    CheckAboutLayout(lang, ref failure);
                }
            });
            th.SetApartmentState(System.Threading.ApartmentState.STA); // окна WinForms требуют STA
            th.IsBackground = true;
            th.Start();
            th.Join();
            Loc.Init(saved);
            AssertTrue(failure == null, "AboutForm: " + failure);
        }

        /// <summary>Собрать окно «О программе» на данном языке и проверить его раскладку.</summary>
        private static void CheckAboutLayout(Lang lang, ref string failure)
        {
                try
                {
                    Loc.Init(lang);
                    using (var about = new AboutForm())
                    {
                        if (about.Handle == IntPtr.Zero) { failure = "окно не создалось"; return; }
                        var box = FindSelectable(about, Loc.T("about.desc"));
                        if (box == null) { failure = "описание не найдено выделяемым полем"; return; }
                        if (box.Bottom > about.ClientSize.Height) { failure = "описание не влезает в окно"; return; }

                        // Владелец просил описание ВЫКЛЮЧЕННЫМ ПО ШИРИНЕ, а такого выравнивания
                        // WinForms не предлагает — оно выставлено формату абзаца напрямую.
                        // Спрашиваем сам элемент, иначе откат к левому краю прошёл бы незаметно.
                        var rich = box as System.Windows.Forms.RichTextBox;
                        if (rich == null) { failure = "описание не RichTextBox — выключка по ширине невозможна"; return; }
                        if (!JustifiedText.IsJustified(rich)) { failure = "описание не выключено по ширине"; return; }

                        // Высота поля считается по замеру текста, поэтому проверяем НЕ формулу,
                        // а факт: спрашиваем у самого поля, где легла последняя строка после
                        // переноса. Если она вылезает за высоту — хвост описания обрезан.
                        System.Drawing.Point last = box.GetPositionFromCharIndex(box.Text.Length - 1);
                        int lineHeight = System.Windows.Forms.TextRenderer.MeasureText("Ag", box.Font).Height;
                        if (last.Y + lineHeight > box.Height)
                        {
                            failure = "последняя строка описания обрезана: низ " + (last.Y + lineHeight) +
                                " при высоте поля " + box.Height;
                            return;
                        }

                        System.Windows.Forms.Label license = null;
                        int lowestOther = 0;
                        foreach (System.Windows.Forms.Control c in about.Controls)
                        {
                            var lbl = c as System.Windows.Forms.Label;
                            if (lbl != null && lbl.Text == Loc.T("about.license")) { license = lbl; continue; }
                            if (c is RoundedButton) continue; // кнопка стоит на той же нижней линии
                            if (c.Bottom > lowestOther) lowestOther = c.Bottom;
                        }
                        if (license == null) { failure = "копирайт с лицензией не найден"; return; }
                        if (license.Left != 24) { failure = "копирайт не прижат влево: " + license.Left; return; }
                        if (license.Top < lowestOther) { failure = "копирайт не в самом низу окна"; return; }
                    }
                }
                catch (Exception ex) { failure = ex.GetType().Name + ": " + ex.Message; }
                if (failure != null)
                    failure = lang + ": " + failure;
        }

        /// <summary>
        /// Единственный экземпляр: первый запуск занимает имя, второй видит его занятым и
        /// потому уходит будить работающий. Хэндл второго закрывается сразу (держать нечего),
        /// а после освобождения первого имя снова свободно — иначе следующий запуск после
        /// нормального выхода считал бы, что приложение всё ещё работает.
        /// </summary>
        private static void TestSingleInstanceName()
        {
            string name = @"Local\iwoHelperDesktop.test." + Guid.NewGuid().ToString("N");
            System.Threading.Mutex first, second, third;

            AssertTrue(SingleInstance.TryAcquire(name, out first), "первый занимает свободное имя");
            AssertTrue(first != null, "первому достался мьютекс");

            AssertTrue(!SingleInstance.TryAcquire(name, out second), "второй видит имя занятым");
            AssertTrue(second == null, "второму мьютекс не отдаётся — хэндл закрыт сразу");

            first.Close(); // имя держалось только этим хэндлом
            AssertTrue(SingleInstance.TryAcquire(name, out third), "после выхода имя снова свободно");
            if (third != null)
                third.Close();
        }

        /// <summary>
        /// Выделяемое поле (read-only) с данным текстом среди контролов окна. Ищем по
        /// TextBoxBase, а не по TextBox: абзац описания — RichTextBox (ему нужна выключка по
        /// ширине), и это тоже поле, которое можно выделить и скопировать.
        /// </summary>
        private static System.Windows.Forms.TextBoxBase FindSelectable(System.Windows.Forms.Control root, string text)
        {
            foreach (System.Windows.Forms.Control c in root.Controls)
            {
                var tb = c as System.Windows.Forms.TextBoxBase;
                if (tb != null && tb.ReadOnly && tb.Text == text)
                    return tb;
                var nested = FindSelectable(c, text);
                if (nested != null)
                    return nested;
            }
            return null;
        }

        private static void TestClampWindow()
        {
            int lo, hi;
            PdfPageGrid.ClampWindow(0, 10, 100, 16, out lo, out hi);
            AssertEqual("0,26", lo + "," + hi, "у начала — не ниже 0");
            PdfPageGrid.ClampWindow(50, 60, 100, 16, out lo, out hi);
            AssertEqual("34,76", lo + "," + hi, "в середине — буфер с обеих сторон");
            PdfPageGrid.ClampWindow(90, 99, 100, 16, out lo, out hi);
            AssertEqual("74,99", lo + "," + hi, "у конца — не выше count-1");
            PdfPageGrid.ClampWindow(0, 5, 6, 16, out lo, out hi);
            AssertEqual("0,5", lo + "," + hi, "мало элементов — весь список");
        }

        private static void TestThemeToBgr()
        {
            AssertEqual(12413967, Theme.ToBgr(System.Drawing.Color.FromArgb(15, 108, 189)), "HubBlue #0F6CBD");
            AssertEqual(4291600, Theme.ToBgr(System.Drawing.Color.FromArgb(16, 124, 65)), "Accent #107C41");
            AssertEqual(16777215, Theme.ToBgr(System.Drawing.Color.White), "белый");
            AssertEqual(255, Theme.ToBgr(System.Drawing.Color.FromArgb(255, 0, 0)), "красный — младший байт");
        }

        private static void TestSheetRef()
        {
            AssertEqual("'Отчет'!A1", TocBuilder.SheetRef("Отчет"), "обычное имя");
            AssertEqual("'it''s'!A1", TocBuilder.SheetRef("it's"), "апостроф удваивается");
            AssertEqual("''!A1", TocBuilder.SheetRef(null), "null — пустое имя");
        }

        private static void TestWindowChromeColorRef()
        {
            // COLORREF кладёт R в младший байт, B в старший (0x00BBGGRR).
            int packed = WindowChrome.ColorRef(System.Drawing.Color.FromArgb(16, 124, 65));
            AssertEqual(0x00417C10, packed, "упаковка акцентного цвета");
            AssertEqual(0x00FFFFFF, WindowChrome.ColorRef(System.Drawing.Color.White), "белый");
            AssertEqual(0x00000000, WindowChrome.ColorRef(System.Drawing.Color.Black), "чёрный");
            AssertEqual(0x000000FF, WindowChrome.ColorRef(System.Drawing.Color.FromArgb(255, 0, 0)), "красный -> младший байт");
        }

        private static void TestHeaderBand()
        {
            using (var band = new HeaderBand("Заголовок", "подпись", Theme.Accent, Theme.AccentPressed))
            {
                AssertTrue(band is System.Windows.Forms.Control, "это контрол");
                band.Width = 400;
                band.Height = 80;
                AssertEqual(400, band.Width, "ширина применяется");
                // Пустая подпись не должна ронять отрисовку логики (конструктор допускает null).
                using (var noSub = new HeaderBand("Только заголовок", null, Theme.PdfRed, Theme.PdfRedDark))
                    AssertTrue(noSub != null, "подпись null допустима");
            }
        }

        // ---------- PdfPageOrder ----------

        private static string OrderSignature(PdfPageOrder order)
        {
            var parts = new List<string>();
            for (int i = 0; i < order.Count; i++)
                parts.Add(order[i].FileName + ":" + (order[i].PageIndex + 1));
            return string.Join("|", parts.ToArray());
        }

        private static PdfPageOrder MakeOrder()
        {
            var order = new PdfPageOrder();
            order.AddDocument(@"C:\in\А.pdf", 2);
            order.AddDocument(@"C:\in\Б.pdf", 1);
            return order; // А:1 | А:2 | Б:1
        }

        private static void TestPdfOrderMoves()
        {
            PdfPageOrder order = MakeOrder();
            AssertEqual("А.pdf:1|А.pdf:2|Б.pdf:1", OrderSignature(order), "исходный порядок");

            AssertEqual(0, order.MoveUp(0), "MoveUp с верхней строки — на месте");
            AssertEqual(2, order.MoveDown(2), "MoveDown с нижней строки — на месте");

            AssertEqual(1, order.MoveUp(2), "MoveUp возвращает новый индекс");
            AssertEqual("А.pdf:1|Б.pdf:1|А.pdf:2", OrderSignature(order), "после MoveUp");
            AssertEqual(2, order.MoveDown(1), "MoveDown возвращает новый индекс");
            AssertEqual("А.pdf:1|А.pdf:2|Б.pdf:1", OrderSignature(order), "MoveDown вернул порядок");
        }

        private static void TestPdfOrderDragMove()
        {
            PdfPageOrder order = MakeOrder();
            order.Move(2, 0); // Б:1 в начало
            AssertEqual("Б.pdf:1|А.pdf:1|А.pdf:2", OrderSignature(order), "перенос вверх");
            order.Move(0, 3); // Б:1 в конец (вставка перед позицией 3)
            AssertEqual("А.pdf:1|А.pdf:2|Б.pdf:1", OrderSignature(order), "перенос вниз");
            order.Move(1, 1); // на себя — без изменений
            AssertEqual("А.pdf:1|А.pdf:2|Б.pdf:1", OrderSignature(order), "перенос на себя");
        }

        private static void TestPdfOrderRemove()
        {
            PdfPageOrder order = MakeOrder();
            order.RemoveAt(new[] { 2, 0 }); // произвольный порядок индексов
            AssertEqual("А.pdf:2", OrderSignature(order), "удаление набора");
            AssertEqual(1, order.Count, "осталась одна строка");
            order.Clear();
            AssertEqual(0, order.Count, "Clear очищает список");
        }

        private static void TestAssemble()
        {
            // Два источника A (2 стр.) и B (1 стр.); собираем в порядке из разных файлов.
            var bysource = new Dictionary<string, List<PdfPageText>>(StringComparer.OrdinalIgnoreCase)
            {
                { "A.pdf", new List<PdfPageText> { new PdfPageText { PageIndex = 0 }, new PdfPageText { PageIndex = 1 } } },
                { "B.pdf", new List<PdfPageText> { new PdfPageText { PageIndex = 0 } } }
            };
            var order = new List<PdfPageRef>
            {
                new PdfPageRef { SourcePath = "B.pdf", PageIndex = 0 },
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = 1 },
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = 0 }
            };
            List<PdfPageText> r = PdfPageExtraction.Assemble(bysource, order);
            AssertEqual(3, r.Count, "собраны все 3 страницы из двух файлов");
            AssertTrue(ReferenceEquals(r[0], bysource["B.pdf"][0]), "первая — стр.0 из B");
            AssertTrue(ReferenceEquals(r[1], bysource["A.pdf"][1]), "вторая — стр.1 из A");
            AssertTrue(ReferenceEquals(r[2], bysource["A.pdf"][0]), "третья — стр.0 из A");
            // Несуществующий источник и индекс вне диапазона — пропускаются.
            var bad = new List<PdfPageRef>
            {
                new PdfPageRef { SourcePath = "нет.pdf", PageIndex = 0 },
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = 9 },
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = -1 },
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = 0 }
            };
            List<PdfPageText> rb = PdfPageExtraction.Assemble(bysource, bad);
            AssertEqual(1, rb.Count, "остаётся только валидная ссылка");
            AssertTrue(ReferenceEquals(rb[0], bysource["A.pdf"][0]), "валидная — A стр.0");
        }

        // ---------- NoteText ----------

        private static void TestNoteBasics()
        {
            MergeResult res = MakePrevious(); // 1 ok, 2 skip
            var options = new MergeOptions();
            options.AddToc = true;
            NoteContent note = NoteText.Build(res, @"C:\in", options,
                new DateTime(2026, 7, 16, 14, 5, 0));

            AssertEqual("СПРАВКА", note.Title, "заголовок");
            string all = string.Join("|", note.Body.ToArray()) + "|" + string.Join("|", note.Tail.ToArray());
            AssertTrue(all.Contains("16 июля 2026 г.") && all.Contains("14:05"), "период: " + note.Body[0]);
            AssertTrue(all.Contains("Обработано файлов: 3"), "всего");
            AssertTrue(all.Contains("Включено листов в сводный файл: 1"), "включено");
            AssertTrue(all.Contains("Пропущено файлов: 2"), "пропущено");
            AssertTrue(all.Contains(res.OutputPath) && all.Contains("(XLSX)"), "файл свода и формат");
            AssertTrue(all.Contains("лист «Содержание» — да"), "параметры");
            AssertTrue(note.Signature.Contains("Исполнитель"), "подпись");
        }

        private static void TestNoteSkippedTable()
        {
            MergeResult res = MakePrevious();
            NoteContent note = NoteText.Build(res, @"C:\in", new MergeOptions(), DateTime.MinValue);

            AssertTrue(note.SkippedIntro != null, "есть вводная к таблице");
            AssertEqual(2, note.SkippedRows.Count, "строк в таблице");
            AssertEqual("1", note.SkippedRows[0][0], "нумерация");
            AssertEqual("Б.xlsx", note.SkippedRows[0][1], "имя файла");
            AssertEqual("битый", note.SkippedRows[0][2], "причина");
        }

        private static void TestNoteClean()
        {
            var res = new MergeResult();
            res.OutputPath = @"C:\out\Свод.xlsb";
            res.Files.Add(MakeResult(@"C:\in\А.xlsx", true, null));
            res.OkCount = 1;
            NoteContent note = NoteText.Build(res, @"C:\in", new MergeOptions(), DateTime.MinValue);

            AssertTrue(note.SkippedIntro == null, "таблица не нужна");
            AssertEqual(0, note.SkippedRows.Count, "нет строк");
            string all = string.Join("|", note.Body.ToArray());
            AssertTrue(all.Contains("Замечания отсутствуют"), "формулировка чистого итога");
        }

        // ---------- CombineRetryResults ----------

        private static FileResult MakeResult(string path, bool ok, string note)
        {
            var fr = new FileResult();
            fr.FullPath = path;
            fr.FileName = Path.GetFileName(path);
            fr.Ok = ok;
            fr.Note = note;
            if (ok)
                fr.SheetName = Path.GetFileNameWithoutExtension(path);
            return fr;
        }

        private static MergeResult MakePrevious()
        {
            var prev = new MergeResult();
            prev.OutputPath = @"C:\out\Свод.xlsx";
            prev.Files.Add(MakeResult(@"C:\in\А.xlsx", true, null));
            prev.Files.Add(MakeResult(@"C:\in\Б.xlsx", false, "битый"));
            prev.Files.Add(MakeResult(@"C:\in\В.xlsx", false, "пароль"));
            prev.OkCount = 1;
            prev.SkipCount = 2;
            return prev;
        }

        private static void TestSheetBaseName()
        {
            AssertEqual("Отчет", MergeService.SheetBaseName("Отчет", "Лист1", false), "только первый — имя файла");
            AssertEqual("Отчет · Лист1", MergeService.SheetBaseName("Отчет", "Лист1", true), "все листы — файл · лист");
        }

        private static void TestFileCount()
        {
            var res = new MergeResult();
            res.Files.Add(MakeResult(@"C:\in\А.xlsx", true, null)); // два листа из одного файла
            res.Files.Add(MakeResult(@"C:\in\А.xlsx", true, null));
            res.Files.Add(MakeResult(@"C:\in\Б.xlsx", true, null));
            AssertEqual(3, res.Files.Count, "листов (строк) — три");
            AssertEqual(2, res.FileCount, "файлов — два");
        }

        private static void TestCombineRetryMultiSheet()
        {
            // Прошлый прогон: А перенесён (1 лист), Б пропущен.
            var prev = new MergeResult();
            prev.OutputPath = @"C:\out\Свод.xlsx";
            prev.Files.Add(MakeResult(@"C:\in\А.xlsx", true, null));
            prev.Files.Add(MakeResult(@"C:\in\Б.xlsx", false, "битый"));
            prev.OkCount = 1;
            prev.SkipCount = 1;

            // Повтор Б в режиме «все листы» дал два листа.
            var b1 = MakeResult(@"C:\in\Б.xlsx", true, null);
            var b2 = MakeResult(@"C:\in\Б.xlsx", true, null);
            var combined = MergeService.CombineRetryResults(prev, new List<FileResult> { b1, b2 });

            AssertEqual(3, combined.Files.Count, "А + два листа Б");
            AssertEqual(3, combined.OkCount, "все перенесены");
            AssertEqual(0, combined.SkipCount, "пропущенных нет");
            AssertEqual(2, combined.FileCount, "файлов — два (А и Б)");
        }

        private static void TestCombineRetryReplaces()
        {
            MergeResult prev = MakePrevious();
            var attempts = new List<FileResult> { MakeResult(@"C:\in\Б.xlsx", true, null) };
            MergeResult combined = MergeService.CombineRetryResults(prev, attempts);

            AssertEqual(3, combined.Files.Count, "число записей");
            AssertEqual(2, combined.OkCount, "перенесено после повтора");
            AssertEqual(1, combined.SkipCount, "осталось пропущенных");
            AssertTrue(combined.Files[1].Ok, "Б теперь перенесён");
            AssertEqual(prev.OutputPath, combined.OutputPath, "путь свода");
        }

        private static void TestCombineRetryKeepsFailed()
        {
            MergeResult prev = MakePrevious();
            var attempts = new List<FileResult> { MakeResult(@"C:\in\Б.xlsx", false, "снова битый") };
            MergeResult combined = MergeService.CombineRetryResults(prev, attempts);

            AssertEqual(1, combined.OkCount, "перенесено не изменилось");
            AssertEqual(2, combined.SkipCount, "пропущенных столько же");
            AssertEqual("снова битый", combined.Files[1].Note, "причина обновлена");
        }

        private static void TestCombineRetryOrder()
        {
            MergeResult prev = MakePrevious();
            var attempts = new List<FileResult>
            {
                MakeResult(@"C:\in\В.xlsx", true, null),
                MakeResult(@"C:\in\Б.xlsx", true, null)
            };
            MergeResult combined = MergeService.CombineRetryResults(prev, attempts);

            AssertEqual("А.xlsx|Б.xlsx|В.xlsx",
                combined.Files[0].FileName + "|" + combined.Files[1].FileName + "|" + combined.Files[2].FileName,
                "порядок исходного прогона сохранён");
            AssertEqual(3, combined.OkCount, "все перенесены");
            AssertEqual(0, combined.SkipCount, "пропущенных не осталось");
            AssertTrue(!ReferenceEquals(prev.Files[0], null) && combined.Files[0].Ok, "успешная запись не тронута");
        }

        // ---------- OutputFormats ----------

        private static void TestOutputFormatCodes()
        {
            AssertEqual(51, OutputFormats.FileFormatFor(@"C:\a\Свод.xlsx"), "xlsx");
            AssertEqual(52, OutputFormats.FileFormatFor("Свод.xlsm"), "xlsm");
            AssertEqual(50, OutputFormats.FileFormatFor("Свод.XLSB"), "xlsb в верхнем регистре");
            AssertEqual(56, OutputFormats.FileFormatFor("Свод.xls"), "xls");
            AssertEqual(0, OutputFormats.FileFormatFor("Свод.pdf"), "чужое расширение");
            AssertEqual(0, OutputFormats.FileFormatFor("Свод"), "без расширения");
        }

        private static void TestStripExtension()
        {
            AssertEqual("Свод", OutputFormats.StripKnownExtension("Свод.xlsx"), "xlsx");
            AssertEqual("Свод", OutputFormats.StripKnownExtension("Свод.XLS"), "xls в верхнем регистре");
            AssertEqual("Свод.pdf", OutputFormats.StripKnownExtension("Свод.pdf"), "чужое расширение не трогаем");
            AssertEqual("Свод", OutputFormats.StripKnownExtension("Свод"), "без расширения");
        }

        // ---------- CrashReport / классификация ошибок ----------

        private static void TestCrashReportFormat()
        {
            string entry = CrashReport.Format(new InvalidOperationException("бум"),
                "1.16.7", new DateTime(2026, 7, 24, 12, 30, 45));
            AssertTrue(entry.StartsWith("[2026-07-24 12:30:45] v1.16.7"), "метка времени и версия");
            AssertTrue(entry.Contains("InvalidOperationException") && entry.Contains("бум"), "тип и сообщение");
            AssertTrue(entry.EndsWith("\r\n\r\n"), "записи разделены пустой строкой");
            AssertTrue(CrashReport.Format(null, "1.0.0", DateTime.Now).Contains("(null)"), "null-исключение не роняет лог");
        }

        private static void TestPermanentOpenError()
        {
            AssertTrue(MergeService.IsPermanentOpenError("Введён неверный пароль."), "ru: пароль");
            AssertTrue(MergeService.IsPermanentOpenError("The password you supplied is not correct."), "en: password");
            AssertTrue(MergeService.IsPermanentOpenError("Файл повреждён и не может быть открыт."), "ru: повреждён");
            AssertTrue(MergeService.IsPermanentOpenError("The file format or file extension is not valid."), "en: not valid");
            AssertTrue(MergeService.IsPermanentOpenError("The file is corrupt and cannot be opened."), "en: corrupt");
            AssertTrue(!MergeService.IsPermanentOpenError("Вызов был отклонён."), "транзиентный COM-сбой — ретраить");
            AssertTrue(!MergeService.IsPermanentOpenError(null), "null — не постоянная ошибка");
        }

        private static void TestShouldWrap()
        {
            AssertTrue(MergeException.ShouldWrap(new IOException("диск")), "обычную ошибку оборачиваем");
            AssertTrue(!MergeException.ShouldWrap(new OutOfMemoryException()), "OOM не маскируем под «файл повреждён»");
        }

        // ---------- CheckOutputWritable ----------

        private static void TestOutputLocked()
        {
            string path = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N") + ".xlsx");
            using (new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                string error = MergeService.CheckOutputWritable(path);
                AssertTrue(error != null, "занятый файл должен давать ошибку");
                AssertTrue(error.Contains("занят"), "сообщение о занятости: " + error);
            }
            File.Delete(path);
        }

        private static void TestOutputWritable()
        {
            string path = Path.Combine(Path.GetTempPath(), "ExcelMergerTests_" + Guid.NewGuid().ToString("N") + ".xlsx");

            // Несуществующий файл: проверка не должна оставлять след
            AssertEqual(null, MergeService.CheckOutputWritable(path), "новый файл");
            AssertTrue(!File.Exists(path), "пробный файл удалён");

            // Существующий свободный файл
            File.WriteAllText(path, "x");
            try
            {
                AssertEqual(null, MergeService.CheckOutputWritable(path), "свободный файл");
            }
            finally
            {
                File.Delete(path);
            }
        }

        private static void TestOutputBadFolder()
        {
            string path = Path.Combine(Path.GetTempPath(),
                "ExcelMergerTests_нет_такой_" + Guid.NewGuid().ToString("N"), "Свод.xlsx");
            string error = MergeService.CheckOutputWritable(path);
            AssertTrue(error != null && error.Contains("не существует"), "папка не существует: " + error);
        }

        // ---------- PDF-сжатие (Ghostscript) ----------

        private static void TestCompressionPreset()
        {
            AssertEqual("/ebook", PdfCompression.Preset(CompressionLevel.Good), "Хорошо -> /ebook");
            AssertEqual("/screen", PdfCompression.Preset(CompressionLevel.Small), "Нормально -> /screen");
            AssertEqual("/default", PdfCompression.Preset(CompressionLevel.VeryGood), "Очень хорошо -> /default");
            AssertEqual(null, PdfCompression.Preset(CompressionLevel.None), "Отлично -> без пресета");
        }

        /// <summary>
        /// Число уровня — идентификатор в settings.txt, позиция в списке — дело показа.
        /// Пин ровно на это: значения ранее существовавших уровней МЕНЯТЬ НЕЛЬЗЯ, иначе у
        /// всех, кто выбрал «Хорошо», после обновления окажется выбран другой уровень.
        /// </summary>
        private static void TestCompressionLevelOrder()
        {
            AssertEqual(0, (int)CompressionLevel.None, "«Отлично» = 0 навсегда");
            AssertEqual(1, (int)CompressionLevel.Good, "«Хорошо» = 1 навсегда");
            AssertEqual(2, (int)CompressionLevel.Small, "«Нормально» = 2 навсегда");
            AssertEqual(3, (int)CompressionLevel.VeryGood, "новый уровень добавлен в КОНЕЦ, а не в середину");

            // Показываем по возрастанию силы сжатия, а не по значению перечисления.
            AssertEqual(4, PdfCompression.LevelLabels().Length, "уровней четыре");
            AssertEqual(CompressionLevel.None, PdfCompression.LevelAt(0), "первым — «без сжатия»");
            AssertEqual(CompressionLevel.VeryGood, PdfCompression.LevelAt(1), "затем «очень хорошо»");
            AssertEqual(CompressionLevel.Good, PdfCompression.LevelAt(2), "затем «хорошо»");
            AssertEqual(CompressionLevel.Small, PdfCompression.LevelAt(3), "последним — самый сильный");

            // Перевод «позиция ↔ уровень» обратим и не падает за границами.
            for (int i = 0; i < PdfCompression.LevelLabels().Length; i++)
                AssertEqual(i, PdfCompression.IndexOf(PdfCompression.LevelAt(i)),
                    "позиция и обратно даёт ту же позицию: " + i);
            AssertEqual(CompressionLevel.None, PdfCompression.LevelAt(-1), "позиция левее списка — без сжатия");
            AssertEqual(CompressionLevel.None, PdfCompression.LevelAt(99), "позиция правее списка — без сжатия");
            AssertEqual(0, PdfCompression.IndexOf((CompressionLevel)77), "неизвестный уровень — позиция «без сжатия»");

            // Каждый уровень перечисления обязан иметь своё место в списке — иначе новый
            // уровень появился бы в настройках, но не в окне.
            foreach (CompressionLevel level in Enum.GetValues(typeof(CompressionLevel)))
                AssertEqual(level, PdfCompression.LevelAt(PdfCompression.IndexOf(level)),
                    "уровень показывается в списке: " + level);
        }

        /// <summary>
        /// Сквозной путь уровня через диск. Настройки хранят ЧИСЛО уровня, а список показывает
        /// его на СВОЁЙ позиции, и это разные вещи: у «Очень хорошо» число 3, а позиция 1.
        /// Сюда же — обещание совместимости: настройка, записанная прежней версией, обязана
        /// означать то же самое и после обновления.
        /// Живой %APPDATA% не трогается: тест работает в изолированном корне AppPaths.
        /// </summary>
        private static void TestCompressionLevelRoundTrip()
        {
            string root = Path.Combine(Path.GetTempPath(), "iwo_level_" + Guid.NewGuid().ToString("N"));
            AppPaths.SetRootForTests(root);
            try
            {
                new UserSettings().SaveView(300, (int)CompressionLevel.VeryGood);
                int stored = UserSettings.Load().CompressionLevel;
                AssertEqual(3, stored, "на диск ушло число уровня, а не его позиция в списке");
                AssertEqual(CompressionLevel.VeryGood, (CompressionLevel)stored, "число вернулось тем же уровнем");
                AssertEqual(1, PdfCompression.IndexOf((CompressionLevel)stored),
                    "а показывается он вторым по счёту — позиция с числом не совпадает намеренно");

                // Настройка прежних версий: 1 — это «Хорошо», 2 — «Нормально», и так навсегда.
                new UserSettings().SaveView(300, 1);
                AssertEqual(CompressionLevel.Good, (CompressionLevel)UserSettings.Load().CompressionLevel,
                    "«1» из настроек прежней версии по-прежнему «Хорошо»");
                new UserSettings().SaveView(300, 2);
                AssertEqual(CompressionLevel.Small, (CompressionLevel)UserSettings.Load().CompressionLevel,
                    "«2» из настроек прежней версии по-прежнему «Нормально»");
            }
            finally
            {
                AppPaths.SetRootForTests(null);
                try { Directory.Delete(root, true); } catch { }
            }
        }

        /// <summary>
        /// Уровень «Очень хорошо» обещает нетронутую чёткость — и держать обещание обязаны
        /// аргументы, а не умолчания чужого движка: портативная сборка работает с тем
        /// Ghostscript, который стоит у пользователя, и его версия нам неизвестна.
        /// </summary>
        private static void TestCompressionKeepsResolution()
        {
            AssertTrue(!PdfCompression.Downsamples(CompressionLevel.VeryGood), "«Очень хорошо» не пересчитывает картинки");
            AssertTrue(!PdfCompression.Downsamples(CompressionLevel.None), "«Отлично» тем более");
            AssertTrue(PdfCompression.Downsamples(CompressionLevel.Good), "«Хорошо» пересчитывает");
            AssertTrue(PdfCompression.Downsamples(CompressionLevel.Small), "«Нормально» пересчитывает");

            string keep = PdfCompression.BuildArguments("in.pdf", "out.pdf", CompressionLevel.VeryGood, null);
            AssertTrue(keep.Contains("-dPDFSETTINGS=/default"), "пресет без пересчёта изображений");
            AssertTrue(keep.Contains("-dDownsampleColorImages=false"), "цветные не пересчитываются — явно");
            AssertTrue(keep.Contains("-dDownsampleGrayImages=false"), "серые не пересчитываются — явно");
            AssertTrue(keep.Contains("-dDownsampleMonoImages=false"), "штриховые не пересчитываются — явно");

            // А уровням с пересчётом эти ключи, наоборот, всё сломали бы.
            foreach (CompressionLevel level in new[] { CompressionLevel.Good, CompressionLevel.Small })
            {
                string args = PdfCompression.BuildArguments("in.pdf", "out.pdf", level, null);
                AssertTrue(!args.Contains("Downsample"), "уровень с пересчётом не выключает его: " + level);
            }
        }

        private static void TestCompressionArgs()
        {
            // Путь с пробелами обязан быть в кавычках, иначе GS примет его за два аргумента.
            string args = PdfCompression.BuildArguments(
                @"C:\Users\test user\вход.pdf", @"C:\Users\test user\выход.pdf",
                CompressionLevel.Good, null);
            AssertTrue(args.Contains("-sDEVICE=pdfwrite"), "устройство pdfwrite");
            AssertTrue(args.Contains("-dCompatibilityLevel=1.4"), "1.4 — читаемо старым PdfSharp");
            AssertTrue(args.Contains("-dPDFSETTINGS=/ebook"), "пресет /ebook");
            AssertTrue(args.Contains("-dSAFER"), "SAFER включён");
            AssertTrue(!args.Contains("-dNOSAFER"), "NOSAFER не должен передаваться");
            AssertTrue(args.Contains("\"C:\\Users\\test user\\вход.pdf\""), "вход в кавычках");
            AssertTrue(args.Contains("-sOutputFile=\"C:\\Users\\test user\\выход.pdf\""), "выход в кавычках");
            AssertTrue(!args.Contains(" -I "), "системный GS — без -I");

            // Вшитый GS: добавляется -I на lib и Resource\Init.
            string bundled = PdfCompression.BuildArguments("in.pdf", "out.pdf", CompressionLevel.Small, @"C:\app\gs");
            AssertTrue(bundled.Contains("-dPDFSETTINGS=/screen"), "пресет /screen");
            AssertTrue(bundled.Contains("-I \"C:\\app\\gs\\lib\""), "-I lib для бандла");
            AssertTrue(bundled.Contains("-I \"C:\\app\\gs\\Resource\\Init\""), "-I Resource\\Init для бандла");
        }

        private static void TestCompressionShouldReplace()
        {
            // Годность вывода к этому моменту уже проверил конвейер, здесь остался размер.
            AssertTrue(PdfCompression.ShouldReplace(1000, 400), "меньше — заменяем");
            AssertTrue(!PdfCompression.ShouldReplace(1000, 1000), "равный размер — оставляем оригинал");
            AssertTrue(!PdfCompression.ShouldReplace(1000, 1500), "больше — оставляем оригинал");
            AssertTrue(!PdfCompression.ShouldReplace(1000, 0), "пустой вывод — не заменяем");
        }

        private static void TestGhostscriptPick()
        {
            var candidates = new List<string> { null, "", @"C:\нет\gs.exe", @"C:\есть\gs.exe", @"C:\тоже\gs.exe" };
            string picked = Ghostscript.PickFirstExisting(candidates,
                delegate(string p) { return p == @"C:\есть\gs.exe" || p == @"C:\тоже\gs.exe"; });
            AssertEqual(@"C:\есть\gs.exe", picked, "первый существующий, пустые пропущены");
            AssertEqual(null, Ghostscript.PickFirstExisting(candidates, delegate { return false; }), "ни одного — null");
            // Бросающий предикат не должен ронять выбор.
            string safe = Ghostscript.PickFirstExisting(new[] { @"C:\a", @"C:\b" },
                delegate(string p) { if (p == @"C:\a") throw new Exception("bad"); return true; });
            AssertEqual(@"C:\b", safe, "исключение в предикате -> пропуск кандидата");
        }

        // ---------- мини-раннер ----------

        // ---------- LruCache ----------

        private static void TestLruEviction()
        {
            var evicted = new List<string>();
            var cache = new LruCache<string>(2, delegate(string v) { evicted.Add(v); });
            cache.Add("a", "a");
            cache.Add("b", "b");
            AssertEqual(2, cache.Count, "две записи в пределах ёмкости");
            cache.Add("c", "c"); // переполнение: вытесняется самый несвежий — «a»
            AssertEqual(2, cache.Count, "ёмкость соблюдена");
            AssertEqual(1, evicted.Count, "ровно одно вытеснение");
            AssertEqual("a", evicted[0], "вытеснен наименее недавно использованный");
            string val;
            AssertTrue(!cache.TryGet("a", out val), "a вытеснен");
            AssertTrue(cache.TryGet("b", out val) && val == "b", "b остался");
            AssertTrue(cache.TryGet("c", out val) && val == "c", "c добавлен");
        }

        private static void TestLruTouchOnGet()
        {
            var evicted = new List<string>();
            var cache = new LruCache<string>(2, delegate(string v) { evicted.Add(v); });
            cache.Add("a", "a");
            cache.Add("b", "b");
            string val;
            AssertTrue(cache.TryGet("a", out val), "обращение к a делает его свежим");
            cache.Add("c", "c"); // теперь несвежий — «b»
            AssertEqual("b", evicted[0], "touch через TryGet сместил вытеснение на b");
            AssertTrue(cache.TryGet("a", out val), "a сохранён");
        }

        private static void TestLruReplace()
        {
            var evicted = new List<string>();
            var cache = new LruCache<string>(2, delegate(string v) { evicted.Add(v); });
            cache.Add("x", "x1");
            cache.Add("x", "x2"); // тот же ключ — замена, не рост
            AssertEqual(1, cache.Count, "замена ключа не увеличивает размер");
            AssertEqual(0, evicted.Count, "замена ничего не вытесняет");
            string val;
            AssertTrue(cache.TryGet("x", out val) && val == "x2", "значение обновлено");
        }

        private static void TestLruCaseInsensitive()
        {
            var cache = new LruCache<string>(2, null);
            cache.Add(@"C:\A.pdf", "doc");
            string val;
            AssertTrue(cache.TryGet(@"c:\a.pdf", out val) && val == "doc", "ключи-пути сравниваются без регистра");
            AssertEqual(1, cache.Count, "разный регистр — тот же ключ");
        }

        private static void TestLruClear()
        {
            var evicted = new List<string>();
            var cache = new LruCache<string>(3, delegate(string v) { evicted.Add(v); });
            cache.Add("a", "a");
            cache.Add("b", "b");
            cache.Clear();
            AssertEqual(0, cache.Count, "после Clear пусто");
            AssertEqual(2, evicted.Count, "Clear освобождает все оставшиеся элементы");
        }

        private static void TestLruCapacityGuard()
        {
            bool threw = false;
            try { new LruCache<string>(0, null); }
            catch (ArgumentOutOfRangeException) { threw = true; }
            AssertTrue(threw, "ёмкость < 1 должна отвергаться");
        }

        // ---------- PdfPageGrid: набор ключей и вытеснение кэша ----------

        private static void TestGridBuildKeySet()
        {
            var pages = new List<PdfPageRef>
            {
                new PdfPageRef { SourcePath = @"C:\a.pdf", PageIndex = 0 },
                new PdfPageRef { SourcePath = @"C:\a.pdf", PageIndex = 1 },
                new PdfPageRef { SourcePath = @"C:\a.pdf", PageIndex = 0 }, // дубль
            };
            HashSet<string> keys = PdfPageGrid.BuildKeySet(pages);
            AssertEqual(2, keys.Count, "дубли схлопываются");
            AssertTrue(keys.Contains(PdfPageGrid.ThumbKey(pages[0])), "ключ страницы 0 присутствует");
            AssertEqual(0, PdfPageGrid.BuildKeySet(null).Count, "null -> пустой набор");
        }

        private static void TestGridStaleKeys()
        {
            var cached = new List<string> { "a|0", "a|1", "b|0" };
            var keep = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "a|0", "b|0", "c|0" };
            List<string> stale = PdfPageGrid.StaleKeys(cached, keep);
            AssertEqual(1, stale.Count, "один устаревший ключ");
            AssertEqual("a|1", stale[0], "вытесняется отсутствующий в keep");

            // Тот же набор -> ничего не устаревает (переупорядочивание не роняет кэш).
            var same = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "a|0", "a|1", "b|0" };
            AssertEqual(0, PdfPageGrid.StaleKeys(cached, same).Count, "полное совпадение — без вытеснения");
        }

        private static void TestLowerBound()
        {
            AssertEqual(3, PdfPageGrid.LowerBound(5, delegate(int i) { return i >= 3; }), "первый индекс >= 3");
            AssertEqual(0, PdfPageGrid.LowerBound(5, delegate(int i) { return true; }), "все истинны -> 0");
            AssertEqual(5, PdfPageGrid.LowerBound(5, delegate(int i) { return false; }), "все ложны -> count");
            AssertEqual(0, PdfPageGrid.LowerBound(0, delegate(int i) { return true; }), "пустой -> 0");
            AssertEqual(1, PdfPageGrid.LowerBound(4, delegate(int i) { return i >= 1; }), "граница у начала");
        }

        private static void TestVisibleRange()
        {
            // 8 плиток высотой 10, один столбец, прокрутка вверх на 25; низ вьюпорта = 30.
            // Top[i]=i*10-25, Bottom[i]=Top+10 -> видимы i=2..5.
            Func<int, int> topOf = delegate(int i) { return i * 10 - 25; };
            Func<int, int> bottomOf = delegate(int i) { return i * 10 - 25 + 10; };
            int first, last;
            PdfPageGrid.VisibleRange(8, topOf, bottomOf, 30, out first, out last);
            AssertEqual(2, first, "первый видимый (Bottom >= 0)");
            AssertEqual(5, last, "последний видимый (Top <= 30)");

            // Всё выше вьюпорта -> ничего целиком не видно (first > last).
            Func<int, int> topHi = delegate(int i) { return i * 10 - 1000; };
            Func<int, int> botHi = delegate(int i) { return i * 10 - 1000 + 10; };
            PdfPageGrid.VisibleRange(8, topHi, botHi, 30, out first, out last);
            AssertTrue(first > last, "всё выше вьюпорта -> пусто");

            // Пустой список.
            PdfPageGrid.VisibleRange(0, topOf, bottomOf, 30, out first, out last);
            AssertTrue(first > last, "нет элементов -> пусто");
        }

        private static void TestSuggestCompression()
        {
            long mb = 1024L * 1024;
            AssertTrue(PdfSplitForm.ShouldSuggestCompression(CompressionLevel.None, 10 * mb, 95 * mb / 10),
                "без сжатия, 9.5МБ из 10МБ -> подсказать");
            AssertTrue(!PdfSplitForm.ShouldSuggestCompression(CompressionLevel.None, 10 * mb, 3 * mb),
                "3МБ из 10МБ -> не подсказывать");
            AssertTrue(!PdfSplitForm.ShouldSuggestCompression(CompressionLevel.Good, 10 * mb, 95 * mb / 10),
                "уже со сжатием -> не подсказывать");
            AssertTrue(!PdfSplitForm.ShouldSuggestCompression(CompressionLevel.None, 700 * 1024, 690 * 1024),
                "мелкий файл (<1МБ) -> не подсказывать");
            AssertTrue(PdfSplitForm.ShouldSuggestCompression(CompressionLevel.None, 10 * mb, 9 * mb),
                "ровно 90% -> подсказать (граница)");
            AssertTrue(!PdfSplitForm.ShouldSuggestCompression(CompressionLevel.None, 0, 5 * mb),
                "размер исходника неизвестен (0) -> не подсказывать");
        }

        // ---------- OcrLayout (порядок чтения born-digital) ----------

        /// <summary>Слово с рамкой (Y вверх): left/bottom — левый нижний угол, +ширина/высота.</summary>
        private static PdfWord W(string text, double left, double bottom, double width, double height)
        {
            return new PdfWord { Text = text, Left = left, Right = left + width, Bottom = bottom, Top = bottom + height };
        }

        private static void TestOcrReadingOrder()
        {
            // Ввод намеренно вперемешку; ожидаем «Hello world» затем «second line».
            var words = new List<PdfWord>
            {
                W("world", 35, 90, 30, 10),
                W("line", 35, 70, 20, 10),
                W("Hello", 0, 90, 30, 10),
                W("second", 0, 70, 30, 10)
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(1, p.Count, "две близкие строки -> один абзац");
            AssertEqual("Hello world second line", p[0], "порядок чтения");
        }

        private static void TestOcrParagraphs()
        {
            // Три плотные строки + одна с большим зазором -> два абзаца.
            var words = new List<PdfWord>
            {
                W("Aaa", 0, 100, 20, 10), // midY 105
                W("Bbb", 0, 88, 20, 10),  // midY 93  (зазор 12)
                W("Ccc", 0, 76, 20, 10),  // midY 81  (зазор 12)
                W("Ddd", 0, 50, 20, 10)   // midY 55  (зазор 26 -> новый абзац)
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(2, p.Count, "разрыв абзаца по большому зазору");
            AssertEqual("Aaa Bbb Ccc", p[0], "первый абзац");
            AssertEqual("Ddd", p[1], "второй абзац");
        }

        private static void TestOcrParagraphsIndent()
        {
            // Justified-документ (строки достают до правого поля 100), абзацы разделены
            // ТОЛЬКО красной строкой (первая строка с отступом) при равном интервале —
            // как в Word-экспорте. Зазор одинаков всюду, поэтому делить должны отступ и
            // «короткая» последняя строка, а не зазор.
            var words = new List<PdfWord>
            {
                W("A1a", 15, 180, 85, 10), // абзац A, 1-я строка с отступом, до правого поля (Right 100)
                W("A2", 0, 168, 100, 10),  // продолжение у левого поля, полная строка
                W("A3", 0, 156, 40, 10),   // короткая последняя строка абзаца
                W("B1", 15, 140, 85, 10),  // абзац B, красная строка (отступ) -> новый абзац
                W("B2", 0, 128, 30, 10)    // короткая последняя строка
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(2, p.Count, "красная строка + короткая строка делят абзацы без зазора");
            AssertEqual("A1a A2 A3", p[0], "первый абзац");
            AssertEqual("B1 B2", p[1], "второй абзац");
        }

        private static void TestOcrIndentDetected()
        {
            // Тот же justified-документ с красной строкой: отступ первых строк 15 pt
            // должен быть измерен (большинство абзацев с отступом).
            var words = new List<PdfWord>
            {
                W("A1a", 15, 180, 85, 10),
                W("A2", 0, 168, 100, 10),
                W("A3", 0, 156, 40, 10),
                W("B1", 15, 140, 85, 10),
                W("B2", 0, 128, 30, 10)
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(2, layout.Paragraphs.Count, "два абзаца");
            AssertEqual(15.0, layout.FirstLineIndentPt, "измерен отступ красной строки (медиана 15 pt)");
        }

        private static void TestOcrNoIndentReported()
        {
            // Документ без отступов (все строки у левого поля) -> красная строка не навязывается.
            var words = new List<PdfWord>
            {
                W("Aaa", 0, 100, 20, 10),
                W("Bbb", 0, 88, 20, 10),
                W("Ccc", 0, 76, 20, 10),
                W("Ddd", 0, 50, 20, 10)
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(0.0, layout.FirstLineIndentPt, "без отступов -> 0 (не портим документ)");
        }

        private static void TestFontNames()
        {
            AssertEqual("Times New Roman", FontNames.Clean("UBWKNX+Times New Roman,Italic"), "subset + курсив-суффикс");
            AssertEqual("Arial", FontNames.Clean("ABCDEF+Arial-BoldMT"), "subset + -BoldMT");
            AssertEqual("Times New Roman", FontNames.Clean("TimesNewRomanPSMT"), "PSMT + слитное имя");
            AssertEqual("Times New Roman", FontNames.Clean("TimesNewRomanPS-BoldMT"), "PS + -BoldMT");
            AssertEqual("Courier New", FontNames.Clean("CourierNewPS-BoldMT"), "Courier New");
            AssertEqual("Arial", FontNames.Clean("ArialMT"), "ArialMT");
            AssertEqual("Calibri", FontNames.Clean("Calibri"), "уже чистое");
            AssertEqual("PT Astra Serif", FontNames.Clean("BBHOZJ+PTAstraSerif-Regular"), "PT Astra Serif: префикс-аббревиатура + слитное");
            AssertEqual("MS Gothic", FontNames.Clean("MSGothic"), "MS Gothic: аббревиатура-префикс");
            AssertTrue(FontNames.Clean(null) == null, "null -> null");
            AssertTrue(FontNames.Clean("  ") == null, "пусто -> null");
        }

        private static void TestResolveFontName()
        {
            var installed = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Times New Roman", "Calibri Light", "Liberation Serif"
            };
            // Установленный шрифт — оставить как есть (в т.ч. без учёта регистра) для латиницы.
            AssertEqual("Calibri Light", FontResolver.ResolveFontName("Calibri Light", "text", installed, "Times New Roman"), "установленный сохранён");
            AssertEqual("times new roman", FontResolver.ResolveFontName("times new roman", "текст", installed, "Times New Roman"), "регистр при поиске не важен");
            // НЕустановленный (напр. PT Astra Serif) -> fallback, иначе Word уводит кириллицу в eastAsia -> разрядка.
            AssertEqual("Times New Roman", FontResolver.ResolveFontName("PT Astra Serif", "текст", installed, "Times New Roman"), "неустановленный -> fallback");
            // УСТАНОВЛЕННЫЙ, но не Word-родной (Liberation Serif): кириллице — fallback (Word
            // ставит hint=eastAsia и разжижает буквы CJK-выключкой), латинице — оставить.
            AssertEqual("Times New Roman", FontResolver.ResolveFontName("Liberation Serif", "Пример", installed, "Times New Roman"), "кириллица вне сейф-листа -> fallback");
            AssertEqual("Liberation Serif", FontResolver.ResolveFontName("Liberation Serif", "latin only", installed, "Times New Roman"), "латиница может остаться");
            AssertEqual("Calibri Light", FontResolver.ResolveFontName("Calibri Light", "кириллица", installed, "Times New Roman"), "сейф-лист держит кириллицу");
            AssertEqual("Times New Roman", FontResolver.ResolveFontName(null, "т", installed, "Times New Roman"), "null -> fallback");
            AssertEqual("Times New Roman", FontResolver.ResolveFontName("X", "т", null, "Times New Roman"), "нет списка -> fallback");
        }

        private static void TestProgressPercent()
        {
            AssertEqual(0, PdfToolFormBase.ProgressPercent(0, 0), "0/0 -> 0 (без деления на ноль)");
            AssertEqual(0, PdfToolFormBase.ProgressPercent(0, 10), "0/10 -> 0");
            AssertEqual(50, PdfToolFormBase.ProgressPercent(5, 10), "5/10 -> 50");
            AssertEqual(100, PdfToolFormBase.ProgressPercent(10, 10), "10/10 -> 100");
            AssertEqual(100, PdfToolFormBase.ProgressPercent(11, 10), "11/10 -> 100 (кламп сверху)");
            AssertEqual(0, PdfToolFormBase.ProgressPercent(-1, 10), "отрицательное сделано -> 0");
            AssertEqual(0, PdfToolFormBase.ProgressPercent(5, -1), "отрицательное всего -> 0");
            AssertEqual(33, PdfToolFormBase.ProgressPercent(1, 3), "1/3 -> 33 (округление вниз)");
            AssertEqual(66, PdfToolFormBase.ProgressPercent(2, 3), "2/3 -> 66");
            AssertEqual(50, PdfToolFormBase.ProgressPercent(1000000, 2000000), "большие числа без переполнения");
        }

        private static void TestOcrRunsFontFamily()
        {
            // Разные семейства -> разные раны.
            var words = new List<PdfWord>
            {
                new PdfWord { Text = "Ариал", Left = 0, Right = 40, Bottom = 0, Top = 8, FontSizePt = 12, FontName = "Arial" },
                new PdfWord { Text = "Таймс", Left = 45, Right = 85, Bottom = 0, Top = 8, FontSizePt = 12, FontName = "Times New Roman" }
            };
            List<OcrRun> runs = OcrLayout.Analyze(words).Paragraphs[0].Runs;
            AssertEqual(2, runs.Count, "смена шрифта -> новый ран");
            AssertEqual("Arial", runs[0].FontName, "первый ран — Arial");
            AssertEqual("Times New Roman", runs[1].FontName, "второй ран — Times New Roman");
        }

        private static void TestOcrParagraphStyle()
        {
            // Курсивная строка кеглем 14, единый формат -> один ран, курсив, кегль 14.
            var words = new List<PdfWord>
            {
                new PdfWord { Text = "Имя:", Left = 0, Right = 30, Bottom = 0, Top = 8, FontSizePt = 14, Italic = true },
                new PdfWord { Text = "_dmarc", Left = 35, Right = 90, Bottom = 0, Top = 8, FontSizePt = 14, Italic = true }
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(1, layout.Paragraphs.Count, "один абзац");
            AssertEqual(1, layout.Paragraphs[0].Runs.Count, "единый формат -> один ран");
            OcrRun r = layout.Paragraphs[0].Runs[0];
            AssertEqual(14.0, r.FontSizePt, "кегль рана");
            AssertTrue(r.Italic, "ран курсивный");
            AssertTrue(!r.Bold, "не полужирный");
            AssertEqual("Имя: _dmarc", r.Text, "текст рана");
        }

        private static void TestOcrRunsMixedFormat()
        {
            // Полужирное слово среди обычных -> три рана; жирным только среднее.
            var words = new List<PdfWord>
            {
                new PdfWord { Text = "обычное", Left = 0, Right = 40, Bottom = 0, Top = 8, FontSizePt = 12 },
                new PdfWord { Text = "жирное", Left = 45, Right = 85, Bottom = 0, Top = 8, FontSizePt = 12, Bold = true },
                new PdfWord { Text = "снова", Left = 90, Right = 130, Bottom = 0, Top = 8, FontSizePt = 12 }
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(1, layout.Paragraphs.Count, "один абзац");
            List<OcrRun> runs = layout.Paragraphs[0].Runs;
            AssertEqual(3, runs.Count, "три рана: обычный / жирный / обычный");
            AssertTrue(!runs[0].Bold && runs[1].Bold && !runs[2].Bold, "жирный только средний ран");
            AssertEqual("обычное жирное снова", layout.Paragraphs[0].Text, "склейка текста");
        }

        private static void TestOcrSuperscript()
        {
            // Мелкое приподнятое слово среди обычных -> надстрочный ран.
            var words = new List<PdfWord>
            {
                W("a", 0, 0, 10, 8),    // обычные, база 0, высота 8
                W("b", 12, 0, 10, 8),
                W("2", 24, 4, 6, 4)     // мельче (4) и приподнято (база 4) -> надстрочный
            };
            List<OcrRun> runs = OcrLayout.Analyze(words).Paragraphs[0].Runs;
            AssertEqual(2, runs.Count, "надстрочный — отдельный ран");
            AssertTrue(runs[1].Super && !runs[1].Sub, "'2' надстрочный");
            AssertEqual("2", runs[1].Text, "текст надстрочного рана");
        }

        private static void TestOcrHyperlinkRun()
        {
            var words = new List<PdfWord>
            {
                new PdfWord { Text = "ссылка", Left = 0, Right = 40, Bottom = 0, Top = 8, FontSizePt = 12, Uri = "https://example.com" },
                new PdfWord { Text = "обычный", Left = 45, Right = 90, Bottom = 0, Top = 8, FontSizePt = 12 }
            };
            List<OcrRun> runs = OcrLayout.Analyze(words).Paragraphs[0].Runs;
            AssertEqual(2, runs.Count, "ссылка — отдельный ран");
            AssertEqual("https://example.com", runs[0].Uri, "URI рана сохранён");
            AssertTrue(runs[1].Uri == null, "обычный ран без ссылки");
        }

        private static void TestOcrColorRun()
        {
            var words = new List<PdfWord>
            {
                new PdfWord { Text = "красный", Left = 0, Right = 40, Bottom = 0, Top = 8, FontSizePt = 12, ColorArgb = 0xFF0000 }
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(1, layout.Paragraphs[0].Runs.Count, "один ран");
            AssertEqual(0xFF0000, layout.Paragraphs[0].Runs[0].ColorArgb, "цвет рана сохранён");
        }

        private static void TestOcrLeftAligned()
        {
            // Рваный справа абзац (строки не достают до правого поля) -> по левому краю,
            // а не насильно по ширине. Полная строка ниже задаёт правое поле (100).
            var words = new List<PdfWord>
            {
                W("aaaa", 0, 100, 60, 8),   // абзац 1, строка 1: Right 60
                W("bbbb", 0, 88, 55, 8),    // абзац 1, строка 2: Right 55 (обе рваные)
                W("cccc", 0, 50, 100, 8)    // абзац 2: полная строка Right 100 (задаёт поле)
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(2, layout.Paragraphs.Count, "два абзаца");
            AssertEqual(OcrAlignment.Left, layout.Paragraphs[0].Alignment, "рваный абзац — по левому краю");
        }

        private static void TestOcrCentered()
        {
            // Короткая центрированная строка («7») над телом -> отдельный центрированный абзац.
            var words = new List<PdfWord>
            {
                W("7", 48, 180, 4, 8),      // центр: слева 48, справа 48 от полей 0..100
                W("Тело", 0, 150, 100, 8),  // полная строка (Left 0, Right 100) — задаёт поля
                W("строка", 0, 138, 100, 8)
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(2, layout.Paragraphs.Count, "два абзаца");
            AssertEqual(OcrAlignment.Center, layout.Paragraphs[0].Alignment, "номер страницы по центру");
            AssertEqual(OcrAlignment.Justify, layout.Paragraphs[1].Alignment, "тело по ширине");
        }

        private static void TestIsCenteredPredicate()
        {
            // Узкая центрированная (номер страницы): большие симметричные зазоры.
            AssertTrue(OcrLayout.IsCentered(48, 48, 8, 100), "узкая по центру");
            // ШИРОКАЯ центрированная строка титула: зазоры малы в долях ширины (~5%), но симметричны —
            // прежний порог 12% ширины её терял, новый (доля кегля) ловит.
            AssertTrue(OcrLayout.IsCentered(31, 22, 16, 453), "широкая строка титула по центру");
            // Красная строка: упирается в правое поле (правый зазор ≈ 0) — не центр.
            AssertTrue(!OcrLayout.IsCentered(35, 0, 16, 453), "красная строка — не центр");
            // Рваная левая: стоит у левого поля (левый зазор ≈ 0) — не центр.
            AssertTrue(!OcrLayout.IsCentered(0, 40, 10, 100), "рваная левая — не центр");
            // Правое выравнивание: у правого поля (правый зазор ≈ 0), большой левый — не центр.
            AssertTrue(!OcrLayout.IsCentered(200, 0, 16, 453), "правое выравнивание — не центр");
            // Несимметричная (левый много больше правого) — не центр, даже если оба > порога.
            AssertTrue(!OcrLayout.IsCentered(120, 20, 16, 453), "асимметрия — не центр");
            // Красная строка, чуть НЕ ДОТЯНУВШАЯ до правого поля. Абсолютная разница зазоров
            // (21 pt) укладывалась в долю ширины A4, и абзац разрывался надвое, а его первая
            // строка ещё и центрировалась. Соизмеримость зазоров это ловит: 14 против 35.
            AssertTrue(!OcrLayout.IsCentered(35, 14, 16, 339), "красная строка в двух пунктах от поля — не центр");
            // А настоящий титул, сбитый с оси, центром остаётся: зазоры соизмеримы.
            AssertTrue(OcrLayout.IsCentered(40, 70, 16, 453), "титул сбит с оси, но зазоры соизмеримы");
        }

        private static void TestOcrCenteredBlock()
        {
            // Типичный центрированный титул документа: ШИРОКИЕ строки дотягиваются до обоих полей
            // (внешне как justified) и лишь короткий хвост «Sh» явно центрирован. Все строки соосны
            // (midX=50) -> ВСЕ распознаются центрированными по общей оси. Каждая — своим абзацем
            // (исходная разбивка сохранена: Word не сольёт их и не перевёрстывает), а не «широкие=тело
            // + сирота-хвост».
            var words = new List<PdfWord>
            {
                W("Wideoneaa", 2, 200, 96, 8),  // Left2 Right98 mid50: дотягивается до полей (не «плавает»)
                W("Widetwoaa", 2, 188, 96, 8),  // Left2 Right98 mid50: тоже широкая
                W("Sh",       35, 176, 30, 8),  // Left35 Right65 mid50: короткая центрированная — доказывает ось
                W("Bodyoneaa", 2, 150, 96, 8),  // тело: полные строки (задают поля, justified), большой зазор
                W("Bodytwoaa", 2, 138, 96, 8),
                W("Bodythree", 2, 126, 96, 8)
            };
            List<OcrParagraph> paras = OcrLayout.Analyze(words).Paragraphs;
            AssertEqual(4, paras.Count, "3 центрированные строки титула (каждая своим абзацем) + тело");
            AssertEqual(OcrAlignment.Center, paras[0].Alignment, "широкая строка титула — по центру (по общей оси)");
            AssertEqual(OcrAlignment.Center, paras[1].Alignment, "вторая широкая строка титула — по центру");
            AssertEqual(OcrAlignment.Center, paras[2].Alignment, "короткий хвост — по центру");
            AssertEqual("Wideoneaa", paras[0].Text, "исходная строка сохранена отдельным абзацем (не слита)");
            AssertEqual(OcrAlignment.Justify, paras[3].Alignment, "тело — выключка, не центр (нет короткой центрированной строки)");
        }

        private static void TestOcrGlueFragments()
        {
            // Почти соприкасающиеся куски одного токена (мизерный зазор < 0.08 кегля) склеиваем
            // без пробела; между словами обычный зазор — пробел.
            var words = new List<PdfWord>
            {
                W("м", 0, 0, 8, 10),      // Right 8
                W("и", 8.5, 0, 8, 10),    // зазор 0.5 (0.05 кегля) -> склеить
                W("р", 17, 0, 8, 10),     // зазор 0.5 (0.05 кегля) -> склеить  => «мир»
                W("тут", 30, 0, 20, 10)   // зазор 5 (0.5 кегля) -> пробел      => «мир тут»
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(1, p.Count, "одна строка — один абзац");
            AssertEqual("мир тут", p[0], "фрагменты склеены, между словами пробел");
        }

        private static void TestOcrNarrowSpaceKept()
        {
            // Регресс: в узких шрифтах (Calibri Light) настоящий межсловный зазор ≈ 0.18 кегля.
            // Прежний порог 0.2 ронял такой пробел и слеплял слова («СЛОВОСЛОВО»);
            // 0.08 сохраняет пробел. Зазор здесь 0.15 кегля — между старым и новым порогом.
            var words = new List<PdfWord>
            {
                W("СЛОВО", 0, 0, 60, 16),      // Right 60
                W("ТЕКСТ", 62.4, 0, 40, 16)        // зазор 2.4 = 0.15 кегля -> пробел
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(1, p.Count, "одна строка — один абзац");
            AssertEqual("СЛОВО ТЕКСТ", p[0], "узкий настоящий пробел сохранён (не склейка)");
        }

        private static void TestOcrThinDashStaysOnLine()
        {
            // Тонкое тире с крошечной рамкой и центром выше базовой линии текста не должно
            // отрываться в отдельную строку/абзац (перекрытие рамок, а не расстояние центров).
            var words = new List<PdfWord>
            {
                W("quarantine", 0, 0, 50, 8),   // [0..8], центр 4
                W("—", 55, 4.3, 8, 0.6),         // [4.3..4.9], центр 4.6 — выше на 0.6
                W("добавляет", 70, 0, 50, 8)     // [0..8], центр 4
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(1, p.Count, "тире не отрывается — одна строка, один абзац");
            AssertEqual("quarantine — добавляет", p[0], "тире осталось между словами");
        }

        private static void TestOcrHyphenation()
        {
            // Строка кончается «wo-», следующая «rld» -> склеить в «world».
            var words = new List<PdfWord>
            {
                W("hello", 0, 90, 30, 10),
                W("wo-", 35, 90, 20, 10),
                W("rld", 0, 75, 20, 10)
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(1, p.Count, "один абзац");
            AssertEqual("hello world", p[0], "дефис-перенос склеен");
        }

        private static void TestOcrHyphenCyrillicKept()
        {
            // Дефис после кириллической буквы на конце строки — составное слово
            // («информационно-коммуникационных»): официальные PDF приходят из Word/LibreOffice,
            // которые по умолчанию слова не переносят, поэтому дефис сохраняется.
            var words = new List<PdfWord>
            {
                W("шёл", 0, 90, 20, 10),
                W("информационно-", 25, 90, 70, 10),   // строка почти полная — не «жёсткий перевод»
                W("коммуникационных", 0, 75, 95, 10)
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(1, p.Count, "один абзац");
            AssertEqual("шёл информационно-коммуникационных", p[0], "кириллический дефис сохранён");
        }

        private static void TestOcrHardLineBreak()
        {
            // «отдела» свободно влезало после «руководитель» — автор сам начал новую
            // строку (подпись), это разные абзацы; иначе Word перевёрстывал бы их произвольно.
            var words = new List<PdfWord>
            {
                W("полная строка на всю ширину колонки текста", 0, 130, 200, 10),
                W("Заместитель", 0, 90, 40, 10),
                W("руководитель", 45, 90, 35, 10),    // Right 80 — заполнено 40% колонки
                W("отдела", 0, 75, 50, 10)            // влезло бы: 80 + 50 + 7.5 <= 200
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(3, p.Count, "полная строка и две строки подписи — три абзаца");
            AssertEqual("Заместитель руководитель", p[1], "первая строка подписи своим абзацем");
            AssertEqual("отдела", p[2], "вторая строка подписи своим абзацем");
        }

        private static void TestOcrSignatureHardBreak()
        {
            // Подпись слева (две короткие строки) и Ф.И.О. справа: рамка левой колонки по её
            // контенту «заполнена», но ДОСТУПНОЕ место тянется до правой колонки — перевод
            // строки умышленный, строки подписи не склеиваются.
            var words = new List<PdfWord>
            {
                W("Руководитель проектной группы", 0, 90, 150, 10),
                W("испытательного отдела", 0, 75, 120, 10),
                W("Иванов", 300, 75, 60, 10)
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(3, p.Count, "две строки подписи и Ф.И.О. — три абзаца");
            AssertEqual("Руководитель проектной группы", p[0], "первая строка подписи");
            AssertEqual("испытательного отдела", p[1], "вторая строка подписи");
            AssertEqual("Иванов", p[2], "правая колонка после левой");
        }

        private static void TestOcrFootnoteDigitSuper()
        {
            // Ink-бокс и подъём цифры-сноски ненадёжны (в некоторых шрифтах она на базовой и
            // почти в высоту строчных): маркер узнаётся по мелкому КЕГЛЮ числа («250²» — 8 pt
            // против 14 pt). Мелкое слово из БУКВ на базовой — не скрипт.
            PdfWord Sized(string text, double left, double bottom, double width, double height, double sizePt)
            {
                PdfWord w = W(text, left, bottom, width, height);
                w.FontSizePt = sizePt;
                return w;
            }
            var words = new List<PdfWord>
            {
                Sized("№250", 0, 0, 30, 7, 14),
                Sized("2", 31, 0, 4, 6, 8),      // кегль 8 против 14, база 0 — всё равно Super
                Sized("далее", 40, 0, 30, 7, 14),
                Sized("ещё", 75, 0, 20, 7, 14),
                Sized("сн", 100, 0, 8, 6, 8)     // мелкое буквенное на базовой — НЕ скрипт
            };
            List<OcrRun> runs = OcrLayout.Analyze(words).Paragraphs[0].Runs;
            AssertEqual(4, runs.Count, "цифра и мелкое слово — отдельные раны (кегль различен)");
            AssertTrue(runs[1].Super && !runs[1].Sub, "«2» — надстрочный маркер сноски");
            AssertEqual("2", runs[1].Text.Trim(), "текст маркера (межсловный пробел уходит в ран)");
            AssertTrue(!runs[3].Super && !runs[3].Sub, "мелкое буквенное слово — не скрипт");
        }

        private static void TestOcrRedIndentNotCentered()
        {
            // Одиночная строка, начатая точно с красной строки документа и случайно почти
            // симметричная (36/42), — не центр; настоящая симметричная строка — центр.
            var words = new List<PdfWord>
            {
                W("первый абзац начинается с красной строки и идёт", 36, 150, 164, 10),
                W("хвост первой строки прижат влево до правого поля", 0, 135, 200, 10),
                W("второй абзац начинается с красной строки и идёт", 36, 120, 164, 10),
                W("хвост второй строки прижат влево до правого поля", 0, 105, 200, 10),
                W("третий абзац начинается с красной строки и идёт", 36, 90, 164, 10),
                W("хвост третьей строки прижат влево до правого", 0, 75, 200, 10),
                W("одиночная строка пункта списка", 36, 60, 122, 10),  // отступы 36/42 — не центр
                W("настоящий центр", 60, 25, 80, 10)                    // 60/60 — центр
            };
            List<OcrParagraph> paras = OcrLayout.Analyze(words).Paragraphs;
            OcrParagraph fake = null, real = null;
            foreach (OcrParagraph p in paras)
            {
                if (p.Text.StartsWith("одиночная")) fake = p;
                if (p.Text == "настоящий центр") real = p;
            }
            AssertTrue(fake != null && fake.Alignment != OcrAlignment.Center,
                "строка с красной строкой не центрирована");
            AssertTrue(real != null && real.Alignment == OcrAlignment.Center,
                "настоящая симметричная строка центрирована");
        }

        private static void TestDocxGapMath()
        {
            // Межблочные интервалы: типичный зазор — нижняя медиана положительных; интервал —
            // лишек сверх типичного с порогом (6 pt) и капом (120 pt).
            var items = new List<PageBlocks.PageItem>
            {
                new PageBlocks.PageItem { Top = 100, Bottom = 90 },
                new PageBlocks.PageItem { Top = 83, Bottom = 70 },   // зазор 7
                new PageBlocks.PageItem { Top = 63, Bottom = 50 },   // зазор 7
                new PageBlocks.PageItem { Top = 20, Bottom = 10 }    // зазор 30
            };
            AssertEqual(7.0, PageBlocks.TypicalItemGap(items), "типичный зазор — медиана {7,7,30}");
            AssertEqual(0.0, PageBlocks.ExtraGapPt(10, 7), "лишек 3 меньше порога — без интервала");
            AssertEqual(23.0, PageBlocks.ExtraGapPt(30, 7), "лишек 23 — интервал 23 pt");
            // Кап 400: пропускает нижний блок у нижнего края почти пустой страницы,
            // а от переливов страхует демпфер FitSpacingToPages.
            AssertEqual(400.0, PageBlocks.ExtraGapPt(900, 7), "кап 400 pt");
        }

        private static void TestWindowBoundsRoundTrip()
        {
            var r = new System.Drawing.Rectangle(120, -40, 800, 660); // отрицательный Y — левый монитор
            System.Drawing.Rectangle back; bool max;
            AssertTrue(WindowPlacement.TryParse(WindowPlacement.Format(r, false), out back, out max), "разбор своей сериализации");
            AssertEqual(r, back, "round-trip прямоугольника");
            AssertTrue(!max, "флаг развёрнуто = false");
            AssertTrue(WindowPlacement.TryParse(WindowPlacement.Format(r, true), out back, out max) && max, "флаг развёрнуто = true");
            System.Drawing.Rectangle junk; bool jm;
            AssertTrue(!WindowPlacement.TryParse(null, out junk, out jm), "null — не разбирается");
            AssertTrue(!WindowPlacement.TryParse("1,2,3", out junk, out jm), "мало полей — брак");
            AssertTrue(!WindowPlacement.TryParse("1,2,0,10,0", out junk, out jm), "нулевая ширина — брак");
            AssertTrue(!WindowPlacement.TryParse("a,b,c,d,e", out junk, out jm), "не числа — брак");
        }

        private static void TestClampToWorkingArea()
        {
            var screen = new System.Drawing.Rectangle(0, 0, 1920, 1040);
            var one = new[] { screen };
            var min = new System.Drawing.Size(700, 560);
            var inside = new System.Drawing.Rectangle(100, 100, 800, 600);
            AssertEqual(inside, WindowPlacement.ClampToWorkingArea(inside, one, min), "внутри экрана — без изменений");
            var offRB = WindowPlacement.ClampToWorkingArea(new System.Drawing.Rectangle(1900, 1000, 800, 600), one, min);
            AssertTrue(offRB.Right <= screen.Right && offRB.Bottom <= screen.Bottom, "не за правым/нижним краем");
            AssertEqual(800, offRB.Width, "ширина сохранена при сдвиге внутрь");
            var offAll = WindowPlacement.ClampToWorkingArea(new System.Drawing.Rectangle(-5000, -5000, 800, 600), one, min);
            AssertTrue(offAll.X >= screen.Left && offAll.Y >= screen.Top, "ушедшее целиком за экран возвращается");
            var tiny = WindowPlacement.ClampToWorkingArea(new System.Drawing.Rectangle(0, 0, 100, 100), one, min);
            AssertTrue(tiny.Width >= min.Width && tiny.Height >= min.Height, "не меньше MinimumSize");
            var huge = WindowPlacement.ClampToWorkingArea(new System.Drawing.Rectangle(0, 0, 5000, 5000), one, min);
            AssertTrue(huge.Width <= screen.Width && huge.Height <= screen.Height, "не больше экрана");
            var two = new[] { screen, new System.Drawing.Rectangle(1920, 0, 1920, 1040) };
            var onSecond = WindowPlacement.ClampToWorkingArea(new System.Drawing.Rectangle(2200, 100, 800, 600), two, min);
            AssertEqual(new System.Drawing.Rectangle(2200, 100, 800, 600), onSecond, "на втором мониторе — без изменений");
        }

        private static void TestMoveActiveLast()
        {
            // Активный снапшот уезжает в конец (его окно показывается последним и остаётся
            // сверху после смены языка); порядок остальных сохраняется; без активного — no-op.
            ShellContext.ToolSnapshot S(string key, bool active)
            {
                return new ShellContext.ToolSnapshot { Key = key, WasActive = active };
            }
            var snap = new List<ShellContext.ToolSnapshot> { S("a", false), S("b", true), S("c", false) };
            ShellContext.MoveActiveLast(snap);
            AssertEqual("a|c|b", snap[0].Key + "|" + snap[1].Key + "|" + snap[2].Key, "активный — последним");
            var none = new List<ShellContext.ToolSnapshot> { S("a", false), S("b", false) };
            ShellContext.MoveActiveLast(none);
            AssertEqual("a|b", none[0].Key + "|" + none[1].Key, "без активного порядок не меняется");
        }

        private static void TestOcrWideGapSplit()
        {
            // Куски одной строки, разделённые гигантским зазором (20 кеглей), — разные зоны
            // (поля слева, боковая метка справа): в потоке рвутся на абзацы, в ячейке — нет.
            var words = new List<PdfWord>
            {
                W("На", 0, 90, 20, 10),
                W("№", 25, 90, 10, 10),
                W("02.07.2026", 40, 90, 60, 10),
                W("REG~MARK", 300, 90, 60, 10)
            };
            List<OcrParagraph> flow = OcrLayout.Analyze(words).Paragraphs;
            AssertEqual(2, flow.Count, "в потоке — два абзаца-зоны");
            AssertEqual("На № 02.07.2026", flow[0].Text, "левая зона");
            AssertEqual("REG~MARK", flow[1].Text, "правая зона");
            List<OcrParagraph> cell = OcrLayout.Analyze(words, false).Paragraphs;
            AssertEqual(1, cell.Count, "в ячейке строка едина");
            AssertEqual("На № 02.07.2026 REG~MARK", cell[0].Text, "текст ячейки не порван");
        }

        private static void TestCoalesceRowBands()
        {
            // Два одиночных абзаца одной «строки» с широким каналом — side-by-side полоса;
            // абзац строкой ниже полосой не становится.
            PageBlocks.PageItem P(double left, double right, double top, double bottom)
            {
                var b = new PageBlocks.Block
                {
                    Paragraph = new OcrParagraph(),
                    Left = left,
                    Right = right,
                    Top = top,
                    Bottom = bottom
                };
                return new PageBlocks.PageItem { Single = b, Top = top, Bottom = bottom };
            }
            var items = new List<PageBlocks.PageItem>
            {
                P(0, 100, 100, 90),    // левая зона строки
                P(200, 300, 99, 89),   // правая зона той же строки (канал 100)
                P(0, 300, 70, 60)      // следующая строка — отдельно
            };
            List<PageBlocks.PageItem> result = PageBlocks.CoalesceRowBands(items);
            AssertEqual(2, result.Count, "полоса + одиночный");
            AssertTrue(result[0].IsBand && result[0].Columns.Count == 2, "первая пара — полоса 1×2");
            AssertEqual(200.0, result[0].ColLeft[1], "левый край правой колонки");
            AssertTrue(!result[1].IsBand, "нижний абзац остался одиночным");
        }

        private static void TestAnchorIndents()
        {
            double fli, li;
            WordDocxWriter.AnchorIndents(303, 510, 0, out fli, out li);
            AssertTrue(fli == 0 && li == 303, "глубокий старт — позиция колонки (боковая метка)");
            WordDocxWriter.AnchorIndents(35, 510, 34.7, out fli, out li);
            AssertTrue(fli == 34.7 && li == 0, "фактический отступ ≈ документному — красная строка");
            WordDocxWriter.AnchorIndents(1, 510, 34.7, out fli, out li);
            AssertTrue(fli == 0 && li == 0, "абзац с края — без ложной красной строки (сноски)");
            WordDocxWriter.AnchorIndents(35, 510, 0, out fli, out li);
            AssertTrue(fli == 35 && li == 0, "документ без общего отступа — отступ по факту");
            WordDocxWriter.AnchorIndents(4, 510, 0, out fli, out li);
            AssertTrue(fli == 0 && li == 0, "шум измерения — не отступ");
        }

        private static void TestOcrHyphenMixedKept()
        {
            // Кириллица с ЛЮБОЙ стороны разрыва — составное слово, дефис остаётся.
            var words = new List<PdfWord>
            {
                W("шла", 0, 90, 20, 10),
                W("Word-", 25, 90, 65, 10),
                W("форма", 0, 75, 90, 10)
            };
            List<string> p = OcrLayout.ToParagraphs(words);
            AssertEqual(1, p.Count, "один абзац");
            AssertEqual("шла Word-форма", p[0], "дефис лат+кириллица сохранён");
        }

        private static void TestCoveredByLowImage()
        {
            var form = new PdfTextExtract.RectPt { Left = 66, Bottom = 570, Right = 260, Top = 577 };
            var stampImg = new List<PdfTextExtract.RectPt>
            {
                new PdfTextExtract.RectPt { Left = 94, Bottom = 567, Right = 254, Top = 581 } // наложение H=14
            };
            AssertTrue(PdfTextExtract.CoveredByLowImage(form, stampImg), "форма под наложением скрыта");
            var tallImg = new List<PdfTextExtract.RectPt>
            {
                new PdfTextExtract.RectPt { Left = 60, Bottom = 400, Right = 300, Top = 590 } // схема H=190
            };
            AssertTrue(!PdfTextExtract.CoveredByLowImage(form, tallImg), "высокая картинка не поглощает текст");
            var aside = new List<PdfTextExtract.RectPt>
            {
                new PdfTextExtract.RectPt { Left = 200, Bottom = 567, Right = 420, Top = 581 } // сбоку, перекрытие < 50%
            };
            AssertTrue(!PdfTextExtract.CoveredByLowImage(form, aside), "слабое перекрытие — не наложение");
            AssertTrue(PdfTextExtract.CoveredByAnyRect(
                new PdfTextExtract.RectPt { Left = 100, Bottom = 570, Right = 120, Top = 578 }, stampImg),
                "слово на подложке — накрыто наполовину по обеим осям");
            AssertTrue(!PdfTextExtract.CoveredByAnyRect(
                new PdfTextExtract.RectPt { Left = 94, Bottom = 581, Right = 104, Top = 585 }, stampImg),
                "касание кромки картинки — не подложка (белая метка у наложения)");
        }

        private static void TestOcrEmpty()
        {
            AssertEqual(0, OcrLayout.ToParagraphs(new List<PdfWord>()).Count, "пусто -> нет абзацев");
            AssertEqual(0, OcrLayout.ToParagraphs(null).Count, "null -> нет абзацев");
        }

        // ---------- ListMarker (маркеры списка) ----------

        private static void TestListMarkerNumbered()
        {
            ListMarker.Result a = ListMarker.Detect("1. Внести изменения");
            AssertEqual((int)ListKind.Numbered, (int)a.Kind, "«1.» -> нумерованный");
            AssertEqual(1, a.Number, "номер 1");
            AssertEqual("Внести изменения", "1. Внести изменения".Substring(a.ContentStart), "содержимое без маркера");

            ListMarker.Result b = ListMarker.Detect("12) пункт");
            AssertEqual((int)ListKind.Numbered, (int)b.Kind, "«12)» -> нумерованный");
            AssertEqual(12, b.Number, "номер 12");
            AssertEqual("пункт", "12) пункт".Substring(b.ContentStart), "содержимое после «12)»");
        }

        private static void TestListMarkerBulleted()
        {
            ListMarker.Result a = ListMarker.Detect("• первый");
            AssertEqual((int)ListKind.Bulleted, (int)a.Kind, "«•» -> маркированный");
            AssertEqual("первый", "• первый".Substring(a.ContentStart), "содержимое без буллета");

            ListMarker.Result b = ListMarker.Detect("— тире-буллет");
            AssertEqual((int)ListKind.Bulleted, (int)b.Kind, "«—» -> маркированный");
            AssertEqual("тире-буллет", "— тире-буллет".Substring(b.ContentStart), "содержимое без тире");
        }

        private static void TestListMarkerNegatives()
        {
            AssertEqual((int)ListKind.None, (int)ListMarker.Detect("2025 год отчёта").Kind, "год не маркер");
            AssertEqual((int)ListKind.None, (int)ListMarker.Detect("12.5% роста").Kind, "проценты не маркер");
            AssertEqual((int)ListKind.None, (int)ListMarker.Detect("1.без пробела").Kind, "без пробела после точки — не маркер");
            AssertEqual((int)ListKind.None, (int)ListMarker.Detect("•безпробела").Kind, "буллет без пробела — не маркер");
            AssertEqual((int)ListKind.None, (int)ListMarker.Detect("Обычный текст").Kind, "обычный текст — не список");
            AssertEqual((int)ListKind.None, (int)ListMarker.Detect("").Kind, "пусто — не список");
        }

        private static void TestOcrNumberedList()
        {
            // Плотный одностроковый список (равный интервал, левый край, без отступа): без деления
            // по маркеру строки слиплись бы в один абзац. Ожидаем два пункта, ListKind=Numbered.
            var words = new List<PdfWord>
            {
                W("1.", 0, 100, 8, 10), W("Первый", 12, 100, 40, 10),
                W("2.", 0, 88, 8, 10),  W("Второй", 12, 88, 40, 10)
            };
            List<OcrParagraph> ps = OcrLayout.Analyze(words).Paragraphs;
            AssertEqual(2, ps.Count, "два пункта -> два абзаца");
            AssertEqual((int)ListKind.Numbered, (int)ps[0].ListKind, "пункт 1 — нумерованный");
            AssertEqual(1, ps[0].ListNumber, "номер первого пункта");
            AssertEqual((int)ListKind.Numbered, (int)ps[1].ListKind, "пункт 2 — нумерованный");
            AssertEqual(2, ps[1].ListNumber, "номер второго пункта");
            AssertEqual("Первый", ps[0].Text.Substring(ps[0].ListContentStart), "содержимое 1 без маркера");
            AssertEqual("Второй", ps[1].Text.Substring(ps[1].ListContentStart), "содержимое 2 без маркера");
        }

        private static void TestOcrBulletedList()
        {
            var words = new List<PdfWord>
            {
                W("•", 0, 100, 6, 10), W("яблоко", 10, 100, 40, 10),
                W("•", 0, 88, 6, 10),  W("груша", 10, 88, 40, 10)
            };
            List<OcrParagraph> ps = OcrLayout.Analyze(words).Paragraphs;
            AssertEqual(2, ps.Count, "два буллета -> два абзаца");
            AssertEqual((int)ListKind.Bulleted, (int)ps[0].ListKind, "буллет 1");
            AssertEqual((int)ListKind.Bulleted, (int)ps[1].ListKind, "буллет 2");
            AssertEqual("яблоко", ps[0].Text.Substring(ps[0].ListContentStart), "содержимое буллета 1 без маркера");
            AssertEqual("груша", ps[1].Text.Substring(ps[1].ListContentStart), "содержимое буллета 2 без маркера");
        }

        // ---------- StampDetector (текстовый штамп) ----------

        /// <summary>Синтетический текстовый штамп (4 строки) в левом нижнем углу + одно слово тела выше.</summary>
        private static List<PdfWord> StampWords()
        {
            return new List<PdfWord>
            {
                // строка-заголовок
                W("Документ", 100, 200, 50, 9), W("подписан", 155, 200, 50, 9),
                W("электронной", 210, 200, 60, 9), W("подписью", 275, 200, 50, 9),
                // сертификат
                W("Сертификат:", 100, 186, 60, 9), W("7f1224cd", 165, 186, 60, 9),
                // владелец
                W("Владелец", 100, 172, 50, 9), W("Иванов", 155, 172, 40, 9), W("Иван", 200, 172, 30, 9),
                // действителен
                W("Действителен", 100, 158, 70, 9), W("с", 175, 158, 8, 9), W("01.01.2025", 188, 158, 55, 9),
                W("по", 248, 158, 15, 9), W("01.01.2026", 268, 158, 55, 9),
                // слово тела заметно выше штампа — в область попасть не должно
                W("Обычныйтекст", 100, 500, 80, 9)
            };
        }

        private static void TestStampDetected()
        {
            List<PdfWord> words = StampWords();
            StampRegion s = StampDetector.Detect(words, 595, 842);
            AssertTrue(s != null, "штамп распознан");
            AssertEqual(14, s.Words.Count, "в область попали все 14 слов штампа (без слова тела)");
            bool bodyIn = false;
            foreach (PdfWord w in s.Words) if (w.Text == "Обычныйтекст") bodyIn = true;
            AssertTrue(!bodyIn, "слово тела вне полосы штампа не захвачено");
            AssertTrue(s.Left <= 100 && s.Right >= 323 && s.Bottom <= 158 && s.Top >= 209, "рамка охватывает все строки штампа");
        }

        private static void TestStampMissingAnchor()
        {
            // Без слова «Действителен» (заменено нейтральным) — одного опорного слова нет.
            List<PdfWord> words = StampWords();
            for (int i = 0; i < words.Count; i++)
                if (words[i].Text == "Действителен") words[i].Text = "Выдан";
            AssertTrue(StampDetector.Detect(words, 595, 842) == null, "нет всех четырёх опорных слов -> не штамп");
        }

        private static void TestStampScatteredRejected()
        {
            // Все четыре слова есть, но раскиданы по всей странице (обычная проза) — не компактно.
            var words = new List<PdfWord>
            {
                W("подписан", 50, 800, 50, 9),
                W("Сертификат", 50, 600, 60, 9),
                W("Владелец", 50, 400, 50, 9),
                W("Действителен", 50, 100, 70, 9),
                W("прочее1", 300, 700, 40, 9), W("прочее2", 300, 300, 40, 9)
            };
            AssertTrue(StampDetector.Detect(words, 595, 842) == null, "разбросанные опорные слова -> не штамп");
        }

        // ---------- Loc (локализация) ----------

        private static void TestLocCatalogComplete()
        {
            int n = 0;
            foreach (string key in Loc.Keys)
            {
                string[] p = Loc.Pair(key);
                AssertTrue(p != null && p.Length == 2, "пара для «" + key + "»");
                AssertTrue(!string.IsNullOrEmpty(p[0]), "ru пусто у «" + key + "»");
                AssertTrue(!string.IsNullOrEmpty(p[1]), "en пусто у «" + key + "»");
                n++;
            }
            AssertTrue(n > 100, "каталог непустой (ключей: " + n + ")");
        }

        /// <summary>
        /// Каждый ключ, который код запрашивает у каталога, обязан в нём быть. Промах виден
        /// только глазами — Loc.T возвращает сам ключ, и в интерфейсе появляется «split.btn.open»
        /// вместо надписи. Это единственная защита от опечатки при переименовании ключей, а
        /// переименования случаются: в 1.17.9 строки одного открытого документа переехали из
        /// «split.*» в «common.*», а шесть операций — в «ops.*».
        ///
        /// Исходники ищем от каталога сборки: тесты компилируются вместе с src и всегда лежат
        /// внутри репозитория. Не нашли — это ОШИБКА, а не повод промолчать: проверка, которая
        /// сама себя отключает, зеленеет ничего не проверив.
        /// </summary>
        private static void TestLocKeysUsedInCodeExist()
        {
            string dir = SourceDir();
            AssertTrue(dir != null, "каталог src не найден рядом с тестами");
            var missing = new List<string>();
            int checked_ = 0;
            var re = new System.Text.RegularExpressions.Regex(
                "Loc\\.T\\(\"([^\"]+)\"\\)", System.Text.RegularExpressions.RegexOptions.Compiled);
            foreach (string file in Directory.GetFiles(dir, "*.cs"))
                foreach (System.Text.RegularExpressions.Match m in re.Matches(File.ReadAllText(file)))
                {
                    string key = m.Groups[1].Value;
                    checked_++;
                    if (Loc.Pair(key) == null)
                        missing.Add(Path.GetFileName(file) + ": " + key);
                }
            AssertTrue(checked_ > 200, "ключи в коде найдены (проверено: " + checked_ + ")");
            AssertTrue(missing.Count == 0, "ключей нет в каталоге: " + string.Join(", ", missing.ToArray()));
        }

        /// <summary>Каталог src репозитория относительно каталога запуска тестов (bin или bin\x86).</summary>
        private static string SourceDir()
        {
            string at = AppDomain.CurrentDomain.BaseDirectory;
            for (int up = 0; up < 5 && at != null; up++)
            {
                string candidate = Path.Combine(at, "src");
                if (Directory.Exists(candidate) && File.Exists(Path.Combine(candidate, "Loc.cs")))
                    return candidate;
                DirectoryInfo parent = Directory.GetParent(at);
                at = parent == null ? null : parent.FullName;
            }
            return null;
        }

        private static void TestLocPlaceholders()
        {
            // {0},{1}… у ru и en должны совпадать — иначе string.Format кинет на одном из языков.
            foreach (string key in Loc.Keys)
            {
                string[] p = Loc.Pair(key);
                AssertEqual(Placeholders(p[0]), Placeholders(p[1]), "плейсхолдеры {N} расходятся у «" + key + "»");
            }
        }

        /// <summary>
        /// Подсказки про перетаскивание: цель броска везде названа «окно программы»
        /// (в английском — the program window), а деепричастие «перетащив» не используется.
        /// По терминологии Microsoft в инструкциях стоит повелительное «перетащите».
        /// Перетаскивание миниатюр ВНУТРИ сетки к окну не относится и под проверку не идёт.
        /// </summary>
        private static void TestLocDragHints()
        {
            int windowTargets = 0;
            foreach (string key in Loc.Keys)
            {
                string[] p = Loc.Pair(key);
                AssertTrue(p[0].IndexOf("перетащив", StringComparison.Ordinal) < 0,
                    "деепричастие «перетащив» осталось у «" + key + "»");
                if (p[0].IndexOf(" в окно", StringComparison.Ordinal) < 0)
                    continue;
                AssertTrue(p[0].Contains("в окно программы"), "цель броска не названа у «" + key + "» (ru)");
                AssertTrue(p[1].Contains("program window"), "цель броска не названа у «" + key + "» (en)");
                windowTargets++;
            }
            AssertTrue(windowTargets >= 12, "подсказок про перетаскивание в окно найдено: " + windowTargets);
        }

        private static string Placeholders(string s)
        {
            var set = new SortedSet<int>();
            foreach (System.Text.RegularExpressions.Match x in
                     System.Text.RegularExpressions.Regex.Matches(s, "\\{(\\d+)\\}"))
                set.Add(int.Parse(x.Groups[1].Value));
            return string.Join(",", set);
        }

        private static void TestLocInit()
        {
            Lang saved = Loc.Current;
            try
            {
                Loc.Init(Lang.En); AssertEqual((int)Lang.En, (int)Loc.Current, "Init En");
                Loc.Init(Lang.Ru); AssertEqual((int)Lang.Ru, (int)Loc.Current, "Init Ru");
                AssertEqual((int)Lang.En, (int)Loc.Parse("en"), "Parse en");
                AssertEqual((int)Lang.En, (int)Loc.Parse("EN"), "Parse EN");
                AssertEqual((int)Lang.Ru, (int)Loc.Parse("ru"), "Parse ru");
                AssertEqual((int)Lang.Ru, (int)Loc.Parse("xx"), "Parse неизвестный -> ru");
                AssertEqual("en", Loc.Code(Lang.En), "Code En");
                AssertEqual("ru", Loc.Code(Lang.Ru), "Code Ru");
                // Дефолт первого запуска по системной локали: русская → ru, прочие → en.
                AssertEqual((int)Lang.Ru, (int)Loc.DefaultForCulture("ru-RU"), "локаль ru-RU -> ru");
                AssertEqual((int)Lang.Ru, (int)Loc.DefaultForCulture("RU"), "локаль RU -> ru");
                AssertEqual((int)Lang.En, (int)Loc.DefaultForCulture("en-US"), "локаль en-US -> en");
                AssertEqual((int)Lang.En, (int)Loc.DefaultForCulture("de-DE"), "локаль de-DE -> en");
                AssertEqual((int)Lang.En, (int)Loc.DefaultForCulture(null), "нет локали -> en");
                AssertEqual((int)Lang.En, (int)Loc.DefaultForCulture(""), "пустая локаль -> en");
            }
            finally { Loc.Init(saved); }
        }

        private static void TestNoCyrillicInEnglishForms()
        {
            Lang saved = Loc.Current;
            var offenders = new List<string>();
            // «Язык / Language» — намеренно двуязычный пункт меню, кириллица там ожидаема.
            // Название банка в пунктах доната — тоже: это реквизит, а не переводимый текст.
            var whitelist = new HashSet<string>
            {
                Loc.Pair("menu.language")[1],
                AboutForm.DonationBank,
                // Подпись глобуса на стартовом экране двуязычна намеренно: её читают как раз
                // тогда, когда язык интерфейса непонятен или выбран по ошибке.
                Loc.Pair("lang.tooltip")[1]
            };
            var th = new System.Threading.Thread(delegate()
            {
                Loc.Init(Lang.En);
                System.Windows.Forms.Form[] forms;
                try
                {
                    // Тул-формы — с непустым showHub, чтобы кнопка «Главная» (условная) создавалась
                    // и попадала в проверку (иначе захардкоженный перевод в ней остался бы незамеченным).
                    Action back = delegate { };
                    forms = new System.Windows.Forms.Form[]
                    {
                        new MainForm(back), new PdfMergeForm(back), new PdfSplitForm(back),
                        new OcrForm(back), new PdfOpsForm(back), new StartForm(),
                        new AboutForm(), new StatsForm(), new SettingsForm(), new HistoryForm()
                    };
                }
                catch (Exception ex) { offenders.Add("ctor: " + ex.Message); return; }
                foreach (System.Windows.Forms.Form f in forms)
                {
                    CheckCyrillic(f.Text, "Form.Text", offenders, whitelist);
                    WalkControls(f, offenders, whitelist);
                    if (f.MainMenuStrip != null)
                        foreach (System.Windows.Forms.ToolStripItem it in f.MainMenuStrip.Items)
                            WalkMenu(it, offenders, whitelist);
                    try { f.Dispose(); } catch { }
                }
            });
            th.SetApartmentState(System.Threading.ApartmentState.STA);
            th.IsBackground = true;
            th.Start();
            th.Join();
            Loc.Init(saved);
            AssertTrue(offenders.Count == 0, "кириллица в EN: " + string.Join(" | ", offenders.ToArray()));
        }

        private static void WalkControls(System.Windows.Forms.Control c, List<string> offenders, HashSet<string> whitelist)
        {
            foreach (System.Windows.Forms.Control child in c.Controls)
            {
                // Пропускаем поля значений: редактируемый ввод (пути, имена) и NumericUpDown.
                // Read-only поля — это подписи, только выделяемые, поэтому они проверяются
                // наравне с Label: иначе перевод описания в «О программе» выпал бы из проверки.
                var editable = child as System.Windows.Forms.TextBoxBase;
                bool isValue = (editable != null && !editable.ReadOnly) ||
                               child is System.Windows.Forms.NumericUpDown;
                if (!isValue)
                {
                    CheckCyrillic(child.Text, child.GetType().Name + ".Text", offenders, whitelist);
                    // Подписи рисуемых вручную карточек живут НЕ в Text (там пусто), а в полях
                    // доступности — без этой проверки заголовки стартового экрана не проверялись
                    // на язык вообще, и английский интерфейс мог показать русскую карточку.
                    CheckCyrillic(child.AccessibleName, child.GetType().Name + ".AccessibleName", offenders, whitelist);
                    CheckCyrillic(child.AccessibleDescription, child.GetType().Name + ".AccessibleDescription", offenders, whitelist);
                    var combo = child as System.Windows.Forms.ComboBox;
                    if (combo != null)
                        foreach (object it in combo.Items)
                            CheckCyrillic(Convert.ToString(it), "ComboItem", offenders, whitelist);
                    var lv = child as System.Windows.Forms.ListView;
                    if (lv != null)
                        foreach (System.Windows.Forms.ColumnHeader col in lv.Columns)
                            CheckCyrillic(col.Text, "Column", offenders, whitelist);
                }
                WalkControls(child, offenders, whitelist);
            }
        }

        private static void WalkMenu(System.Windows.Forms.ToolStripItem item, List<string> offenders, HashSet<string> whitelist)
        {
            CheckCyrillic(item.Text, "Menu", offenders, whitelist);
            var mi = item as System.Windows.Forms.ToolStripMenuItem;
            if (mi != null)
                foreach (System.Windows.Forms.ToolStripItem sub in mi.DropDownItems)
                    WalkMenu(sub, offenders, whitelist);
        }

        private static void CheckCyrillic(string text, string where, List<string> offenders, HashSet<string> whitelist)
        {
            if (string.IsNullOrEmpty(text) || whitelist.Contains(text))
                return;
            foreach (char ch in text)
                if (ch >= 'Ѐ' && ch <= 'ӿ')
                {
                    offenders.Add(where + ": «" + text + "»");
                    return;
                }
        }

        // ---------- TableDetector ----------
        // Хелперы строят линовку и слова в координатах PDF (pt, ось Y вверх).

        private static PdfLine HLine(double y, double x1, double x2)
        {
            return new PdfLine { Orientation = LineOrientation.Horizontal, X1 = x1, Y1 = y, X2 = x2, Y2 = y };
        }

        private static PdfLine VLine(double x, double y1, double y2)
        {
            return new PdfLine { Orientation = LineOrientation.Vertical, X1 = x, Y1 = y1, X2 = x, Y2 = y2 };
        }

        /// <summary>Слово шириной 10×10 с центром (cx, cy).</summary>
        private static PdfWord Word(string text, double cx, double cy)
        {
            return new PdfWord { Text = text, Left = cx - 5, Right = cx + 5, Bottom = cy - 5, Top = cy + 5, FontSizePt = 10 };
        }

        private static string TableCellText(OcrTable t, int row, int col)
        {
            OcrTableCell cell = t.Rows[row].Cells[col];
            return cell.Paragraphs.Count > 0 ? cell.Paragraphs[0].Text : "";
        }

        private static void TestTable2x2()
        {
            // Колонки X=0,50,100; строки Y=0,50,100 (полная сетка 2x2).
            var lines = new List<PdfLine>
            {
                HLine(100, 0, 100), HLine(50, 0, 100), HLine(0, 0, 100),
                VLine(0, 0, 100), VLine(50, 0, 100), VLine(100, 0, 100)
            };
            var words = new List<PdfWord>
            {
                Word("A", 25, 75), Word("B", 75, 75), // верхняя строка
                Word("C", 25, 25), Word("D", 75, 25)  // нижняя строка
            };
            TableDetectResult res = TableDetector.Detect(lines, words, 200, 200);
            AssertEqual(1, res.Tables.Count, "одна таблица");
            OcrTable t = res.Tables[0];
            AssertEqual(2, t.Rows.Count, "2 строки");
            AssertEqual(2, t.ColumnCount, "2 колонки");
            AssertEqual("A", TableCellText(t, 0, 0), "ячейка (0,0)");
            AssertEqual("B", TableCellText(t, 0, 1), "ячейка (0,1)");
            AssertEqual("C", TableCellText(t, 1, 0), "ячейка (1,0)");
            AssertEqual("D", TableCellText(t, 1, 1), "ячейка (1,1)");
            AssertEqual(0, res.RemainingWords.Count, "все слова в таблице");
        }

        private static void TestTableRowSpan()
        {
            // Внутренняя горизонталь Y=50 есть только в правой колонке -> левая ячейка на 2 строки.
            var lines = new List<PdfLine>
            {
                HLine(100, 0, 100), HLine(0, 0, 100), HLine(50, 50, 100),
                VLine(0, 0, 100), VLine(50, 0, 100), VLine(100, 0, 100)
            };
            var words = new List<PdfWord> { Word("L", 25, 50), Word("TR", 75, 75), Word("BR", 75, 25) };
            TableDetectResult res = TableDetector.Detect(lines, words, 200, 200);
            AssertEqual(1, res.Tables.Count, "одна таблица");
            OcrTable t = res.Tables[0];
            AssertEqual(2, t.Rows[0].Cells[0].RowSpan, "левая ячейка на 2 строки");
            AssertTrue(t.Rows[1].Cells[0].Covered, "накрытая позиция под объединением");
            AssertEqual("L", TableCellText(t, 0, 0), "текст объединённой ячейки");
            AssertEqual("TR", TableCellText(t, 0, 1), "правая верхняя");
            AssertEqual("BR", TableCellText(t, 1, 1), "правая нижняя");
        }

        private static void TestTableColSpan()
        {
            // Внутренняя вертикаль X=50 есть только в нижней строке -> верхняя ячейка на 2 колонки.
            var lines = new List<PdfLine>
            {
                HLine(100, 0, 100), HLine(50, 0, 100), HLine(0, 0, 100),
                VLine(0, 0, 100), VLine(100, 0, 100), VLine(50, 0, 50)
            };
            var words = new List<PdfWord> { Word("Top", 50, 75), Word("BL", 25, 25), Word("BR", 75, 25) };
            TableDetectResult res = TableDetector.Detect(lines, words, 200, 200);
            AssertEqual(1, res.Tables.Count, "одна таблица");
            OcrTable t = res.Tables[0];
            AssertEqual(2, t.Rows[0].Cells[0].ColSpan, "верхняя ячейка на 2 колонки");
            AssertTrue(t.Rows[0].Cells[1].Covered, "накрытая позиция справа");
            AssertEqual("Top", TableCellText(t, 0, 0), "текст объединённой ячейки");
            AssertEqual("BL", TableCellText(t, 1, 0), "нижняя левая");
            AssertEqual("BR", TableCellText(t, 1, 1), "нижняя правая");
        }

        private static void TestTableStrayLines()
        {
            // Два подчёркивания (горизонтали без вертикалей) — не таблица.
            var lines = new List<PdfLine> { HLine(60, 0, 40), HLine(20, 0, 40) };
            var words = new List<PdfWord> { Word("x", 20, 60) };
            TableDetectResult res = TableDetector.Detect(lines, words, 200, 200);
            AssertEqual(0, res.Tables.Count, "нет таблиц");
            AssertEqual(1, res.RemainingWords.Count, "слово осталось в потоке");
        }

        private static void TestTableSingleBox()
        {
            // Рамка без внутренних линий — 1x1, не таблица.
            var lines = new List<PdfLine>
            {
                HLine(100, 0, 100), HLine(0, 0, 100), VLine(0, 0, 100), VLine(100, 0, 100)
            };
            var words = new List<PdfWord> { Word("x", 50, 50) };
            TableDetectResult res = TableDetector.Detect(lines, words, 200, 200);
            AssertEqual(0, res.Tables.Count, "рамка 1x1 — не таблица");
            AssertEqual(1, res.RemainingWords.Count, "слово осталось в потоке");
        }

        private static void TestTableWordsOutside()
        {
            var lines = new List<PdfLine>
            {
                HLine(100, 0, 100), HLine(50, 0, 100), HLine(0, 0, 100),
                VLine(0, 0, 100), VLine(50, 0, 100), VLine(100, 0, 100)
            };
            var words = new List<PdfWord> { Word("in", 25, 75), Word("out", 300, 300) };
            TableDetectResult res = TableDetector.Detect(lines, words, 500, 500);
            AssertEqual(1, res.Tables.Count, "одна таблица");
            AssertEqual(1, res.RemainingWords.Count, "внешнее слово в потоке");
            AssertEqual("out", res.RemainingWords[0].Text, "именно внешнее слово");
        }

        private static void TestTableNoLines()
        {
            var words = new List<PdfWord> { Word("a", 10, 10), Word("b", 20, 20) };
            TableDetectResult res = TableDetector.Detect(new List<PdfLine>(), words, 200, 200);
            AssertEqual(0, res.Tables.Count, "нет линий — нет таблиц");
            AssertEqual(2, res.RemainingWords.Count, "все слова в потоке");
        }

        private static void TestUnderlineMarks()
        {
            // Слово [Left..Right]=[10..40], низ Bottom=50; линия у самой базовой линии на всю ширину.
            var w = Word("под", 25, 55); // центр 25 -> Left 20, Right 30; переопределим ниже вручную
            w.Left = 10; w.Right = 40; w.Bottom = 50; w.Top = 60;
            var lines = new List<PdfLine> { HLine(48, 10, 40) }; // на 2 pt ниже низа, вся ширина
            UnderlineDetector.Mark(new List<PdfWord> { w }, lines);
            AssertTrue(w.Underline, "линия у базовой линии на всю ширину -> подчёркнуто");
        }

        private static void TestUnderlineIgnores()
        {
            var far = new PdfWord { Text = "far", Left = 10, Right = 40, Bottom = 50, Top = 60 };
            UnderlineDetector.Mark(new List<PdfWord> { far }, new List<PdfLine> { HLine(30, 10, 40) }); // далеко внизу
            AssertTrue(!far.Underline, "далёкая линия -> не подчёркнуто");

            var shortLine = new PdfWord { Text = "sh", Left = 10, Right = 40, Bottom = 50, Top = 60 };
            UnderlineDetector.Mark(new List<PdfWord> { shortLine }, new List<PdfLine> { HLine(48, 10, 18) }); // покрытие ~27%
            AssertTrue(!shortLine.Underline, "короткая линия -> не подчёркнуто");

            var noLines = new PdfWord { Text = "no", Left = 10, Right = 40, Bottom = 50, Top = 60 };
            UnderlineDetector.Mark(new List<PdfWord> { noLines }, new List<PdfLine>());
            AssertTrue(!noLines.Underline, "нет линий -> не подчёркнуто");
        }

        private static void TestUnderlineWideRule()
        {
            // Полноширинный разделитель под меткой не должен подчёркивать её.
            var w = new PdfWord { Text = "метка", Left = 57, Right = 120, Bottom = 50, Top = 60, FontSizePt = 10 };
            var rule = HLine(48, 57, 520); // длина 463 >> ширины слова 63 (×7)
            UnderlineDetector.Mark(new List<PdfWord> { w }, new List<PdfLine> { rule });
            AssertTrue(!w.Underline, "разделитель во всю ширину -> не подчёркивание");
        }

        /// <summary>Слово с явной рамкой (для тестов колонок/подчёркивания).</summary>
        private static PdfWord WordBox(string text, double left, double right, double bottom)
        {
            return new PdfWord { Text = text, Left = left, Right = right, Bottom = bottom, Top = bottom + 10, FontSizePt = 10 };
        }

        private static void TestSidebarSeparation()
        {
            // 3 строки: слева узкая метка (X57-90), справа тело (X150-350), большой зазор между ними.
            var words = new List<PdfWord>();
            double[] ys = { 220, 200, 180, 160 };
            for (int i = 0; i < ys.Length; i++)
            {
                double y = ys[i];
                words.Add(WordBox("SIDE" + i, 57, 90, y));            // сайдбар-сегмент
                words.Add(WordBox("BODYa" + i, 150, 190, y));         // тело: слова с обычным интервалом
                words.Add(WordBox("BODYb" + i, 195, 235, y));
                words.Add(WordBox("BODYc" + i, 240, 280, y));
                words.Add(WordBox("BODYd" + i, 285, 350, y));
            }
            List<OcrParagraph> paras = OcrLayout.Analyze(words).Paragraphs;
            // Ни один абзац не должен содержать одновременно сайдбар- и тело-токен (нет перемешивания).
            bool anyMixed = false, anySidebarOnly = false;
            foreach (OcrParagraph p in paras)
            {
                bool hasSide = p.Text.Contains("SIDE");
                bool hasBody = p.Text.Contains("BODY");
                if (hasSide && hasBody) anyMixed = true;
                if (hasSide && !hasBody) anySidebarOnly = true;
            }
            AssertTrue(!anyMixed, "сайдбар и тело не смешаны в одном абзаце");
            AssertTrue(anySidebarOnly, "метка сайдбара — отдельным абзацем");
        }

        private static void TestNoSidebarSingleColumn()
        {
            // Одноколоночная строка без больших зазоров: слова обязаны остаться в одном абзаце.
            var words = new List<PdfWord>
            {
                WordBox("one", 50, 90, 200), WordBox("two", 95, 135, 200),
                WordBox("three", 140, 190, 200), WordBox("four", 195, 235, 200)
            };
            List<OcrParagraph> paras = OcrLayout.Analyze(words).Paragraphs;
            AssertEqual(1, paras.Count, "одна строка -> один абзац (сайдбар не сработал)");
            AssertTrue(paras[0].Text.Contains("one") && paras[0].Text.Contains("four"), "все слова в одном абзаце");
        }

        private static void TestMarginsWithImages()
        {
            // Логотип НАД первым словом и штамп НИЖЕ последней строки должны входить в поля:
            // только по словам верхнее поле сдвигало весь вывод вниз на высоту логотипа, а нижнее
            // раздувалось до полустраницы (и выталкивало счёт на второй лист) — его ждёт кап.
            var pt = new PdfPageText();
            var words = new List<PdfWord> { W("x", 100, 700, 200, 10) };       // 100..300, 700..710
            var images = new List<OcrImage>
            {
                new OcrImage { LeftPt = 180, TopPt = 785, WidthPt = 72, HeightPt = 77 },  // логотип выше слова
                new OcrImage { LeftPt = 120, TopPt = 325, WidthPt = 113, HeightPt = 25 }  // печать на середине листа
            };
            PdfTextExtract.SetMargins(pt, words, images, 595, 842);
            AssertEqual(57.0, pt.TopMarginPt, "верхнее поле — от логотипа, а не от первого слова");
            AssertEqual(100.0, pt.LeftMarginPt, "левое поле — минимум по контенту");
            AssertEqual(295.0, pt.RightMarginPt, "правое поле — от самого правого края контента");
            AssertEqual(90.0, pt.BottomMarginPt, "нижнее поле ограничено капом");
        }

        private static void TestColumnConfine()
        {
            double li, ri;
            // Правая колонка на странице textLeft=84..textRight=567: колонка 340..543.
            bool c1 = WordDocxWriter.ColumnConfineIndents(true, 340, 543, 84, 567, out li, out ri);
            AssertTrue(c1, "правая колонка конфайнится");
            AssertEqual(256.0, li, "левый отступ = 340-84");
            AssertEqual(24.0, ri, "правый отступ = 567-543");
            // Левая колонка 104..285.
            WordDocxWriter.ColumnConfineIndents(true, 104, 285, 84, 567, out li, out ri);
            AssertEqual(20.0, li, "левый отступ = 104-84");
            AssertEqual(282.0, ri, "правый отступ = 567-285");
            // Полноширинный блок (тело) 90..560 — конфайна нет.
            bool c3 = WordDocxWriter.ColumnConfineIndents(true, 90, 560, 84, 567, out li, out ri);
            AssertTrue(!c3, "полная ширина -> без конфайна");
            AssertEqual(0.0, li, "нет левого"); AssertEqual(0.0, ri, "нет правого");
            // Не центрированный (eligible=false) — не трогаем.
            bool c4 = WordDocxWriter.ColumnConfineIndents(false, 340, 543, 84, 567, out li, out ri);
            AssertTrue(!c4, "не центрированный -> без конфайна");
            // Титул на всю страницу, центрированный узкий блок с почти равными зазорами: остаётся центрированным.
            WordDocxWriter.ColumnConfineIndents(true, 250, 345, 84, 567, out li, out ri);
            AssertTrue(Math.Abs(li - ri) < 60, "симметричный центр остаётся ~по центру");
        }

        private static void TestHasCyrillic()
        {
            AssertTrue(FontResolver.HasCyrillic("Текст"), "кириллица");
            AssertTrue(FontResolver.HasCyrillic("mix Текст 2"), "смешанное — есть кириллица");
            AssertTrue(!FontResolver.HasCyrillic("Latin 123 %"), "латиница — нет");
            AssertTrue(!FontResolver.HasCyrillic(""), "пусто — нет");
        }

        private static void TestGridColSpan()
        {
            // Три строки пар + итоговая строка одним широким сегментом во всю ширину -> ColSpan=2.
            var words = new List<PdfWord>
            {
                WS("LblA", 0, 200, 40, 8, 8),  WS("ValA", 100, 200, 40, 8, 8),
                WS("LblB", 0, 180, 40, 8, 8),  WS("ValB", 100, 180, 40, 8, 8),
                WS("LblC", 0, 160, 40, 8, 8),  WS("ValC", 100, 160, 40, 8, 8),
                WS("ИтогоОдно", 0, 140, 140, 8, 8) // одна широкая ячейка от левой до правой колонки
            };
            GridDetectResult r = GridDetector.Detect(words);
            AssertEqual(1, r.Tables.Count, "одна сетка");
            OcrTable t = r.Tables[0];
            OcrTableRow last = t.Rows[t.Rows.Count - 1];
            AssertEqual(2, last.Cells[0].ColSpan, "широкий ряд -> ColSpan=2");
            AssertTrue(last.Cells[1].Covered, "вторая ячейка накрыта");
        }

        private static void TestGridRowSpacing()
        {
            // Пять строк пар: шаг 20, но перед последней зазор 32 (пустой промежуток группы, но в
            // пределах одного окна сетки). Ожидаем интервал после 4-й строки (~зазор минус шаг), у остальных 0.
            var words = new List<PdfWord>
            {
                WS("LblA", 0, 300, 40, 8, 8),  WS("ValA", 100, 300, 40, 8, 8),
                WS("LblB", 0, 280, 40, 8, 8),  WS("ValB", 100, 280, 40, 8, 8),
                WS("LblC", 0, 260, 40, 8, 8),  WS("ValC", 100, 260, 40, 8, 8),
                WS("LblD", 0, 240, 40, 8, 8),  WS("ValD", 100, 240, 40, 8, 8),
                WS("LblE", 0, 208, 40, 8, 8),  WS("ValE", 100, 208, 40, 8, 8) // зазор 32 перед этой
            };
            GridDetectResult r = GridDetector.Detect(words);
            AssertEqual(1, r.Tables.Count, "одна сетка");
            OcrTable t = r.Tables[0];
            AssertEqual(5, t.Rows.Count, "пять строк");
            AssertTrue(t.Rows[3].SpaceAfterPt > 10, "перед пустым промежутком — интервал после строки: " + t.Rows[3].SpaceAfterPt);
            AssertEqual(0.0, t.Rows[0].SpaceAfterPt, "у плотных строк интервала нет");
            AssertEqual(0.0, t.Rows[4].SpaceAfterPt, "последняя строка без интервала");
        }

        private static void TestLoneCollinearRule()
        {
            // Прочерк поля, нарисованный тремя кусками на ОДНОЙ оси (дырки под «№»/«от»):
            // компонент коллинеарен -> куски идут в LoneLines (станут прочерком), а не таблицей.
            var lines = new List<PdfLine>
            {
                new PdfLine { Orientation = LineOrientation.Horizontal, X1 = 10,  Y1 = 100, X2 = 60,  Y2 = 100, Thickness = 1 },
                new PdfLine { Orientation = LineOrientation.Horizontal, X1 = 80,  Y1 = 100, X2 = 140, Y2 = 100, Thickness = 1 },
                new PdfLine { Orientation = LineOrientation.Horizontal, X1 = 160, Y1 = 100, X2 = 220, Y2 = 100, Thickness = 1 }
            };
            TableDetectResult r = TableDetector.Detect(lines, new List<PdfWord>(), 600, 800);
            AssertEqual(0, r.Tables.Count, "коллинеарные куски — не таблица");
            AssertEqual(3, r.LoneLines.Count, "все три куска в LoneLines");
        }

        private static void TestRuleInsideTableExcluded()
        {
            // Сетка-рамка 2x2 (замкнутая) + линия ВНУТРИ неё (поле подписи в ячейке): линия не
            // должна попасть в LoneLines (иначе прочерк ляжет поверх таблицы отдельным абзацем).
            var lines = new List<PdfLine>
            {
                new PdfLine { Orientation = LineOrientation.Horizontal, X1 = 0,  Y1 = 200, X2 = 200, Y2 = 200, Thickness = 1 },
                new PdfLine { Orientation = LineOrientation.Horizontal, X1 = 0,  Y1 = 100, X2 = 200, Y2 = 100, Thickness = 1 },
                new PdfLine { Orientation = LineOrientation.Horizontal, X1 = 0,  Y1 = 0,   X2 = 200, Y2 = 0,   Thickness = 1 },
                new PdfLine { Orientation = LineOrientation.Vertical,   X1 = 0,  Y1 = 0,   X2 = 0,   Y2 = 200, Thickness = 1 },
                new PdfLine { Orientation = LineOrientation.Vertical,   X1 = 100, Y1 = 0,  X2 = 100, Y2 = 200, Thickness = 1 },
                new PdfLine { Orientation = LineOrientation.Vertical,   X1 = 200, Y1 = 0,  X2 = 200, Y2 = 200, Thickness = 1 },
                // короткая одиночная линия внутри ячейки (поле для подписи)
                new PdfLine { Orientation = LineOrientation.Horizontal, X1 = 20, Y1 = 50, X2 = 80, Y2 = 50, Thickness = 1 }
            };
            TableDetectResult r = TableDetector.Detect(lines, new List<PdfWord>(), 600, 800);
            AssertEqual(1, r.Tables.Count, "рамка распознана таблицей");
            AssertEqual(0, r.LoneLines.Count, "линия внутри таблицы не даёт прочерк");
        }

        // ---------- GlyphDedup (сдвоенные глифы псевдо-жирного) ----------

        private static GlyphInfo G(string value, double cx, double size)
        {
            return new GlyphInfo { Value = value, CenterX = cx, CenterY = 100, SizePt = size };
        }

        private static string KeptText(IList<GlyphInfo> glyphs, List<int> keep)
        {
            var sb = new System.Text.StringBuilder();
            foreach (int k in keep) sb.Append(glyphs[k].Value);
            return sb.ToString();
        }

        private static void TestGlyphDedupDoubled()
        {
            // «№74» отрисовано дважды со смещением 0.3 pt (как в ж/д билете, кегль 8.4):
            // буквы в слове идут парами. Остаётся «№74», слово — жирное.
            var glyphs = new List<GlyphInfo>
            {
                G("№", 10.0, 8.4), G("№", 10.3, 8.4),
                G("7", 16.0, 8.4), G("7", 16.3, 8.4),
                G("4", 21.0, 8.4), G("4", 21.3, 8.4)
            };
            int dropped;
            List<int> keep = GlyphDedup.Keep(glyphs, out dropped);
            AssertEqual("№74", KeptText(glyphs, keep), "дубли схлопнуты");
            AssertEqual(3, dropped, "выброшено три дубля");
            AssertTrue(GlyphDedup.LooksBold(keep.Count, dropped), "массовая двойная отрисовка = жирный");
        }

        private static void TestGlyphDedupRealPair()
        {
            // Настоящее «77»: одинаковые цифры на шаге ~половины кегля — НЕ дубль.
            var glyphs = new List<GlyphInfo> { G("7", 10.0, 8.4), G("7", 14.2, 8.4) };
            int dropped;
            List<int> keep = GlyphDedup.Keep(glyphs, out dropped);
            AssertEqual(2, keep.Count, "обе семёрки на месте");
            AssertEqual(0, dropped, "дублей нет");
        }

        private static void TestGlyphDedupSparse()
        {
            // Один случайный дубль в обычном слове: текст чистится, но слово НЕ жирное.
            var glyphs = new List<GlyphInfo>
            {
                G("с", 10, 10), G("л", 15, 10), G("л", 15.2, 10), G("о", 20, 10),
                G("в", 25, 10), G("а", 30, 10), G("м", 35, 10), G("и", 40, 10)
            };
            int dropped;
            List<int> keep = GlyphDedup.Keep(glyphs, out dropped);
            AssertEqual("словами", KeptText(glyphs, keep), "дубль вычищен");
            AssertEqual(1, dropped, "один дубль");
            AssertTrue(!GlyphDedup.LooksBold(keep.Count, dropped), "единичный дубль — не жирный");
        }

        private static void TestXyCutOneSubstantialColumn()
        {
            // Блок слева (3 строки) + одинокая пометка справа одной строкой: колонка-
            // крошка допустима рядом с существенной — иначе пометка вклинивается в строку блока.
            var boxes = new[]
            {
                CB(0, 0, 200, 60, 10), CB(1, 300, 200, 50, 10),
                CB(2, 0, 180, 60, 10),
                CB(4, 0, 160, 60, 10)
            };
            List<CutLeaf> leaves = XyCut.Order(boxes, 1000, 30, 25, 2);
            AssertEqual(2, leaves.Count, "блок и пометка разрезаны");
            AssertEqual("0,2,4", TagsOf(leaves[0]), "блок целиком");
            AssertEqual("1", TagsOf(leaves[1]), "пометка отдельно");
        }

        private static PdfWord WS(string text, double left, double bottom, double width, double height, double size)
        {
            return new PdfWord { Text = text, Left = left, Right = left + width, Bottom = bottom, Top = bottom + height, FontSizePt = size };
        }

        private static void TestGridReceipt()
        {
            // Форма «метка … значение»: четыре строки с широким зазором, левые края выровнены
            // в две колонки. Данные полностью синтетические (Lbl/Val), не из реальных документов.
            var words = new List<PdfWord>
            {
                WS("LblA", 0, 200, 50, 8, 8),  WS("ValA", 100, 200, 40, 8, 8),
                WS("LblB", 0, 180, 35, 8, 8),  WS("ValB", 100, 180, 35, 8, 8),
                WS("LblC", 0, 160, 30, 8, 8),  WS("ValC", 100, 160, 40, 8, 8),
                WS("LblD", 0, 140, 35, 8, 8),  WS("ValD", 100, 140, 40, 8, 8)
            };
            GridDetectResult r = GridDetector.Detect(words);
            AssertEqual(1, r.Tables.Count, "одна сетка-таблица");
            OcrTable t = r.Tables[0];
            AssertTrue(t.Borderless, "сетка без границ");
            AssertEqual(2, t.ColumnCount, "две колонки");
            AssertEqual(4, t.Rows.Count, "четыре ряда");
            AssertEqual(0, r.RemainingWords.Count, "все слова в таблице");
            AssertTrue(t.Rows[0].Cells[0].Paragraphs[0].Text.Contains("LblA"), "метка в первой колонке");
            AssertTrue(t.Rows[0].Cells[1].Paragraphs[0].Text.Contains("ValA"), "значение во второй колонке");
        }

        private static void TestGridJustifiedNegative()
        {
            // Обычный текст: зазоры между словами меньше порога — сегмент один, сетки нет.
            var words = new List<PdfWord>();
            for (int row = 0; row < 4; row++)
            {
                double y = 200 - row * 20;
                for (int i = 0; i < 6; i++)
                    words.Add(WS("w" + row + i, i * 45, y, 40, 10, 10));
            }
            GridDetectResult r = GridDetector.Detect(words);
            AssertEqual(0, r.Tables.Count, "justified-текст не считается сеткой");
            AssertEqual(words.Count, r.RemainingWords.Count, "слова остались в потоке");
        }

        private static void TestGridTwoRowsNegative()
        {
            // Две строки пар («подпись … дата» и подобные) — мало строк для сетки.
            var words = new List<PdfWord>
            {
                WS("a", 0, 200, 30, 10, 10), WS("b", 200, 200, 30, 10, 10),
                WS("c", 0, 180, 30, 10, 10), WS("d", 200, 180, 30, 10, 10)
            };
            GridDetectResult r = GridDetector.Detect(words);
            AssertEqual(0, r.Tables.Count, "двух строк недостаточно");
        }

        private static void TestRuleWords()
        {
            // Одиночная линия-прочерк становится «____» по кеглю окружения; линия-подчёркивание
            // и толстая полоса — нет.
            var words = new List<PdfWord> { WS("№", 100, 100, 10, 10, 10) };
            var rule = new PdfLine { Orientation = LineOrientation.Horizontal, X1 = 10, Y1 = 100, X2 = 60, Y2 = 100, Thickness = 1 };
            var under = new PdfLine { Orientation = LineOrientation.Horizontal, X1 = 200, Y1 = 100, X2 = 240, Y2 = 100, Thickness = 1 };
            var fat = new PdfLine { Orientation = LineOrientation.Horizontal, X1 = 300, Y1 = 100, X2 = 350, Y2 = 100, Thickness = 8 };
            var lone = new List<PdfLine> { rule, under, fat };
            var used = new HashSet<PdfLine> { under };
            PdfTextExtract.AddRuleWords(words, lone, used);
            AssertEqual(2, words.Count, "добавлен один прочерк");
            PdfWord w = words[1];
            AssertEqual("__________", w.Text, "50pt при кегле 10 -> десять «_»");
            AssertEqual(10.0, w.Left, "прочерк на месте линии");
            AssertTrue(w.Top > w.Bottom, "прочерку дана высота");
        }

        private static void TestUnderscoreAndPua()
        {
            AssertTrue(PdfTextExtract.IsUnderscoreOnly("____"), "прочерк распознан");
            AssertTrue(!PdfTextExtract.IsUnderscoreOnly("_а_"), "буква внутри — не прочерк");
            AssertTrue(!PdfTextExtract.IsUnderscoreOnly(""), "пусто — не прочерк");
            AssertTrue(PdfTextExtract.IsPrivateUseOnly(new string(new[] { (char)0xF0A7, (char)0x20, (char)0xF0B7 })), "PUA-мусор распознан");
            AssertTrue(!PdfTextExtract.IsPrivateUseOnly("текст"), "обычный текст не PUA");
            AssertTrue(!PdfTextExtract.IsPrivateUseOnly("  "), "одни пробелы — не PUA");
        }

        // ---------- XyCut (порядок чтения многоколоночной вёрстки) ----------

        /// <summary>Бокс разреза: left/bottom — левый нижний угол (Y вверх), +ширина/высота.</summary>
        private static CutBox CB(int tag, double left, double bottom, double width, double height)
        {
            return new CutBox { Tag = tag, Left = left, Right = left + width, Bottom = bottom, Top = bottom + height };
        }

        private static string TagsOf(CutLeaf leaf)
        {
            var parts = new List<string>();
            foreach (int t in leaf.Tags) parts.Add(t.ToString());
            return string.Join(",", parts);
        }

        private static void TestXyCutColumns()
        {
            // Две колонки, строки на ОБЩИХ базовых линиях (сортировка по Top их перемежает).
            var boxes = new[]
            {
                CB(0, 0, 200, 40, 10),   CB(1, 100, 200, 40, 10),
                CB(2, 0, 180, 40, 10),   CB(3, 100, 180, 40, 10),
                CB(4, 0, 160, 45, 10),   CB(5, 100, 160, 40, 10)
            };
            List<CutLeaf> leaves = XyCut.Order(boxes, 1000, 30, 25, 2);
            AssertEqual(2, leaves.Count, "два листа-колонки");
            AssertEqual("0,2,4", TagsOf(leaves[0]), "левая колонка целиком раньше");
            AssertEqual("1,3,5", TagsOf(leaves[1]), "правая колонка после");
            AssertEqual(0.0, leaves[0].ColumnLeft, "рамка левой колонки: левый край");
            AssertEqual(45.0, leaves[0].ColumnRight, "рамка левой колонки: правый край");
            AssertEqual(100.0, leaves[1].ColumnLeft, "рамка правой колонки: левый край");
        }

        private static void TestXyCutFloorsThenColumns()
        {
            // Колонки + полноширинная строка ниже: сначала обе колонки, потом нижний этаж.
            var boxes = new[]
            {
                CB(0, 0, 200, 40, 10),   CB(1, 100, 200, 40, 10),
                CB(2, 0, 180, 40, 10),   CB(3, 100, 180, 40, 10),
                CB(4, 0, 160, 45, 10),   CB(5, 100, 160, 40, 10),
                CB(6, 0, 100, 140, 10)   // пустая полоса 160-110=50 над ней -> свой этаж
            };
            List<CutLeaf> leaves = XyCut.Order(boxes, 20, 30, 25, 2);
            AssertEqual(3, leaves.Count, "две колонки + нижний этаж");
            AssertEqual("0,2,4", TagsOf(leaves[0]), "левая колонка");
            AssertEqual("1,3,5", TagsOf(leaves[1]), "правая колонка");
            AssertEqual("6", TagsOf(leaves[2]), "нижний этаж последним");
        }

        private static void TestXyCutGuardSingleLine()
        {
            // «(подпись) …| (дата)»: широкий пробел ЕСТЬ, но колонки высотой в одну строку —
            // разрез отклоняется, страница не читается «столбиками из одиночных слов».
            var boxes = new[] { CB(0, 0, 100, 60, 10), CB(1, 300, 100, 40, 10) };
            List<CutLeaf> leaves = XyCut.Order(boxes, 1000, 30, 25, 2);
            AssertEqual(1, leaves.Count, "одна строка не делится на колонки");
            AssertEqual("0,1", TagsOf(leaves[0]), "оба элемента в одном листе");
        }

        private static void TestXyCutGuardThinColumn()
        {
            // Крошка (один бокс) рядом с СУЩЕСТВЕННОЙ колонкой — валидный разрез (пометка
            // у блока); защиту от «(подпись) … (дата)» держит TestXyCutGuardSingleLine.
            var boxes = new[]
            {
                CB(0, 0, 200, 40, 10),
                CB(2, 0, 180, 40, 10),
                CB(1, 100, 190, 40, 10)
            };
            List<CutLeaf> leaves = XyCut.Order(boxes, 1000, 30, 25, 2);
            AssertEqual(2, leaves.Count, "существенная колонка + крошка режутся");
        }

        private static void TestOcrTwoColumnsSeparated()
        {
            // Двухколоночная шапка: строки колонок на общих базовых линиях. Абзацы не должны
            // перемешаться (прежняя сортировка по Top читала «через строку»), левая — раньше.
            var words = new List<PdfWord>
            {
                W("L1a", 0, 200, 35, 10), W("L1b", 40, 200, 35, 10),
                W("R1a", 200, 200, 35, 10), W("R1b", 240, 200, 35, 10),
                W("L2a", 0, 180, 35, 10), W("L2b", 40, 180, 35, 10),
                W("R2a", 200, 180, 35, 10), W("R2b", 240, 180, 35, 10)
            };
            List<OcrParagraph> paras = OcrLayout.Analyze(words).Paragraphs;
            // Короткие строки левой колонки — умышленные переводы (влезали бы до правой
            // колонки): каждая своим абзацем. Главное — колонки не перемешаны, левая раньше.
            string all = "";
            int lastLeft = -1, firstRight = int.MaxValue;
            for (int i = 0; i < paras.Count; i++)
            {
                string t = paras[i].Text;
                all += t + "|";
                AssertTrue(!(t.Contains("L") && t.Contains("R")), "в абзаце слова одной колонки: " + t);
                if (t.Contains("L")) lastLeft = i;
                if (t.Contains("R") && i < firstRight) firstRight = i;
            }
            AssertTrue(all.Contains("L1a") && all.Contains("L2b") && all.Contains("R1a") && all.Contains("R2b"),
                "все слова обеих колонок на месте");
            AssertTrue(lastLeft < firstRight, "левая колонка целиком раньше правой");
        }

        private static void TestOcrCellNoColumns()
        {
            // Ячейка «Итого: … 112» с широким зазором между меткой и числом на КАЖДОЙ строке:
            // без запрета колонок вертикальный разрез растащил бы метки и числа в два столбика.
            var words = new List<PdfWord>
            {
                W("Итого:", 0, 200, 40, 10),   W("112", 200, 200, 30, 10),
                W("НДС:", 0, 180, 40, 10),     W("20", 200, 180, 30, 10),
                W("Всего:", 0, 160, 40, 10),   W("113", 200, 160, 30, 10)
            };
            List<OcrParagraph> paras = OcrLayout.Analyze(words, false).Paragraphs;
            // Пустой результат раньше проходил молча — тело foreach просто не выполнялось.
            AssertTrue(paras.Count > 0, "разбор дал хотя бы один абзац");
            var all = new System.Text.StringBuilder();
            foreach (OcrParagraph p in paras)
            {
                bool label = p.Text.Contains("Итого") || p.Text.Contains("НДС") || p.Text.Contains("Всего");
                bool number = p.Text.Contains("112") || p.Text.Contains("20") || p.Text.Contains("113");
                AssertTrue(label == number, "метка и её число не разлучены: " + p.Text);
                all.Append(p.Text).Append(' ');
            }
            // Предикат выше истинен и когда в абзаце НЕТ ни метки, ни числа, поэтому отдельно
            // требуем, чтобы ни одна метка и ни одно число не потерялись при разборе.
            foreach (string token in new[] { "Итого", "НДС", "Всего", "112", "20", "113" })
                AssertTrue(all.ToString().Contains(token), "в разборе сохранилось «" + token + "»");
        }

        private static void TestOcrIndentWithHeader()
        {
            // Шапка (две колонки коротких рваных строк) над justified-телом с красной строкой:
            // строки шапки не должны размыть долю отступов — отступ тела обязан сохраниться
            // (гейт по justified-группам; раньше многострочная шапка обнуляла красную строку).
            var words = new List<PdfWord>
            {
                // шапка: левая колонка (рваная справа)
                W("H1a", 0, 300, 35, 10), W("H1b", 40, 300, 20, 10),
                W("H2a", 0, 280, 35, 10), W("H2b", 40, 280, 35, 10),
                // шапка: правая колонка (рваная)
                W("A1a", 200, 300, 35, 10), W("A1b", 240, 300, 20, 10),
                W("A2a", 200, 280, 35, 10), W("A2b", 240, 280, 30, 10),
                // тело: justified с отступом 15 во всю ширину страницы (как в настоящем письме
                // тело — самая широкая зона; рамка колонки этажа тела = рамка страницы)
                W("B1a", 15, 200, 255, 10),
                W("B2", 0, 188, 270, 10),
                W("B3", 0, 176, 90, 10),
                W("C1", 15, 160, 255, 10),
                W("C2", 0, 148, 80, 10)
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEqual(15.0, layout.FirstLineIndentPt, "красная строка тела пережила шапку");
        }

        private static OcrParagraph Para(string text, double left, double bottom, double width, double height)
        {
            var p = new OcrParagraph
            {
                TopPt = bottom + height,
                LeftPt = left,
                RightPt = left + width,
                BottomPt = bottom
            };
            p.Runs.Add(new OcrRun { Text = text });
            return p;
        }

        private static string ColText(List<PageBlocks.Block> col)
        {
            var t = new List<string>();
            foreach (PageBlocks.Block b in col) t.Add(b.Paragraph != null ? b.Paragraph.Text : "<img>");
            return string.Join(",", t);
        }

        private static void TestOrderedItemsColumns()
        {
            // Две колонки абзацев (перекрываются по вертикали, широкий просвет между ними) +
            // полноширинный абзац ниже. Ожидаем: одна side-by-side полоса (левая|правая колонка),
            // затем одиночный нижний блок — а НЕ перемешанные строки.
            var page = new PdfPageText();
            page.Paragraphs.Add(Para("L1", 100, 650, 180, 50)); // правый край 280
            page.Paragraphs.Add(Para("R1", 340, 640, 200, 50)); // просвет 340-280=60
            page.Paragraphs.Add(Para("L2", 100, 590, 180, 50));
            page.Paragraphs.Add(Para("R2", 340, 580, 200, 50));
            page.Paragraphs.Add(Para("F", 100, 500, 440, 40)); // этаж ниже (зазор 580-540=40)
            List<PageBlocks.PageItem> items = PageBlocks.OrderedItems(page);
            AssertEqual(2, items.Count, "полоса + нижний блок");
            AssertTrue(items[0].IsBand, "первый элемент — side-by-side полоса");
            AssertEqual(2, items[0].Columns.Count, "две колонки в полосе");
            AssertEqual("L1,L2", ColText(items[0].Columns[0]), "левая колонка целиком");
            AssertEqual("R1,R2", ColText(items[0].Columns[1]), "правая колонка целиком");
            AssertTrue(!items[1].IsBand && items[1].Single.Paragraph.Text == "F", "нижний — одиночный блок");
        }

        private static void TestCropRect()
        {
            // Страница 500×1000 pt отрендерена в 1000×2000 px (2 px/pt). Картинка Y-вверх:
            // left=100, top=800, w=200, h=100 (занимает Y 700..800). Верх в пикселях = (1000-800)*2.
            System.Drawing.Rectangle r = PageRasterizer.CropRect(1000, 2000, 500, 1000, 100, 800, 200, 100);
            AssertEqual(200, r.X, "X");
            AssertEqual(400, r.Y, "Y (ось вниз)");
            AssertEqual(400, r.Width, "ширина в px");
            AssertEqual(200, r.Height, "высота в px");
            // Кламп: картинка выходит за правый/нижний край — обрезается по границам страницы.
            System.Drawing.Rectangle c = PageRasterizer.CropRect(1000, 2000, 500, 1000, 400, 1000, 200, 100);
            AssertTrue(c.X + c.Width <= 1000, "правый край не за пределами");
            AssertTrue(c.Y >= 0 && c.Y + c.Height <= 2000, "по вертикали в пределах");
        }

        private static void TestBandColumnWidths()
        {
            // Колонки [100..280] и [340..540] на текстовой области 84..567: граница = середина
            // зазора (280+340)/2=310; ширины ячеек = 310-84 и 567-310.
            var band = new PageBlocks.PageItem
            {
                Columns = new List<List<PageBlocks.Block>> { new List<PageBlocks.Block>(), new List<PageBlocks.Block>() },
                ColLeft = new double[] { 100, 340 },
                ColRight = new double[] { 280, 540 }
            };
            double[] w = WordDocxWriter.BandColumnWidths(band, 84, 567);
            AssertEqual(2, w.Length, "две ширины");
            AssertEqual(226.0, w[0], "левая ячейка = 310-84");
            AssertEqual(257.0, w[1], "правая ячейка = 567-310");
        }

        private static void TestOrderedItemsWithinLeaf()
        {
            // Внутри ОДНОГО листа (узкий просвет — не колонки): соседние строки строго сверху вниз
            // (не переставлять по X), а перекрытые бок о бок блоки с узким зазором — слева направо.
            var page = new PdfPageText();
            page.Paragraphs.Add(Para("upper", 20, 592, 200, 8));  // 592..600
            page.Paragraphs.Add(Para("lower", 10, 580, 200, 8));  // 580..588 (не перекрыты по вертикали)
            page.Paragraphs.Add(Para("right", 232, 450, 100, 100)); // левый край 232
            page.Paragraphs.Add(Para("left", 20, 450, 200, 100));   // правый край 220, зазор 12 < порога колонки
            List<PageBlocks.PageItem> items = PageBlocks.OrderedItems(page);
            var order = new List<string>();
            foreach (PageBlocks.PageItem it in items)
            {
                AssertTrue(!it.IsBand, "узкий зазор — не полоса, всё одиночными блоками");
                order.Add(it.Single.Paragraph.Text);
            }
            AssertEqual("upper,lower,left,right", string.Join(",", order),
                "строки сверху вниз, перекрытые бок о бок (узкий зазор) — слева направо");
        }

        private static void TestImageCentered()
        {
            // Логотип сверху по центру A4 (595 pt): left=281, width=64 -> зазоры 281/250, почти равны.
            AssertTrue(WordDocxWriter.IsImageCentered(281, 64, 595), "логотип по центру -> центрируем");
            // Врезка у левого поля: left=72, width=100 -> правый зазор огромный, асимметрия -> нет.
            AssertTrue(!WordDocxWriter.IsImageCentered(72, 100, 595), "врезка слева -> не центр");
            // Печать сбоку справа: left=450, width=100 -> левый зазор огромный -> нет.
            AssertTrue(!WordDocxWriter.IsImageCentered(450, 100, 595), "печать справа -> не центр");
            // Изображение во всю ширину -> не центрируем (нет зазоров).
            AssertTrue(!WordDocxWriter.IsImageCentered(0, 595, 595), "во всю ширину -> не центр");
            // Вырожденная ширина страницы -> не центрируем (защита от деления/мусора).
            AssertTrue(!WordDocxWriter.IsImageCentered(10, 20, 0), "ширина страницы 0 -> не центр");
        }

        private static void TestHasExtractableContent()
        {
            // Пустая страница — нет текста.
            AssertTrue(!PdfPageExtraction.HasExtractableContent(new PdfPageText()), "пустая страница — нет текста");
            // Только абзац — есть текст.
            var withPar = new PdfPageText();
            withPar.Paragraphs.Add(new OcrParagraph());
            AssertTrue(PdfPageExtraction.HasExtractableContent(withPar), "абзац — есть текст");
            // Только таблица с текстом в ячейке — есть текст (иначе ложный «скан»).
            var withTable = new PdfPageText();
            var cell = new OcrTableCell();
            cell.Paragraphs.Add(new OcrParagraph());
            var row = new OcrTableRow();
            row.Cells.Add(cell);
            var table = new OcrTable();
            table.Rows.Add(row);
            withTable.Tables.Add(table);
            AssertTrue(PdfPageExtraction.HasExtractableContent(withTable), "таблица с текстом — есть текст");
        }

        // ---------- PDF → PowerPoint ----------

        /// <summary>Страница-заготовка нужного размера (содержимое добавляют сами тесты).</summary>
        private static PdfPageText PptxPage(double widthPt, double heightPt)
        {
            return new PdfPageText { WidthPt = widthPt, HeightPt = heightPt };
        }

        private static void TestXmlEscape()
        {
            AssertEqual("a &amp; b", XmlText.Escape("a & b"), "амперсанд");
            AssertEqual("&lt;tag&gt;", XmlText.Escape("<tag>"), "угловые скобки");
            AssertEqual("&quot;q&quot;", XmlText.Escape("\"q\""), "кавычка (ею же пишутся значения атрибутов)");
            AssertEqual("", XmlText.Escape(null), "null — пустая строка, а не падение");
            AssertEqual("a=1&amp;b=2", XmlText.Encode("a=1&b=2"), "адрес ссылки: & обязан экранироваться");
        }

        private static void TestXmlSanitize()
        {
            AssertEqual("ab", XmlText.Sanitize("a\0b"), "нулевой символ снят");
            AssertEqual("ab", XmlText.Sanitize("a\u001Bb"), "управляющий символ снят");
            AssertEqual("a\tb", XmlText.Sanitize("a\tb"), "табуляция допустима в XML");
            AssertEqual("a b", XmlText.Sanitize("a\nb"), "перевод строки → пробел");
            AssertEqual("a b", XmlText.Sanitize("a\rb"), "возврат каретки → пробел");
            AssertEqual("ab", XmlText.Sanitize("a\uD83Db"), "непарная верхняя половина суррогата снята");
            AssertEqual("ab", XmlText.Sanitize("a\uDE00b"), "непарная нижняя половина суррогата снята");
            AssertEqual("a\uD83D\uDE00b", XmlText.Sanitize("a\uD83D\uDE00b"), "корректная пара сохранена целиком");
            AssertEqual("ab", XmlText.Sanitize("a\uFFFEb"), "«не символ» U+FFFE снят");
        }

        private static void TestPptxEmu()
        {
            AssertEqual(914400L, PptxGeometry.Emu(72), "72 pt — дюйм — 914400 EMU");
            AssertEqual(12700L, PptxGeometry.Emu(1), "пункт — 12700 EMU");
            AssertEqual(-12700L, PptxGeometry.Emu(-1), "координата сохраняет знак");
            AssertEqual(0L, PptxGeometry.EmuSize(-5), "отрицательного размера не бывает");
            AssertEqual(0L, PptxGeometry.Emu(double.NaN), "мусор — ноль, а не исключение");
            AssertEqual(13L, PptxGeometry.Emu(0.001), "округление к ближайшему");
        }

        private static void TestPptxHundredths()
        {
            AssertEqual(1200, PptxGeometry.Hundredths(12), "12 pt — 1200 сотых");
            AssertEqual(1050, PptxGeometry.Hundredths(10.5), "дробный кегль");
            AssertEqual(100, PptxGeometry.Hundredths(0.2), "ниже предела разметки — 1 pt");
            AssertEqual(400000, PptxGeometry.Hundredths(9000), "выше предела разметки — 4000 pt");
            // Кегль, УМЕНЬШЕННЫЙ вписыванием, обязан стать меньше, а не «отскочить» к умолчанию.
            SlideFit fit = PptxGeometry.Fit(1200, 675, 960, 540);
            AssertEqual(480, fit.FontHundredths(6), "6 pt при масштабе 0,8 — это 4,8 pt");
        }

        private static void TestPptxSlideSize()
        {
            double w, h;
            PptxGeometry.SlideSizePt(new List<PdfPageText>
            {
                PptxPage(960, 540), PptxPage(595, 842), PptxPage(960, 540), PptxPage(960, 540)
            }, out w, out h);
            AssertEqual(960.0, w, "ширина — по самому частому размеру страницы");
            AssertEqual(540.0, h, "высота — по самому частому размеру страницы");

            PptxGeometry.SlideSizePt(new List<PdfPageText>(), out w, out h);
            AssertEqual(595.0, w, "пустой набор — лист A4");

            PptxGeometry.SlideSizePt(new List<PdfPageText> { PptxPage(720, 540), PptxPage(960, 540) }, out w, out h);
            AssertEqual(720.0, w, "при равенстве побеждает встретившийся раньше");

            PptxGeometry.SlideSizePt(new List<PdfPageText> { PptxPage(99999, 540) }, out w, out h);
            AssertEqual(PptxGeometry.MaxSlidePt, w, "размер сверх предела PowerPoint обрезан");
        }

        private static void TestPptxFit()
        {
            SlideFit same = PptxGeometry.Fit(960, 540, 960, 540);
            AssertEqual(1.0, same.Scale, "размер совпал — масштаб 1");
            AssertEqual(0L, same.X(0), "совпал — и сдвига нет");

            SlideFit noise = PptxGeometry.Fit(960.0001, 540, 960, 540);
            AssertEqual(1.0, noise.Scale, "разница в тысячные — тот же размер, а не повод масштабировать");

            // Книжная A4 на широком слайде: вписывается по высоте, центрируется по ширине.
            SlideFit fit = PptxGeometry.Fit(595, 842, 960, 540);
            AssertTrue(Math.Abs(fit.Scale - 540.0 / 842.0) < 1e-9, "масштаб по меньшей стороне");
            double scaledW = 595 * fit.Scale;
            AssertEqual(PptxGeometry.Emu((960 - scaledW) / 2), fit.X(0), "левый край — по центру слайда");
            AssertEqual(0L, fit.Y(842), "верх страницы совпал с верхом слайда");
        }

        private static void TestPptxFlipY()
        {
            SlideFit fit = PptxGeometry.Fit(600, 800, 600, 800);
            AssertEqual(0L, fit.Y(800), "верх страницы (ось PDF вверх) — ноль сверху слайда");
            AssertEqual(PptxGeometry.Emu(100), fit.Y(700), "на 100 pt ниже верха");
            AssertEqual(PptxGeometry.Emu(800), fit.Y(0), "низ страницы");
            AssertEqual(PptxGeometry.Emu(20), fit.Size(20), "длина без масштаба не меняется");
        }

        private static void TestOoxmlContentTypes()
        {
            var pkg = new OoxmlPackage();
            pkg.AddXml("_rels/.rels", "<a/>");
            pkg.AddXml("ppt/presentation.xml", "<a/>", PptxParts.CtPresentation);
            pkg.AddBinary("ppt/media/image1.png", new byte[] { 1, 2, 3 }, PptxParts.CtPng);

            string types = pkg.BuildContentTypes();
            AssertTrue(types.Contains("<Default Extension=\"png\" ContentType=\"image/png\"/>"),
                "картинки покрыты правилом по расширению — одной строкой на сотню файлов");
            AssertTrue(types.Contains("<Default Extension=\"rels\""), "связи — тоже правилом");
            AssertTrue(types.Contains("<Override PartName=\"/ppt/presentation.xml\""),
                "поимённый тип пишется С ведущим слэшем");
            AssertTrue(!types.Contains("PartName=\"ppt/"), "без слэша — та самая ошибка «файл повреждён»");

            using (var ms = new MemoryStream(pkg.ToArray()))
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
            {
                AssertEqual("[Content_Types].xml", zip.Entries[0].FullName, "типы содержимого — ПЕРВОЙ записью архива");
                AssertEqual(4, zip.Entries.Count, "три части плюс типы содержимого");
            }
        }

        private static void TestOoxmlPartNames()
        {
            var pkg = new OoxmlPackage();
            AssertThrowsAny("ведущий слэш в имени части", delegate { pkg.AddXml("/ppt/a.xml", "<a/>"); });
            AssertThrowsAny("обратный слэш в имени части", delegate { pkg.AddXml("ppt\\a.xml", "<a/>"); });
            AssertThrowsAny("не-ASCII в имени части", delegate { pkg.AddXml("ppt/слайд.xml", "<a/>"); });
            AssertThrowsAny("пустой сегмент в имени части", delegate { pkg.AddXml("ppt//a.xml", "<a/>"); });
            pkg.AddXml("ppt/a.xml", "<a/>");
            AssertThrowsAny("дубль имени части", delegate { pkg.AddXml("ppt/a.xml", "<a/>"); });
            AssertThrowsAny("двоичная часть без типа содержимого", delegate { pkg.AddBinary("ppt/media/i.png", new byte[1], null); });
            // Правило по расширению одно на весь пакет: второй тип для «png» выразить нечем.
            pkg.AddBinary("ppt/media/i1.png", new byte[1], PptxParts.CtPng);
            AssertThrowsAny("одно расширение — два разных типа", delegate { pkg.AddBinary("ppt/media/i2.png", new byte[1], "image/jpeg"); });
        }

        private static void TestPptxPackageConsistency()
        {
            OoxmlPackage pkg = PptxWriter.BuildPackage(
                new List<PdfPageText> { PptxPage(960, 540), PptxPage(960, 540) },
                new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc));
            var failures = new List<string>();
            AssertOoxmlSelfConsistent(pkg.ToArray(), failures);
            AssertTrue(failures.Count == 0, "пакет несогласован: " + string.Join(" | ", failures.ToArray()));
        }

        /// <summary>
        /// Самопроверка пакета Open XML — та, которую читатель формата делает молча, а при
        /// провале говорит только «файл повреждён»: типы содержимого покрывают каждую часть,
        /// каждая связь ведёт в существующую часть, каждая ссылка r:id объявлена в связях
        /// СВОЕЙ части, XML-части без BOM. Ошибки собираются списком — чтобы видеть все сразу.
        /// </summary>
        private static void AssertOoxmlSelfConsistent(byte[] package, List<string> failures)
        {
            using (var ms = new MemoryStream(package))
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
            {
                if (zip.Entries.Count == 0 || zip.Entries[0].FullName != "[Content_Types].xml")
                    failures.Add("[Content_Types].xml не первой записью");

                var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var bytes = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                foreach (ZipArchiveEntry e in zip.Entries)
                {
                    names.Add(e.FullName);
                    using (Stream s = e.Open())
                    using (var buf = new MemoryStream())
                    {
                        s.CopyTo(buf);
                        bytes[e.FullName] = buf.ToArray();
                    }
                }

                // 1. XML-части — без метки порядка байтов: читатель ждёт «<» первым байтом.
                foreach (KeyValuePair<string, byte[]> kv in bytes)
                    if ((kv.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                        || kv.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                        && (kv.Value.Length < 1 || kv.Value[0] != (byte)'<'))
                        failures.Add("BOM или мусор в начале: " + kv.Key);

                // 2. Каждая часть описана: правилом по расширению или поимённо.
                XDocument types = XDocument.Parse(Encoding.UTF8.GetString(bytes["[Content_Types].xml"]));
                XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
                var defaults = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var overrides = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (XElement d in types.Root.Elements(ct + "Default"))
                    defaults.Add((string)d.Attribute("Extension"));
                foreach (XElement o in types.Root.Elements(ct + "Override"))
                {
                    string part = (string)o.Attribute("PartName");
                    if (part == null || !part.StartsWith("/", StringComparison.Ordinal))
                        failures.Add("Override без ведущего слэша: " + part);
                    else if (!names.Contains(part.Substring(1)))
                        failures.Add("Override указывает на несуществующую часть: " + part);
                    if (part != null)
                        overrides.Add(part.TrimStart('/'));
                }
                foreach (string name in names)
                {
                    if (name == "[Content_Types].xml" || overrides.Contains(name))
                        continue;
                    int dot = name.LastIndexOf('.');
                    string ext = dot >= 0 ? name.Substring(dot + 1) : string.Empty;
                    if (!defaults.Contains(ext))
                        failures.Add("часть без типа содержимого: " + name);
                }

                // 3. Связи ведут в существующие части, ссылки r:id объявлены в связях своей части.
                XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
                XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                foreach (string name in new List<string>(names))
                {
                    if (!name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                        continue;
                    XDocument rels = XDocument.Parse(Encoding.UTF8.GetString(bytes[name]));
                    var declared = new HashSet<string>(StringComparer.Ordinal);
                    foreach (XElement r in rels.Root.Elements(rel + "Relationship"))
                    {
                        declared.Add((string)r.Attribute("Id"));
                        if ((string)r.Attribute("TargetMode") == "External")
                            continue;
                        string target = ResolveRelTarget(name, (string)r.Attribute("Target"));
                        if (!names.Contains(target))
                            failures.Add("связь ведёт в никуда: " + name + " → " + target);
                    }
                    string owner = OwnerOfRels(name);
                    if (owner.Length == 0 || !bytes.ContainsKey(owner))
                        continue; // корневые связи пакета части-владельца не имеют
                    XDocument doc = XDocument.Parse(Encoding.UTF8.GetString(bytes[owner]));
                    foreach (XElement el in doc.Descendants())
                        foreach (XAttribute a in el.Attributes())
                            if (a.Name.Namespace == rNs && !declared.Contains(a.Value))
                                failures.Add("ссылка " + a.Value + " в " + owner + " не объявлена в связях");
                }
            }
        }

        /// <summary>«ppt/slideMasters/_rels/slideMaster1.xml.rels» + «../theme/theme1.xml» → «ppt/theme/theme1.xml».</summary>
        private static string ResolveRelTarget(string relsPart, string target)
        {
            if (string.IsNullOrEmpty(target))
                return string.Empty;
            if (target.StartsWith("/", StringComparison.Ordinal))
                return target.Substring(1);
            int marker = relsPart.LastIndexOf("_rels/", StringComparison.Ordinal);
            string baseDir = marker <= 0 ? string.Empty : relsPart.Substring(0, marker);
            var parts = new List<string>((baseDir + target).Split('/'));
            var stack = new List<string>();
            foreach (string p in parts)
            {
                if (p.Length == 0 || p == ".")
                    continue;
                if (p == ".." && stack.Count > 0)
                    stack.RemoveAt(stack.Count - 1);
                else if (p != "..")
                    stack.Add(p);
            }
            return string.Join("/", stack.ToArray());
        }

        /// <summary>Часть-владелец файла связей: «ppt/_rels/presentation.xml.rels» → «ppt/presentation.xml».</summary>
        private static string OwnerOfRels(string relsPart)
        {
            int marker = relsPart.LastIndexOf("_rels/", StringComparison.Ordinal);
            if (marker < 0)
                return string.Empty;
            string dir = relsPart.Substring(0, marker);
            string file = relsPart.Substring(marker + "_rels/".Length);
            if (!file.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                return string.Empty;
            file = file.Substring(0, file.Length - ".rels".Length);
            return file.Length == 0 ? string.Empty : dir + file;
        }

        /// <summary>Абзац-заготовка: один ран с заданным текстом в заданной рамке.</summary>
        private static OcrParagraph PptxPar(string text, double leftPt, double topPt, double rightPt, double bottomPt,
            double fontPt = 12)
        {
            var p = new OcrParagraph
            {
                LeftPt = leftPt, TopPt = topPt, RightPt = rightPt, BottomPt = bottomPt,
                Alignment = OcrAlignment.Left
            };
            p.Runs.Add(new OcrRun { Text = text, FontSizePt = fontPt });
            return p;
        }

        /// <summary>Первый элемент выборки или null: тесты читают разметку без LINQ.</summary>
        private static XElement PptxFirst(IEnumerable<XElement> items)
        {
            foreach (XElement e in items)
                return e;
            return null;
        }

        /// <summary>Весь видимый текст слайда — склейка узлов a:t по порядку.</summary>
        private static string PptxAllText(XDocument slide)
        {
            var sb = new StringBuilder();
            foreach (XElement t in slide.Descendants(PptxA + "t"))
                sb.Append(t.Value);
            return sb.ToString();
        }

        /// <summary>XML одного слайда собранного пакета (1-based номер).</summary>
        private static XDocument PptxSlideXml(OoxmlPackage package, int slideNumber)
        {
            using (var ms = new MemoryStream(package.ToArray()))
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
            using (Stream s = zip.GetEntry("ppt/slides/slide" + slideNumber + ".xml").Open())
            using (var reader = new StreamReader(s, Encoding.UTF8))
                return XDocument.Parse(reader.ReadToEnd());
        }

        /// <summary>Часть пакета как текст — для проверок самой разметки, а не разобранного XML.</summary>
        private static string PptxPartText(OoxmlPackage package, string part)
        {
            using (var ms = new MemoryStream(package.ToArray()))
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
            using (Stream s = zip.GetEntry(part).Open())
            using (var reader = new StreamReader(s, Encoding.UTF8))
                return reader.ReadToEnd();
        }

        private static readonly XNamespace PptxA = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace PptxP = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private static readonly XNamespace PptxR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly DateTime PptxStamp = new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc);

        private static void TestPptxTextBox()
        {
            PdfPageText page = PptxPage(600, 800);
            OcrParagraph par = PptxPar("Привет", 100, 700, 300, 680, 14);
            page.Paragraphs.Add(par);
            XDocument slide = PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp), 1);

            List<XElement> shapes = new List<XElement>(slide.Descendants(PptxP + "sp"));
            AssertEqual(1, shapes.Count, "один абзац — одна надпись");

            XElement off = PptxFirst(shapes[0].Descendants(PptxA + "off"));
            AssertEqual(PptxGeometry.Emu(100).ToString(), (string)off.Attribute("x"), "левый край — из рамки абзаца");
            // Рамка поднята над базовой линией первой строки: PowerPoint ставит буквы не вплотную
            // к верхнему краю надписи, и без поправки текст сел бы ниже, чем стоял в источнике.
            double boxTop = PptxWriter.FirstBaselinePt(par, 14) + PptxWriter.FirstBaselineDropPt(14);
            AssertEqual(PptxGeometry.Emu(800 - boxTop).ToString(), (string)off.Attribute("y"),
                "верх = высота страницы − базовая линия первой строки − её место под верхом рамки");

            XElement bodyPr = PptxFirst(shapes[0].Descendants(PptxA + "bodyPr"));
            AssertEqual("0", (string)bodyPr.Attribute("lIns"), "нулевое поле слева — иначе текст уезжает");
            AssertEqual("0", (string)bodyPr.Attribute("tIns"), "нулевое поле сверху");
            AssertTrue(bodyPr.Element(PptxA + "noAutofit") != null, "автоподбор кегля выключен");
            AssertTrue(PptxFirst(shapes[0].Descendants(PptxA + "noFill")) != null, "надпись без заливки");
            AssertEqual("Привет", PptxAllText(slide), "текст на месте");
        }

        private static void TestPptxRunProps()
        {
            PdfPageText page = PptxPage(600, 800);
            var par = new OcrParagraph { LeftPt = 10, TopPt = 700, RightPt = 500, BottomPt = 680 };
            par.Runs.Add(new OcrRun
            {
                Text = "Жирный", FontSizePt = 13.5, Bold = true, Italic = true, Underline = true,
                ColorArgb = 0xD32F2F, FontName = "Arial", Super = true
            });
            page.Paragraphs.Add(par);
            XDocument slide = PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp), 1);

            XElement rPr = PptxFirst(slide.Descendants(PptxA + "rPr"));
            AssertEqual("1350", (string)rPr.Attribute("sz"), "кегль в сотых пункта");
            AssertEqual("1", (string)rPr.Attribute("b"), "полужирный");
            AssertEqual("1", (string)rPr.Attribute("i"), "курсив");
            AssertEqual("sng", (string)rPr.Attribute("u"), "подчёркивание");
            AssertEqual("30000", (string)rPr.Attribute("baseline"), "надстрочный");
            AssertEqual("ru-RU", (string)rPr.Attribute("lang"), "кириллице — свой язык");
            AssertEqual("D32F2F", (string)PptxFirst(rPr.Descendants(PptxA + "srgbClr")).Attribute("val"), "цвет");

            // Порядок детей задан схемой: заливка, шрифты, ссылка. Проверяем именно порядок.
            var order = new List<string>();
            foreach (XElement child in rPr.Elements())
                order.Add(child.Name.LocalName);
            AssertTrue(order.IndexOf("solidFill") < order.IndexOf("latin"), "заливка раньше шрифта");
            AssertTrue(order.IndexOf("latin") < order.IndexOf("ea") && order.IndexOf("ea") < order.IndexOf("cs"),
                "latin → ea → cs");
        }

        private static void TestPptxAlignAndBullet()
        {
            PdfPageText page = PptxPage(600, 800);
            OcrParagraph centered = PptxPar("По центру", 10, 700, 500, 680);
            centered.Alignment = OcrAlignment.Center;
            OcrParagraph item = PptxPar("1. Пункт", 10, 650, 500, 630);
            item.ListKind = ListKind.Numbered;
            item.ListNumber = 1;
            item.ListContentStart = 3;
            page.Paragraphs.Add(centered);
            page.Paragraphs.Add(item);
            XDocument slide = PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp), 1);

            var aligns = new List<string>();
            foreach (XElement pPr in slide.Descendants(PptxA + "pPr"))
                aligns.Add((string)pPr.Attribute("algn"));
            AssertTrue(aligns.Contains("ctr"), "центрирование перенесено");
            AssertTrue(aligns.Contains("l"), "выравнивание влево перенесено");

            AssertTrue(PptxFirst(slide.Descendants(PptxA + "buNone")) != null,
                "родной маркер отключён — иначе нумерация в каждой надписи начнётся заново");
            string all = PptxAllText(slide);
            AssertTrue(all.Contains("1. Пункт"), "маркер остаётся текстом, номер не теряется");
        }

        private static void TestPptxEmptyParagraph()
        {
            PdfPageText page = PptxPage(600, 800);
            page.Paragraphs.Add(PptxPar("   ", 10, 700, 500, 680));
            page.Paragraphs.Add(PptxPar("", 10, 650, 500, 630));
            XDocument slide = PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp), 1);
            AssertTrue(PptxFirst(slide.Descendants(PptxP + "sp")) == null,
                "из пробелов фигура не делается — пустые рамки мешают править слайд");
        }

        private static void TestPptxWidthSlack()
        {
            AssertEqual(104.0, PptxWriter.SlackWidthPt(0, 100, 600), "запас 4% от ширины");
            AssertEqual(12.0, PptxWriter.SlackWidthPt(0, 10, 600), "короткой строке — минимум 2 pt");
            AssertEqual(100.0, PptxWriter.SlackWidthPt(500, 100, 600), "запас не выходит за край страницы");
            AssertTrue(PptxWriter.SlackWidthPt(100, 0, 600) > 0, "вырожденная ширина — до правого края");
        }

        private static void TestPptxMediaDedup()
        {
            byte[] png = PptxTinyPng();
            var pages = new List<PdfPageText>();
            for (int i = 0; i < 3; i++)
            {
                PdfPageText page = PptxPage(600, 800);
                page.Images.Add(new OcrImage { Png = png, LeftPt = 10, TopPt = 700, WidthPt = 50, HeightPt = 40 });
                pages.Add(page);
            }
            OoxmlPackage package = PptxWriter.BuildPackage(pages, PptxStamp);
            byte[] bytes = package.ToArray();

            int mediaParts = 0;
            using (var ms = new MemoryStream(bytes))
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
                foreach (ZipArchiveEntry e in zip.Entries)
                    if (e.FullName.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase))
                        mediaParts++;
            AssertEqual(1, mediaParts, "одинаковые байты кладутся в пакет один раз");

            for (int i = 1; i <= 3; i++)
            {
                XDocument slide = PptxSlideXml(package, i);
                AssertTrue(PptxFirst(slide.Descendants(PptxA + "blip")) != null, "картинка на слайде " + i);
            }
            var failures = new List<string>();
            AssertOoxmlSelfConsistent(bytes, failures);
            AssertTrue(failures.Count == 0, "пакет с картинками согласован: " + string.Join(" | ", failures.ToArray()));
        }

        private static void TestPptxHyperlink()
        {
            PdfPageText page = PptxPage(600, 800);
            var par = new OcrParagraph { LeftPt = 10, TopPt = 700, RightPt = 500, BottomPt = 680 };
            par.Runs.Add(new OcrRun { Text = "первый", FontSizePt = 12, Uri = "https://example.org/a?b=1&c=2" });
            par.Runs.Add(new OcrRun { Text = "второй", FontSizePt = 12, Uri = "https://example.org/a?b=1&c=2" });
            page.Paragraphs.Add(par);
            OoxmlPackage package = PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp);
            byte[] bytes = package.ToArray();

            string rels;
            using (var ms = new MemoryStream(bytes))
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
            using (Stream s = zip.GetEntry("ppt/slides/_rels/slide1.xml.rels").Open())
            using (var reader = new StreamReader(s, Encoding.UTF8))
                rels = reader.ReadToEnd();

            AssertTrue(rels.Contains("TargetMode=\"External\""), "внешняя ссылка помечена как внешняя");
            AssertTrue(rels.Contains("b=1&amp;c=2"), "амперсанд в адресе экранирован — иначе файл не откроется");
            int hyperlinks = 0, at = 0;
            while ((at = rels.IndexOf("/hyperlink", at, StringComparison.Ordinal) + 1) > 0)
                hyperlinks++;
            AssertEqual(1, hyperlinks, "два рана с одним адресом — одна связь");

            XDocument slide = PptxSlideXml(package, 1);
            var ids = new List<string>();
            foreach (XElement click in slide.Descendants(PptxA + "hlinkClick"))
                ids.Add((string)click.Attribute(PptxR + "id"));
            AssertEqual(2, ids.Count, "ссылка стоит у обоих ранов");
            AssertEqual(ids[0], ids[1], "и указывает на одну связь");

            var failures = new List<string>();
            AssertOoxmlSelfConsistent(bytes, failures);
            AssertTrue(failures.Count == 0, "пакет со ссылкой согласован: " + string.Join(" | ", failures.ToArray()));
        }

        private static void TestPptxFillerRule()
        {
            AssertTrue(PptxWriter.IsFillerRule("____"), "линия, ставшая словом из подчёркиваний");
            AssertTrue(PptxWriter.IsFillerRule("____ ____"), "две линии в одном ране — тоже заполнитель");
            AssertTrue(!PptxWriter.IsFillerRule("_x_"), "подчёркивания вокруг текста — не заполнитель");
            AssertTrue(!PptxWriter.IsFillerRule("_"), "одиночное подчёркивание — не линия");
            AssertTrue(!PptxWriter.IsFillerRule(""), "пустая строка — не заполнитель");
            AssertTrue(!PptxWriter.IsFillerRule("   "), "пробелы — не заполнитель");

            PdfPageText page = PptxPage(600, 800);
            page.Paragraphs.Add(PptxPar("__________", 40, 700, 300, 692)); // линия сетки диаграммы
            var mixed = new OcrParagraph { LeftPt = 40, TopPt = 650, RightPt = 300, BottomPt = 640 };
            mixed.Runs.Add(new OcrRun { Text = "12,3", FontSizePt = 10 });
            mixed.Runs.Add(new OcrRun { Text = "______", FontSizePt = 10 });
            page.Paragraphs.Add(mixed);

            XDocument slide = PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp), 1);
            AssertEqual(1, new List<XElement>(slide.Descendants(PptxP + "sp")).Count,
                "абзац из одних прочерков фигуры не даёт");
            AssertEqual("12,3", PptxAllText(slide), "в смешанном абзаце остаётся только текст");
        }

        /// <summary>Таблица rows×cols из пустых ячеек в заданной рамке.</summary>
        private static OcrTable PptxTable(int rows, int cols, double leftPt, double topPt, double rightPt, double bottomPt)
        {
            var table = new OcrTable
            {
                LeftPt = leftPt, TopPt = topPt, RightPt = rightPt, BottomPt = bottomPt
            };
            double colWidth = (rightPt - leftPt) / cols;
            for (int c = 0; c < cols; c++)
                table.ColumnWidthsPt.Add(colWidth);
            for (int r = 0; r < rows; r++)
            {
                var row = new OcrTableRow();
                for (int c = 0; c < cols; c++)
                    row.Cells.Add(new OcrTableCell());
                table.Rows.Add(row);
            }
            return table;
        }

        private static void TestSlideGapBreak()
        {
            OcrLayout.Line first = OcrLine("Заголовок", 60, 500, 300, 480);
            OcrLayout.Line near = OcrLine("подзаголовок", 60, 470, 300, 455);
            OcrLayout.Line far = OcrLine("текст ниже", 60, 300, 300, 285);

            AssertTrue(!OcrLayout.SlideGapBreak(PageLayoutMode.Document, first, far, 12),
                "у документа абсолютного порога нет — там зазор сравнивается с типичным");
            AssertTrue(!OcrLayout.SlideGapBreak(PageLayoutMode.Slide, first, near, 12),
                "соседние строки блока остаются одним абзацем");
            AssertTrue(OcrLayout.SlideGapBreak(PageLayoutMode.Slide, first, far, 12),
                "строки, разнесённые на треть слайда, — разные блоки");
            AssertTrue(!OcrLayout.SlideGapBreak(PageLayoutMode.Slide, first, far, 0),
                "кегль неизвестен — не гадаем");
        }

        /// <summary>Строка разбора из одного слова в заданной рамке.</summary>
        private static OcrLayout.Line OcrLine(string text, double left, double top, double right, double bottom)
        {
            var line = new OcrLayout.Line();
            line.Words.Add(new PdfWord
            {
                Text = text, Left = left, Right = right, Top = top, Bottom = bottom, FontSizePt = 12
            });
            line.MidY = (top + bottom) / 2;
            return line;
        }

        private static void TestPptxSingleLine()
        {
            AssertTrue(PptxWriter.IsSingleLine(12, 12), "высота в кегль — одна строка");
            AssertTrue(PptxWriter.IsSingleLine(18, 12), "полтора кегля — всё ещё одна строка");
            AssertTrue(!PptxWriter.IsSingleLine(30, 12), "две с половиной — уже абзац");
            AssertTrue(PptxWriter.IsSingleLine(0, 12), "вырожденная рамка — не многострочная");

            PdfPageText page = PptxPage(600, 800);
            page.Paragraphs.Add(PptxPar("короткая", 40, 700, 120, 690, 10));      // одна строка
            // Высокая рамка = абзац, который в источнике занимал несколько строк.
            page.Paragraphs.Add(PptxPar("многострочный абзац", 40, 600, 500, 540, 10));

            XDocument slide = PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp), 1);
            var wraps = new List<string>();
            foreach (XElement bodyPr in slide.Descendants(PptxA + "bodyPr"))
                wraps.Add((string)bodyPr.Attribute("wrap"));
            AssertTrue(wraps.Contains("none"), "однострочная — без переноса (иначе слово уедет вниз)");
            AssertTrue(wraps.Contains("square"), "многострочная — с переносом по своей ширине");
        }

        private static void TestPptxParagraphKeptWhole()
        {
            // Абзац из трёх строк с НЕРАВНОМЕРНЫМ шагом: 20 pt между первой и второй, 30 —
            // между второй и третьей (между смысловыми частями отступ больше).
            PdfPageText page = PptxPage(600, 800);
            var par = new OcrParagraph
            {
                LeftPt = 60, TopPt = 700, RightPt = 500, BottomPt = 640,
                Alignment = OcrAlignment.Left, LineCount = 3
            };
            par.LineTopsPt.Add(700);
            par.LineTopsPt.Add(680);
            par.LineTopsPt.Add(650);
            par.MinLeftPt = 40; // вторая и третья строки левее первой — красная строка
            par.Runs.Add(new OcrRun { Text = "первая", FontSizePt = 12 });
            par.Runs.Add(new OcrRun { Text = "вторая", FontSizePt = 12, StartsLine = true });
            par.Runs.Add(new OcrRun { Text = "третья", FontSizePt = 12, StartsLine = true });
            page.Paragraphs.Add(par);

            XDocument slide = PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp), 1);
            var shapes = new List<XElement>(slide.Descendants(PptxP + "sp"));
            AssertEqual(1, shapes.Count, "абзац остаётся ОДНОЙ надписью — иначе его не выделить целиком");

            var paras = new List<XElement>(shapes[0].Descendants(PptxA + "p"));
            AssertEqual(3, paras.Count, "строка источника = свой абзац разметки (у каждой свой шаг)");

            var pitches = new List<string>();
            foreach (XElement spc in slide.Descendants(PptxA + "spcPts"))
                pitches.Add((string)spc.Attribute("val"));
            AssertTrue(pitches.Contains("2000"), "шаг 20 pt перенесён");
            AssertTrue(pitches.Contains("3000"), "увеличенный шаг 30 pt тоже — иначе низ абзаца уедет вверх");

            XElement pPr = PptxFirst(shapes[0].Descendants(PptxA + "pPr"));
            AssertEqual(PptxGeometry.Emu(20).ToString(), (string)pPr.Attribute("indent"),
                "красная строка — отступом первой строки, а не сдвигом всей рамки");
            AssertEqual("перваявтораятретья", PptxAllText(slide), "текст на месте и в порядке");
        }

        private static void TestPptxAscentGap()
        {
            // Значение измерено настоящим PowerPoint (см. PptxWriter.FirstBaselineFactor): без
            // поправки текст садится ниже, чем стоял, и линовка подложки пересекает буквы.
            AssertEqual(9.9, Math.Round(PptxWriter.FirstBaselineDropPt(10), 6), "десятый кегль");
            AssertEqual(23.76, Math.Round(PptxWriter.FirstBaselineDropPt(24), 6), "двадцать четвёртый");
            AssertEqual(PptxWriter.FirstBaselineDropPt(FontResolver.DefaultFontSize), PptxWriter.FirstBaselineDropPt(0),
                "неизвестный кегль — как у кегля по умолчанию");
            // Место ПЕРВОЙ строки от шага не зависит: точный шаг ей не задаётся вовсе, иначе
            // при щедром интерлиньяже абзац садится ниже своего места на треть кегля.
            AssertEqual(19.8, Math.Round(PptxWriter.FirstBaselineDropPt(20), 6), "кегль 20 — и только он");

            // Строка ИЗ ОДНИХ СТРОЧНЫХ букв стоит там же, где строка с заглавной: место берётся
            // от базовой линии, а верх чернил у них разный. Это и был дефект — абзац без
            // заглавных всплывал над своим местом, и подчёркивание подложки резало его пополам.
            var withCaps = new OcrParagraph { TopPt = 500, BottomPt = 480, LeftPt = 10, RightPt = 300 };
            withCaps.LineBaselinesPt.Add(486);
            var lowerOnly = new OcrParagraph { TopPt = 494, BottomPt = 480, LeftPt = 10, RightPt = 300 };
            lowerOnly.LineBaselinesPt.Add(486);
            AssertEqual(PptxWriter.FirstBaselinePt(withCaps, 20), PptxWriter.FirstBaselinePt(lowerOnly, 20),
                "базовая линия одна — и место одно, каким бы ни был верх чернил");

            // Базовой линии нет (линовка, прочерки) — прежний отсчёт от верха чернил.
            var noBaseline = new OcrParagraph { TopPt = 500, BottomPt = 480, LeftPt = 10, RightPt = 300 };
            AssertEqual(500 - 14.0, Math.Round(PptxWriter.FirstBaselinePt(noBaseline, 20), 6),
                "запасной вариант — верх чернил минус высота прописных");
        }


        /// <summary>
        /// Хвостовой пробел рана обязан пережить запись. Без xml:space="preserve" его снимает
        /// уже разбор XML, и соседние слова слипаются: ран кончается там, где меняется формат,
        /// а это ровно граница подчёркнутого куска посреди предложения.
        /// </summary>
        private static void TestPptxKeepsSpaces()
        {
            PdfPageText page = PptxPage(600, 800);
            var par = new OcrParagraph { LeftPt = 10, TopPt = 700, RightPt = 500, BottomPt = 680 };
            par.Runs.Add(new OcrRun { Text = "первое ", FontSizePt = 12, Bold = true });
            par.Runs.Add(new OcrRun { Text = "второе", FontSizePt = 12 });
            page.Paragraphs.Add(par);

            OoxmlPackage package = PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp);
            XDocument slide = PptxSlideXml(package, 1);
            AssertEqual("первое второе", PptxAllText(slide), "пробел на стыке ранов остался");

            string xml = PptxPartText(package, "ppt/slides/slide1.xml");
            AssertTrue(xml.Contains("<a:t xml:space=\"preserve\">"),
                "у текста рана стоит xml:space=\"preserve\" — иначе крайние пробелы снимает разбор");
        }

        /// <summary>
        /// Шаг строк и место абзаца считаются по БАЗОВЫМ линиям. Строка из одних строчных букв
        /// начинается ниже строки с заглавной при той же базовой линии, и отсчёт от верха чернил
        /// поднимал такую строку над своим местом — подчёркивание подложки резало её пополам.
        /// </summary>
        private static void TestPptxPitchFromBaselines()
        {
            PdfPageText page = PptxPage(600, 800);
            var par = new OcrParagraph
            {
                LeftPt = 40, TopPt = 700, RightPt = 500, BottomPt = 660,
                Alignment = OcrAlignment.Left, LineCount = 3, MinLeftPt = 40
            };
            // Верх чернил у второй строки на 4 pt ниже (в ней нет заглавных), базовые линии же
            // идут ровным шагом 20 pt — по ним и должно получиться 20 pt между всеми строками.
            par.LineTopsPt.Add(700);
            par.LineTopsPt.Add(676);
            par.LineTopsPt.Add(660);
            par.LineBaselinesPt.Add(692);
            par.LineBaselinesPt.Add(672);
            par.LineBaselinesPt.Add(652);
            par.Runs.Add(new OcrRun { Text = "Первая", FontSizePt = 12 });
            par.Runs.Add(new OcrRun { Text = "вторая", FontSizePt = 12, StartsLine = true });
            par.Runs.Add(new OcrRun { Text = "третья", FontSizePt = 12, StartsLine = true });
            page.Paragraphs.Add(par);

            XDocument slide = PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp), 1);
            var spacings = new List<XElement>(slide.Descendants(PptxA + "spcPts"));
            AssertEqual(2, spacings.Count, "шаг задан ВТОРОЙ и третьей строкам, а первой — нет");
            foreach (XElement spc in spacings)
                AssertEqual("2000", (string)spc.Attribute("val"),
                    "шаг ровный: чернила второй строки ниже, а базовая линия — на своём месте");

            var shapes = new List<XElement>(slide.Descendants(PptxP + "sp"));
            XElement off = PptxFirst(shapes[0].Descendants(PptxA + "off"));
            double boxTop = 692 + PptxWriter.FirstBaselineDropPt(12);
            AssertEqual(PptxGeometry.Emu(800 - boxTop).ToString(), (string)off.Attribute("y"),
                "верх рамки отсчитан от базовой линии первой строки");
        }

        /// <summary>
        /// Базовая линия слова — медиана по буквам (надстрочный знак не должен её уводить),
        /// и поворот страницы переносит её как ТОЧКУ: у боковой строки она вертикальна.
        /// </summary>
        private static void TestWordBaseline()
        {
            var line = new OcrLayout.Line();
            line.Words.Add(new PdfWord { Text = "раз", Left = 10, Right = 40, Bottom = 100, Top = 112, BaselineYPt = 100 });
            line.Words.Add(new PdfWord { Text = "два", Left = 45, Right = 75, Bottom = 100, Top = 112, BaselineYPt = 100 });
            line.Words.Add(new PdfWord { Text = "1", Left = 76, Right = 80, Bottom = 108, Top = 116, BaselineYPt = 108 });
            AssertEqual(100.0, line.BaselinePt, "надстрочный знак медиану не уводит");

            var empty = new OcrLayout.Line();
            empty.Words.Add(new PdfWord { Text = "___", Left = 0, Right = 30, Bottom = 50, Top = 52 });
            AssertEqual(0.0, empty.BaselinePt, "букв нет — базовой линии нет, а не ноль по ошибке");

            // Поворот на 90°: точка (x, y) страницы 600×800 переходит в (y, x) страницы 800×600.
            var words = new List<PdfWord>
            {
                new PdfWord { Text = "бок", Left = 100, Right = 130, Bottom = 200, Top = 212,
                    BaselineXPt = 100, BaselineYPt = 200 }
            };
            double w = 600, h = 800;
            PageRotation.RotatePage(words, null, null, 90, ref w, ref h);
            AssertTrue(words[0].BaselineXPt != 100 || words[0].BaselineYPt != 200,
                "базовая линия повёрнута вместе со словом, а не осталась в старом пространстве");
            AssertTrue(words[0].BaselineYPt >= words[0].Bottom - 0.001 && words[0].BaselineYPt <= words[0].Top + 0.001,
                "после разворота базовая линия лежит внутри рамки слова");
        }


        /// <summary>
        /// Значки инструментов различаются НЕ ТОЛЬКО рисунком внутри листа: карточки стоят
        /// рядом, и первым читается цветовое пятно. Тёмный оранжево-красный PowerPoint рядом
        /// с красным PDF выглядел «ещё одним красным» — цвета обязаны расходиться заметно.
        /// Порог по сумме модулей каналов: 90 из 765 — примерно граница, на которой два пятна
        /// перестают читаться как один цвет.
        /// </summary>
        private static void TestGlyphColoursDistinct()
        {
            var palette = new List<KeyValuePair<string, System.Drawing.Color>>
            {
                new KeyValuePair<string, System.Drawing.Color>("PDF", System.Drawing.Color.FromArgb(211, 47, 47)),
                new KeyValuePair<string, System.Drawing.Color>("Excel", Theme.Accent),
                new KeyValuePair<string, System.Drawing.Color>("Word", Theme.WordViolet),
                new KeyValuePair<string, System.Drawing.Color>("PowerPoint", Theme.PowerPointOrange),
                new KeyValuePair<string, System.Drawing.Color>("разделы", Theme.HubBlue)
            };
            var offenders = new List<string>();
            for (int i = 0; i < palette.Count; i++)
                for (int j = i + 1; j < palette.Count; j++)
                {
                    System.Drawing.Color a = palette[i].Value, b = palette[j].Value;
                    int distance = Math.Abs(a.R - b.R) + Math.Abs(a.G - b.G) + Math.Abs(a.B - b.B);
                    if (distance < 90)
                        offenders.Add(palette[i].Key + " и " + palette[j].Key + ": " + distance);
                }
            AssertTrue(offenders.Count == 0, "значки сливаются по цвету: " + string.Join(" | ", offenders.ToArray()));

            // Полоса шапки инструмента — своя, ТЁМНАЯ: по ней идёт белый текст, и подпись под
            // заголовком требует контраста, которого яркий оранжевый не даёт.
            AssertTrue(Brightness(Theme.PowerPointBand) < Brightness(Theme.PowerPointOrange),
                "полоса шапки темнее значка — иначе белая подпись на ней не читается");
        }

        /// <summary>Воспринимаемая яркость цвета (ITU-R BT.601) — для проверок контраста.</summary>
        private static double Brightness(System.Drawing.Color c)
        {
            return 0.299 * c.R + 0.587 * c.G + 0.114 * c.B;
        }


        /// <summary>
        /// Размеченный PDF несёт готовые границы абзацев, и спорить с ними зазорами не нужно:
        /// строки ОДНОГО блока остаются одним абзацем, даже когда между ними отбивка, по которой
        /// геометрия разорвала бы. Обратного разметка не делает никогда — блок на строку (её
        /// пишут и такую) не должен дробить сильнее прежнего. И геометрия сохраняет право вето:
        /// строки, разнесённые на полстраницы, разные абзацы при любом номере.
        /// </summary>
        private static void TestStructureBlocksGroup()
        {
            // Красная строка у третьей строки: без разметки это граница абзаца.
            AssertEqual(2, ParagraphsOf(-1, -1, -1, -1), "без разметки красная строка делит абзац — как и раньше");
            AssertEqual(1, ParagraphsOf(7, 7, 7, 7), "один блок разметки — один абзац, отступ не в счёт");
            AssertEqual(2, ParagraphsOf(7, 7, 8, 8), "разные блоки — разные абзацы");
            AssertEqual(ParagraphsOf(-1, -1, -1, -1), ParagraphsOf(1, 2, 3, 4),
                "блок на строку разрывов НЕ добавляет: результат тот же, что и вовсе без разметки");

            // Вето геометрии — чистое условие, проверяем его напрямую: до трёх кеглей разметка
            // разрыв снимает, дальше решает расстояние.
            OcrLayout.Line a = OcrLine("первая", 0, 200, 60, 190);
            OcrLayout.Line near = OcrLine("вторая", 0, 176, 60, 166);
            OcrLayout.Line far = OcrLine("третья", 0, 140, 60, 130);
            a.Words[0].BlockId = near.Words[0].BlockId = far.Words[0].BlockId = 5;
            AssertTrue(OcrLayout.SameBlock(a, near, 10), "24 pt при кегле 10 — тот же блок");
            AssertTrue(!OcrLayout.SameBlock(a, far, 10),
                "60 pt при кегле 10 — геометрия главнее: разметка описывает логику, а не место");
            far.Words[0].BlockId = 6;
            AssertTrue(!OcrLayout.SameBlock(a, far, 10), "разные номера — и подавно не один блок");

            // Куски ОДНОЙ физической строки (широкий зазор внутри строки) тоже несут общий
            // номер. Слить их значит поставить друг под друга то, что стояло в строку.
            OcrLayout.Line beside = OcrLine("рядом", 300, 200, 360, 190);
            beside.Words[0].BlockId = 5;
            AssertTrue(!OcrLayout.SameBlock(a, beside, 10), "куски одной строки стоят рядом, а не одна под другой");
        }

        /// <summary>Четыре строки, у третьей красная строка; номера блоков задаются по порядку.</summary>
        private static int ParagraphsOf(int a, int b, int c, int d)
        {
            var words = new List<PdfWord>
            {
                W("первая", 0, 200, 60, 10),
                W("вторая", 0, 188, 60, 10),
                W("третья", 15, 176, 45, 10), // красная строка — сигнал нового абзаца
                W("четвёртая", 0, 164, 60, 10)
            };
            int[] ids = { a, b, c, d };
            for (int i = 0; i < words.Count; i++)
                words[i].BlockId = ids[i];
            return OcrLayout.Analyze(words).Paragraphs.Count;
        }

        /// <summary>
        /// На слайде место строки задаёт её собственный отступ, а не выключка: рамка одна на
        /// весь абзац, её ширина известна с точностью до запаса на чужие метрики, и «по центру»
        /// ставило бы строку в середину этой неточной рамки. Именно так центрированный
        /// заголовок внутри абзаца уезжал в сторону.
        /// </summary>
        private static void TestPptxPerLineIndent()
        {
            PdfPageText page = PptxPage(600, 800);
            var par = new OcrParagraph
            {
                LeftPt = 100, TopPt = 700, RightPt = 400, BottomPt = 660,
                Alignment = OcrAlignment.Center, LineCount = 3, MinLeftPt = 100
            };
            par.LineTopsPt.Add(700); par.LineTopsPt.Add(680); par.LineTopsPt.Add(660);
            par.LineBaselinesPt.Add(692); par.LineBaselinesPt.Add(672); par.LineBaselinesPt.Add(652);
            par.LineLeftsPt.Add(100); par.LineLeftsPt.Add(140); par.LineLeftsPt.Add(120);
            par.Runs.Add(new OcrRun { Text = "первая", FontSizePt = 12 });
            par.Runs.Add(new OcrRun { Text = "вторая", FontSizePt = 12, StartsLine = true });
            par.Runs.Add(new OcrRun { Text = "третья", FontSizePt = 12, StartsLine = true });
            page.Paragraphs.Add(par);

            XDocument slide = PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp), 1);
            var props = new List<XElement>(slide.Descendants(PptxA + "pPr"));
            AssertEqual(3, props.Count, "у каждой строки свои свойства абзаца");
            foreach (XElement p in props)
                AssertEqual("l", (string)p.Attribute("algn"),
                    "выключки нет: место задаёт отступ, а не середина неточной рамки");
            AssertEqual(null, (string)props[0].Attribute("indent"), "первая строка стоит у края рамки");
            AssertEqual(PptxGeometry.Emu(40).ToString(), (string)props[1].Attribute("indent"),
                "вторая строка отодвинута на свои 40 pt");
            AssertEqual(PptxGeometry.Emu(20).ToString(), (string)props[2].Attribute("indent"),
                "третья — на свои 20 pt");
        }


        /// <summary>
        /// «Нет места на диске» без указания диска — половина ответа: результат пишется в одну
        /// папку, временные файлы — в другую, на другом диске, и переполниться может любой.
        /// </summary>
        private static void TestDiskFullMessage()
        {
            var full = new IOException("There is not enough space on the disk.");
            SetHResult(full, unchecked((int)0x80070070));
            AssertTrue(DiskSpace.IsFull(full), "код 0x70 — это нехватка места");
            AssertTrue(DiskSpace.IsFull(new IOException("обёртка", full)), "причина ищется и во вложенных");
            AssertTrue(!DiskSpace.IsFull(new IOException("файл занят")), "обычная ошибка ввода-вывода — не про место");
            AssertTrue(!DiskSpace.IsFull(new InvalidOperationException()), "и не всякое исключение вообще");

            string path = Path.Combine(Path.GetTempPath(), "iwo_disk.pdf");
            string message = DiskSpace.Describe(full, path);
            string root = Path.GetPathRoot(Path.GetFullPath(path));
            AssertTrue(message.Contains(root), "в сообщении назван диск: " + message);
            AssertTrue(message != full.Message, "сообщение системы дополнено, а не повторено");
            AssertEqual("файл занят", DiskSpace.Describe(new IOException("файл занят"), path),
                "прочие ошибки передаются как есть");
            AssertTrue(DiskSpace.FreeMegabytes(path) >= 0, "остаток на своём же диске измеряется");
        }

        /// <summary>HResult у Exception доступен на запись только защищённо — ставим отражением.</summary>
        private static void SetHResult(Exception ex, int value)
        {
            typeof(Exception).GetProperty("HResult",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic |
                System.Reflection.BindingFlags.Public).SetValue(ex, value, null);
        }


        /// <summary>
        /// Смешанный документ — часть слайдов «напечатана» картинкой целиком. Страница в
        /// результате есть, а текста на ней нет, и без объяснения это выглядит как поломка
        /// перевода. Скан целиком отклоняется заранее, смешанный документ через тот заслон
        /// проходит — значит сказать должен результат.
        /// </summary>
        private static void TestConvertDoneStatus()
        {
            const string fmt = "Готово: страниц {0}.";
            AssertEqual("Готово: страниц 13.",
                PdfConvertFormBase.DoneStatus(new ConvertResult { Pages = 13, PagesWithText = 13 }, fmt),
                "текст нашёлся везде — лишнего не говорим");
            string mixed = PdfConvertFormBase.DoneStatus(new ConvertResult { Pages = 13, PagesWithText = 6 }, fmt);
            AssertTrue(mixed.StartsWith("Готово: страниц 13."), "результат по-прежнему успешный");
            AssertTrue(mixed.Contains("7"), "названо число страниц без текста: " + mixed);
            AssertTrue(mixed.Length > fmt.Length + 6, "объяснение добавлено, а не только цифра");
        }


        /// <summary>
        /// Отмена на фазе СЖАТИЯ. Сорок частей — сорок запусков Ghostscript, дольше самой
        /// нарезки; раньше на этой фазе кнопку «Отмена» убирали, считая точкой невозврата момент
        /// записи файлов. Раз остановка убирает за собой, невозврата нет. Проверяем через
        /// настоящее окно: тут важна не арифметика, а то, что кнопка на месте и нажатие доходит.
        /// </summary>
        private static void TestSplitCancelDuringCompressionLive()
        {
            var offenders = new List<string>();
            string dir = Path.Combine(Path.GetTempPath(), "iwo_splitcancel_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = MakePagesPdf(dir, 8);
                string outDir = Path.Combine(dir, "out");
                Directory.CreateDirectory(outDir);

                RunSta(delegate
                {
                    using (var form = new PdfSplitForm())
                    {
                        IntPtr handle = form.Handle;
                        if (handle == IntPtr.Zero) { offenders.Add("окно не создалось"); return; }
                        form.Show();
                        Pump(form, 400);

                        // Работа уже ЗАВЕРШЕНА к моменту вызова: проверяем именно фазу сжатия —
                        // отмену, нажатую после того, как файлы записаны.
                        var files = new List<string>();
                        for (int i = 0; i < 4; i++)
                        {
                            string part = Path.Combine(outDir, "part" + i + ".pdf");
                            File.Copy(src, part, true);
                            files.Add(part);
                        }
                        System.Reflection.MethodInfo run = typeof(PdfSplitForm).GetMethod("RunSplit",
                            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                        if (run == null) { offenders.Add("RunSplit не найден"); return; }

                        // Нажимаем «Отмена» до старта: предикат читается перед КАЖДОЙ частью,
                        // поэтому не сжата будет ни одна, а папка останется пустой.
                        System.Reflection.FieldInfo flag = typeof(PdfToolFormBase).GetField("_cancelRequested",
                            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                        if (flag == null) { offenders.Add("флага отмены нет"); return; }
                        Func<Action<int, int>, List<string>> work = delegate(Action<int, int> pr)
                        {
                            flag.SetValue(form, true); // «Отмена» нажата ровно на границе фаз
                            return files;
                        };
                        run.Invoke(form, new object[] { work, CompressionLevel.Small,
                            outDir, true, (Action)delegate { }, 8 });
                        Pump(form, 4000);

                        string[] left = Directory.GetFiles(outDir);
                        if (left.Length != 0)
                            offenders.Add("после отмены на сжатии осталось файлов: " + left.Length);
                    }
                });
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
            AssertTrue(offenders.Count == 0, "отмена на сжатии: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>Прокачать очередь сообщений окна — фоновая работа возвращается через неё.</summary>
        private static void Pump(System.Windows.Forms.Form form, int ms)
        {
            for (int i = 0; i < ms; i += 30)
            {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(30);
            }
        }


        /// <summary>
        /// В репозитории не должно остаться ни рабочих упоминаний, ни следов машины, на
        /// которой он собирался: имя учётной записи в пути, название рабочей папки, рабочее
        /// слово в комментарии.
        ///
        /// Слова сверяем ПО ОТПЕЧАТКАМ: список из тех самых слов, лежащий в открытом
        /// репозитории, — это ровно то, чему в нём быть и не должно. Отпечаток берётся от
        /// начала слова, поэтому одна запись накрывает все склонения и производные.
        /// </summary>
        private static void TestNoWorkReferencesInRepo()
        {
            string root = RepoRoot();
            if (root == null) { AssertTrue(true, "исходники рядом не найдены — пропускаем"); return; }
            var banned = new HashSet<string>(new[]
            {
                "610a4b77c77b3ffd",
                "2c0cda61e536f2c5",
                "fe4e39d2d1e3c6b5",
                "2821bc2655a8343f",
                "d74a2fa10ca4be11",
                "f3dcd5e6ce37de3e",
                "d028554a0a7b4514",
                "e994a1c965fcf540",
                "6d54019640cdd227",
                "41ad327734850b89",
                "d197f4a34fadf25e",
                "dab70bc2d883a287"
            });
            var offenders = new List<string>();
            string[] folders = { "src", "tests", "tools", "docs", "installer", ".github" };
            string[] extensions = { ".cs", ".ps1", ".py", ".md", ".cmd", ".yml", ".iss", ".json", ".txt" };
            foreach (string folder in folders)
            {
                string dir = Path.Combine(root, folder);
                if (!Directory.Exists(dir)) continue;
                foreach (string file in Directory.GetFiles(dir, "*.*", SearchOption.AllDirectories))
                {
                    if (Skipped(file)) continue;
                    bool wanted = false;
                    foreach (string ext in extensions)
                        if (file.EndsWith(ext, StringComparison.OrdinalIgnoreCase)) { wanted = true; break; }
                    if (!wanted) continue;
                    string word = FirstBannedWord(File.ReadAllText(file).ToLowerInvariant(), banned);
                    if (word != null)
                        offenders.Add(file.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar) +
                            ": " + Mask(word));
                }
            }
            AssertTrue(offenders.Count == 0,
                "рабочие упоминания или путь машины в репозитории: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>Служебные и собранные каталоги проверять незачем — их на гите нет.</summary>
        private static bool Skipped(string file)
        {
            string[] parts = { "bin", "obj", "shots-ru", "shots-en", "samples-ru", "samples-en", "out" };
            foreach (string part in parts)
                if (file.IndexOf(Path.DirectorySeparatorChar + part + Path.DirectorySeparatorChar,
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            return false;
        }

        /// <summary>
        /// Первое слово текста, начало которого совпало с запретным отпечатком (или null).
        /// Проверяем именно НАЧАЛА слов: так одна запись накрывает все склонения и
        /// производные разом, а «администратор» не выдаётся за что-то запретное.
        /// </summary>
        private static string FirstBannedWord(string text, HashSet<string> banned)
        {
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                var word = new System.Text.StringBuilder();
                for (int i = 0; i <= text.Length; i++)
                {
                    char c = i < text.Length ? text[i] : ' ';
                    if (char.IsLetter(c) || c == '.') { word.Append(c); continue; }
                    if (word.Length >= 3)
                    {
                        string whole = word.ToString();
                        for (int size = 3; size <= Math.Min(whole.Length, 14); size++)
                        {
                            byte[] hash = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(whole.Substring(0, size)));
                            var hex = new System.Text.StringBuilder();
                            for (int b = 0; b < 8; b++)
                                hex.Append(hash[b].ToString("x2"));
                            if (banned.Contains(hex.ToString()))
                                return whole;
                        }
                    }
                    word.Length = 0;
                }
            }
            return null;
        }

        /// <summary>Найденное слово в отчёте показываем ЗАКРЫТЫМ: иначе отчёт сам его и разгласит.</summary>
        private static string Mask(string word)
        {
            return word.Substring(0, 1) + new string('*', Math.Max(1, word.Length - 1));
        }

        /// <summary>
        /// В исходниках не должно быть управляющих символов, кроме табуляции и перевода строки.
        /// Повод не выдуманный: в пути к проверке PPTX внутри полного прогона однажды оказалась
        /// ВЕРТИКАЛЬНАЯ ТАБУЛЯЦИЯ вместо « обратной косой с «v» » — путь перестал существовать, и этот шаг пирамиды
        /// не выполнялся вовсе, молча. Такой символ не виден ни в редакторе, ни в обзоре правок,
        /// поэтому ловить его должен тест, а не глаз.
        /// </summary>
        private static void TestNoControlCharactersInSources()
        {
            string root = RepoRoot();
            if (root == null) { AssertTrue(true, "исходники рядом не найдены — пропускаем"); return; }
            var offenders = new List<string>();
            string[] folders = { "src", "tests", "tools", "docs" };
            string[] extensions = { ".cs", ".ps1", ".py", ".md", ".cmd", ".yml", ".iss" };
            foreach (string folder in folders)
            {
                string dir = Path.Combine(root, folder);
                if (!Directory.Exists(dir)) continue;
                foreach (string file in Directory.GetFiles(dir, "*.*", SearchOption.AllDirectories))
                {
                    if (file.IndexOf(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar,
                            StringComparison.OrdinalIgnoreCase) >= 0
                        || file.IndexOf(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar,
                            StringComparison.OrdinalIgnoreCase) >= 0)
                        continue;
                    bool wanted = false;
                    foreach (string ext in extensions)
                        if (file.EndsWith(ext, StringComparison.OrdinalIgnoreCase)) { wanted = true; break; }
                    if (!wanted) continue;
                    string text = File.ReadAllText(file);
                    for (int i = 0; i < text.Length; i++)
                    {
                        char c = text[i];
                        if (c >= ' ' || c == (char)9 || c == (char)13 || c == (char)10) continue;
                        offenders.Add(Path.GetFileName(file) + ": U+" + ((int)c).ToString("X4"));
                        break;
                    }
                }
            }
            AssertTrue(offenders.Count == 0, "управляющие символы в исходниках: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>Корень репозитория относительно собранного теста; null — не нашли.</summary>
        private static string RepoRoot()
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            for (int up = 0; up < 6 && dir != null; up++)
            {
                if (Directory.Exists(Path.Combine(dir, "src")) && Directory.Exists(Path.Combine(dir, "tests")))
                    return dir;
                DirectoryInfo parent = Directory.GetParent(dir);
                dir = parent == null ? null : parent.FullName;
            }
            return null;
        }


        /// <summary>
        /// Оба перевода — бета, и сказано об этом должно быть на ОБОИХ языках и в обоих
        /// местах: на карточке инструмента (там выбирают) и в самом окне (там работают).
        /// Причина названа прямо: берётся текстовый слой PDF, а отсканированную страницу
        /// читать нечем — без этой строки пустой слайд читается как поломка программы.
        /// </summary>
        private static void TestBetaNotice()
        {
            Lang was = Loc.Current;
            try
            {
                var offenders = new List<string>();
                foreach (Lang lang in new[] { Lang.Ru, Lang.En })
                {
                    Loc.Init(lang);
                    string mark = lang == Lang.Ru ? "(бета)" : "(beta)";
                    foreach (string key in new[] { "hub.ocr.name", "hub.pptx.name" })
                        if (!Loc.T(key).EndsWith(mark))
                            offenders.Add(lang + " " + key + ": нет пометки " + mark);
                    string notice = Loc.T("convert.beta");
                    if (notice.Length < 40)
                        offenders.Add(lang + ": примечание пустое или слишком короткое");
                    // Названо и чего не хватает, и что это не навсегда.
                    if (notice.IndexOf("OCR", StringComparison.OrdinalIgnoreCase) < 0)
                        offenders.Add(lang + ": не сказано, что нужно распознавание");
                    string future = lang == Lang.Ru ? "версия" : "version";
                    if (notice.IndexOf(future, StringComparison.OrdinalIgnoreCase) < 0)
                        offenders.Add(lang + ": не сказано, что это планируется");
                }
                AssertTrue(offenders.Count == 0, "пометка беты: " + string.Join(" | ", offenders.ToArray()));
            }
            finally { Loc.Init(was); }
        }


        /// <summary>
        /// Примечание «бета» — длинная строка, и на узком окне ей нужно больше строк, чем на
        /// широком. Раньше высота была задана числом (две строки), и на минимальной ширине хвост
        /// уходил под сетку: человек читал обрезанную фразу и не узнавал главного — что скан
        /// переводить нечем. Проверяем сам расчёт: чем уже, тем выше, и тело окна съезжает ровно
        /// на столько, сколько заняло примечание.
        /// </summary>
        private static void TestBetaNoteWraps()
        {
            Lang was = Loc.Current;
            try
            {
                using (var font = new System.Drawing.Font("Segoe UI", 9.75f))
                {
                    var offenders = new List<string>();
                    foreach (Lang lang in new[] { Lang.Ru, Lang.En })
                    {
                        Loc.Init(lang);
                        string notice = Loc.T("convert.beta");
                        int wide = PdfConvertFormBase.NoteHeight(notice, font, 1200);
                        int narrow = PdfConvertFormBase.NoteHeight(notice, font, 420);
                        int tiny = PdfConvertFormBase.NoteHeight(notice, font, 240);
                        if (narrow <= wide)
                            offenders.Add(lang + ": узкое окно не добавило строк (" + wide + " -> " + narrow + ")");
                        if (tiny <= narrow)
                            offenders.Add(lang + ": совсем узкое окно не добавило строк");
                        if (wide < font.Height)
                            offenders.Add(lang + ": высота меньше строки текста");
                        // Тело окна съезжает вниз ровно на высоту примечания: иначе сетка
                        // наехала бы на него (или под ним осталась бы пустая полоса).
                        int topNarrow = PdfConvertFormBase.ContentTop(24, narrow);
                        int topWide = PdfConvertFormBase.ContentTop(24, wide);
                        if (topNarrow - topWide != narrow - wide)
                            offenders.Add(lang + ": тело окна съехало не на высоту примечания");
                    }
                    AssertTrue(offenders.Count == 0, "перенос примечания: " + string.Join(" | ", offenders.ToArray()));
                    AssertEqual(0, PdfConvertFormBase.NoteHeight(null, null, 100), "без текста и шрифта — ноль");
                }
            }
            finally { Loc.Init(was); }
        }

        /// <summary>
        /// То же, но на ЖИВЫХ окнах и на их минимальном размере: в расчёте можно ошибиться
        /// шрифтом или шириной, и единственная надёжная проверка — спросить у самой надписи,
        /// поместился ли в неё текст. Заодно убеждаемся, что сетка стоит НИЖЕ примечания.
        ///
        /// Второй проход — с КРУПНЫМ шрифтом: именно так дефект и увидели. При постоянной
        /// высоте в две строки достаточно шрифта побольше (масштаб экрана, системная настройка
        /// размера текста), чтобы третья строка ушла под сетку, а окно ни в чём не признавалось.
        /// </summary>
        private static void TestBetaNoteFitsLive()
        {
            var offenders = new List<string>();
            InIsolatedSettings("iwo_beta_", delegate
            {
                Action back = delegate { };
                foreach (System.Windows.Forms.Form f in new System.Windows.Forms.Form[] { new OcrForm(back), new PptxForm(back) })
                    using (f)
                    {
                        string where = f.GetType().Name;
                        try
                        {
                            f.Show();
                            f.Size = f.MinimumSize; // пользователь дотащил рамку до упора
                            f.PerformLayout();
                            CheckBetaNote(f, where + " (обычный шрифт)", offenders);

                            using (var big = new System.Drawing.Font("Segoe UI", 14f))
                            {
                                f.Font = big; // крупный системный шрифт: строк становится больше
                                f.PerformLayout();
                                CheckBetaNote(f, where + " (крупный шрифт)", offenders);
                            }
                            f.Close();
                        }
                        catch (Exception ex) { offenders.Add(where + " — " + Root(ex).Message); }
                    }
            });
            AssertTrue(offenders.Count == 0, "полоска «бета»: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>Примечание «бета» цело и не перекрыто сеткой — на текущей раскладке окна.</summary>
        private static void CheckBetaNote(System.Windows.Forms.Form f, string where, List<string> offenders)
        {
            System.Windows.Forms.Label note = FindLabel(f, Loc.T("convert.beta"));
            if (note == null)
            {
                offenders.Add(where + ": примечания «бета» нет в окне");
                return;
            }
            int needed = System.Windows.Forms.TextRenderer.MeasureText(note.Text, note.Font,
                new System.Drawing.Size(note.Width, int.MaxValue),
                System.Windows.Forms.TextFormatFlags.WordBreak).Height;
            if (note.Height < needed)
                offenders.Add(where + ": текст обрезан (высота " + note.Height +
                    " при нужных " + needed + ", ширина " + note.Width + ")");
            System.Windows.Forms.Control grid = FindPageGrid(f);
            if (grid == null)
                offenders.Add(where + ": сетки страниц нет в окне");
            else if (grid.Top < note.Bottom)
                offenders.Add(where + ": сетка наехала на примечание (" +
                    grid.Top + " < " + note.Bottom + ")");
        }

        /// <summary>Надпись с таким текстом (вглубь по дереву) или null.</summary>
        private static System.Windows.Forms.Label FindLabel(System.Windows.Forms.Control root, string text)
        {
            foreach (System.Windows.Forms.Control c in root.Controls)
            {
                var l = c as System.Windows.Forms.Label;
                if (l != null && l.Text == text)
                    return l;
                System.Windows.Forms.Label deeper = FindLabel(c, text);
                if (deeper != null)
                    return deeper;
            }
            return null;
        }

        /// <summary>Сетка страниц окна (вглубь по дереву) или null.</summary>
        private static PdfPageGrid FindPageGrid(System.Windows.Forms.Control root)
        {
            foreach (System.Windows.Forms.Control c in root.Controls)
            {
                var grid = c as PdfPageGrid;
                if (grid != null)
                    return grid;
                PdfPageGrid deeper = FindPageGrid(c);
                if (deeper != null)
                    return deeper;
            }
            return null;
        }

        /// <summary>
        /// «Прочие операции» применяют действия к документу ТАКИМ, КАК ОН СОБРАН в сетке, и весь
        /// выбор «копировать исходник или собрать заново» держится на этой проверке. Ошибись она
        /// в сторону «как есть» — и убранная страница вернулась бы в результат, а поворот пропал.
        /// </summary>
        private static void TestAssembledPristine()
        {
            const string src = @"C:\docs\a.pdf";
            var whole = new List<PdfPageRef>
            {
                new PdfPageRef { SourcePath = src, PageIndex = 0 },
                new PdfPageRef { SourcePath = src, PageIndex = 1 },
                new PdfPageRef { SourcePath = src, PageIndex = 2 }
            };
            AssertTrue(PdfSingleDocFormBase.IsPristine(whole, src, 3), "не тронутый набор — это сам файл");
            AssertTrue(PdfSingleDocFormBase.IsPristine(whole, src.ToUpperInvariant(), 3),
                "регистр пути не делает файл другим");

            var removed = new List<PdfPageRef> { whole[0], whole[1] };
            AssertTrue(!PdfSingleDocFormBase.IsPristine(removed, src, 3), "убрали страницу — уже не файл");

            var reordered = new List<PdfPageRef> { whole[1], whole[0], whole[2] };
            AssertTrue(!PdfSingleDocFormBase.IsPristine(reordered, src, 3), "переставили — уже не файл");

            var rotated = new List<PdfPageRef>
            {
                new PdfPageRef { SourcePath = src, PageIndex = 0, Rotation = 90 },
                whole[1], whole[2]
            };
            AssertTrue(!PdfSingleDocFormBase.IsPristine(rotated, src, 3), "повернули — уже не файл");

            var foreign = new List<PdfPageRef>
            {
                whole[0], whole[1], new PdfPageRef { SourcePath = @"C:\docs\b.pdf", PageIndex = 0 }
            };
            AssertTrue(!PdfSingleDocFormBase.IsPristine(foreign, src, 3), "чужая страница в наборе — уже не файл");
            AssertTrue(!PdfSingleDocFormBase.IsPristine(null, src, 3), "пустой набор не выдаём за файл");
            AssertTrue(!PdfSingleDocFormBase.IsPristine(whole, null, 3), "без открытого файла сравнивать нечего");
        }

        /// <summary>
        /// Постраничные действия (картинки) читают прямо из исходника, пока страницы идут из него
        /// по возрастанию и без поворотов: так в именах файлов остаются НАСТОЯЩИЕ номера страниц.
        /// Стоит переставить или повернуть — читать нужно из собранного, иначе на диск ляжет
        /// не то, что показано.
        /// </summary>
        private static void TestPlainSubset()
        {
            const string src = @"C:\docs\a.pdf";
            Func<int, int, PdfPageRef> page = delegate(int index, int rotation)
            {
                return new PdfPageRef { SourcePath = src, PageIndex = index, Rotation = rotation };
            };
            AssertTrue(PdfOpsForm.IsPlainSubset(new List<PdfPageRef> { page(0, 0), page(2, 0) }, src),
                "выделенные подряд возрастающие страницы читаются из исходника");
            AssertTrue(!PdfOpsForm.IsPlainSubset(new List<PdfPageRef> { page(2, 0), page(0, 0) }, src),
                "переставленные — только через сборку");
            AssertTrue(!PdfOpsForm.IsPlainSubset(new List<PdfPageRef> { page(0, 90) }, src),
                "повёрнутая — только через сборку");
            AssertTrue(!PdfOpsForm.IsPlainSubset(new List<PdfPageRef> { page(1, 0), page(1, 0) }, src),
                "повторённая страница — только через сборку");
            AssertTrue(!PdfOpsForm.IsPlainSubset(new List<PdfPageRef>
                { page(0, 0), new PdfPageRef { SourcePath = @"C:\docs\b.pdf", PageIndex = 0 } }, src),
                "страница из другого файла — только через сборку");
            AssertTrue(!PdfOpsForm.IsPlainSubset(new List<PdfPageRef>(), src), "пустой набор читать нечем");
        }

        /// <summary>
        /// Что окна принимают перетаскиванием. «Прочие операции» собирают документ и из картинок,
        /// а остальные инструменты — только из PDF: расширение снимка, попавшее в общий набор,
        /// молча уронило бы дроп в «Объединении» на разборе «PDF».
        /// </summary>
        private static void TestDropExtensions()
        {
            AssertTrue(PdfDrop.Matches("a.pdf", PdfDrop.PdfOnly), "PDF принимается всегда");
            AssertTrue(PdfDrop.Matches("A.PDF", PdfDrop.PdfOnly), "регистр расширения не важен");
            AssertTrue(!PdfDrop.Matches("photo.jpg", PdfDrop.PdfOnly), "картинка не PDF-инструментам");
            AssertTrue(PdfDrop.Matches("photo.JPG", PdfDrop.PdfAndImages), "картинка — «Прочим операциям»");
            AssertTrue(PdfDrop.Matches("scan.tiff", PdfDrop.PdfAndImages), "многостраничный TIFF тоже");
            AssertTrue(PdfDrop.Matches("a.pdf", PdfDrop.PdfAndImages), "и PDF в том же наборе");
            AssertTrue(!PdfDrop.Matches("photo.heic", PdfDrop.PdfAndImages),
                "HEIC не обещаем: GDI+ его не читает");
            AssertTrue(!PdfDrop.Matches(null, PdfDrop.PdfAndImages), "пустой путь не подходит");
            AssertTrue(ImageToPdfService.IsImage("a.png") && !ImageToPdfService.IsImage("a.pdf"),
                "картинка отличается от документа по расширению");
        }

        /// <summary>
        /// Картинка на листе: лист берёт ориентацию картинки, картинка вписана в поля целиком,
        /// по центру и БЕЗ искажения пропорций. Растянутый снимок — брак, который видно сразу.
        /// </summary>
        private static void TestImageLayoutOnPage()
        {
            double pageW, pageH, x, y, w, h;
            ImageToPdfService.Layout(1200, 1600, out pageW, out pageH, out x, out y, out w, out h);
            AssertTrue(Math.Abs(pageW - ImageToPdfService.A4WidthPt) < 0.01 &&
                Math.Abs(pageH - ImageToPdfService.A4HeightPt) < 0.01, "высокая картинка — книжный лист");
            AssertTrue(Math.Abs((w / h) - 0.75) < 0.001, "пропорции сохранены (1200x1600)");
            AssertTrue(x >= ImageToPdfService.MarginPt - 0.01 && y >= ImageToPdfService.MarginPt - 0.01,
                "картинка не заезжает в поля");
            AssertTrue(Math.Abs(x - (pageW - w) / 2) < 0.01 && Math.Abs(y - (pageH - h) / 2) < 0.01,
                "картинка по центру листа");
            // Вписана ЦЕЛИКОМ: одна сторона занимает поле полностью, вторая в него не выходит.
            // У снимка 3:4 на A4 (1:1,41) упор приходится в ширину, а не в высоту: лист
            // «длиннее» кадра, и требовать полной высоты значило бы требовать растянуть кадр.
            AssertTrue(FillsMarginBox(w, h, pageW, pageH), "картинка вписана в поля целиком (книжный лист)");

            ImageToPdfService.Layout(1600, 1200, out pageW, out pageH, out x, out y, out w, out h);
            AssertTrue(pageW > pageH, "широкая картинка — альбомный лист");
            AssertTrue(FillsMarginBox(w, h, pageW, pageH), "картинка вписана в поля целиком (альбомный лист)");

            // Узкая полоска упирается в высоту — тот же инвариант с другой стороны.
            ImageToPdfService.Layout(200, 1600, out pageW, out pageH, out x, out y, out w, out h);
            AssertTrue(FillsMarginBox(w, h, pageW, pageH), "узкая полоска вписана в поля целиком");
            AssertTrue(Math.Abs(h - (pageH - 2 * ImageToPdfService.MarginPt)) < 0.01,
                "у высокой полоски упор в высоту листа");

            ImageToPdfService.Layout(0, 0, out pageW, out pageH, out x, out y, out w, out h);
            AssertTrue(w > 0 && h > 0 && !double.IsNaN(w), "вырожденный кадр не делит на ноль");
        }

        /// <summary>
        /// Картинка вписана в поля листа целиком: хотя бы одна сторона упирается в поле, и ни
        /// одна за него не выходит. Именно это и значит «во весь лист без искажения пропорций».
        /// </summary>
        private static bool FillsMarginBox(double w, double h, double pageW, double pageH)
        {
            double boxW = pageW - 2 * ImageToPdfService.MarginPt;
            double boxH = pageH - 2 * ImageToPdfService.MarginPt;
            if (w > boxW + 0.01 || h > boxH + 0.01)
                return false;
            return Math.Abs(w - boxW) < 0.01 || Math.Abs(h - boxH) < 0.01;
        }

        /// <summary>
        /// EXIF-ориентация: GDI+ её НЕ применяет, и снимок с телефона ложится в PDF боком.
        /// Таблица из стандарта — восемь значений, всё непонятное считается «как есть».
        /// </summary>
        private static void TestImageExifRotation()
        {
            AssertEqual(System.Drawing.RotateFlipType.RotateNoneFlipNone.ToString(),
                ImageToPdfService.ExifRotation(1).ToString(), "1 — как снято");
            AssertEqual(System.Drawing.RotateFlipType.Rotate90FlipNone.ToString(),
                ImageToPdfService.ExifRotation(6).ToString(), "6 — вправо на 90");
            AssertEqual(System.Drawing.RotateFlipType.Rotate270FlipNone.ToString(),
                ImageToPdfService.ExifRotation(8).ToString(), "8 — влево на 90");
            AssertEqual(System.Drawing.RotateFlipType.Rotate180FlipNone.ToString(),
                ImageToPdfService.ExifRotation(3).ToString(), "3 — вверх ногами");
            AssertEqual(System.Drawing.RotateFlipType.RotateNoneFlipNone.ToString(),
                ImageToPdfService.ExifRotation(0).ToString(), "нуля в стандарте нет — как есть");
            AssertEqual(System.Drawing.RotateFlipType.RotateNoneFlipNone.ToString(),
                ImageToPdfService.ExifRotation(42).ToString(), "мусор в теге — как есть");
        }

        /// <summary>
        /// Живая сборка PDF из настоящих картинок: JPEG должен уйти в документ КАК ЕСТЬ (иначе
        /// снимок раздувает файл и теряет качество), PNG с прозрачностью — лечь на белое, а
        /// многостраничный TIFF — дать столько страниц, сколько в нём кадров. Плюс проверка, что
        /// битый файл объясняется человеку, а не роняет поток.
        /// </summary>
        private static void TestImageToPdfLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "iwo_img2pdf_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                var offenders = new List<string>();

                // JPEG с шумом: такой почти не сжимается, поэтому по размеру видно, перенесли его
                // как есть или пересобрали.
                string jpg = Path.Combine(dir, "photo.jpg");
                using (var bmp = new System.Drawing.Bitmap(900, 600))
                {
                    var rnd = new Random(4242);
                    for (int yy = 0; yy < bmp.Height; yy += 3)
                        for (int xx = 0; xx < bmp.Width; xx += 3)
                            bmp.SetPixel(xx, yy, System.Drawing.Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256)));
                    bmp.Save(jpg, System.Drawing.Imaging.ImageFormat.Jpeg);
                }
                string fromJpg = Path.Combine(dir, "from_jpg.pdf");
                int pages = ImageToPdfService.WritePages(jpg, fromJpg);
                if (pages != 1)
                    offenders.Add("JPEG дал страниц: " + pages);
                using (PdfDocument doc = PdfReader.Open(fromJpg, PdfDocumentOpenMode.InformationOnly))
                {
                    if (doc.PageCount != 1)
                        offenders.Add("в PDF из JPEG страниц: " + doc.PageCount);
                    else
                    {
                        double w = doc.Pages[0].Width.Point, h = doc.Pages[0].Height.Point;
                        if (Math.Abs(w - ImageToPdfService.A4HeightPt) > 1 ||
                            Math.Abs(h - ImageToPdfService.A4WidthPt) > 1)
                            offenders.Add("широкий снимок дал лист " + Math.Round(w) + "x" + Math.Round(h) +
                                " вместо альбомного A4");
                    }
                }
                long jpgSize = new FileInfo(jpg).Length, pdfSize = new FileInfo(fromJpg).Length;
                if (pdfSize > jpgSize * 1.5 + 4096)
                    offenders.Add("JPEG пересобрали: снимок " + jpgSize + ", PDF " + pdfSize);
                // Нижняя граница не менее важна верхней: страница нужного размера получается и
                // тогда, когда картинки в ней НЕТ, — а такой PDF весит пару килобайт.
                if (pdfSize < jpgSize)
                    offenders.Add("картинки в PDF нет: снимок " + jpgSize + ", PDF " + pdfSize);
                if (!HasImageXObject(fromJpg))
                    offenders.Add("в странице из JPEG нет картинки (XObject)");

                // PNG с прозрачностью: страница должна получиться, а прозрачное — стать белым.
                string png = Path.Combine(dir, "shot.png");
                using (var bmp = new System.Drawing.Bitmap(400, 300,
                    System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                {
                    using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bmp))
                    {
                        g.Clear(System.Drawing.Color.Transparent);
                        g.FillRectangle(System.Drawing.Brushes.Black, 10, 10, 100, 50);
                    }
                    bmp.Save(png, System.Drawing.Imaging.ImageFormat.Png);
                }
                string fromPng = Path.Combine(dir, "from_png.pdf");
                if (ImageToPdfService.WritePages(png, fromPng) != 1)
                    offenders.Add("PNG дал не одну страницу");
                if (!HasImageXObject(fromPng))
                    offenders.Add("в странице из PNG нет картинки (XObject)");

                // Многостраничный TIFF — норма для сканов.
                string tif = Path.Combine(dir, "scan.tif");
                SaveMultipageTiff(tif, 3);
                string fromTif = Path.Combine(dir, "from_tif.pdf");
                int tifPages = ImageToPdfService.WritePages(tif, fromTif);
                if (tifPages != 3)
                    offenders.Add("трёхкадровый TIFF дал страниц: " + tifPages);
                using (PdfDocument doc = PdfReader.Open(fromTif, PdfDocumentOpenMode.InformationOnly))
                    if (doc.PageCount != 3)
                        offenders.Add("в PDF из TIFF страниц: " + doc.PageCount);
                if (!HasImageXObject(fromTif))
                    offenders.Add("в странице из TIFF нет картинки (XObject)");

                // Собранный из картинок документ — обычный PDF: его читает то же объединение,
                // которым окно и сохраняет результат.
                string merged = Path.Combine(dir, "merged.pdf");
                var order = new List<PdfPageRef>
                {
                    new PdfPageRef { SourcePath = fromJpg, PageIndex = 0 },
                    new PdfPageRef { SourcePath = fromPng, PageIndex = 0, Rotation = 90 }
                };
                PdfMergeService.Merge(order, merged);
                using (PdfDocument doc = PdfReader.Open(merged, PdfDocumentOpenMode.InformationOnly))
                    if (doc.PageCount != 2)
                        offenders.Add("сборка из двух картинок дала страниц: " + doc.PageCount);

                // Битый файл: понятное сообщение вместо падения потока.
                string broken = Path.Combine(dir, "broken.png");
                File.WriteAllText(broken, "это не картинка");
                try
                {
                    ImageToPdfService.WritePages(broken, Path.Combine(dir, "broken.pdf"));
                    offenders.Add("битая картинка не вызвала ошибки");
                }
                catch (MergeException ex)
                {
                    if (ex.Message.IndexOf("broken.png", StringComparison.OrdinalIgnoreCase) < 0)
                        offenders.Add("в сообщении об ошибке нет имени файла: " + ex.Message);
                }

                AssertTrue(offenders.Count == 0, "картинки в PDF: " + string.Join(" | ", offenders.ToArray()));
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        /// <summary>
        /// В первой странице документа лежит хотя бы одна картинка. Проверка именно ресурсов, а
        /// не размера файла: пустая страница нужного формата выглядит в отчётах так же успешно,
        /// как страница с картинкой, и отличить их можно только заглянув внутрь.
        /// </summary>
        private static bool HasImageXObject(string pdfPath)
        {
            using (PdfDocument doc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.InformationOnly))
            {
                var resources = doc.Pages[0].Elements.GetDictionary("/Resources");
                if (resources == null)
                    return false;
                var xobjects = resources.Elements.GetDictionary("/XObject");
                if (xobjects == null)
                    return false;
                foreach (string key in xobjects.Elements.Keys)
                {
                    var obj = xobjects.Elements.GetDictionary(key);
                    if (obj != null && obj.Elements.GetName("/Subtype") == "/Image")
                        return true;
                }
                return false;
            }
        }

        /// <summary>TIFF из нескольких кадров — как его отдаёт сканер.</summary>
        private static void SaveMultipageTiff(string path, int frames)
        {
            System.Drawing.Imaging.ImageCodecInfo codec = null;
            foreach (System.Drawing.Imaging.ImageCodecInfo c in System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders())
                if (c.FormatID == System.Drawing.Imaging.ImageFormat.Tiff.Guid)
                    codec = c;
            var first = new System.Drawing.Bitmap(300, 400);
            try
            {
                using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(first))
                    g.Clear(System.Drawing.Color.White);
                using (var ps = new System.Drawing.Imaging.EncoderParameters(1))
                {
                    ps.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                        System.Drawing.Imaging.Encoder.SaveFlag,
                        (long)System.Drawing.Imaging.EncoderValue.MultiFrame);
                    first.Save(path, codec, ps);
                    for (int i = 1; i < frames; i++)
                        using (var next = new System.Drawing.Bitmap(300, 400))
                        using (var page = new System.Drawing.Imaging.EncoderParameters(1))
                        {
                            using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(next))
                                g.Clear(System.Drawing.Color.White);
                            page.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                                System.Drawing.Imaging.Encoder.SaveFlag,
                                (long)System.Drawing.Imaging.EncoderValue.FrameDimensionPage);
                            first.SaveAdd(next, page);
                        }
                    using (var flush = new System.Drawing.Imaging.EncoderParameters(1))
                    {
                        flush.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                            System.Drawing.Imaging.Encoder.SaveFlag,
                            (long)System.Drawing.Imaging.EncoderValue.Flush);
                        first.SaveAdd(flush);
                    }
                }
            }
            finally { first.Dispose(); }
        }

        /// <summary>
        /// «Прочие операции» (живое окно): сетка стала полноправной — страницы переставляют,
        /// поворачивают и убирают, а действия смотрят на СОБРАННЫЕ страницы, не на открытый файл
        /// (набор можно собрать одними картинками). Правая панель при этом обязана поместиться
        /// на своём минимуме — её высота теперь считается по факту, а не числом.
        /// </summary>
        private static void TestOpsGridEditableLive()
        {
            var offenders = new List<string>();
            InIsolatedSettings("iwo_opsgrid_", delegate
            {
                try
                {
                    using (var f = new PdfOpsForm(delegate { }))
                    {
                        f.Show();
                        PdfPageGrid grid = FindPageGrid(f);
                        if (grid == null)
                        {
                            offenders.Add("сетки нет в окне");
                        }
                        else
                        {
                            if (!grid.AllowReorder)
                                offenders.Add("порядок страниц править нельзя");
                            if (!grid.AllowRotate)
                                offenders.Add("страницы поворачивать нельзя");
                            if (!PdfDrop.Matches("photo.jpg", grid.DropExtensions))
                                offenders.Add("сетка не принимает картинки перетаскиванием");
                        }
                        // Без страниц действия недоступны — нечего делать.
                        System.Windows.Forms.Button save = FindVisibleButton(f, Loc.T("ops.btn.savePdf"));
                        System.Windows.Forms.Button addImages = FindVisibleButton(f, Loc.T("ops.btn.addImages"));
                        System.Windows.Forms.Button print = FindVisibleButton(f, Loc.T("common.btn.print"));
                        if (save == null || addImages == null || print == null)
                            offenders.Add("нет кнопки: сохранение/картинки/печать");
                        else
                        {
                            if (save.Enabled)
                                offenders.Add("«Сохранить PDF…» доступно на пустом наборе");
                            if (!addImages.Enabled)
                                offenders.Add("«Добавить картинки…» недоступно на пустом наборе");
                            if (print.Enabled)
                                offenders.Add("«Печать…» доступна на пустом наборе");
                        }
                        // На минимуме окна панель не должна наезжать на нижний строй.
                        f.Size = f.MinimumSize;
                        f.PerformLayout();
                        string where = "PdfOpsForm (Min=" + f.MinimumSize.Height + ")";
                        CheckFits(f, where, offenders);
                        CheckControlsDoNotOverlap(f, where, offenders);
                        f.Close();
                    }
                }
                catch (Exception ex) { offenders.Add("не собралось: " + Root(ex).Message); }
            });
            AssertTrue(offenders.Count == 0, "«Прочие операции»: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// Открывается то руководство, на языке которого человек читает интерфейс:
        /// английскому читателю русский документ инструкцией не станет. Что оба вшиты в
        /// настоящий exe, проверяет --selftest — у тестов своя сборка без ресурсов.
        /// </summary>
        private static void TestManualBothLanguages()
        {
            Lang was = Loc.Current;
            try
            {
                var offenders = new List<string>();
                foreach (Lang lang in new[] { Lang.Ru, Lang.En })
                {
                    Loc.Init(lang);
                    // Вшито ли руководство в НАСТОЯЩИЙ exe, проверяет --selftest: у тестов своя
                    // сборка, и ресурсов приложения в ней нет. Здесь — выбор языка.
                    string path = UserManual.FilePath;
                    string name = System.IO.Path.GetFileName(path);
                    bool english = name.IndexOf("User", StringComparison.OrdinalIgnoreCase) >= 0;
                    if ((lang == Lang.En) != english)
                        offenders.Add(lang + ": открылось бы «" + name + "»");
                }
                AssertTrue(offenders.Count == 0, "руководство: " + string.Join(" | ", offenders.ToArray()));
            }
            finally { Loc.Init(was); }
        }

        private static void TestPptxBackground()
        {
            byte[] png = PptxTinyPng();
            var backgrounds = new List<Background>
            {
                new Background { Data = png, IsJpeg = false },
                new Background { Data = png, IsJpeg = false }
            };
            var pages = new List<PdfPageText> { PptxPage(960, 540), PptxPage(595, 842) };
            pages[0].Paragraphs.Add(PptxPar("текст", 40, 500, 300, 480));
            pages[1].Paragraphs.Add(PptxPar("текст", 40, 800, 300, 780));

            OoxmlPackage package = PptxWriter.BuildPackage(pages, PptxStamp, null, null, backgrounds);
            XDocument first = PptxSlideXml(package, 1);
            AssertTrue(PptxFirst(first.Descendants(PptxP + "bg")) != null,
                "страница размером со слайд — подложка становится фоном (её нельзя сдвинуть)");

            XDocument second = PptxSlideXml(package, 2);
            AssertTrue(PptxFirst(second.Descendants(PptxP + "bg")) == null,
                "вписанная страница фоном быть не может — фон растянулся бы на весь слайд");
            XElement locks = PptxFirst(second.Descendants(PptxA + "picLocks"));
            AssertTrue(locks != null && (string)locks.Attribute("noMove") == "1"
                && (string)locks.Attribute("noSelect") == "1", "подложка-картинка закреплена");

            byte[] bytes = package.ToArray();
            int media = 0;
            using (var ms = new MemoryStream(bytes))
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
                foreach (ZipArchiveEntry e in zip.Entries)
                    if (e.FullName.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase))
                        media++;
            AssertEqual(1, media, "одинаковые подложки кладутся один раз");

            var failures = new List<string>();
            AssertOoxmlSelfConsistent(bytes, failures);
            AssertTrue(failures.Count == 0, "пакет с подложками согласован: " + string.Join(" | ", failures.ToArray()));
            AssertPptxValid(pages);
        }

        private static void TestPptxBackgroundDpi()
        {
            AssertEqual(150, PageBackgrounds.Dpi(1), "короткой колоде — подробная подложка");
            AssertEqual(150, PageBackgrounds.Dpi(50), "граница пятидесяти страниц");
            AssertEqual(120, PageBackgrounds.Dpi(51), "дальше подробность снижается");
            AssertEqual(120, PageBackgrounds.Dpi(150), "полторы сотни страниц");
            AssertEqual(96, PageBackgrounds.Dpi(151), "и всё, что больше — иначе файл раздуется");
        }

        private static void TestPptxMergeGrid()
        {
            // Владелец (1,1) накрывает 2×2: справа, снизу и по диагонали.
            OcrTable table = PptxTable(3, 3, 0, 100, 300, 0);
            table.Rows[1].Cells[1].ColSpan = 2;
            table.Rows[1].Cells[1].RowSpan = 2;
            table.Rows[1].Cells[2].Covered = true;
            table.Rows[2].Cells[1].Covered = true;
            table.Rows[2].Cells[2].Covered = true;

            PptxWriter.CellMerge[,] grid = PptxWriter.MergeGrid(table);
            AssertEqual(2, grid[1, 1].GridSpan, "владелец объявляет пролёт по колонкам");
            AssertEqual(2, grid[1, 1].RowSpan, "владелец объявляет пролёт по строкам");
            AssertTrue(!grid[1, 1].HMerge && !grid[1, 1].VMerge, "владелец не помечен накрытым");

            AssertTrue(grid[1, 2].HMerge, "правая ячейка полосы накрыта слева");
            AssertEqual(1, grid[1, 2].GridSpan, "и не повторяет пролёт по колонкам");
            AssertEqual(2, grid[1, 2].RowSpan, "но несёт пролёт по строкам своей полосы");

            AssertTrue(grid[2, 1].VMerge, "нижняя ячейка накрыта сверху");
            AssertEqual(2, grid[2, 1].GridSpan, "и несёт пролёт по колонкам");
            AssertEqual(1, grid[2, 1].RowSpan, "не повторяя пролёт по строкам");

            AssertTrue(grid[2, 2].HMerge && grid[2, 2].VMerge, "диагональная накрыта с обеих сторон");

            AssertTrue(!grid[0, 0].HMerge && grid[0, 0].GridSpan == 1, "нетронутая ячейка — без объединения");
        }

        private static void TestPptxTable()
        {
            PdfPageText page = PptxPage(600, 800);
            OcrTable table = PptxTable(2, 3, 50, 700, 350, 640);
            table.Rows[0].Cells[0].Paragraphs.Add(PptxPar("Ячейка", 50, 700, 150, 690));
            page.Tables.Add(table);

            OoxmlPackage package = PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp);
            XDocument slide = PptxSlideXml(package, 1);

            AssertTrue(PptxFirst(slide.Descendants(PptxA + "tbl")) != null, "таблица переносится таблицей, а не надписями");
            AssertEqual(3, new List<XElement>(slide.Descendants(PptxA + "gridCol")).Count, "три колонки в сетке");
            var rows = new List<XElement>(slide.Descendants(PptxA + "tr"));
            AssertEqual(2, rows.Count, "две строки");
            foreach (XElement row in rows)
                AssertEqual(3, new List<XElement>(row.Elements(PptxA + "tc")).Count,
                    "в строке ровно столько ячеек, сколько колонок — иначе файл «повреждён»");
            foreach (XElement tr in rows)
                AssertTrue((string)tr.Attribute("h") != null, "высота строки обязательна по схеме");
            AssertTrue(PptxAllText(slide).Contains("Ячейка"), "текст ячейки на месте");
            AssertTrue(PptxFirst(slide.Descendants(PptxA + "lnL")) != null, "у таблицы по линовке есть границы");

            // Сетка, восстановленная без линовки, границ получать не должна.
            PdfPageText page2 = PptxPage(600, 800);
            OcrTable borderless = PptxTable(1, 2, 50, 700, 350, 680);
            borderless.Borderless = true;
            page2.Tables.Add(borderless);
            XDocument slide2 = PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { page2 }, PptxStamp), 1);
            XElement lnL = PptxFirst(slide2.Descendants(PptxA + "lnL"));
            AssertTrue(lnL != null && lnL.Element(PptxA + "noFill") != null,
                "безлиновочная сетка рисуется без границ — иначе на слайде появится таблица, которой не было");

            AssertPptxValid(new List<PdfPageText> { page, page2 });
        }

        private static void TestPptxValidatorWithContent()
        {
            PdfPageText page = PptxPage(960, 540);
            OcrParagraph title = PptxPar("Заголовок", 40, 500, 900, 470, 28);
            title.Alignment = OcrAlignment.Center;
            page.Paragraphs.Add(title);
            OcrParagraph item = PptxPar("• пункт списка", 60, 440, 900, 420);
            item.ListKind = ListKind.Bulleted;
            page.Paragraphs.Add(item);
            var linked = new OcrParagraph { LeftPt = 60, TopPt = 400, RightPt = 900, BottomPt = 380 };
            linked.Runs.Add(new OcrRun { Text = "ссылка", FontSizePt = 12, Uri = "https://example.org/x?y=1&z=2" });
            linked.Runs.Add(new OcrRun { Text = " и цвет", FontSizePt = 12, ColorArgb = 0x00A0FF, Bold = true });
            page.Paragraphs.Add(linked);
            page.Images.Add(new OcrImage { Png = PptxTinyPng(), LeftPt = 700, TopPt = 520, WidthPt = 60, HeightPt = 40 });

            AssertPptxValid(new List<PdfPageText> { page });
        }

        /// <summary>Крошечный настоящий PNG (1×1) — картинка для тестов, а не набор байтов.</summary>
        private static byte[] PptxTinyPng()
        {
            using (var bmp = new System.Drawing.Bitmap(1, 1))
            using (var ms = new MemoryStream())
            {
                bmp.SetPixel(0, 0, System.Drawing.Color.Red);
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                return ms.ToArray();
            }
        }

        /// <summary>Собрать .pptx во временный файл и проверить официальным валидатором Open XML.</summary>
        private static void AssertPptxValid(IList<PdfPageText> pages)
        {
            string path = Path.Combine(Path.GetTempPath(), "iwo_pptx_" + Guid.NewGuid().ToString("N") + ".pptx");
            try
            {
                PptxWriter.Write(pages, path, PptxStamp);
                var problems = new List<string>();
                using (var doc = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(path, false))
                {
                    var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                        DocumentFormat.OpenXml.FileFormatVersions.Office2013);
                    foreach (DocumentFormat.OpenXml.Validation.ValidationErrorInfo info in validator.Validate(doc))
                    {
                        problems.Add(info.Description + " @ " + (info.Path == null ? "?" : info.Path.XPath));
                        if (problems.Count >= 5)
                            break;
                    }
                }
                AssertTrue(problems.Count == 0, "валидатор нашёл: " + string.Join(" | ", problems.ToArray()));
            }
            finally
            {
                try { File.Delete(path); } catch { }
            }
        }

        private static void TestPptxValidator()
        {
            string path = Path.Combine(Path.GetTempPath(), "iwo_pptx_" + Guid.NewGuid().ToString("N") + ".pptx");
            try
            {
                PptxWriter.Write(new List<PdfPageText> { PptxPage(960, 540) }, path,
                    new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc));
                var problems = new List<string>();
                using (var doc = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(path, false))
                {
                    var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                        DocumentFormat.OpenXml.FileFormatVersions.Office2013);
                    foreach (DocumentFormat.OpenXml.Validation.ValidationErrorInfo info in validator.Validate(doc))
                    {
                        problems.Add(info.Description + " @ " + (info.Path == null ? "?" : info.Path.XPath));
                        if (problems.Count >= 5)
                            break;
                    }
                }
                AssertTrue(problems.Count == 0, "валидатор нашёл: " + string.Join(" | ", problems.ToArray()));
            }
            finally
            {
                try { File.Delete(path); } catch { }
            }
        }

        /// <summary>
        /// Валидатор — снаряжение тестов, а не приложения: попади он в src, один exe вырос бы
        /// на 8,6 МБ ради проверки, которая в бою не нужна. Стережём обе двери: исходники и
        /// файл проекта приложения.
        /// </summary>
        private static void TestPptxValidatorNotInApp()
        {
            string dir = SourceDir();
            AssertTrue(dir != null, "каталог src не найден рядом с тестами");
            var offenders = new List<string>();
            foreach (string file in Directory.GetFiles(dir, "*.cs"))
                if (File.ReadAllText(file).IndexOf("DocumentFormat.OpenXml", StringComparison.OrdinalIgnoreCase) >= 0)
                    offenders.Add(Path.GetFileName(file));
            AssertTrue(offenders.Count == 0, "валидатор просочился в приложение: " + string.Join(", ", offenders.ToArray()));

            string csproj = Path.Combine(Path.GetDirectoryName(dir), "iwoHelperDesktop.csproj");
            AssertTrue(File.Exists(csproj), "файл проекта приложения найден");
            AssertTrue(File.ReadAllText(csproj).IndexOf("DocumentFormat.OpenXml", StringComparison.OrdinalIgnoreCase) < 0,
                "в проекте приложения ссылки на валидатор нет");
        }

        private static void Run(string name, Action test)
        {
            try
            {
                test();
                _passed++;
                Console.WriteLine("PASS  " + name);
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("FAIL  " + name + " — " + ex.Message);
            }
        }

        // ---------- буфер страниц, поворот, кэш миниатюр (1.17.0) ----------

        private static void TestLruPeekRemove()
        {
            var evicted = new List<string>();
            var lru = new LruCache<string>(2, delegate(string val) { evicted.Add(val); });
            lru.Add("a", "A");
            lru.Add("b", "B");
            string v;
            AssertTrue(lru.TryPeek("a", out v) && v == "A", "TryPeek видит значение");
            lru.Add("c", "C"); // peek НЕ освежил «a» — вытесняется именно он
            AssertEqual(1, evicted.Count, "одно вытеснение");
            AssertEqual("A", evicted[0], "вытеснен несвежий после peek");
            AssertTrue(!lru.TryPeek("a", out v), "вытесненный изъят");

            AssertTrue(lru.Remove("b", out v) && v == "B", "Remove возвращает значение");
            AssertEqual(1, evicted.Count, "Remove не зовёт onEvict");
            AssertEqual(1, lru.Count, "после Remove остался один");
            List<string> keys = lru.KeysSnapshot();
            AssertTrue(keys.Count == 1 && string.Equals(keys[0], "c", StringComparison.OrdinalIgnoreCase),
                "KeysSnapshot отражает содержимое");
            AssertTrue(!lru.Remove("nope", out v), "Remove отсутствующего — false");
        }

        private static void TestRotationCompose()
        {
            AssertEqual(90, PdfPageRef.ComposeRotation(0, 90), "0+90");
            AssertEqual(0, PdfPageRef.ComposeRotation(270, 90), "270+90 → 0");
            AssertEqual(270, PdfPageRef.ComposeRotation(0, -90), "0−90 → 270");
            AssertEqual(0, PdfPageRef.ComposeRotation(180, 180), "180+180 → 0");
            AssertEqual(180, PdfPageRef.ComposeRotation(90, 90), "90+90");
            AssertEqual(90, PdfPageRef.ComposeRotation(90, 45), "не кратный 90 срезается вниз");
            AssertEqual(270, PdfPageRef.ComposeRotation(-90, 0), "отрицательный исходный");
        }

        private static void TestPageRefClone()
        {
            var page = new PdfPageRef { SourcePath = "C:\\a.pdf", PageIndex = 3, Rotation = 90 };
            PdfPageRef clone = page.Clone();
            AssertTrue(!ReferenceEquals(page, clone), "новый экземпляр");
            AssertEqual(page.SourcePath, clone.SourcePath, "путь");
            AssertEqual(page.PageIndex, clone.PageIndex, "индекс");
            AssertEqual(page.Rotation, clone.Rotation, "поворот скопирован");
            clone.Rotation = 180;
            AssertEqual(90, page.Rotation, "оригинал не задет");
        }

        private static List<string> FiveLetters()
        {
            return new List<string> { "a", "b", "c", "d", "e" };
        }

        private static void TestMoveRange()
        {
            List<string> list = FiveLetters();
            AssertEqual(0, ListReorder.MoveRange(list, new[] { 2, 3 }, 0), "перенос к началу: индекс");
            AssertEqual("c,d,a,b,e", string.Join(",", list), "перенос к началу: порядок");

            list = FiveLetters();
            AssertEqual(3, ListReorder.MoveRange(list, new[] { 0, 1 }, 5), "в конец: индекс");
            AssertEqual("c,d,e,a,b", string.Join(",", list), "в конец: порядок");

            list = FiveLetters();
            AssertEqual(1, ListReorder.MoveRange(list, new[] { 1, 2 }, 2), "вставка внутри набора: индекс");
            AssertEqual("a,b,c,d,e", string.Join(",", list), "вставка внутри набора: порядок не меняется");

            list = FiveLetters();
            AssertEqual(0, ListReorder.MoveRange(list, new[] { 3, 2 }, 0), "несортированные индексы");
            AssertEqual("c,d,a,b,e", string.Join(",", list), "несортированные: порядок");

            list = FiveLetters();
            AssertEqual(0, ListReorder.MoveRange(list, new[] { 2, 2, 3 }, 0), "дубли игнорируются");
            AssertEqual("c,d,a,b,e", string.Join(",", list), "дубли: порядок");

            list = FiveLetters();
            AssertEqual(-1, ListReorder.MoveRange(list, new int[0], 2), "пустой набор — -1");
            AssertEqual(-1, ListReorder.MoveRange(list, new[] { -1, 9 }, 2), "все вне диапазона — -1");
            AssertEqual("a,b,c,d,e", string.Join(",", list), "список не тронут");

            list = FiveLetters();
            AssertEqual(4, ListReorder.MoveRange(list, new[] { 0 }, 99), "insertAt клампится к концу");
            AssertEqual("b,c,d,e,a", string.Join(",", list), "кламп: порядок");
        }

        private static void TestMoveRangeHelpers()
        {
            AssertEqual(3, ListReorder.AdjustedInsertIndex(new List<int> { 1, 3 }, 5), "два изъятых левее");
            AssertEqual(0, ListReorder.AdjustedInsertIndex(new List<int> { 0, 1, 2 }, 1), "изъятые вокруг цели");
            AssertEqual(2, ListReorder.AdjustedInsertIndex(new List<int> { 3, 4 }, 2), "изъятые правее не сдвигают");
            List<int> norm = ListReorder.NormalizeIndices(new[] { 3, 1, 3, -2, 9 }, 5);
            AssertEqual("1,3", string.Join(",", norm), "сортировка, дубли и границы");
            AssertEqual(0, ListReorder.NormalizeIndices(null, 5).Count, "null — пусто");
        }

        private static void TestPdfOrderInsertAt()
        {
            var order = new PdfPageOrder();
            order.AddDocument("x.pdf", 3);
            AssertEqual(1, order.InsertDocument(1, "y.pdf", 2), "вставка в середину: индекс");
            AssertEqual(5, order.Count, "стало пять страниц");
            AssertEqual("y.pdf", order[1].SourcePath, "первая вставленная");
            AssertEqual("y.pdf", order[2].SourcePath, "вторая вставленная");
            AssertEqual(1, order[3].PageIndex, "прежняя страница сдвинулась");
            AssertEqual(5, order.InsertDocument(99, "z.pdf", 1), "кламп к концу");
            AssertEqual("z.pdf", order[5].SourcePath, "в конец");

            var pages = new List<PdfPageRef> { new PdfPageRef { SourcePath = "p.pdf", PageIndex = 0 } };
            AssertEqual(0, order.InsertAt(-5, pages), "отрицательный insertAt — в начало");
            AssertEqual("p.pdf", order[0].SourcePath, "вставлено в начало");

            int landed = order.MoveRange(new[] { 0 }, order.Count);
            AssertEqual(order.Count - 1, landed, "перенос в конец через модель");
            AssertEqual("p.pdf", order[order.Count - 1].SourcePath, "перенесён");
        }

        private static void TestRenderWidthFor()
        {
            // Базовый рендер (до обычного масштаба BaseWidth = 190).
            AssertEqual(300, ThumbZoom.RenderWidthFor(96, ThumbZoom.BaseWidth), "база, 100% DPI — прежние 300");
            AssertEqual(357, ThumbZoom.RenderWidthFor(120, ThumbZoom.BaseWidth), "база, 125% DPI");
            AssertEqual(428, ThumbZoom.RenderWidthFor(144, ThumbZoom.BaseWidth), "база, 150% DPI");
            // Увеличенный рендер (крупный зум до MaxWidth = 400).
            AssertEqual(600, ThumbZoom.RenderWidthFor(96, ThumbZoom.MaxWidth), "крупный, 100% DPI");
            AssertEqual(640, ThumbZoom.RenderWidthFor(144, ThumbZoom.MaxWidth), "крупный, 150% DPI — потолок 640");
            AssertEqual(300, ThumbZoom.RenderWidthFor(0, ThumbZoom.BaseWidth), "мусорный DPI — минимум");
            AssertTrue(ThumbZoom.RenderWidthFor(96, ThumbZoom.MaxWidth) >= ThumbZoom.MaxWidth,
                "рендер не меньше самой крупной плитки");
        }

        private static void TestThumbZoomCell()
        {
            System.Drawing.Size tile = ThumbZoom.TileSize(200);
            System.Drawing.Size cell = ThumbZoom.CellSize(200);
            AssertTrue(cell.Width > tile.Width, "ячейка шире плитки (поля)");
            AssertTrue(cell.Height > tile.Height + 15, "в ячейке есть полоса под номер");
            System.Drawing.Size max = ThumbZoom.TileSize(ThumbZoom.MaxWidth);
            AssertEqual(ThumbZoom.MaxWidth, max.Width, "максимум 400 доступен (лимита ImageList больше нет)");
            AssertTrue(max.Height > 256, "высота плитки на максимуме больше прежнего лимита 256");
        }

        private static void TestTileRectFromIcon()
        {
            // Иконная зона 256×256 в ячейке; плитка 300×390 крупнее — рисуется вокруг её центра
            // по X и от её верха по Y (хитбокс остаётся 256, клики по кольцу компенсируются).
            var icon = new System.Drawing.Rectangle(100, 50, 256, 256);
            System.Drawing.Rectangle tile = PdfPageGrid.TileRectFromIcon(icon, new System.Drawing.Size(300, 390));
            AssertEqual(icon.Left + 128 - 150, tile.Left, "плитка центрирована по иконной зоне");
            AssertEqual(icon.Top, tile.Top, "плитка от верха иконной зоны");
            AssertEqual(300, tile.Width, "ширина плитки");
            AssertEqual(390, tile.Height, "высота плитки");

            // Плитка меньше иконной зоны (обычный масштаб) — тоже по центру.
            System.Drawing.Rectangle small = PdfPageGrid.TileRectFromIcon(icon, new System.Drawing.Size(100, 130));
            AssertEqual(icon.Left + 128 - 50, small.Left, "малая плитка по центру");
        }

        private static void TestHoverRotateButtons()
        {
            var tile = new System.Drawing.Rectangle(100, 50, 132, 172);
            System.Drawing.Rectangle[] b = PdfPageGrid.HoverRotateButtons(tile, 24, 6);
            AssertEqual(2, b.Length, "две кнопки");
            AssertTrue(!b[0].IntersectsWith(b[1]), "кнопки не пересекаются");
            AssertTrue(tile.Contains(b[0]) && tile.Contains(b[1]), "кнопки внутри плитки");
            AssertEqual(b[0].Y, b[1].Y, "на одной высоте");
            AssertTrue(b[0].Right < b[1].Left, "левая левее правой");
            int center = tile.Left + tile.Width / 2;
            AssertTrue(Math.Abs((b[0].Left + b[1].Right) / 2 - center) <= 1, "пара по центру плитки");
            AssertTrue(b[0].Bottom < tile.Bottom, "над нижней кромкой (зазор)");
        }

        /// <summary>
        /// Регрессия «чипы поворота съедаются при увеличении»: нативные границы элемента
        /// ограничены иконной зоной 256, поэтому инвалидация идёт полной ЯЧЕЙКОЙ —
        /// она обязана накрывать плитку целиком и hover-кнопки у её нижней кромки
        /// даже на максимальном зуме (плитка 400 больше хитбокса 256).
        /// </summary>
        private static void TestCellRectCoversChips()
        {
            var icon = new System.Drawing.Rectangle(100, 50, 256, 256); // нативная иконная зона (потолок 256)
            System.Drawing.Size tileSz = ThumbZoom.TileSize(ThumbZoom.MaxWidth);   // 400×520
            System.Drawing.Size cellSz = ThumbZoom.CellSize(ThumbZoom.MaxWidth);
            System.Drawing.Rectangle tile = PdfPageGrid.TileRectFromIcon(icon, tileSz);
            System.Drawing.Rectangle cell = PdfPageGrid.CellRectFromTile(tile, cellSz, 4);

            AssertTrue(cell.Contains(tile), "ячейка накрывает плитку целиком");
            AssertTrue(tile.Bottom > icon.Bottom, "плитка на максимуме выходит за нативные границы (сама причина бага)");
            System.Drawing.Rectangle[] chips = PdfPageGrid.HoverRotateButtons(tile, 24, 6);
            AssertTrue(cell.Contains(chips[0]) && cell.Contains(chips[1]),
                "hover-кнопки внутри инвалидируемой ячейки");
            AssertEqual(cellSz.Width, cell.Width, "ширина ячейки — из ThumbZoom.CellSize");
            AssertEqual(cellSz.Height, cell.Height, "высота ячейки — из ThumbZoom.CellSize");
            AssertEqual(tile.Top - 4, cell.Top, "верхний отступ ячейки соблюдён");

            // И на обычном масштабе геометрия согласована так же.
            System.Drawing.Size tileDef = ThumbZoom.TileSize(ThumbZoom.DefaultWidth);
            System.Drawing.Rectangle tileD = PdfPageGrid.TileRectFromIcon(
                new System.Drawing.Rectangle(0, 0, tileDef.Width, tileDef.Height), tileDef);
            System.Drawing.Rectangle cellD = PdfPageGrid.CellRectFromTile(tileD, ThumbZoom.CellSize(ThumbZoom.DefaultWidth), 4);
            AssertTrue(cellD.Contains(PdfPageGrid.HoverRotateButtons(tileD, 24, 6)[1]),
                "чипы внутри ячейки и на обычном масштабе");
        }

        private static void TestPdfOrderUndo()
        {
            var order = new PdfPageOrder();
            order.AddDocument("x.pdf", 3);
            AssertTrue(!order.CanUndo, "изначально откатывать нечего");
            AssertTrue(!order.Undo(), "Undo без снимков — false");

            order.Checkpoint();
            order.Move(0, 3); // [1,2,0]
            AssertEqual(0, order[2].PageIndex, "порядок изменился (страница 0 в конце)");
            AssertEqual(1, order[0].PageIndex, "порядок изменился (страница 1 в начале)");
            AssertTrue(order.CanUndo, "есть снимок");
            AssertTrue(order.Undo(), "откат удался");
            AssertEqual(0, order[0].PageIndex, "порядок восстановлен");
            AssertTrue(!order.CanUndo, "снимок израсходован");

            // Повороты входят в снимок: откат возвращает и порядок, и углы (в общие ссылки).
            order.Checkpoint();
            order[0].Rotation = 90;
            order.RemoveAt(new[] { 2 });
            order.Undo();
            AssertEqual(3, order.Count, "состав восстановлен");
            AssertEqual(0, order[0].Rotation, "поворот откатился вместе с жестом");
            AssertTrue(order.CanRedo, "после отката доступен возврат");
            AssertTrue(order.Redo(), "возврат удался");
            AssertEqual(2, order.Count, "состав вернулся");
            AssertEqual(90, order[0].Rotation, "поворот вернулся");
            order.Undo(); // назад к трём страницам без поворота — для следующих проверок

            // Новый жест обнуляет ветку возврата (как в любом редакторе).
            order.Checkpoint();
            order.Move(0, 2);
            order.Undo();
            AssertTrue(order.CanRedo, "возврат доступен после отката");
            order.Checkpoint();
            order.Move(0, 2);
            AssertTrue(!order.CanRedo, "новый жест очистил ветку возврата");
            order.Undo();

            // Лимит стека: старые снимки уходят, свежие живут.
            for (int i = 0; i < 60; i++)
                order.Checkpoint();
            int undone = 0;
            while (order.Undo())
                undone++;
            AssertEqual(50, undone, "стек ограничен 50 снимками");

            order.Checkpoint();
            order.Clear();
            AssertTrue(!order.CanUndo, "Clear очищает и историю отката");
        }

        private static void TestPageCacheCapacity()
        {
            AssertEqual(96, ThumbZoom.PageCacheCapacity(48L << 20, 300), "48 МБ при 300 px");
            AssertEqual(24, ThumbZoom.PageCacheCapacity(1L << 20, 300), "пол — 24");
            AssertEqual(512, ThumbZoom.PageCacheCapacity(8L << 30, 300), "потолок — 512");
            AssertTrue(ThumbZoom.PageCacheCapacity(192L << 20, 640) >= 24, "большой рендер не роняет пол");
        }

        private static void TestGridLabelTileKey()
        {
            var page = new PdfPageRef { SourcePath = "C:\\Docs\\a.pdf", PageIndex = 4 };
            AssertEqual("3", PdfPageGrid.PageLabel(page, 2, true), "режим позиции");
            AssertEqual("5", PdfPageGrid.PageLabel(page, 2, false), "режим исходной страницы");
            AssertTrue(PdfPageGrid.TileKey(page).EndsWith("|r0"), "плитка без поворота");
            page.Rotation = 90;
            AssertTrue(PdfPageGrid.TileKey(page).EndsWith("|r90"), "плитка с поворотом");
            AssertTrue(PdfPageGrid.TileKey(page).StartsWith(PdfPageGrid.ThumbKey(page)),
                "ключ плитки начинается с ключа страницы");
            AssertEqual(System.Drawing.RotateFlipType.Rotate90FlipNone, PageRotation.FlipFor(90), "90");
            AssertEqual(System.Drawing.RotateFlipType.Rotate180FlipNone, PageRotation.FlipFor(180), "180");
            AssertEqual(System.Drawing.RotateFlipType.Rotate270FlipNone, PageRotation.FlipFor(270), "270");
            AssertEqual(System.Drawing.RotateFlipType.RotateNoneFlipNone, PageRotation.FlipFor(0), "0");
        }

        private static void TestPasteIndex()
        {
            AssertEqual(3, PdfPageGrid.PasteIndex(2, true, new int[0], 10), "каретка после плитки 2");
            AssertEqual(2, PdfPageGrid.PasteIndex(2, false, new int[0], 10), "каретка до плитки 2");
            AssertEqual(10, PdfPageGrid.PasteIndex(9, true, new int[0], 10), "каретка в конце");
            AssertEqual(5, PdfPageGrid.PasteIndex(-1, false, new[] { 1, 4 }, 10), "после последнего выбранного");
            AssertEqual(10, PdfPageGrid.PasteIndex(-1, false, new int[0], 10), "без каретки и выбора — в конец");
        }

        private static void TestClassifyClipboardKeys()
        {
            var K = new Func<System.Windows.Forms.Keys, PdfToolFormBase.PageKeyAction>(PdfToolFormBase.ClassifyPageKey);
            AssertEqual(PdfToolFormBase.PageKeyAction.Cut, K(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.X), "Ctrl+X");
            AssertEqual(PdfToolFormBase.PageKeyAction.Copy, K(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.C), "Ctrl+C");
            AssertEqual(PdfToolFormBase.PageKeyAction.Paste, K(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.V), "Ctrl+V");
            AssertEqual(PdfToolFormBase.PageKeyAction.GoTo, K(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.G), "Ctrl+G");
            AssertEqual(PdfToolFormBase.PageKeyAction.RotateRight,
                K(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift | System.Windows.Forms.Keys.Oemplus), "Ctrl+Shift+«+»");
            AssertEqual(PdfToolFormBase.PageKeyAction.RotateRight,
                K(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift | System.Windows.Forms.Keys.Add), "Ctrl+Shift+Num+");
            AssertEqual(PdfToolFormBase.PageKeyAction.RotateLeft,
                K(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift | System.Windows.Forms.Keys.OemMinus), "Ctrl+Shift+«−»");
            AssertEqual(PdfToolFormBase.PageKeyAction.RotateLeft,
                K(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift | System.Windows.Forms.Keys.Subtract), "Ctrl+Shift+Num−");
            AssertEqual(PdfToolFormBase.PageKeyAction.CancelClipboard, K(System.Windows.Forms.Keys.Escape), "Esc");
            AssertEqual(PdfToolFormBase.PageKeyAction.Undo, K(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Z), "Ctrl+Z");
            AssertEqual(PdfToolFormBase.PageKeyAction.Redo, K(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Y), "Ctrl+Y");
            AssertEqual(PdfToolFormBase.PageKeyAction.Redo,
                K(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift | System.Windows.Forms.Keys.Z), "Ctrl+Shift+Z");
            AssertEqual(PdfToolFormBase.PageKeyAction.None, K(System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Q), "чужая клавиша");
        }

        private static void TestRotationAt()
        {
            AssertEqual(0, PdfSplitService.RotationAt(null, 0), "нет карты — 0");
            AssertEqual(90, PdfSplitService.RotationAt(new[] { 0, 90 }, 1), "по индексу");
            AssertEqual(0, PdfSplitService.RotationAt(new[] { 0, 90 }, 5), "за пределами — 0");
            AssertEqual(0, PdfSplitService.RotationAt(new[] { 0, 90 }, -1), "отрицательный — 0");
        }

        /// <summary>Двухстраничный PDF для проверок поворота; firstPageRotate — собственный /Rotate первой страницы.</summary>
        private static string MakeTwoPagePdf(string dir, int firstPageRotate)
        {
            string path = Path.Combine(dir, "rotsrc.pdf");
            using (var doc = new PdfDocument())
            {
                PdfPage first = doc.AddPage();
                if (firstPageRotate != 0)
                    first.Rotate = firstPageRotate;
                doc.AddPage();
                doc.Save(path);
            }
            return path;
        }

        private static void TestMergeRotationLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "iwo_rot_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = MakeTwoPagePdf(dir, 270);
                string outPath = Path.Combine(dir, "out.pdf");
                var order = new List<PdfPageRef>
                {
                    new PdfPageRef { SourcePath = src, PageIndex = 0, Rotation = 90 },
                    new PdfPageRef { SourcePath = src, PageIndex = 1, Rotation = 180 },
                    new PdfPageRef { SourcePath = src, PageIndex = 1 }
                };
                PdfMergeService.Merge(order, outPath);
                using (PdfDocument doc = PdfReader.Open(outPath, PdfDocumentOpenMode.Import))
                {
                    AssertEqual(3, doc.PageCount, "страниц в итоге");
                    AssertEqual(0, doc.Pages[0].Rotate, "исходные 270 + 90 пользователя = 0");
                    AssertEqual(180, doc.Pages[1].Rotate, "поворот на 180");
                    AssertEqual(0, doc.Pages[2].Rotate, "без поворота — как в исходнике");
                }
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        private static void TestSplitRotationLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "iwo_rot_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = MakeTwoPagePdf(dir, 0);
                var ranges = new List<PageRange> { new PageRange(0, 1) };
                List<string> files = PdfSplitService.SplitByRanges(src, ranges, dir, "part", null, new[] { 90, 0 });
                AssertEqual(1, files.Count, "один файл диапазона");
                using (PdfDocument doc = PdfReader.Open(files[0], PdfDocumentOpenMode.Import))
                {
                    AssertEqual(90, doc.Pages[0].Rotate, "поворот из карты");
                    AssertEqual(0, doc.Pages[1].Rotate, "вторая без поворота");
                }
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        private static void TestPageRotationMap()
        {
            double x, y;
            // Страница 400×600, нижний-левый угол (0,0): при 90° CW уходит в верхний-левый.
            PageRotation.MapPoint(0, 0, 90, 400, 600, out x, out y);
            AssertTrue(x == 0 && y == 400, "90°: НЛ -> ВЛ (0,400), получено (" + x + "," + y + ")");
            PageRotation.MapPoint(0, 0, 180, 400, 600, out x, out y);
            AssertTrue(x == 400 && y == 600, "180°: НЛ -> ВП");
            PageRotation.MapPoint(0, 0, 270, 400, 600, out x, out y);
            AssertTrue(x == 600 && y == 0, "270°: НЛ -> НП");
            PageRotation.MapPoint(10, 20, 0, 400, 600, out x, out y);
            AssertTrue(x == 10 && y == 20, "0°: тождество");

            // Прямоугольник нормализуется (min/max по повёрнутым углам).
            double l, b, r, t;
            PageRotation.MapBox(10, 20, 110, 70, 90, 400, 600, out l, out b, out r, out t);
            AssertTrue(l == 20 && r == 70, "90°: X из прежних Y");
            AssertTrue(b == 290 && t == 390, "90°: Y из W−X");

            // Обратный поворот возвращает точку на место (композиция — тождество).
            double ix, iy;
            PageRotation.MapPoint(33, 44, 90, 400, 600, out x, out y);
            PageRotation.MapPoint(x, y, PageRotation.Inverse(90), 600, 400, out ix, out iy);
            AssertTrue(Math.Abs(ix - 33) < 1e-9 && Math.Abs(iy - 44) < 1e-9, "инверсия 90°");

            AssertEqual(270, PageRotation.Inverse(90), "Inverse(90)");
            AssertEqual(180, PageRotation.Inverse(180), "Inverse(180)");
            AssertEqual(0, PageRotation.Inverse(0), "Inverse(0)");
            AssertTrue(PageRotation.SwapsDimensions(90) && PageRotation.SwapsDimensions(270), "свап 90/270");
            AssertTrue(!PageRotation.SwapsDimensions(0) && !PageRotation.SwapsDimensions(180), "нет свапа 0/180");
            AssertEqual(90, PageRotation.At(new[] { 0, 90 }, 1), "At по индексу");
            AssertEqual(0, PageRotation.At(null, 3), "At без карты");
        }

        private static byte[] MakeTinyPng(int w, int h)
        {
            using (var bmp = new System.Drawing.Bitmap(w, h))
            {
                bmp.SetPixel(0, 0, System.Drawing.Color.Red);
                bmp.SetPixel(w - 1, h - 1, System.Drawing.Color.Blue);
                using (var ms = new MemoryStream())
                {
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    return ms.ToArray();
                }
            }
        }

        private static void TestPageRotationRotatePage()
        {
            var word = new PdfWord { Text = "w", Left = 10, Right = 110, Bottom = 20, Top = 70, FontSizePt = 10 };
            var line = new PdfLine { Orientation = LineOrientation.Horizontal, X1 = 10, Y1 = 500, X2 = 200, Y2 = 500, Thickness = 1 };
            var img = new OcrImage { Png = MakeTinyPng(20, 10), LeftPt = 50, TopPt = 400, WidthPt = 100, HeightPt = 40 };
            var words = new List<PdfWord> { word };
            var lines = new List<PdfLine> { line };
            var images = new List<OcrImage> { img };
            double w = 400, h = 600;

            PageRotation.RotatePage(words, lines, images, 90, ref w, ref h);
            AssertTrue(w == 600 && h == 400, "размеры страницы поменялись местами");
            AssertTrue(word.Left == 20 && word.Right == 70, "слово: X из прежних Y");
            AssertTrue(word.Bottom == 290 && word.Top == 390, "слово: Y из W−X");
            AssertEqual(LineOrientation.Vertical, line.Orientation, "горизонталь стала вертикалью");
            AssertTrue(Math.Abs(line.Position - 500) < 1e-9, "постоянная координата линии сохранилась");
            AssertTrue(img.WidthPt == 40 && img.HeightPt == 100, "рамка картинки повёрнута");
            using (var ms = new MemoryStream(img.Png))
            using (var bmp = new System.Drawing.Bitmap(ms))
                AssertTrue(bmp.Width == 10 && bmp.Height == 20, "пиксели картинки повёрнуты");

            // rotation == 0 — строгий no-op.
            var word2 = new PdfWord { Left = 1, Right = 2, Bottom = 3, Top = 4 };
            double w2 = 100, h2 = 200;
            PageRotation.RotatePage(new List<PdfWord> { word2 }, null, null, 0, ref w2, ref h2);
            AssertTrue(word2.Left == 1 && word2.Top == 4 && w2 == 100 && h2 == 200, "0°: ничего не тронуто");
        }

        private static void TestUnderscoreHeights()
        {
            var rule = new PdfWord { Text = "____", Left = 0, Right = 50, Bottom = 100, Top = 101, FontSizePt = 12 };
            var normal = new PdfWord { Text = "№", Left = 0, Right = 10, Bottom = 100, Top = 101, FontSizePt = 12 };
            PdfTextExtract.ApplyUnderscoreHeights(new List<PdfWord> { rule, normal });
            AssertTrue(rule.Top > 100 + 0.3 * 12, "прочерку дана виртуальная высота");
            AssertTrue(Math.Abs(normal.Top - 101) < 1e-9, "обычное слово не тронуто");
        }

        private static void TestBuildRotations()
        {
            AssertTrue(PdfPageExtraction.BuildRotations(new List<PdfPageRef>()) == null, "пусто — null");
            var order = new List<PdfPageRef>
            {
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = 0 },                  // первый экземпляр: без поворота
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = 0, Rotation = 90 },   // дубль игнорируется (первый решил)
                new PdfPageRef { SourcePath = "a.pdf", PageIndex = 2, Rotation = 180 },  // регистр пути не важен
                new PdfPageRef { SourcePath = "B.pdf", PageIndex = 1, Rotation = 270 }
            };
            var maps = PdfPageExtraction.BuildRotations(order);
            AssertTrue(maps != null, "карта построена");
            AssertTrue(maps.ContainsKey("A.pdf"), "источник A");
            AssertEqual(0, PageRotation.At(maps["A.pdf"], 0), "первый экземпляр (без поворота) решил");
            AssertEqual(180, PageRotation.At(maps["A.pdf"], 2), "страница 3 источника A");
            AssertEqual(270, PageRotation.At(maps["B.pdf"], 1), "источник B");
            AssertTrue(PdfPageExtraction.BuildRotations(
                new List<PdfPageRef> { new PdfPageRef { SourcePath = "x.pdf", PageIndex = 5 } }) == null,
                "нет ненулевых поворотов — null");
        }

        /// <summary>
        /// Живая калибровка поворота «PDF → Word»: PdfSharp рисует текст, повёрнутый на
        /// странице по часовой (+90 в Y-вниз GDI = по часовой в вьюере) и против; пользователь
        /// выправляет такой текст поворотом страницы в ПРОТИВОПОЛОЖНУЮ сторону. Проверяется и
        /// прежнее поведение: без поворота боковой текст в поток не попадает.
        /// </summary>
        private static void TestExtractRotationLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "iwo_rotw_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string path = Path.Combine(dir, "sideways.pdf");
                using (var doc = new PdfDocument())
                {
                    PdfPage page = doc.AddPage();
                    page.Width = 400;
                    page.Height = 600;
                    using (XGraphics gfx = XGraphics.FromPdfPage(page))
                    {
                        var font = new XFont("Arial", 14);
                        gfx.RotateAtTransform(90, new XPoint(200, 300)); // текст «лежит» по часовой
                        gfx.DrawString("SIDEWAYSCW", font, XBrushes.Black, 120, 300);
                    }
                    PdfPage page2 = doc.AddPage();
                    page2.Width = 400;
                    page2.Height = 600;
                    using (XGraphics gfx2 = XGraphics.FromPdfPage(page2))
                    {
                        var font = new XFont("Arial", 14);
                        gfx2.RotateAtTransform(-90, new XPoint(200, 300)); // «лежит» против часовой
                        gfx2.DrawString("SIDEWAYSCCW", font, XBrushes.Black, 120, 300);
                    }
                    doc.Save(path);
                }

                // Сравнение без пробелов: PdfPig сегментирует БОКОВОЙ текст на слова чуть
                // иначе (рамка первого глифа шире — «S IDEWAYSCW»); буквы и порядок целы,
                // а лишний пробел внутри слова — артефакт сегментации PdfPig, не поворота.
                Func<PdfPageText, string> squash = delegate(PdfPageText p) { return p.Text.Replace(" ", ""); };

                // Без поворота: боковой текст отфильтрован (прежнее поведение — регрессии нет).
                List<PdfPageText> plain = PdfTextExtract.Extract(path);
                AssertTrue(squash(plain[0]).IndexOf("SIDEWAYSCW", StringComparison.Ordinal) < 0,
                    "без поворота боковой текст не извлекается");

                // Текст, лежащий ПО часовой, выправляется поворотом страницы ПРОТИВ часовой (270° CW).
                List<PdfPageText> fixedCw = PdfTextExtract.Extract(path, null, new[] { 270, 0 });
                AssertTrue(squash(fixedCw[0]).IndexOf("SIDEWAYSCW", StringComparison.Ordinal) >= 0,
                    "270°: текст по часовой стал строками, получено: «" + fixedCw[0].Text + "»");
                AssertTrue(Math.Abs(fixedCw[0].WidthPt - 600) < 1 && Math.Abs(fixedCw[0].HeightPt - 400) < 1,
                    "270°: размеры страницы поменялись местами");
                AssertTrue(squash(fixedCw[1]).IndexOf("SIDEWAYSCCW", StringComparison.Ordinal) < 0,
                    "вторая страница без поворота — её боковой текст не извлечён");

                // Текст против часовой — поворотом по часовой (90° CW).
                List<PdfPageText> fixedCcw = PdfTextExtract.Extract(path, null, new[] { 0, 90 });
                AssertTrue(squash(fixedCcw[1]).IndexOf("SIDEWAYSCCW", StringComparison.Ordinal) >= 0,
                    "90°: текст против часовой стал строками, получено: «" + fixedCcw[1].Text + "»");
                AssertTrue(Math.Abs(fixedCcw[1].WidthPt - 600) < 1 && Math.Abs(fixedCcw[1].HeightPt - 400) < 1,
                    "90°: размеры страницы поменялись местами");
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        /// <summary>
        /// Фильтр перетаскиваемых файлов: берутся только существующие .pdf (регистр
        /// расширения не важен), прочее отсекается; данные не-FileDrop → пусто.
        /// Регрессия «файлы над сеткой мертвы» держится именно на этом фильтре.
        /// DataObject/DragEventArgs — managed, STA-контрол не создаётся.
        /// </summary>
        private static void TestPdfDropExtract()
        {
            string dir = Path.Combine(Path.GetTempPath(), "iwo_drop_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string pdf = Path.Combine(dir, "a.pdf");
                File.WriteAllBytes(pdf, new byte[] { 1 });
                string pdfUpper = Path.Combine(dir, "B.PDF");
                File.WriteAllBytes(pdfUpper, new byte[] { 1 });
                string txt = Path.Combine(dir, "c.txt");
                File.WriteAllBytes(txt, new byte[] { 1 });
                string missing = Path.Combine(dir, "ghost.pdf"); // не создаём

                string[] got = PdfDrop.ExtractPaths(MakeFileDrop(new[] { pdf, pdfUpper, txt, missing }));
                Array.Sort(got);
                AssertEqual(2, got.Length, "только существующие .pdf (регистр не важен)");
                AssertTrue(Array.IndexOf(got, pdf) >= 0, "нижний регистр .pdf");
                AssertTrue(Array.IndexOf(got, pdfUpper) >= 0, "верхний регистр .PDF");
                AssertTrue(Array.IndexOf(got, txt) < 0, ".txt отсеян");
                AssertTrue(Array.IndexOf(got, missing) < 0, "несуществующий отсеян");

                AssertEqual(0, PdfDrop.ExtractPaths(MakeFileDrop(new string[0])).Length, "пустой список");
                AssertEqual(0, PdfDrop.ExtractPaths(null).Length, "null-аргумент");

                var text = new System.Windows.Forms.DataObject(System.Windows.Forms.DataFormats.Text, "hi");
                var e = new System.Windows.Forms.DragEventArgs(text, 0, 0, 0,
                    System.Windows.Forms.DragDropEffects.Copy, System.Windows.Forms.DragDropEffects.None);
                AssertEqual(0, PdfDrop.ExtractPaths(e).Length, "не-FileDrop — пусто");
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        private static System.Windows.Forms.DragEventArgs MakeFileDrop(string[] paths)
        {
            var data = new System.Windows.Forms.DataObject();
            data.SetData(System.Windows.Forms.DataFormats.FileDrop, paths);
            return new System.Windows.Forms.DragEventArgs(data, 0, 0, 0,
                System.Windows.Forms.DragDropEffects.Copy, System.Windows.Forms.DragDropEffects.None);
        }

        private static void TestDropInsertIndex()
        {
            // Сетка с порядком: вставка по метке (до/после плитки).
            AssertEqual(2, PdfPageGrid.DropInsertIndex(true, 2, false, 10), "до плитки 2");
            AssertEqual(3, PdfPageGrid.DropInsertIndex(true, 2, true, 10), "после плитки 2");
            // Нет метки (курсор в пустоте) — в конец.
            AssertEqual(10, PdfPageGrid.DropInsertIndex(true, -1, false, 10), "нет метки — в конец");
            // Сетка без порядка («Разделение») — позиция не важна, всегда в конец.
            AssertEqual(10, PdfPageGrid.DropInsertIndex(false, 3, true, 10), "без порядка — в конец");
            AssertEqual(0, PdfPageGrid.DropInsertIndex(true, -1, false, 0), "пустая сетка — 0");
        }

        // ---------- быстрые победы UX 1.17.2 ----------

        private static void TestHintFor()
        {
            AssertEqual(null, PdfPageGrid.HintFor(3, false, "empty", "drop"), "непустой список — без подсказки");
            AssertEqual(null, PdfPageGrid.HintFor(3, true, "empty", "drop"), "непустой во время дропа — без подсказки");
            AssertEqual("empty", PdfPageGrid.HintFor(0, false, "empty", "drop"), "пусто, не дроп — подсказка добавления");
            AssertEqual("drop", PdfPageGrid.HintFor(0, true, "empty", "drop"), "пусто, дроп — «отпустите»");
            AssertEqual("empty", PdfPageGrid.HintFor(0, true, "empty", null), "нет drop-текста — падаем на empty");
        }

        private static void TestRestingStatus()
        {
            AssertEqual("3 / 10", PdfToolFormBase.RestingStatus(3, 10, "idle", "{0} / {1}"), "есть выделение");
            AssertEqual("idle", PdfToolFormBase.RestingStatus(0, 10, "idle", "{0} / {1}"), "нет выделения — idle");
            AssertEqual("1 / 1", PdfToolFormBase.RestingStatus(1, 1, "idle", "{0} / {1}"), "выбрана одна из одной");
            AssertEqual("", PdfToolFormBase.RestingStatus(0, 0, "", "{0} / {1}"), "пусто — пустой idle");
        }

        /// <summary>
        /// Пунктуацию статуса собирает одно место: галочка спереди, « · » между частями,
        /// точка в конце. Пустые и null-части выпадают, поэтому «сжато» просто не
        /// добавляется, когда сжатия не было, и лишних разделителей не остаётся.
        /// </summary>
        private static void TestSuccessStatus()
        {
            AssertEqual("✓ Сохранено страниц: 12.", PdfToolFormBase.SuccessStatus("Сохранено страниц: 12"),
                "одна часть — галочка и точка");
            AssertEqual("✓ Сохранено страниц: 12 · сжато.", PdfToolFormBase.SuccessStatus("Сохранено страниц: 12", "сжато"),
                "две части через разделитель");
            AssertEqual("✓ a · b.", PdfToolFormBase.SuccessStatus("a", null, "", "b"),
                "null и пустые части выпадают без лишних разделителей");
            AssertEqual("", PdfToolFormBase.SuccessStatus(), "нет частей — пустая строка");
            AssertEqual("", PdfToolFormBase.SuccessStatus(null, ""), "все части пустые — пустая строка");
            AssertEqual("", PdfToolFormBase.SuccessStatus(null), "null вместо массива — пустая строка");
        }

        /// <summary>
        /// Разрешение сжатия берётся из одного места и совпадает с пресетами Ghostscript
        /// (Resource\Init\gs_pdfwr.ps: /ebook — 150, /screen — 72). Уровень без сжатия
        /// разрешения не имеет, и часть статуса про сжатие для него не строится.
        /// </summary>
        private static void TestCompressionDpi()
        {
            AssertEqual(150, PdfCompression.ImageDpi(CompressionLevel.Good), "/ebook — 150 dpi");
            AssertEqual(72, PdfCompression.ImageDpi(CompressionLevel.Small), "/screen — 72 dpi");
            AssertEqual(0, PdfCompression.ImageDpi(CompressionLevel.None), "без сжатия — разрешения нет");
            AssertEqual(0, PdfCompression.ImageDpi(CompressionLevel.VeryGood),
                "изображения не пересчитываются — разрешения у уровня нет");
            // Уровни и пресеты не должны разъезжаться: у кого есть разрешение, у того есть и
            // пресет. Обратное с 1.18.2 неверно — «Очень хорошо» имеет пресет без пересчёта.
            foreach (CompressionLevel level in Enum.GetValues(typeof(CompressionLevel)))
                if (PdfCompression.ImageDpi(level) > 0)
                    AssertTrue(PdfCompression.Preset(level) != null, "разрешение без пресета у " + level);
            AssertEqual(null, PdfToolFormBase.CompressedPart(false, CompressionLevel.None),
                "сжатия не заказывали — части нет");
            // А вот заказанное сжатие, не давшее выигрыша, обязано отчитаться словами:
            // прежде статус в этом случае молчал, и человек не знал, сработало ли оно.
            foreach (CompressionLevel level in new[]
                { CompressionLevel.VeryGood, CompressionLevel.Good, CompressionLevel.Small })
            {
                string quiet = PdfToolFormBase.CompressedPart(false, level);
                AssertTrue(!string.IsNullOrEmpty(quiet), "заказанное сжатие без выигрыша не молчит: " + level);
                AssertTrue(quiet != PdfToolFormBase.CompressedPart(true, level),
                    "«не уменьшило» и «сжато» — разные слова: " + level);
            }
            AssertTrue(PdfToolFormBase.CompressedPart(true, CompressionLevel.Good).Contains("150"),
                "часть про сжатие называет разрешение уровня");
            AssertTrue(PdfToolFormBase.CompressedPart(true, CompressionLevel.Small).Contains("72"),
                "минимальный размер — 72 dpi");
            // А уровень без пересчёта не должен называть ни своего «нуля», ни чужих чисел.
            string keptPart = PdfToolFormBase.CompressedPart(true, CompressionLevel.VeryGood);
            AssertTrue(!keptPart.Contains("0 dpi") && !keptPart.Contains("150") && !keptPart.Contains("72"),
                "часть про сжатие без пересчёта не называет разрешение: " + keptPart);
        }

        /// <summary>
        /// Догон поворота в предпросмотре: картинка хранит уже впечённый угол, а нужный
        /// берётся из страницы, поэтому доворачивать надо ровно на их разницу. Разница
        /// нормализована в {0, 90, 180, 270}, совпадение углов даёт 0 (лишнего поворота нет),
        /// и после доворота впечённый угол равен нужному при любой паре углов.
        /// </summary>
        private static void TestPreviewRotationDelta()
        {
            AssertEqual(90, PdfPageRef.ComposeRotation(90, -0), "с 0 на 90");
            AssertEqual(180, PdfPageRef.ComposeRotation(270, -90), "с 90 на 270");
            AssertEqual(270, PdfPageRef.ComposeRotation(0, -90), "с 90 на 0 — доворот на 270, а не -90");
            AssertEqual(0, PdfPageRef.ComposeRotation(180, -180), "углы совпали — доворачивать нечего");

            int[] angles = { 0, 90, 180, 270 };
            foreach (int applied in angles)
                foreach (int desired in angles)
                {
                    int delta = PdfPageRef.ComposeRotation(desired, -applied);
                    AssertEqual(desired, PdfPageRef.ComposeRotation(applied, delta),
                        "доворот " + applied + " на " + delta + " должен дать " + desired);
                    AssertEqual(applied == desired, delta == 0,
                        "доворот нужен тогда и только тогда, когда углы разошлись (" + applied + "→" + desired + ")");
                    AssertEqual(applied == desired,
                        PageRotation.FlipFor(delta) == System.Drawing.RotateFlipType.RotateNoneFlipNone,
                        "ненулевой доворот обязан дать реальный RotateFlip (" + delta + ")");
                }
        }

        /// <summary>
        /// Готовые строки результата на обоих языках: значения подставляются в нужные места
        /// (иначе string.Format свалился бы уже в бою), разрешение называется только когда
        /// сжатие сработало, разбиение считает файлы, а извлечение — страницы.
        /// </summary>
        private static void TestDoneStatusLines()
        {
            Lang saved = Loc.Current;
            try
            {
                foreach (Lang lang in new[] { Lang.Ru, Lang.En })
                {
                    Loc.Init(lang);
                    string merged = PdfMergeForm.DoneStatus(12, true, CompressionLevel.Good);
                    AssertTrue(merged.StartsWith("✓ "), "строка начинается с галочки (" + lang + ")");
                    AssertTrue(merged.EndsWith("."), "строка кончается точкой (" + lang + ")");
                    AssertTrue(merged.Contains("12"), "названо число страниц (" + lang + ")");
                    AssertTrue(merged.Contains("150"), "названо разрешение сжатия (" + lang + ")");

                    string plain = PdfMergeForm.DoneStatus(3, false, CompressionLevel.None);
                    AssertTrue(plain.Contains("3"), "число страниц без сжатия (" + lang + ")");
                    AssertTrue(!plain.Contains("dpi"), "без сжатия разрешение не упоминается (" + lang + ")");

                    string parts = PdfSplitForm.DoneStatus(true, 4, 99, 4, CompressionLevel.Small);
                    AssertTrue(parts.Contains("4"), "названо число файлов (" + lang + ")");
                    AssertTrue(parts.Contains("72"), "названо разрешение /screen (" + lang + ")");
                    AssertTrue(!parts.Contains("99"), "страницы в разбиении на части не показываем (" + lang + ")");

                    string extracted = PdfSplitForm.DoneStatus(false, 1, 7, 1, CompressionLevel.Good);
                    AssertTrue(extracted.Contains("7"), "названо число извлечённых страниц (" + lang + ")");
                    AssertTrue(extracted.Contains("150"), "названо разрешение при извлечении (" + lang + ")");

                    string extractedPlain = PdfSplitForm.DoneStatus(false, 1, 7, 0, CompressionLevel.None);
                    AssertTrue(!extractedPlain.Contains("dpi"), "без сжатия строка чистая (" + lang + ")");
                    AssertTrue(!extractedPlain.Contains(" · "), "единственная часть — без разделителя (" + lang + ")");
                }
            }
            finally { Loc.Init(saved); }
        }

        private static void TestProgressItem()
        {
            AssertEqual(0, PdfToolFormBase.ProgressItem(50, 0), "нет единиц — 0");
            AssertEqual(1, PdfToolFormBase.ProgressItem(0, 10), "0% — минимум 1");
            AssertEqual(5, PdfToolFormBase.ProgressItem(50, 10), "50% из 10 — 5");
            AssertEqual(10, PdfToolFormBase.ProgressItem(100, 10), "100% — всё");
            AssertEqual(3, PdfToolFormBase.ProgressItem(37, 8), "37% из 8 ≈ 3");
            AssertEqual(10, PdfToolFormBase.ProgressItem(150, 10), "процент сверх 100 клампится");
            AssertEqual(2, PdfToolFormBase.ProgressItem(1, 200), "1% из 200 — пропорционально 2");
            AssertEqual(1, PdfToolFormBase.ProgressItem(0, 200), "0% — минимум 1, не 0");
        }

        private static void TestBuildShortcuts()
        {
            Func<string, string> t = delegate(string k) { return k; }; // «переводчик» возвращает ключ
            string split = PdfToolFormBase.BuildShortcuts(false, true, t);   // разделение: без reorder, с rotate
            AssertTrue(split.Contains("shortcuts.zoom") && split.Contains("shortcuts.goto"), "общие клавиши всегда");
            AssertTrue(split.Contains("shortcuts.rotate"), "поворот при rotate");
            AssertTrue(!split.Contains("shortcuts.cutcopy"), "без reorder — нет буфера страниц");
            // Поворот попадает в историю, значит Ctrl+Z работает и без права переставлять
            // страницы: раньше шпаргалка о нём молчала, а клавиша была заперта.
            AssertTrue(split.Contains("shortcuts.undo"), "откат при rotate — поворот откатывается");

            string merge = PdfToolFormBase.BuildShortcuts(true, true, t);    // объединение: всё
            AssertTrue(merge.Contains("shortcuts.cutcopy") && merge.Contains("shortcuts.paste") &&
                merge.Contains("shortcuts.undo") && merge.Contains("shortcuts.move"), "reorder — буфер/перемещение/отмена");

            string plain = PdfToolFormBase.BuildShortcuts(false, false, t);  // ни reorder, ни rotate
            AssertTrue(!plain.Contains("shortcuts.rotate") && !plain.Contains("shortcuts.cutcopy"), "минимальный набор");
            AssertTrue(!plain.Contains("shortcuts.undo"), "нечего править — нечего и откатывать");
            AssertTrue(plain.Contains("shortcuts.selectAll"), "выделить всё — общее");
            AssertTrue(!plain.EndsWith("\n") && !plain.EndsWith("\r"), "хвостовые переводы срезаны");
        }

        private static void TestCardFilter()
        {
            string dir = Path.Combine(Path.GetTempPath(), "iwo_card_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string pdf = Path.Combine(dir, "a.pdf"); File.WriteAllBytes(pdf, new byte[] { 1 });
                string pdfUp = Path.Combine(dir, "B.PDF"); File.WriteAllBytes(pdfUp, new byte[] { 1 });
                string txt = Path.Combine(dir, "c.txt"); File.WriteAllBytes(txt, new byte[] { 1 });
                string ghost = Path.Combine(dir, "no.pdf"); // не создаём

                string[] got = ChoiceCard.FilterByExtension(new[] { pdf, pdfUp, txt, ghost }, new[] { ".pdf" });
                Array.Sort(got);
                AssertEqual(2, got.Length, "только существующие .pdf (регистр не важен)");
                AssertTrue(Array.IndexOf(got, pdf) >= 0 && Array.IndexOf(got, pdfUp) >= 0, "оба регистра .pdf");
                AssertEqual(0, ChoiceCard.FilterByExtension(new[] { txt }, new[] { ".pdf" }).Length, ".txt отсеян");
                AssertEqual(0, ChoiceCard.FilterByExtension(null, new[] { ".pdf" }).Length, "null-пути — пусто");
                AssertEqual(0, ChoiceCard.FilterByExtension(new[] { pdf }, null).Length, "null-расширения — пусто");
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        // ---------- 1.17.3: масштаб-%, полоса номера, отмена операций ----------

        private static void TestZoomPercent()
        {
            AssertEqual(100, ThumbZoom.Percent(ThumbZoom.DefaultWidth), "умолчание — 100%");
            AssertTrue(ThumbZoom.Percent(ThumbZoom.MinWidth) < 100, "минимум < 100%");
            AssertTrue(ThumbZoom.Percent(ThumbZoom.MaxWidth) > 100, "максимум > 100%");
            AssertEqual(ThumbZoom.Percent(ThumbZoom.MaxWidth), ThumbZoom.Percent(10000), "сверх максимума клампится");
        }

        private static void TestWidthFromPercent()
        {
            AssertEqual(ThumbZoom.DefaultWidth, ThumbZoom.WidthFromPercent(100), "100% → ширина по умолчанию");
            // Round-trip на границах: процент границы обратно даёт ту же граничную ширину.
            AssertEqual(ThumbZoom.MinWidth, ThumbZoom.WidthFromPercent(ThumbZoom.Percent(ThumbZoom.MinWidth)), "round-trip: минимум");
            AssertEqual(ThumbZoom.MaxWidth, ThumbZoom.WidthFromPercent(ThumbZoom.Percent(ThumbZoom.MaxWidth)), "round-trip: максимум");
            AssertEqual(ThumbZoom.MinWidth, ThumbZoom.WidthFromPercent(1), "ниже минимума клампится");
            AssertEqual(ThumbZoom.MaxWidth, ThumbZoom.WidthFromPercent(9999), "выше максимума клампится");
            AssertTrue(ThumbZoom.WidthFromPercent(50) < ThumbZoom.DefaultWidth, "50% < умолчания");
            AssertEqual(ThumbZoom.Percent(ThumbZoom.MinWidth), ThumbZoom.MinPercent, "MinPercent = Percent(MinWidth)");
            AssertEqual(ThumbZoom.Percent(ThumbZoom.MaxWidth), ThumbZoom.MaxPercent, "MaxPercent = Percent(MaxWidth)");
        }

        private static void TestIsOnLabel()
        {
            var tile = new System.Drawing.Rectangle(100, 50, 132, 172); // низ плитки = 222
            var cell = new System.Drawing.Rectangle(90, 46, 152, 210);  // низ ячейки = 256
            AssertTrue(PdfPageGrid.IsOnLabel(new System.Drawing.Point(150, 235), tile, cell), "точка под плиткой — номер");
            AssertTrue(!PdfPageGrid.IsOnLabel(new System.Drawing.Point(150, 120), tile, cell), "точка в плитке — не номер");
            AssertTrue(!PdfPageGrid.IsOnLabel(new System.Drawing.Point(150, 300), tile, cell), "ниже ячейки — не номер");
            AssertTrue(!PdfPageGrid.IsOnLabel(new System.Drawing.Point(50, 235), tile, cell), "левее ячейки — не номер");
        }

        /// <summary>
        /// Регресс: масштаб/сжатие принадлежат PDF-окнам. Долгоживущий экземпляр (MainForm
        /// грузит настройки при старте) не должен затирать их устаревшим значением при своём
        /// Save — общий Save обязан перечитать эти поля с диска, а явно писать их лишь SaveView.
        /// Живой %APPDATA% не трогается: тест работает в изолированном корне AppPaths.
        /// </summary>
        private static void TestSettingsViewNotClobbered()
        {
            string root = Path.Combine(Path.GetTempPath(), "iwo_settings_" + Guid.NewGuid().ToString("N"));
            AppPaths.SetRootForTests(root);
            try
            {
                // 1. PDF-окно сохраняет вид: масштаб 300, сжатие 2.
                new UserSettings().SaveView(300, 2);
                AssertEqual(300, UserSettings.Load().ZoomWidth, "SaveView записал масштаб");
                AssertEqual(2, UserSettings.Load().CompressionLevel, "SaveView записал сжатие");

                // 2. Устаревший экземпляр (как у MainForm: масштаб 100) сохраняет СВОИ поля.
                var stale = new UserSettings { LastInputFolder = "X", ZoomWidth = 100, CompressionLevel = 0 };
                stale.Save();

                // Масштаб/сжатие взяты с диска (300/2), а не из устаревшего экземпляра (100/0).
                UserSettings after = UserSettings.Load();
                AssertEqual(300, after.ZoomWidth, "общий Save НЕ затёр масштаб");
                AssertEqual(2, after.CompressionLevel, "общий Save НЕ затёр сжатие");
                AssertEqual("X", after.LastInputFolder, "общий Save сохранил собственные поля экземпляра");
            }
            finally
            {
                AppPaths.SetRootForTests(null);
                try { Directory.Delete(root, true); } catch { }
            }
        }

        /// <summary>
        /// Настройки обновлений живут по тому же правилу, что масштаб и границы окон: своего
        /// окна-владельца у них нет, поэтому общий Save из долгоживущей формы обязан оставить
        /// их в покое. Иначе просьба «не напоминать» отменялась бы сама собой при следующем
        /// закрытии любого окна — дефект, который вручную не поймать никогда.
        /// </summary>
        private static void TestUpdatePrefsNotClobbered()
        {
            string root = Path.Combine(Path.GetTempPath(), "iwo_upd_" + Guid.NewGuid().ToString("N"));
            AppPaths.SetRootForTests(root);
            try
            {
                AssertTrue(UserSettings.Load().UpdateCheckOnStart, "по умолчанию проверка при запуске включена");
                AssertEqual(null, UserSettings.Load().SkippedVersion, "по умолчанию ничего не пропущено");

                UserSettings.SaveSkippedVersion("1.18.0");
                UserSettings.SaveUpdateCheckOnStart(false);
                AssertEqual("1.18.0", UserSettings.Load().SkippedVersion, "пропущенная версия записана");
                AssertTrue(!UserSettings.Load().UpdateCheckOnStart, "проверка при запуске выключена");

                // Устаревший экземпляр (создан ДО этих правок) сохраняет свои поля.
                var stale = new UserSettings { LastInputFolder = "Y" };
                stale.Save();

                UserSettings after = UserSettings.Load();
                AssertEqual("1.18.0", after.SkippedVersion, "общий Save НЕ затёр пропущенную версию");
                AssertTrue(!after.UpdateCheckOnStart, "общий Save НЕ включил проверку обратно");
                AssertEqual("Y", after.LastInputFolder, "общий Save сохранил собственные поля экземпляра");

                // Узкие методы не мешают друг другу: запись версии не трогает флаг и наоборот.
                UserSettings.SaveSkippedVersion("1.18.1");
                AssertTrue(!UserSettings.Load().UpdateCheckOnStart, "запись версии не включила проверку");
                UserSettings.SaveUpdateCheckOnStart(true);
                AssertEqual("1.18.1", UserSettings.Load().SkippedVersion, "запись флага не стёрла версию");
            }
            finally
            {
                AppPaths.SetRootForTests(null);
                try { Directory.Delete(root, true); } catch { }
            }
        }

        private static void TestWindowBoundsPersistence()
        {
            string root = Path.Combine(Path.GetTempPath(), "iwo_wnd_" + Guid.NewGuid().ToString("N"));
            AppPaths.SetRootForTests(root);
            try
            {
                // Границы двух окон сохраняются и читаются обратно.
                UserSettings.SaveWindowBounds("PdfMergeForm", "10,20,780,660,0");
                UserSettings.SaveWindowBounds("MainForm", "0,0,760,725,1");
                UserSettings loaded = UserSettings.Load();
                string a, b;
                AssertTrue(loaded.WindowBounds.TryGetValue("PdfMergeForm", out a) && a == "10,20,780,660,0", "границы Merge сохранены");
                AssertTrue(loaded.WindowBounds.TryGetValue("MainForm", out b) && b == "0,0,760,725,1", "границы Main сохранены");

                // Устаревший экземпляр (загружен ДО правки) своим Save() НЕ затирает чужие границы.
                UserSettings stale = UserSettings.Load();
                UserSettings.SaveWindowBounds("PdfSplitForm", "5,5,700,560,0"); // другое окно записало свои
                stale.LastInputFolder = "Y";
                stale.Save();
                UserSettings after = UserSettings.Load();
                string c;
                AssertTrue(after.WindowBounds.TryGetValue("PdfSplitForm", out c) && c == "5,5,700,560,0", "чужие границы не затёрты устаревшим Save");
                AssertEqual("Y", after.LastInputFolder, "устаревший экземпляр сохранил свои поля");
            }
            finally
            {
                AppPaths.SetRootForTests(null);
                try { Directory.Delete(root, true); } catch { }
            }
        }

        /// <summary>Пустое окно со своим именем типа: ключ хранения берётся от имени класса.</summary>
        private sealed class PlacementProbeForm : System.Windows.Forms.Form
        {
        }

        /// <summary>
        /// ЖИВАЯ проводка запоминания окон: чистые функции проверены отдельно, а здесь
        /// настоящее окно показывается и закрывается, и границы обязаны записаться и
        /// восстановиться. Без такого теста подключение можно случайно оторвать, и отказ
        /// будет тихим — окно просто перестанет помнить место или откроется за краем экрана.
        /// Заодно проверяется, что БЕЗ подключения не сохраняется ничего, иначе тест мог бы
        /// зеленеть по чужой причине.
        /// </summary>
        private static void TestWindowPlacementAttachLive()
        {
            string root = Path.Combine(Path.GetTempPath(), "iwo_attach_" + Guid.NewGuid().ToString("N"));
            AppPaths.SetRootForTests(root);
            string failure = null;
            var th = new System.Threading.Thread(delegate()
            {
                try
                {
                    const string key = "PlacementProbeForm";
                    var moved = new System.Drawing.Rectangle(40, 30, 500, 400);

                    // Контроль: без Attach окно ничего о себе не пишет.
                    using (var bare = new PlacementProbeForm())
                    {
                        bare.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                        bare.Bounds = moved;
                        bare.Show();
                        bare.Close();
                    }
                    string ignored;
                    if (UserSettings.Load().WindowBounds.TryGetValue(key, out ignored))
                    {
                        failure = "без подключения границы всё равно сохранились";
                        return;
                    }

                    using (var f = new PlacementProbeForm())
                    {
                        f.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                        WindowPlacement.Attach(f);
                        f.Show();
                        f.Bounds = moved; // пользователь подвинул и изменил размер
                        f.Close();
                    }
                    string saved;
                    if (!UserSettings.Load().WindowBounds.TryGetValue(key, out saved))
                    {
                        failure = "границы не сохранились при закрытии";
                        return;
                    }
                    if (saved != WindowPlacement.Format(moved, false))
                    {
                        failure = "сохранено «" + saved + "», ожидалось «" + WindowPlacement.Format(moved, false) + "»";
                        return;
                    }

                    // Новое окно того же типа обязано встать туда же (с поправкой на кламп,
                    // потому что экран прогонщика может быть меньше сохранённого места).
                    using (var again = new PlacementProbeForm())
                    {
                        WindowPlacement.Attach(again);
                        again.Show();
                        System.Drawing.Rectangle[] areas = ScreenWorkAreas();
                        System.Drawing.Rectangle want =
                            WindowPlacement.ClampToWorkingArea(moved, areas, again.MinimumSize);
                        if (again.Bounds != want)
                            failure = "восстановлено " + again.Bounds + ", ожидалось " + want;
                        else if (again.StartPosition != System.Windows.Forms.FormStartPosition.Manual)
                            failure = "восстановление не перевело окно в ручное позиционирование";
                        again.Close();
                    }
                    if (failure != null)
                        return;

                    // Подключение живёт одной строкой в конструкторе, оторвать его легко и
                    // незаметно, поэтому проверяем КАЖДОЕ настоящее окно, которое обязано
                    // помнить своё место: главное, любой PDF-инструмент (за всю базу) и
                    // полноэкранный просмотр.
                    failure = RoundTripPlacement("MainForm", new MainForm(null));
                    if (failure != null) return;
                    failure = RoundTripPlacement("PdfMergeForm", new PdfMergeForm(null));
                    if (failure != null) return;
                    failure = RoundTripPlacement("PagePreviewForm", NewPreviewForm());
                }
                catch (Exception ex) { failure = ex.GetType().Name + ": " + ex.Message; }
            });
            th.SetApartmentState(System.Threading.ApartmentState.STA); // окна WinForms требуют STA
            th.IsBackground = true;
            th.Start();
            th.Join();
            AppPaths.SetRootForTests(null);
            try { Directory.Delete(root, true); } catch { }
            AssertTrue(failure == null, "WindowPlacement.Attach: " + failure);
        }

        /// <summary>
        /// Показать окно, подвинуть, закрыть — и убедиться, что оно записало свои границы под
        /// своим именем. Возвращает null при успехе или причину отказа. Окно освобождается.
        /// </summary>
        private static string RoundTripPlacement(string key, System.Windows.Forms.Form form)
        {
            var moved = new System.Drawing.Rectangle(60, 50, 620, 480);
            try
            {
                using (form)
                {
                    form.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                    form.Show();
                    form.Bounds = moved;
                    form.Close();
                }
            }
            catch (Exception ex) { return key + ": окно не отработало — " + ex.GetType().Name + " " + ex.Message; }
            string saved;
            if (!UserSettings.Load().WindowBounds.TryGetValue(key, out saved))
                return key + ": окно не запомнило своё место (потеряно подключение WindowPlacement)";
            // Размер мог не примениться целиком из-за MinimumSize окна, поэтому сверяем
            // с тем, что окно реально имело: важно, что записано именно оно, а не пусто.
            if (string.IsNullOrEmpty(saved))
                return key + ": записана пустая строка границ";
            return null;
        }

        /// <summary>
        /// Полноэкранный просмотр для проверки запоминания места. Конструктор приватный, а путь
        /// к PDF намеренно несуществующий: фоновый рендер тихо не найдёт файл и покажет
        /// «недоступно», чего для проверки границ окна достаточно.
        /// </summary>
        private static System.Windows.Forms.Form NewPreviewForm()
        {
            object page = new PdfPageRef { SourcePath = "нет-такого.pdf", PageIndex = 0, Rotation = 0 };
            return (System.Windows.Forms.Form)Activator.CreateInstance(
                typeof(PagePreviewForm),
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                null,
                new object[] { page, "placement", new System.Drawing.Size(640, 520), null },
                null);
        }

        /// <summary>Рабочие области всех экранов — тот же набор, что использует восстановление.</summary>
        private static System.Drawing.Rectangle[] ScreenWorkAreas()
        {
            System.Windows.Forms.Screen[] screens = System.Windows.Forms.Screen.AllScreens;
            var areas = new System.Drawing.Rectangle[screens.Length];
            for (int i = 0; i < screens.Length; i++)
                areas[i] = screens[i].WorkingArea;
            return areas;
        }

        private static void TestShouldOfferCancel()
        {
            AssertTrue(!PdfToolFormBase.ShouldOfferCancel(1), "1 страница — без отмены");
            AssertTrue(!PdfToolFormBase.ShouldOfferCancel(4), "4 — без отмены");
            AssertTrue(PdfToolFormBase.ShouldOfferCancel(5), "5 — отмена доступна (порог)");
            AssertTrue(PdfToolFormBase.ShouldOfferCancel(500), "много — отмена доступна");
        }

        /// <summary>Многостраничный PDF во временной папке (для живых проверок отмены).</summary>
        private static string MakePagesPdf(string dir, int pages)
        {
            string path = Path.Combine(dir, "cancel_src.pdf");
            using (var doc = new PdfDocument())
            {
                for (int i = 0; i < pages; i++)
                    doc.AddPage();
                doc.Save(path);
            }
            return path;
        }

        private static void TestCancelMergeLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "iwo_cx_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = MakePagesPdf(dir, 6);
                string outPath = Path.Combine(dir, "out.pdf");
                var order = new List<PdfPageRef>();
                for (int i = 0; i < 6; i++)
                    order.Add(new PdfPageRef { SourcePath = src, PageIndex = i });

                bool threw = false;
                try { PdfMergeService.Merge(order, outPath, null, delegate { return true; }); }
                catch (OperationCanceledException) { threw = true; }
                AssertTrue(threw, "отмена бросает OperationCanceledException");
                AssertTrue(!File.Exists(outPath), "при отмене файл не создан (сохранение — только в конце)");
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        private static void TestCancelSplitLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "iwo_cs_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = MakePagesPdf(dir, 3);
                // Отмена ПОСЛЕ первой части: первый pre-check пропускает, второй — отменяет.
                int calls = 0;
                Func<bool> cancel = delegate { return calls++ >= 1; };
                bool threw = false;
                try { PdfSplitService.SplitEveryN(src, 1, dir, "part", null, null, cancel); }
                catch (OperationCanceledException) { threw = true; }
                AssertTrue(threw, "отмена разбиения бросает OperationCanceledException");
                string[] parts = Directory.GetFiles(dir, "part*.pdf");
                AssertEqual(0, parts.Length, "частичные файлы удалены при отмене (осталось 0)");
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        /// <summary>
        /// Тот же инвариант для сохранения страниц картинками: этот экспорт пишет ОТДЕЛЬНЫЙ
        /// файл на каждую страницу, поэтому одного «бросило исключение» мало — папка обязана
        /// остаться пустой. Прогон БЕЗ отмены здесь не для полноты: без него «файлов 0»
        /// зеленело бы и на сломанном коде (например, если бы рендер не создавал вообще
        /// ничего), то есть тест не мог бы упасть по той причине, ради которой написан.
        /// </summary>
        private static void TestCancelExportImagesLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "iwo_ce_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = MakePagesPdf(dir, 3);
                var pages = new List<int> { 0, 1, 2 };

                string whole = Path.Combine(dir, "whole");
                PdfExportService.ToImages(src, pages, whole, NameTemplate.Default, ImageExportFormat.Png, 96);
                AssertEqual(3, Directory.GetFiles(whole).Length, "без отмены сохраняется файл на страницу");

                // Отмена ПОСЛЕ первой картинки: первый pre-check пропускает, второй — отменяет.
                string part = Path.Combine(dir, "part");
                int calls = 0;
                Func<bool> cancel = delegate { return calls++ >= 1; };
                bool threw = false;
                try
                {
                    PdfExportService.ToImages(src, pages, part, NameTemplate.Default,
                        ImageExportFormat.Png, 96, null, cancel);
                }
                catch (OperationCanceledException) { threw = true; }
                AssertTrue(threw, "отмена экспорта бросает OperationCanceledException");
                AssertEqual(0, Directory.GetFiles(part).Length, "сохранённые картинки удалены при отмене (осталось 0)");
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        /// <summary>
        /// «Прочие операции» с доступом к прологу и эпилогу операции: сама форма начинает их
        /// изнутри своих кнопок, за диалогами выбора папки и формата, а для проверки проводки
        /// отмены нужен вход без диалогов и без настоящей работы.
        /// </summary>
        private sealed class OpsCancelProbe : PdfOpsForm
        {
            public OpsCancelProbe() : base(delegate { }) { }
            public void Begin(int units) { BeginOperation(Loc.T("ops.status.exporting"), units); }
            public void End() { EndOperation(); }
            public bool Canceled { get { return CancelToken()(); } }
        }

        /// <summary>
        /// Проводка отмены в «Прочих операциях» (живое окно). Удаление частичного результата
        /// в сервисе бессмысленно, если окно не предложит отмену, а у этого окна ветка показа
        /// СВОЯ: единственной кнопки действия у него нет, и «Отмена» живёт в отдельном углу
        /// (RegisterCancelArea) — то есть проверенная на других инструментах подмена кнопки
        /// здесь ничего не доказывает. Порог берём у самого приложения, а не литералом.
        /// </summary>
        private static void TestOpsCancelWiringLive()
        {
            var offenders = new List<string>();
            int enough = 1;
            while (!PdfToolFormBase.ShouldOfferCancel(enough) && enough < 1000)
                enough++;
            InIsolatedSettings("iwo_opscancel_", delegate
            {
                try
                {
                    using (var f = new OpsCancelProbe())
                    {
                        f.Show();
                        if (FindVisibleButton(f, Loc.T("common.cancel")) != null)
                            offenders.Add("«Отмена» видна до начала операции");

                        f.Begin(enough - 1); // короче порога — отмену не предлагаем
                        if (FindVisibleButton(f, Loc.T("common.cancel")) != null)
                            offenders.Add("«Отмена» показана для операции короче порога");
                        f.End();

                        f.Begin(enough);
                        System.Windows.Forms.Button cancel = FindVisibleButton(f, Loc.T("common.cancel"));
                        if (cancel == null)
                        {
                            offenders.Add("«Отмена» не показана на операции от порога");
                        }
                        else
                        {
                            if (f.Canceled)
                                offenders.Add("токен отмены поднят ещё до нажатия");
                            cancel.PerformClick();
                            if (!f.Canceled)
                                offenders.Add("нажатие «Отмены» не поднимает токен для сервиса");
                            if (cancel.Enabled)
                                offenders.Add("нажатая «Отмена» осталась доступной (повторные нажатия)");
                        }
                        f.End();
                        if (FindVisibleButton(f, Loc.T("common.cancel")) != null ||
                            FindVisibleButton(f, Loc.T("common.canceling")) != null)
                            offenders.Add("«Отмена» осталась на экране после операции");
                        f.Close();
                    }
                }
                catch (Exception ex) { offenders.Add("не собралось: " + Root(ex).Message); }
            });
            AssertTrue(offenders.Count == 0, "проводка отмены: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>Видимая кнопка с такой подписью (вглубь по дереву контролов) или null.</summary>
        private static System.Windows.Forms.Button FindVisibleButton(System.Windows.Forms.Control root, string text)
        {
            foreach (System.Windows.Forms.Control c in root.Controls)
            {
                var b = c as System.Windows.Forms.Button;
                if (b != null && b.Visible && b.Text == text)
                    return b;
                System.Windows.Forms.Button deeper = FindVisibleButton(c, text);
                if (deeper != null)
                    return deeper;
            }
            return null;
        }

        // ---------- 1.17.4 ----------

        /// <summary>
        /// Регресс 1.17.3: двойной клик по hover-кнопке поворота открывал предпросмотр
        /// вместо второго поворота. Приоритет: чип поворота > номер под плиткой > плитка.
        /// </summary>
        private static void TestClassifyDoubleClick()
        {
            var tile = new System.Drawing.Rectangle(20, 10, 200, 260);
            var cell = new System.Drawing.Rectangle(10, 6, 220, 290);
            System.Drawing.Rectangle[] chips = PdfPageGrid.HoverRotateButtons(tile, 24, 6);
            var onChip = new System.Drawing.Point(chips[1].X + 5, chips[1].Y + 5);
            var onTile = new System.Drawing.Point(tile.X + 30, tile.Y + 30);
            var onLabel = new System.Drawing.Point(cell.X + cell.Width / 2, tile.Bottom + 8);

            AssertEqual(PdfPageGrid.DoubleClickAction.RotateChip,
                PdfPageGrid.ClassifyDoubleClick(onChip, tile, cell, chips, true), "чип поворота — проглотить");
            AssertEqual(PdfPageGrid.DoubleClickAction.Preview,
                PdfPageGrid.ClassifyDoubleClick(onTile, tile, cell, chips, true), "плитка — предпросмотр");
            AssertEqual(PdfPageGrid.DoubleClickAction.MoveAfter,
                PdfPageGrid.ClassifyDoubleClick(onLabel, tile, cell, chips, true), "номер — перенос");
            AssertEqual(PdfPageGrid.DoubleClickAction.Preview,
                PdfPageGrid.ClassifyDoubleClick(onLabel, tile, cell, chips, false), "номер без reorder — предпросмотр");
            AssertEqual(PdfPageGrid.DoubleClickAction.Preview,
                PdfPageGrid.ClassifyDoubleClick(onChip, tile, cell, null, true),
                "кнопки не показаны (не hot/Locked) — обычная плитка");
        }

        private static void TestNumberDialogWidth()
        {
            AssertEqual(300, NumberPromptDialog.DialogWidth(100, 16, 300), "короткая подпись — базовая ширина");
            AssertEqual(300, NumberPromptDialog.DialogWidth(268, 16, 300), "ровно впритык — базовая");
            AssertEqual(333, NumberPromptDialog.DialogWidth(301, 16, 300), "длинная подпись расширяет окно с полями");
        }

        private static void TestAppIconCached()
        {
            System.Drawing.Icon first = Ui.AppIcon();
            System.Drawing.Icon second = Ui.AppIcon();
            // Без этой проверки тест не мог упасть: у неудачной выборки кэшируется null, а
            // ReferenceEquals(null, null) — истина. Из своего exe иконка достаётся всегда.
            AssertTrue(first != null, "иконка приложения извлечена из exe");
            AssertTrue(ReferenceEquals(first, second), "оба вызова возвращают ОДИН экземпляр (HICON не плодится)");
        }

        private static void TestFontCached()
        {
            AssertTrue(ReferenceEquals(Ui.Font(9.75f), Ui.Font(9.75f)), "один экземпляр на одинаковый запрос");
            AssertTrue(!ReferenceEquals(Ui.Font(9.75f), Ui.Font(9.75f, System.Drawing.FontStyle.Bold)),
                "стиль различает записи кэша");
            AssertTrue(!ReferenceEquals(Ui.Font(9.75f), Ui.Font(11f)), "размер различает записи кэша");
            System.Drawing.Font symbol = Ui.Font(9.5f, System.Drawing.FontStyle.Regular, "Segoe UI Symbol");
            AssertTrue(ReferenceEquals(symbol, Ui.Font(9.5f, System.Drawing.FontStyle.Regular, "Segoe UI Symbol")),
                "семейство — часть ключа");
            AssertEqual(9.75f, Ui.Font(9.75f).Size, "размер шрифта соответствует запросу");
        }

        /// <summary>
        /// Ctrl+колесо в сетке применяется троттлингом регулятора: пока цель не применена,
        /// следующие щелчки шагают от неё — быстрое вращение не теряет шаги.
        /// </summary>
        private static void TestWheelBasis()
        {
            AssertEqual(132, PdfPageGrid.WheelBasis(0, 132), "без цели — текущая ширина");
            AssertEqual(148, PdfPageGrid.WheelBasis(148, 132), "неприменённая цель старше");
            // Последовательность из двух быстрых щелчков «+»: вторая цель считается от первой.
            int first = ThumbZoom.StepFromWheel(PdfPageGrid.WheelBasis(0, 132), 120);
            int second = ThumbZoom.StepFromWheel(PdfPageGrid.WheelBasis(first, 132), 120);
            AssertEqual(148, first, "первый щелчок: +16");
            AssertEqual(164, second, "второй щелчок шагает от цели, а не от устаревшей ширины");
        }

        /// <summary>
        /// Median — общий знаменатель ВСЕХ порогов разбора PDF (14 мест вызова: кегли, зазоры,
        /// шаг сетки). Контракт: НИЖНЯЯ медиана при чётном числе значений, 0 на пустом входе,
        /// исходный список не переставляется. Без теста «уточнение» до среднего или до верхней
        /// медианы сдвинуло бы разом все пороги, а фикстуры остались бы зелёными.
        /// </summary>
        private static void TestMedian()
        {
            AssertEqual(0.0, MathUtil.Median(null), "null → 0");
            AssertEqual(0.0, MathUtil.Median(new List<double>()), "пустой список → 0");
            AssertEqual(5.0, MathUtil.Median(new List<double> { 5 }), "один элемент");
            AssertEqual(2.0, MathUtil.Median(new List<double> { 1, 2, 3, 4 }), "чётное число — НИЖНЯЯ медиана");
            AssertEqual(3.0, MathUtil.Median(new List<double> { 5, 1, 3 }), "порядок на входе не важен");
            var src = new List<double> { 9, 1, 5 };
            MathUtil.Median(src);
            AssertEqual("9,1,5", string.Join(",", src), "исходный список не переставлен");
        }

        /// <summary>
        /// Номер списка распознаём только по ASCII-цифрам. char.IsDigit пропускает и другие
        /// десятичные цифры Юникода, а вычитание '0' превращало «１.» в номер 65249 — Word такой
        /// начальный номер отвергает, исключение глушится, и пункт молча остаётся без номера.
        /// </summary>
        private static void TestListMarkerNonAsciiDigits()
        {
            AssertEqual(ListKind.None, ListMarker.Detect("１. текст").Kind, "полноширинная «１» — не номер");
            AssertEqual(ListKind.None, ListMarker.Detect("١. текст").Kind, "арабско-индийская «١» — не номер");
            ListMarker.Result ok = ListMarker.Detect("1. текст");
            AssertEqual(ListKind.Numbered, ok.Kind, "обычная ASCII-цифра по-прежнему номер");
            AssertEqual(1, ok.Number, "номер разобран");
        }

        /// <summary>
        /// Отсутствующий ключ Loc.T возвращает сам ключ — на этом держится проверка каталога:
        /// переименовали ключ, а в коде забыли — и в интерфейс попадёт «compress.level.none».
        /// </summary>
        private static void TestLocMissingKey()
        {
            AssertEqual("нет.такого.ключа", Loc.T("нет.такого.ключа"), "нет ключа — возвращается сам ключ");
            AssertEqual(null, Loc.T(null), "null-ключ отдаётся как есть, без исключения");
            AssertEqual("", Loc.T(""), "пустой ключ отдаётся как есть, без исключения");
        }

        /// <summary>
        /// Подписи уровней сжатия идут в порядке показа и уходят прямо в выпадающий список.
        /// Промах по ключу Loc.T не бросает, а тихо отдаёт сам ключ, поэтому
        /// «compress.level.none» уехал бы в интерфейс при всех зелёных тестах.
        /// </summary>
        private static void TestCompressionLevelLabels()
        {
            Lang saved = Loc.Current;
            try
            {
                foreach (Lang lang in new[] { Lang.Ru, Lang.En })
                {
                    Loc.Init(lang);
                    string[] labels = PdfCompression.LevelLabels();
                    AssertEqual(PdfCompression.LevelLabels().Length, labels.Length,
                        "по подписи на каждый уровень (" + lang + ")");
                    var seen = new HashSet<string>();
                    for (int i = 0; i < labels.Length; i++)
                    {
                        AssertTrue(!string.IsNullOrEmpty(labels[i]), "подпись уровня " + i + " не пуста (" + lang + ")");
                        AssertTrue(labels[i].IndexOf("compress.level", StringComparison.Ordinal) < 0,
                            "подпись уровня " + i + " — перевод, а не сам ключ (" + lang + ")");
                        AssertTrue(!labels[i].Contains("{0}"),
                            "подстановка выполнена, а не показана как есть (" + lang + "): " + labels[i]);
                        AssertTrue(seen.Add(labels[i]), "уровни различимы на глаз (" + lang + "): " + labels[i]);
                    }
                }
            }
            finally { Loc.Init(saved); }
        }

        /// <summary>
        /// ThrowIf — единственная точка кооперативной отмены во всех сервисах. Отдельно проверяем
        /// ветку null: она означает «отмена не предусмотрена» и не должна ничего бросать.
        /// </summary>
        private static void TestCancellationThrowIf()
        {
            Cancellation.ThrowIf(null); // «отмены нет» — молча продолжаем
            Cancellation.ThrowIf(delegate { return false; });
            bool thrown = false;
            try { Cancellation.ThrowIf(delegate { return true; }); }
            catch (OperationCanceledException) { thrown = true; }
            AssertTrue(thrown, "поднятый флаг бросает OperationCanceledException");
        }

        /// <summary>
        /// Вторая половина того же договора: сервис, пишущий файл на каждую единицу работы, при
        /// отмене не оставляет частичного результата. Здесь же закреплено и обратное решение —
        /// ОШИБКА созданное не удаляет: то, что успело записаться до сбоя, остаётся, как было
        /// до появления обёртки (иначе половина разбиения молча исчезала бы из-за одного
        /// неверного диапазона).
        /// </summary>
        private static void TestNoPartialOutput()
        {
            string dir = Path.Combine(Path.GetTempPath(), "iwo_np_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                Func<string, string> make = delegate(string name)
                {
                    string p = Path.Combine(dir, name);
                    File.WriteAllText(p, "x");
                    return p;
                };

                List<string> done = Cancellation.NoPartialOutput(delegate(List<string> created)
                {
                    created.Add(make("ok1.txt"));
                    created.Add(make("ok2.txt"));
                });
                AssertEqual(2, done.Count, "без отмены возвращаются все созданные пути");
                AssertTrue(File.Exists(done[0]) && File.Exists(done[1]), "без отмены файлы остаются");

                bool canceled = false;
                try
                {
                    Cancellation.NoPartialOutput(delegate(List<string> created)
                    {
                        created.Add(make("part1.txt"));
                        created.Add(make("part2.txt"));
                        throw new OperationCanceledException();
                    });
                }
                catch (OperationCanceledException) { canceled = true; }
                AssertTrue(canceled, "отмена уходит наверх, а не проглатывается");
                AssertEqual(0, Directory.GetFiles(dir, "part*.txt").Length, "созданное до отмены удалено");

                // Сбой убирает за собой ровно так же, как отмена: из папки с семью частями от
                // сорока не видно, что она неполная, — половина набора выглядит как набор.
                bool failed = false;
                try
                {
                    Cancellation.NoPartialOutput(delegate(List<string> created)
                    {
                        created.Add(make("err1.txt"));
                        throw new MergeException("сбой");
                    });
                }
                catch (MergeException) { failed = true; }
                AssertTrue(failed, "ошибка уходит наверх");
                AssertEqual(0, Directory.GetFiles(dir, "err*.txt").Length, "после сбоя частей не остаётся");

                // Нехватка памяти — единственное исключение: там уже не до уборки.
                bool oom = false;
                try
                {
                    Cancellation.NoPartialOutput(delegate(List<string> created)
                    {
                        created.Add(make("oom1.txt"));
                        throw new OutOfMemoryException();
                    });
                }
                catch (OutOfMemoryException) { oom = true; }
                AssertTrue(oom, "нехватка памяти уходит наверх нетронутой");
                AssertEqual(1, Directory.GetFiles(dir, "oom*.txt").Length, "при нехватке памяти уборку не затеваем");
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        /// <summary>
        /// Экран для восстановления окна: тот, с которым пересечение больше. Отдельно нужна
        /// ветка «пересечения нет вовсе» (монитор отключили) — иначе окно осталось бы за краем.
        /// </summary>
        private static void TestBestWorkArea()
        {
            var left = new System.Drawing.Rectangle(0, 0, 1000, 800);
            var right = new System.Drawing.Rectangle(1000, 0, 1000, 800);
            var areas = new[] { left, right };
            AssertEqual(right, WindowPlacement.BestWorkArea(new System.Drawing.Rectangle(1200, 100, 400, 300), areas),
                "окно целиком на втором экране");
            AssertEqual(left, WindowPlacement.BestWorkArea(new System.Drawing.Rectangle(850, 100, 200, 300), areas),
                "окно на стыке (150 px слева, 50 справа) — где пересечение больше");
            AssertEqual(right, WindowPlacement.BestWorkArea(new System.Drawing.Rectangle(950, 100, 200, 300), areas),
                "перевес в другую сторону (50 px слева, 150 справа)");
            AssertEqual(left, WindowPlacement.BestWorkArea(new System.Drawing.Rectangle(5000, 5000, 400, 300), areas),
                "монитора больше нет — первый экран, а не пустота");
        }

        /// <summary>
        /// Дерево XY-разреза (не плоский порядок чтения): писатель по нему выводит колонки
        /// таблицей, поэтому проверяем ИМЕННО вложенность и AvailRight. Изменение, сохраняющее
        /// плоский порядок, но схлопывающее уровень дерева, ломает шапки при зелёном Order.
        /// </summary>
        private static void TestXyCutOrderTreeShape()
        {
            // Два блока бок о бок сверху и один во всю ширину снизу. Высота верхних — 40 pt:
            // колонка ниже minColumnExtent считается ложной и в колонки не режется.
            var boxes = new[]
            {
                CB(0, 0, 200, 80, 40), CB(1, 300, 200, 80, 40),
                CB(2, 0, 100, 380, 20)
            };
            CutNode root = XyCut.OrderTree(boxes, 30, 30, 25, 1);
            AssertTrue(root != null && !root.IsLeaf, "корень — внутренний узел");
            AssertTrue(!root.SideBySide, "верхний уровень — этажи (стек сверху вниз)");
            AssertEqual(2, root.Children.Count, "два этажа: пара колонок и нижний блок");
            CutNode floor = root.Children[0];
            AssertTrue(!floor.IsLeaf && floor.SideBySide, "верхний этаж — колонки бок о бок");
            AssertEqual(2, floor.Children.Count, "в этаже две колонки");
            AssertTrue(floor.Children[0].AvailRight <= floor.Children[1].ColumnLeft,
                "доступное место левой колонки упирается в содержимое правой");
            AssertTrue(root.Children[1].IsLeaf, "нижний этаж — один блок");
        }

        /// <summary>
        /// Ui.OnUi — общий guard доставки в UI-поток. В его собственном комментарии сказано, что
        /// ручные копии guard'а уже дважды теряли catch, а теста на него не было.
        /// </summary>
        private static void TestOnUiGuard()
        {
            var results = new List<string>();
            InIsolatedSettings("iwo_onui_", delegate
            {
                AssertLocal(!Ui.OnUi(null, delegate { }), "null-контрол — доставки нет", results);
                using (var f = new System.Windows.Forms.Form())
                {
                    AssertLocal(!Ui.OnUi(f, delegate { }), "окно без хэндла — доставки нет", results);
                    f.Show();
                    AssertLocal(Ui.OnUi(f, delegate { }), "живое окно принимает делегат", results);
                    f.Close();
                }
                var dead = new System.Windows.Forms.Form();
                dead.Show();
                dead.Dispose();
                AssertLocal(!Ui.OnUi(dead, delegate { }), "разрушенное окно — доставки нет, без исключения", results);
            });
            AssertTrue(results.Count == 0, "Ui.OnUi: " + string.Join(" | ", results.ToArray()));
        }

        /// <summary>Проверка внутри чужого потока: копим отказы, а бросаем уже в основном.</summary>
        private static void AssertLocal(bool condition, string what, List<string> failures)
        {
            if (!condition)
                failures.Add(what);
        }

        /// <summary>
        /// Язык, выбранный флагом в установщике, приезжает отдельным ASCII-маркером и
        /// применяется ОДИН раз. До 1.17.9 установщик писал язык прямо в settings.txt и только
        /// при первой установке — поэтому переустановка с другим флагом до программы не
        /// доезжала. Здесь проверяются обе половины: маркер читается и маркер снимается.
        /// </summary>
        private static void TestSetupLanguageMarker()
        {
            string root = Path.Combine(Path.GetTempPath(), "iwo_setuplang_" + Guid.NewGuid().ToString("N"));
            AppPaths.SetRootForTests(root);
            try
            {
                Directory.CreateDirectory(root);
                AssertTrue(SetupLanguage.Read() == null, "без файла выбора установщика нет");

                File.WriteAllText(AppPaths.SetupLanguageFile, "en");
                AssertEqual("en", SetupLanguage.Read(), "маркер прочитан");

                // Установщик пишет строками, поэтому перевод строки в конце — норма.
                File.WriteAllText(AppPaths.SetupLanguageFile, "ru\r\n");
                AssertEqual("ru", SetupLanguage.Read(), "перевод строки не мешает");

                SetupLanguage.Consume();
                AssertTrue(SetupLanguage.Read() == null, "снятый маркер больше не применяется");

                // Чужое содержимое (правил руками, попал мусор) языком не считается.
                File.WriteAllText(AppPaths.SetupLanguageFile, "deutsch");
                AssertTrue(SetupLanguage.Read() == null, "неизвестный код игнорируется");
            }
            finally
            {
                AppPaths.SetRootForTests(null);
                try { Directory.Delete(root, true); } catch { }
            }
        }

        /// <summary>
        /// Выбранный язык обязан пережить запись настроек чужим окном: WriteAll берёт язык из
        /// ЖИВОГО Loc, а не из своего поля, ровно потому, что устаревший экземпляр однажды уже
        /// стирал выбор пользователя. Ветки zoom/compression и границ окон покрыты отдельно.
        /// </summary>
        private static void TestSettingsLanguageNotClobbered()
        {
            string root = Path.Combine(Path.GetTempPath(), "iwo_lang_" + Guid.NewGuid().ToString("N"));
            Lang saved = Loc.Current;
            AppPaths.SetRootForTests(root);
            try
            {
                UserSettings stale = UserSettings.Load(); // «старое» окно держит копию с прежним языком
                Loc.Set(Lang.En);
                AssertEqual("en", UserSettings.Load().Language, "Loc.Set записал выбор на диск");

                stale.Save(); // устаревший экземпляр сохраняет свои поля
                AssertEqual("en", UserSettings.Load().Language, "чужая запись не вернула прежний язык");
            }
            finally
            {
                Loc.Init(saved);
                AppPaths.SetRootForTests(null);
                try { Directory.Delete(root, true); } catch { }
            }
        }

        /// <summary>
        /// Каждое настоящее окно, сжатое до своего МИНИМАЛЬНОГО размера, обязано остаться
        /// целым: ни один контрол не свисает за клиентскую область и никакие две кнопки не
        /// накладываются. Глазами это ловится плохо (нужно дотащить рамку до упора и ещё
        /// угадать, при каком состоянии появится скрытая кнопка), а пользователь упирается
        /// сразу — и WindowPlacement потом ЗАПОМНИТ сломанный размер. Невидимые кнопки
        /// проверяются наравне с видимыми: кнопка, которая появится позже (например
        /// «Повторить пропущенные»), обязана иметь своё место уже сейчас.
        /// </summary>
        private static void TestWindowsSurviveMinimumSize()
        {
            var offenders = new List<string>();
            InIsolatedSettings("iwo_minsize_", delegate
            {
                foreach (System.Windows.Forms.Form f in NewAllToolWindows())
                    using (f)
                    {
                        string where = f.GetType().Name;
                        try
                        {
                            f.Show();
                            f.Size = f.MinimumSize; // пользователь дотащил рамку до упора
                            f.PerformLayout();
                            // Размеры — прямо в описание отказа: рамка окна на другой машине
                            // другая, и без этих чисел причина падения не восстанавливается.
                            where += " (Min=" + f.MinimumSize.Height + ", рамка=" +
                                (f.Height - f.ClientSize.Height) + ")";
                            CheckFits(f, where, offenders);
                            CheckControlsDoNotOverlap(f, where, offenders);
                            // Кнопки — ещё и невидимые: скрытая сейчас («Повторить пропущенные»)
                            // появится позже и обязана иметь своё место уже теперь.
                            CheckButtonsDoNotOverlap(f, where, offenders);
                            f.Close();
                        }
                        catch (Exception ex) { offenders.Add(where + " — раскладка: " + ex.Message); }
                    }
            });
            AssertTrue(offenders.Count == 0, "при минимальном размере окна: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// Ни одна подпись кнопки не должна обрезаться на СОБСТВЕННОМ размере окна — том, каким
        /// пользователь видит его при первом открытии. Проверяются и окна-инструменты, и
        /// диалоги: обрезка выглядит как опечатка в программе, а найти её можно только глазами
        /// на каждой кнопке каждого окна.
        /// </summary>
        private static void TestButtonCaptionsFit()
        {
            var offenders = new List<string>();
            InIsolatedSettings("iwo_captions_", delegate
            {
                var windows = new List<System.Windows.Forms.Form>(NewAllToolWindows());
                windows.AddRange(NewAllDialogs(offenders));
                foreach (System.Windows.Forms.Form f in windows)
                    using (f)
                    {
                        string where = f.GetType().Name;
                        try
                        {
                            f.Show();
                            f.PerformLayout();
                            CheckButtonCaptionsFit(f, where, offenders);
                            f.Close();
                        }
                        catch (Exception ex) { offenders.Add(where + " — подписи: " + ex.Message); }
                    }
            });
            AssertTrue(offenders.Count == 0, "подписи кнопок: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// Модальные диалоги обязаны быть целыми на СВОЁМ размере: ничего не свисает из окна и
        /// никакие два элемента не наложены. Сжимать их до MinimumSize, как окна-инструменты,
        /// нельзя — он у них не задан и равен нулю: «О программе» схлопывается в 136×39, и все
        /// девятнадцать контролов оказываются снаружи, то есть тест ловил бы не дефект, а сам
        /// себя. Проверка нужна: наложение кнопок на поле в «Свойствах документа» прожило до
        /// 1.17.9 именно потому, что диалоги в раскладочные тесты не входили.
        /// </summary>
        private static void TestDialogsLayoutIsSound()
        {
            var offenders = new List<string>();
            InIsolatedSettings("iwo_dialogs_", delegate
            {
                foreach (System.Windows.Forms.Form f in NewAllDialogs(offenders))
                    using (f)
                    {
                        string where = f.GetType().Name;
                        try
                        {
                            f.Show();
                            f.PerformLayout();
                            where += " (" + f.ClientSize.Width + "×" + f.ClientSize.Height + ")";
                            CheckFits(f, where, offenders);
                            CheckControlsDoNotOverlap(f, where, offenders);
                            CheckButtonsDoNotOverlap(f, where, offenders);
                            f.Close();
                        }
                        catch (Exception ex) { offenders.Add(where + " — раскладка: " + ex.Message); }
                    }
            });
            AssertTrue(offenders.Count == 0, "раскладка диалогов: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// Двойной клик по полю процентов возвращает масштаб к 100%. Раньше этот жест штатно
        /// ВЫДЕЛЯЛ число, и выделение принимали за «скопировалось в буфер»: жест выглядел
        /// сработавшим, не делая ничего. Проверяем на настоящем окне и через внутреннее поле
        /// ввода — события мыши получает именно оно, а не сам NumericUpDown.
        /// </summary>
        private static void TestZoomDoubleClickResetsLive()
        {
            var offenders = new List<string>();
            RunSta(delegate
            {
                using (var form = new OcrForm())
                {
                    IntPtr handle = form.Handle;
                    if (handle == IntPtr.Zero) { offenders.Add("окно не создалось"); return; }
                    System.Windows.Forms.NumericUpDown input = null;
                    foreach (System.Windows.Forms.Control c in form.Controls)
                    {
                        input = c as System.Windows.Forms.NumericUpDown;
                        if (input != null) break;
                    }
                    if (input == null) { offenders.Add("поля процентов нет"); return; }

                    // Значение заведомо допустимое и ЗАВЕДОМО не 100%: границы масштаба
                    // заданы ThumbZoom, и литерал вроде 50 может оказаться вне диапазона.
                    input.Value = input.Maximum;
                    System.Windows.Forms.Control inner = null;
                    foreach (System.Windows.Forms.Control c in input.Controls)
                        if (c is System.Windows.Forms.TextBox) { inner = c; break; }
                    if (inner == null) { offenders.Add("внутреннего поля ввода нет"); return; }

                    // Двойной клик приходит внутреннему полю — поднимаем событие именно у него.
                    System.Reflection.MethodInfo raise = typeof(System.Windows.Forms.Control).GetMethod(
                        "OnDoubleClick", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                    if (raise == null) { offenders.Add("не добраться до OnDoubleClick"); return; }
                    raise.Invoke(inner, new object[] { EventArgs.Empty });

                    if (input.Value != 100)
                        offenders.Add("после двойного клика масштаб " + input.Value + "%, а должен быть 100%");

                    // А вот двойной щелчок по СТРЕЛКАМ сбрасывать не должен: это два шага
                    // масштаба подряд, а не просьба вернуть 100%.
                    System.Windows.Forms.Control buttons = null;
                    foreach (System.Windows.Forms.Control c in input.Controls)
                        if (!(c is System.Windows.Forms.TextBox)) { buttons = c; break; }
                    if (buttons != null)
                    {
                        input.Value = input.Maximum;
                        raise.Invoke(buttons, new object[] { EventArgs.Empty });
                        if (input.Value != input.Maximum)
                            offenders.Add("двойной щелчок по стрелкам сбросил масштаб, а не должен");
                    }
                }
            });
            AssertTrue(offenders.Count == 0, "сброс масштаба: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// Размер шрифта, выбранный в системе («размер текста» в специальных возможностях),
        /// не должен доставать до нашей раскладки: все окна задают шрифт явно, а полоса меню
        /// имеет фиксированную высоту, и системным шрифтом её подписи обрезало бы снизу.
        /// Проверяем оба конца: окна и все их контролы живут разрешёнными шрифтами, а меню —
        /// шрифтом приложения, и его подписи помещаются в высоту полосы.
        /// </summary>
        private static void TestSystemFontIgnoredLive()
        {
            var offenders = new List<string>();
            // Символьные шрифты законны: ими рисуются глифы значков и стрелок.
            var allowed = new List<string> { "Segoe UI", "Segoe MDL2 Assets", "Segoe UI Symbol", "Segoe UI Emoji" };
            RunSta(delegate
            {
                var windows = new List<System.Windows.Forms.Form>
                {
                    new StartForm(), new SettingsForm(), new AboutForm(), new StatsForm()
                };
                try
                {
                    foreach (System.Windows.Forms.Form form in windows)
                    {
                        IntPtr handle = form.Handle;
                        if (handle == IntPtr.Zero) { offenders.Add(form.GetType().Name + ": окно не создалось"); continue; }
                        if (Math.Abs(form.Font.SizeInPoints - 9.75f) > 0.01f)
                            offenders.Add(form.GetType().Name + ": шрифт окна " + form.Font.SizeInPoints + " pt");
                        WalkFonts(form, form.GetType().Name, allowed, offenders);
                    }
                }
                finally
                {
                    foreach (System.Windows.Forms.Form form in windows)
                        try { form.Dispose(); } catch { }
                }

                // Полоса меню: свой шрифт и подписи, помещающиеся в её фиксированную высоту.
                using (var owner = new System.Windows.Forms.Form())
                {
                    System.Windows.Forms.MenuStrip menu = HelpMenu.Create(owner, delegate { });
                    if (Math.Abs(menu.Font.SizeInPoints - 9.75f) > 0.01f)
                        offenders.Add("полоса меню: шрифт " + menu.Font.SizeInPoints + " pt (системный?)");
                    foreach (System.Windows.Forms.ToolStripItem item in menu.Items)
                    {
                        int need = System.Windows.Forms.TextRenderer.MeasureText(item.Text, menu.Font).Height;
                        if (need > HelpMenu.Height - 6)
                            offenders.Add("пункт «" + item.Text + "» не помещается в полосу: нужно " + need);
                    }
                    menu.Dispose();
                }
            });
            AssertTrue(offenders.Count == 0, "системный шрифт: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>Обойти дерево контролов и собрать те, чей шрифт не из разрешённых семейств.</summary>
        private static void WalkFonts(System.Windows.Forms.Control root, string where,
            List<string> allowed, List<string> offenders)
        {
            foreach (System.Windows.Forms.Control c in root.Controls)
            {
                if (c.Font != null && !allowed.Contains(c.Font.FontFamily.Name))
                    offenders.Add(where + ": " + c.GetType().Name + " шрифтом «" + c.Font.FontFamily.Name + "»");
                WalkFonts(c, where, allowed, offenders);
            }
        }

        /// <summary>Выполнить действие на STA-потоке: окна WinForms другого не терпят.</summary>
        private static void RunSta(Action action)
        {
            Exception error = null;
            var th = new System.Threading.Thread(delegate()
            {
                try { action(); }
                catch (Exception ex) { error = ex; }
            });
            th.SetApartmentState(System.Threading.ApartmentState.STA);
            th.IsBackground = true;
            th.Start();
            th.Join();
            if (error != null)
                throw new Exception(error.GetType().Name + ": " + error.Message);
        }

        /// <summary>
        /// Три действия истории стоят В РЯД и подписи в них ПОМЕЩАЮТСЯ. Обрезанная подпись
        /// здесь опаснее обычного: «Очистить историю» — необратимое действие, и от соседей
        /// оно должно отличаться словами, а не догадкой.
        /// </summary>
        private static void TestSettingsHistoryRowLive()
        {
            var offenders = new List<string>();
            RunSta(delegate
            {
                using (var f = new SettingsForm())
                {
                    IntPtr handle = f.Handle; // создать окно и разложить контролы
                    if (handle == IntPtr.Zero) { offenders.Add("окно не создалось"); return; }
                    var row = new List<RoundedButton>();
                    foreach (System.Windows.Forms.Control c in f.Controls)
                    {
                        var b = c as RoundedButton;
                        if (b == null) continue;
                        if (b.Text == Loc.T("settings.btn.historyShow")
                            || b.Text == Loc.T("settings.btn.historyClear")
                            || b.Text == Loc.T("settings.btn.stats"))
                            row.Add(b);
                    }
                    if (row.Count != 3) { offenders.Add("кнопок истории: " + row.Count); return; }
                    foreach (RoundedButton b in row)
                    {
                        if (b.Top != row[0].Top)
                            offenders.Add("«" + b.Text + "» не в одном ряду с остальными");
                        int room = b.Width - 2 * RoundedButton.TextPadFor(b.Width);
                        int need = System.Windows.Forms.TextRenderer.MeasureText(b.Text, b.Font).Width;
                        if (need > room)
                            offenders.Add("«" + b.Text + "» не помещается: нужно " + need + ", есть " + room);
                        if (b.Right > f.ClientSize.Width)
                            offenders.Add("«" + b.Text + "» вылезает за край окна");
                    }
                }
            });
            AssertTrue(offenders.Count == 0, "ряд действий истории: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// Значки нижнего ряда хаба (шестерня и справка) остались ДОСТУПНЫ: у кнопки без
        /// подписи это единственный способ узнать, что она делает — и для экранного диктора,
        /// и для человека, который наводит мышь.
        /// </summary>
        private static void TestHubIconButtonsLive()
        {
            var offenders = new List<string>();
            RunSta(delegate
            {
                using (var hub = new StartForm())
                {
                    IntPtr handle = hub.Handle;
                    if (handle == IntPtr.Zero) { offenders.Add("хаб не создался"); return; }
                    var named = new List<string>();
                    var icons = new List<System.Windows.Forms.Control>();
                    foreach (System.Windows.Forms.Control c in hub.Controls)
                        if (c.AccessibleName != null && c.AccessibleName.Length > 0 && c.TabStop)
                        {
                            named.Add(c.AccessibleName);
                            if (c.AccessibleName == Loc.T("settings.title") || c.AccessibleName == Loc.T("hub.about"))
                                icons.Add(c);
                        }
                    if (!named.Contains(Loc.T("settings.title")))
                        offenders.Add("у шестерни нет имени для диктора");
                    if (!named.Contains(Loc.T("hub.about")))
                        offenders.Add("у справки нет имени для диктора");
                    // Цвет и отклик на курсор. Два одинаково серых значка по краям белого ряда
                    // сливались и с фоном, и между собой, а значок без рамки ничем не отвечал
                    // на наведение — «нажимается ли это» выяснялось нажатием.
                    if (icons.Count != 2)
                        offenders.Add("значков нижнего ряда не два, а " + icons.Count);
                    else
                    {
                        if (icons[0].ForeColor.ToArgb() == icons[1].ForeColor.ToArgb())
                            offenders.Add("шестерня и справка одного цвета — их не различить");
                        foreach (System.Windows.Forms.Control icon in icons)
                        {
                            if (icon.ForeColor.ToArgb() == Theme.TextPrimary.ToArgb())
                                offenders.Add("значок цвета обычного текста — на белом он теряется");
                            System.Reflection.FieldInfo fill = icon.GetType().GetField("HoverFill");
                            if (fill == null) { offenders.Add("у значка нет подсветки при наведении"); continue; }
                            var color = (System.Drawing.Color)fill.GetValue(icon);
                            if (color.IsEmpty || color.A == 0)
                                offenders.Add("подсветка значка прозрачна — наведение ничего не показывает");
                        }
                    }
                }
            });
            AssertTrue(offenders.Count == 0, "значки нижнего ряда: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// Окно «Настройки» на НАСТОЯЩИХ контролах: флажок показывает то, что на диске, его
        /// переключение сразу пишется, а кнопка «снова напоминать» есть ровно тогда, когда
        /// отменять есть что. Последнее — не украшение: без этой кнопки галочка «больше не
        /// напоминать» была бы необратимой, а необратимое действие без выхода — это дефект.
        /// </summary>
        private static void TestSettingsUpdateControlsLive()
        {
            var offenders = new List<string>();
            InIsolatedSettings("iwo_setwin_", delegate
            {
                try
                {
                    // 1. Чистые настройки: проверка включена, отменять нечего.
                    using (var f = new SettingsForm())
                    {
                        f.Show();
                        AccentCheckBox box = FindCheck(f);
                        RoundedButton unskip = FindUnskip(f);
                        if (box == null || unskip == null)
                            offenders.Add("флажок или кнопка не собрались");
                        else
                        {
                            if (!box.Checked) offenders.Add("на чистых настройках проверка должна быть включена");
                            if (unskip.Visible) offenders.Add("отменять нечего, а кнопка «снова напоминать» видна");

                            // 2. Снятие флажка пишется на диск сразу, без кнопки «Сохранить».
                            box.Checked = false;
                            if (UserSettings.Load().UpdateCheckOnStart)
                                offenders.Add("снятый флажок не записался");
                            box.Checked = true;
                            if (!UserSettings.Load().UpdateCheckOnStart)
                                offenders.Add("возвращённый флажок не записался");
                        }
                        f.Close();
                    }

                    // 3. Настройки, изменённые СНАРУЖИ, подхватываются при возврате фокуса:
                    // окно модально только для своего владельца, и вторые «Настройки» из
                    // меню инструмента (или окно обновления поверх) меняют файл под ним.
                    using (var f = new SettingsForm())
                    {
                        f.Show();
                        AccentCheckBox box = FindCheck(f);
                        UserSettings.SaveUpdateCheckOnStart(false); // «кто-то другой» выключил
                        // Нужен НАСТОЯЩИЙ цикл «потерял фокус — вернул»: Activate() на уже
                        // активном окне ничего не делает, и событие не приходит.
                        using (var other = new System.Windows.Forms.Form())
                        {
                            other.ShowInTaskbar = false;
                            other.Show();
                            other.Activate();
                            System.Windows.Forms.Application.DoEvents();
                            f.Activate();
                            System.Windows.Forms.Application.DoEvents();
                            other.Close();
                        }
                        if (box != null && box.Checked)
                            offenders.Add("внешнее изменение не подхвачено при возврате фокуса");
                        f.Close();
                    }
                    UserSettings.SaveUpdateCheckOnStart(true);

                    // 4. Есть пропущенная версия — кнопка появилась и называет её.
                    UserSettings.SaveSkippedVersion("1.18.0");
                    using (var f = new SettingsForm())
                    {
                        f.Show();
                        RoundedButton unskip = FindUnskip(f);
                        if (unskip == null || !unskip.Visible)
                            offenders.Add("пропущенная версия есть, а кнопка «снова напоминать» не видна");
                        else
                        {
                            if (unskip.Text.IndexOf("1.18.0", StringComparison.Ordinal) < 0)
                                offenders.Add("кнопка не называет версию: " + unskip.Text);
                            // Кнопка автоширины: длинный перевод мог бы вынести её за окно.
                            // Проверяем на ОБОИХ языках — английская подпись длиннее русской.
                            foreach (Lang lang in new[] { Lang.Ru, Lang.En })
                            {
                                Lang was = Loc.Current;
                                Loc.Init(lang);
                                using (var g = new SettingsForm())
                                {
                                    g.Show();
                                    RoundedButton u = FindUnskip(g);
                                    if (u != null && u.Right > g.ClientSize.Width - 24)
                                        offenders.Add(lang + ": кнопка «" + u.Text + "» выходит за окно (" +
                                                      u.Right + " > " + (g.ClientSize.Width - 24) + ")");
                                    g.Close();
                                }
                                Loc.Init(was);
                            }
                            unskip.PerformClick();
                            if (!string.IsNullOrEmpty(UserSettings.Load().SkippedVersion))
                                offenders.Add("нажатие не стёрло пропущенную версию");
                            if (unskip.Visible)
                                offenders.Add("версия стёрта, а кнопка осталась видна");
                        }
                        f.Close();
                    }
                }
                catch (Exception ex) { offenders.Add("не собралось: " + Root(ex).Message); }
            });
            AssertTrue(offenders.Count == 0, "настройки обновлений: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// История в «Настройках» на НАСТОЯЩИХ контролах: запись появляется, счётчик её
        /// показывает, очистка опустошает, а выключение стирает накопленное. Последнее — не
        /// украшение: оставить перечень путей тому, кто только что отказался от их хранения,
        /// значит сделать не то, о чём просили.
        /// </summary>
        private static void TestSettingsHistoryLive()
        {
            var offenders = new List<string>();
            InIsolatedSettings("iwo_hist_", delegate
            {
                try
                {
                    OperationHistory.Record("hist.op.merge", @"C:\где-то\итог.pdf");
                    OperationHistory.Record("hist.op.excel", @"C:\где-то\свод.xlsx");
                    if (OperationHistory.Load().Entries.Count != 2)
                        offenders.Add("операции не записались");

                    using (var f = new SettingsForm())
                    {
                        f.Show();
                        AccentCheckBox keep = null;
                        foreach (System.Windows.Forms.Control c in f.Controls)
                            if (c is AccentCheckBox && c.Text == Loc.T("settings.chk.history"))
                                keep = (AccentCheckBox)c;
                        if (keep == null)
                        {
                            offenders.Add("флажок истории не собрался");
                        }
                        else
                        {
                            if (!keep.Checked) offenders.Add("история должна быть включена по умолчанию");
                            // Выключение стирает накопленное (и спрашивает — диалог тут не
                            // показать, поэтому проверяем через само хранилище).
                            OperationHistory.SetEnabled(false);
                            OperationHistory.Data off = OperationHistory.Load();
                            if (off.Entries.Count != 0)
                                offenders.Add("выключение не стёрло накопленное");
                            // Выключенная история молчит и файл не растит.
                            OperationHistory.Record("hist.op.merge", @"C:\ещё\файл.pdf");
                            if (OperationHistory.Load().Entries.Count != 0)
                                offenders.Add("выключенная история всё равно записала операцию");
                        }
                        f.Close();
                    }

                    // Подтверждённое ВЫКЛЮЧЕНИЕ обязано сработать. Здесь жила ошибка: пока
                    // модальный вопрос на экране, окно теряет и возвращает фокус, а на
                    // возврате перечитывает настройки и ставит флажок обратно по диску —
                    // и обработчик сохранял прочитанное ПОСЛЕ диалога, то есть прежнее
                    // значение. Человек соглашался выключить историю, а она оставалась.
                    OperationHistory.SetEnabled(true);
                    OperationHistory.Record("hist.op.merge", @"C:\есть\что терять.pdf");
                    using (var f = new SettingsForm())
                    {
                        f.Show();
                        AccentCheckBox keep = null;
                        foreach (System.Windows.Forms.Control c in f.Controls)
                            if (c is AccentCheckBox && c.Text == Loc.T("settings.chk.history"))
                                keep = (AccentCheckBox)c;
                        if (keep == null)
                            offenders.Add("флажок истории не найден");
                        else
                        {
                            // Заглушка обязана ВОСПРОИЗВЕСТИ то, что делает настоящий диалог:
                            // пока он на экране, окно теряет и возвращает фокус, а на возврате
                            // перечитывает настройки и ставит флажок обратно по диску. Простое
                            // «вернуть true» убирало бы само условие ошибки, и проверка зеленела
                            // бы на сломанном коде — что и случилось с первой её версией.
                            f.ConfirmClearHistory = delegate { keep.Checked = true; return true; };
                            keep.Checked = false; // снимаем — идёт ветка с подтверждением
                            if (OperationHistory.Load().Enabled)
                                offenders.Add("подтверждённое выключение истории не сохранилось");
                            if (OperationHistory.Load().Entries.Count != 0)
                                offenders.Add("выключение не стёрло накопленное");
                        }
                        f.Close();
                    }

                    // Отказ от вопроса ничего не меняет.
                    OperationHistory.SetEnabled(true);
                    OperationHistory.Record("hist.op.merge", @"C:\остаться\должно.pdf");
                    using (var f = new SettingsForm())
                    {
                        f.ConfirmClearHistory = delegate { return false; }; // «нет, передумал»
                        f.Show();
                        foreach (System.Windows.Forms.Control c in f.Controls)
                            if (c is AccentCheckBox && c.Text == Loc.T("settings.chk.history"))
                                ((AccentCheckBox)c).Checked = false;
                        if (!OperationHistory.Load().Enabled)
                            offenders.Add("отказ от вопроса всё равно выключил историю");
                        if (OperationHistory.Load().Entries.Count != 1)
                            offenders.Add("отказ от вопроса стёр записи");
                        f.Close();
                    }

                    OperationHistory.Clear();
                    OperationHistory.SetEnabled(true);
                    OperationHistory.Record("hist.op.split", @"C:\папка частей");
                    if (OperationHistory.Load().Entries.Count != 1)
                        offenders.Add("после включения запись не пошла");
                    OperationHistory.Clear();
                    if (OperationHistory.Load().Entries.Count != 0)
                        offenders.Add("очистка не опустошила историю");
                }
                catch (Exception ex) { offenders.Add("не собралось: " + Root(ex).Message); }
            });
            AssertTrue(offenders.Count == 0, "история в настройках: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// Окно истории на настоящих контролах: записи показаны СВЕЖИМИ СВЕРХУ, в колонке
        /// «что сделано» стоит подпись на текущем языке, а не хранимый ключ, и кнопки открытия
        /// погашены, пока ничего не выбрано. Ключ вместо подписи — самая вероятная ошибка
        /// здесь: в файле лежит именно он, и показать его целиком ничего не стоит.
        /// </summary>
        private static void TestHistoryWindowLive()
        {
            var offenders = new List<string>();
            InIsolatedSettings("iwo_histwin_", delegate
            {
                try
                {
                    OperationHistory.Clear();
                    OperationHistory.SetEnabled(true);
                    OperationHistory.Record("hist.op.merge", @"C:\первая\раньше.pdf");
                    OperationHistory.Record("hist.op.excel", @"C:\вторая\позже.xlsx");

                    using (var f = new HistoryForm())
                    {
                        f.Show();
                        System.Windows.Forms.ListView list = null;
                        foreach (System.Windows.Forms.Control c in f.Controls)
                            if (c is System.Windows.Forms.ListView)
                                list = (System.Windows.Forms.ListView)c;
                        if (list == null)
                        {
                            offenders.Add("список не собрался");
                        }
                        else
                        {
                            if (list.Items.Count != 2)
                                offenders.Add("строк в списке: " + list.Items.Count + ", ожидалось 2");
                            else
                            {
                                if (list.Items[0].SubItems[2].Text.IndexOf("позже", StringComparison.Ordinal) < 0)
                                    offenders.Add("сверху не самая свежая запись");
                                string what = list.Items[0].SubItems[1].Text;
                                if (what == "hist.op.excel")
                                    offenders.Add("в колонке показан КЛЮЧ, а не подпись");
                                if (what != Loc.T("hist.op.excel"))
                                    offenders.Add("подпись операции не из каталога: " + what);
                            }
                            // Ничего не выбрано — открывать нечего.
                            list.SelectedItems.Clear();
                            foreach (System.Windows.Forms.Control c in f.Controls)
                            {
                                var b = c as RoundedButton;
                                if (b != null && b.Text == Loc.T("excel.link.openFile") && b.Enabled)
                                    offenders.Add("кнопка «Открыть файл» доступна без выбора");
                            }
                        }
                        f.Close();
                    }
                }
                catch (Exception ex) { offenders.Add("не собралось: " + Root(ex).Message); }
            });
            AssertTrue(offenders.Count == 0, "окно истории: " + string.Join(" | ", offenders.ToArray()));
        }

        private static AccentCheckBox FindCheck(System.Windows.Forms.Form f)
        {
            foreach (System.Windows.Forms.Control c in f.Controls)
                if (c is AccentCheckBox)
                    return (AccentCheckBox)c;
            return null;
        }

        /// <summary>
        /// Кнопка «снова напоминать» — по ИМЕНИ, а не по подписи: пока отменять нечего,
        /// подписи у неё нет, и поиск «первой кнопки без текста» находил бы любую следующую,
        /// которую когда-нибудь добавят рядом, то есть проверял бы не то.
        /// </summary>
        private static RoundedButton FindUnskip(System.Windows.Forms.Form f)
        {
            foreach (System.Windows.Forms.Control c in f.Controls)
                if (c is RoundedButton && c.Name == SettingsForm.UnskipName)
                    return (RoundedButton)c;
            return null;
        }

        /// <summary>
        /// Подпись флажка «больше не напоминать» обязана помещаться в диалог ЦЕЛИКОМ на обоих
        /// языках. Ширина флажка ограничена шириной текстовой колонки, поэтому подпись длиннее
        /// неё не вылезет за окно — она молча обрежется, а это ровно тот дефект, который глазами
        /// не замечают: текст выглядит просто коротким. Меряем настоящим контролом в настоящем
        /// диалоге, а не прикидкой по числу букв.
        /// </summary>
        private static void TestUpdateSkipCaptionFits()
        {
            var offenders = new List<string>();
            Lang before = Loc.Current;
            InIsolatedSettings("iwo_skipcap_", delegate
            {
                try
                {
                    foreach (Lang lang in new[] { Lang.Ru, Lang.En })
                    {
                        Loc.Init(lang);
                        using (var f = (System.Windows.Forms.Form)Construct(typeof(MessageForm)))
                        {
                            f.Show();
                            f.PerformLayout();
                            AccentCheckBox box = null;
                            foreach (System.Windows.Forms.Control c in f.Controls)
                                if (c is AccentCheckBox)
                                    box = (AccentCheckBox)c;
                            if (box == null)
                            {
                                offenders.Add(lang + ": флажок не собрался");
                                continue;
                            }
                            // Доступная ширина — от левого края флажка до правого поля диалога.
                            // Поле берём из отступа значка слева: раскладка симметрична, и так
                            // проверка не разъедется, если поле однажды поменяют.
                            int pad = box.Left;
                            foreach (System.Windows.Forms.Control c in f.Controls)
                                if (c is System.Windows.Forms.PictureBox)
                                    pad = c.Left;
                            int room = f.ClientSize.Width - box.Left - pad;
                            box.Text = Loc.T("update.skip"); // настоящая подпись вместо нейтральной
                            int want = box.GetPreferredSize(System.Drawing.Size.Empty).Width;
                            if (want > room)
                                offenders.Add(lang + ": подпись «" + box.Text + "» просит " + want +
                                              " px, а колонка " + room);
                            // Фокус обязан стоять НА КНОПКЕ: флажок добавлен раньше кнопок и
                            // иначе забирает его первым, а пробел, которым диалог привычно
                            // закрывают, вместо кнопки включал бы «больше не напоминать».
                            // Утверждение положительное: проверка «не на флажке» зеленела бы
                            // и тогда, когда фокус не достался никому, то есть не доказывала
                            // бы ничего.
                            if (!(f.ActiveControl is RoundedButton))
                                offenders.Add(lang + ": фокус не на кнопке, а на " +
                                              (f.ActiveControl == null ? "ничём" : f.ActiveControl.GetType().Name));
                            f.Close();
                        }
                    }
                }
                catch (Exception ex) { offenders.Add("не собралось: " + Root(ex).Message); }
            });
            Loc.Init(before);
            AssertTrue(offenders.Count == 0, "флажок обновления: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// Полноэкранный просмотр обязан и сворачиваться, и иметь СВОЮ кнопку на панели задач.
        /// Это пара, порознь нельзя: окно модальное, и пока оно открыто, владелец отключён —
        /// свёрнутое окно без кнопки на панели задач вернуть было бы нечем, а приложение
        /// выглядело бы зависшим. Проверка дешёвая, а потеря любой из половин молчаливая.
        /// </summary>
        private static void TestPreviewCanMinimize()
        {
            bool minimize = false, taskbar = false, maximize = false;
            string failure = null;
            InIsolatedSettings("iwo_preview_", delegate
            {
                try
                {
                    using (var f = (System.Windows.Forms.Form)Construct(typeof(PagePreviewForm)))
                    {
                        minimize = f.MinimizeBox;
                        maximize = f.MaximizeBox;
                        taskbar = f.ShowInTaskbar;
                    }
                }
                catch (Exception ex) { failure = Root(ex).Message; }
            });
            AssertTrue(failure == null, "окно просмотра не собралось: " + failure);
            AssertTrue(minimize, "у просмотра есть кнопка свёртки");
            AssertTrue(maximize, "у просмотра есть кнопка разворачивания");
            AssertTrue(taskbar, "свёрнутый просмотр видно на панели задач (иначе его не вернуть)");
        }

        /// <summary>
        /// Все модальные диалоги приложения. Конструкторы у них приватные (окна показываются
        /// статическими методами), поэтому берём самый короткий конструктор и подставляем
        /// нейтральные значения ПО ТИПАМ параметров — так проверка переживает смену сигнатуры
        /// вместо того, чтобы разваливаться на ней.
        /// </summary>
        private static List<System.Windows.Forms.Form> NewAllDialogs(List<string> offenders)
        {
            var types = new[]
            {
                typeof(AboutForm), typeof(StatsForm), typeof(MetadataForm),
                typeof(MessageForm), typeof(NumberPromptDialog), typeof(SettingsForm), typeof(HistoryForm)
            };
            var forms = new List<System.Windows.Forms.Form>();
            foreach (Type t in types)
            {
                try { forms.Add((System.Windows.Forms.Form)Construct(t)); }
                catch (Exception ex) { offenders.Add(t.Name + " — не собрался: " + Root(ex).Message); }
            }
            return forms;
        }

        /// <summary>Создать объект самым коротким конструктором, подставив значения по типам.</summary>
        private static object Construct(Type t)
        {
            System.Reflection.ConstructorInfo best = null;
            foreach (System.Reflection.ConstructorInfo c in t.GetConstructors(
                         System.Reflection.BindingFlags.Instance |
                         System.Reflection.BindingFlags.Public |
                         System.Reflection.BindingFlags.NonPublic))
                if (best == null || c.GetParameters().Length < best.GetParameters().Length)
                    best = c;
            if (best == null)
                throw new Exception("нет конструктора");
            System.Reflection.ParameterInfo[] ps = best.GetParameters();
            var args = new object[ps.Length];
            for (int i = 0; i < ps.Length; i++)
                args[i] = Neutral(ps[i].ParameterType);
            return best.Invoke(args);
        }

        /// <summary>Нейтральное значение параметра: текст, единица, первый элемент перечисления, пустой объект.</summary>
        private static object Neutral(Type t)
        {
            if (t == typeof(string)) return "Проверка раскладки";
            if (t.IsEnum) return Enum.GetValues(t).GetValue(0);
            if (t == typeof(bool)) return false;
            if (t == typeof(int)) return 1;
            if (t.IsValueType) return Activator.CreateInstance(t);
            // Ссылочный параметр (например PdfMetadata): null уронил бы конструктор, поэтому
            // создаём пустой экземпляр, если это возможно.
            try { return Activator.CreateInstance(t); }
            catch { return null; }
        }

        private static Exception Root(Exception ex)
        {
            return ex.InnerException != null ? Root(ex.InnerException) : ex;
        }

        /// <summary>Рекурсивно: каждый контрол лежит внутри клиентской области своего родителя.</summary>
        private static void CheckFits(System.Windows.Forms.Control parent, string where, List<string> offenders)
        {
            System.Drawing.Rectangle box = parent.ClientRectangle;
            foreach (System.Windows.Forms.Control c in parent.Controls)
            {
                if (!box.Contains(c.Bounds))
                    offenders.Add(where + " → " + Describe(c) + " " + c.Bounds + " вне " + box);
                CheckFits(c, where, offenders);
            }
        }

        /// <summary>
        /// Соседние элементы управления одного окна не должны перекрывать друг друга.
        /// Раньше проверялись только кнопки — и мимо прошёл флажок, наехавший на сетку
        /// страниц на 20 px. Берём всё, что пользователь видит и с чем работает: кнопки,
        /// флажки, поля ввода, списки, сетку. Подписи и полосы пропускаем — они лежат
        /// фоном и накладываться друг на друга им не запрещено.
        /// </summary>
        private static void CheckControlsDoNotOverlap(System.Windows.Forms.Control parent, string where, List<string> offenders)
        {
            // Только ВИДИМЫЕ: поля разных режимов («по диапазонам» и «каждые N») намеренно
            // делят одно место и показываются по очереди — это не наложение, а один слот.
            var items = new List<System.Windows.Forms.Control>();
            foreach (System.Windows.Forms.Control c in parent.Controls)
                if (c.Visible && IsInteractive(c))
                    items.Add(c);
            for (int i = 0; i < items.Count; i++)
                for (int j = i + 1; j < items.Count; j++)
                    if (items[i].Bounds.IntersectsWith(items[j].Bounds))
                        offenders.Add(where + " → " + Describe(items[i]) + " накрывает " + Describe(items[j]));
        }

        /// <summary>Кнопки одного контейнера не перекрываются — включая пока скрытые.</summary>
        private static void CheckButtonsDoNotOverlap(System.Windows.Forms.Control parent, string where, List<string> offenders)
        {
            var buttons = new List<System.Windows.Forms.Control>();
            foreach (System.Windows.Forms.Control c in parent.Controls)
                if (c is System.Windows.Forms.ButtonBase)
                    buttons.Add(c);
            for (int i = 0; i < buttons.Count; i++)
                for (int j = i + 1; j < buttons.Count; j++)
                    if (buttons[i].Bounds.IntersectsWith(buttons[j].Bounds))
                        offenders.Add(where + " → " + Describe(buttons[i]) + " накрывает " + Describe(buttons[j]));
        }

        /// <summary>
        /// Подпись кнопки обязана помещаться в кнопку. <see cref="RoundedButton"/> режет длинную
        /// подпись многоточием, и обрезка молчалива: «Прочие операции» показывались как «Прочие
        /// опер…» ровно до тех пор, пока окно не попало на снимок для инструкции. Меряем тем же
        /// шрифтом и с тем же полем, каким кнопка рисует.
        /// </summary>
        private static void CheckButtonCaptionsFit(System.Windows.Forms.Control parent, string where,
            List<string> offenders)
        {
            foreach (System.Windows.Forms.Control c in parent.Controls)
            {
                CheckButtonCaptionsFit(c, where, offenders); // кнопки бывают и внутри панелей
                var b = c as System.Windows.Forms.ButtonBase;
                if (b == null || string.IsNullOrEmpty(b.Text) || b.Width <= 0)
                    continue;
                // Флажок рисует квадрат слева от подписи, а не заливку под ней — своя ширина.
                if (b is System.Windows.Forms.CheckBox || b is System.Windows.Forms.RadioButton)
                    continue;
                int pad = b is RoundedButton ? RoundedButton.TextPadFor(b.Width) : 4;
                int need = Ui.TextWidth(b.Text, b.Font) + 2 * pad; // тем же замером, что и окна
                if (need > b.Width)
                    offenders.Add(where + " → «" + b.Text + "» не помещается: нужно " + need +
                        " px, а кнопка " + b.Width);
            }
        }

        /// <summary>Элемент, с которым работает пользователь (а не фоновая подпись).</summary>
        private static bool IsInteractive(System.Windows.Forms.Control c)
        {
            // ВАЖНО: собственные составные элементы (выбор сжатия — UserControl со списком
            // внутри) под стандартные типы не подходят и раньше выпадали из проверки — из-за
            // этого кнопки панели наехали на список сжатия и это никто не заметил.
            return c is System.Windows.Forms.ButtonBase ||
                   c is System.Windows.Forms.TextBoxBase ||
                   c is System.Windows.Forms.ComboBox ||
                   c is System.Windows.Forms.NumericUpDown ||
                   c is System.Windows.Forms.ListView ||
                   c is System.Windows.Forms.TrackBar ||
                   c is PdfPageGrid ||
                   c is CompressionPicker;
        }

        private static string Describe(System.Windows.Forms.Control c)
        {
            return c.GetType().Name + (string.IsNullOrEmpty(c.Text) ? "" : "«" + c.Text + "»");
        }

        /// <summary>
        /// Кнопка «Главная» на шапке обязана быть ПОСЛЕДНЕЙ в обходе Tab, а не первой:
        /// иначе окно открывается с фокусом на ней и первый же Enter выкидывает
        /// пользователя обратно в меню, ничего не сделав. Ловушка тонкая:
        /// ControlCollection.Add раздаёт следующий свободный TabIndex, поэтому «большой»
        /// индекс, заданный шапке ПРИ ДОБАВЛЕНИИ, ставит её, наоборот, первой.
        /// </summary>
        private static void TestHomeHeaderLastInTabOrder()
        {
            var offenders = new List<string>();
            InIsolatedSettings("iwo_tab_", delegate
            {
                // Стартовый экран тоже: его шапка несёт глобус выбора языка, доступный с
                // клавиатуры, и без переноса в конец фокус при открытии доставался бы ему,
                // а не карточке инструмента.
                var windows = new List<System.Windows.Forms.Form>(NewAllToolWindows());
                windows.Add(new StartForm());
                foreach (System.Windows.Forms.Form f in windows)
                    using (f)
                    {
                        string where = f.GetType().Name;
                        try
                        {
                            f.Show();
                            HeaderBand header = null;
                            System.Windows.Forms.Control first = null;
                            foreach (System.Windows.Forms.Control c in f.Controls)
                            {
                                if (c is HeaderBand)
                                    header = (HeaderBand)c;
                                if ((c.TabStop || c.Controls.Count > 0) &&
                                    (first == null || c.TabIndex < first.TabIndex))
                                    first = c;
                            }
                            if (header == null)
                                offenders.Add(where + ": шапки нет");
                            else if (ReferenceEquals(first, header))
                                offenders.Add(where + ": шапка первая в обходе Tab (TabIndex=" +
                                    header.TabIndex + ") — фокус попадёт на «Главная»");
                            f.Close();
                        }
                        catch (Exception ex) { offenders.Add(where + ": " + ex.Message); }
                    }
            });
            AssertTrue(offenders.Count == 0, "обход Tab: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>Все окна-инструменты (с непустой «Главной») — по ним идут живые проверки раскладки.</summary>
        private static System.Windows.Forms.Form[] NewAllToolWindows()
        {
            Action back = delegate { };
            return new System.Windows.Forms.Form[]
            {
                new MainForm(back), new PdfMergeForm(back), new PdfSplitForm(back), new OcrForm(back),
                new PptxForm(back), new PdfOpsForm(back)
            };
        }

        /// <summary>
        /// Выполнить тело на STA-потоке с настройками в отдельной временной папке: живые окна
        /// пишут размеры и положение при закрытии, и без изоляции тест затирал бы НАСТОЯЩИЕ
        /// настройки пользователя. Общая обвязка (DRY: три живых теста делали это копией).
        /// </summary>
        private static void InIsolatedSettings(string prefix, Action body)
        {
            string root = Path.Combine(Path.GetTempPath(), prefix + Guid.NewGuid().ToString("N"));
            AppPaths.SetRootForTests(root);
            try
            {
                var th = new System.Threading.Thread(delegate() { body(); });
                th.SetApartmentState(System.Threading.ApartmentState.STA); // окна WinForms требуют STA
                th.IsBackground = true;
                th.Start();
                th.Join();
            }
            finally
            {
                AppPaths.SetRootForTests(null);
                try { Directory.Delete(root, true); } catch { }
            }
        }

        /// <summary>
        /// Строки шапки считаются от РЕАЛЬНЫХ высот шрифтов: шрифт задан в пунктах и растёт
        /// вместе с масштабом экрана, а прежние литералы 26/20 не росли — на 125% и выше низ
        /// заголовка срезался. Проверяем весь диапазон масштабов (96…192 dpi).
        /// </summary>
        private static void TestHeaderRowsFitAnyDpi()
        {
            foreach (double scale in new double[] { 1.0, 1.25, 1.5, 1.75, 2.0 })
            {
                // Полоса шапки и шрифты масштабируются вместе (AutoScaleMode.Dpi).
                int band = (int)(76 * scale);
                int titleH = (int)Math.Ceiling(28 * scale);    // 15 pt Bold
                int subtitleH = (int)Math.Ceiling(18 * scale); // 9.75 pt
                int titleY, subtitleY;
                HeaderBand.TextRows(band, titleH, subtitleH, out titleY, out subtitleY);
                string at = " при масштабе " + scale;
                AssertTrue(titleY >= 0, "заголовок не выше шапки" + at);
                AssertTrue(titleY + titleH <= subtitleY, "заголовок не наезжает на подпись" + at);
                AssertTrue(subtitleY + subtitleH <= band, "подпись не свисает из шапки" + at);
                // Блок стоит ПО ЦЕНТРУ полосы: сверху и снизу остаётся поровну (в пределах
                // округления на нечётной высоте). Прижатый к низу блок оставлял пустую синеву.
                int above = titleY, below = band - (subtitleY + subtitleH);
                AssertTrue(Math.Abs(above - below) <= 1,
                    "текст по центру полосы" + at + " (сверху " + above + ", снизу " + below + ")");

                // Глобус языка выравнивается по центру текстового блока: он обязан оказаться
                // внутри блока и целиком внутри шапки на любом масштабе, иначе иконка снова
                // повиснет выше текста (как было до 1.17.9) или уедет за нижнюю грань.
                int center = HeaderBand.TextBlockCenter(band, titleH, subtitleH);
                AssertTrue(center > titleY && center < subtitleY + subtitleH,
                    "центр текстового блока внутри самого блока" + at);
                int iconH = (int)Math.Ceiling(30 * scale);
                int iconTop = center - iconH / 2;
                AssertTrue(iconTop >= 0 && iconTop + iconH <= band,
                    "иконка по этому центру целиком внутри шапки" + at + " (top=" + iconTop + ", h=" + iconH + ")");

                // Шапка без подписи: центрируется ОДИН заголовок, пустая строка места не занимает.
                int soloTitleY, soloSubY;
                HeaderBand.TextRows(band, titleH, 0, out soloTitleY, out soloSubY);
                AssertTrue(Math.Abs(soloTitleY - (band - titleH) / 2) <= 1,
                    "без подписи по центру стоит сам заголовок" + at);
            }
        }

        // ---------- Масштаб просмотра ----------

        /// <summary>
        /// Положение страницы внутри прокручиваемой области. Смещение прокрутки прибавляется
        /// как есть: WinForms хранит положение дочернего элемента в координатах СОДЕРЖИМОГО, а
        /// AutoScrollPosition читается отрицательным. Без этой поправки начало содержимого
        /// уползало на величину прокрутки при каждом шаге лупы, и страница пряталась в сером
        /// поле — жалоба, с которой началась правка.
        /// </summary>
        private static void TestPreviewCentered()
        {
            var view = new System.Drawing.Size(800, 600);
            // Мельче окна и без прокрутки — ровно по центру.
            AssertEqual(new System.Drawing.Point(275, 200),
                PreviewZoom.Centered(new System.Drawing.Size(250, 200), view, System.Drawing.Point.Empty),
                "мелкая страница по центру");
            // Крупнее окна — в начало содержимого, без отрицательных полей.
            AssertEqual(new System.Drawing.Point(0, 0),
                PreviewZoom.Centered(new System.Drawing.Size(1600, 1200), view, System.Drawing.Point.Empty),
                "крупная страница от начала области");
            // Прокрутка сдвигает НАЧАЛО содержимого: центр остаётся тем же местом страницы.
            AssertEqual(new System.Drawing.Point(275 - 120, 200 - 340),
                PreviewZoom.Centered(new System.Drawing.Size(250, 200), view,
                    new System.Drawing.Point(-120, -340)),
                "прокрутка учтена в координатах содержимого");
            AssertEqual(new System.Drawing.Point(-120, -340),
                PreviewZoom.Centered(new System.Drawing.Size(1600, 1200), view,
                    new System.Drawing.Point(-120, -340)),
                "крупная страница при прокрутке стоит ровно на смещении");
        }

        private static void TestPreviewZoomSteps()
        {
            AssertEqual(1.25, PreviewZoom.Next(1.00, +1), "следующая ступень вверх от натуральной");
            AssertEqual(0.75, PreviewZoom.Next(1.00, -1), "следующая ступень вниз от натуральной");
            // Подгонка по окну даёт масштаб МЕЖДУ ступенями — шаг обязан вести к соседней.
            AssertEqual(0.50, PreviewZoom.Next(0.42, +1), "вверх от промежуточного значения");
            AssertEqual(0.33, PreviewZoom.Next(0.42, -1), "вниз от промежуточного значения");
            AssertEqual(PreviewZoom.Max, PreviewZoom.Next(PreviewZoom.Max, +1), "выше предела не уходим");
            AssertEqual(PreviewZoom.Min, PreviewZoom.Next(PreviewZoom.Min, -1), "ниже предела не уходим");
            AssertEqual(100, PreviewZoom.Percent(1.0), "проценты для подписи");
            AssertEqual(33, PreviewZoom.Percent(0.33), "дробный масштаб округляется");
        }

        private static void TestPreviewZoomFit()
        {
            // Широкая страница в узком окне ограничена шириной, высокая — высотой.
            AssertEqual(0.5, PreviewZoom.Fit(new System.Drawing.Size(1000, 500), new System.Drawing.Size(500, 500)),
                "вписывание по ширине");
            AssertEqual(0.5, PreviewZoom.Fit(new System.Drawing.Size(500, 1000), new System.Drawing.Size(500, 500)),
                "вписывание по высоте");
            // Мелкую страницу не растягиваем: увеличенное мыло хуже честного размера.
            AssertEqual(1.0, PreviewZoom.Fit(new System.Drawing.Size(100, 100), new System.Drawing.Size(900, 900)),
                "мелкая страница не растягивается");
            AssertEqual(1.0, PreviewZoom.Fit(new System.Drawing.Size(0, 0), new System.Drawing.Size(500, 500)),
                "пустая картинка не роняет расчёт");
        }

        /// <summary>
        /// Ctrl+колесо обязано увеличивать К ТОЧКЕ ПОД КУРСОРОМ, а не к центру: иначе
        /// интересное место уезжает за край, и пользователь его догоняет прокруткой.
        /// </summary>
        private static void TestPreviewZoomAnchor()
        {
            // Курсор в 100 px от края области, прокрутка 200 -> под курсором точка 300.
            // При удвоении масштаба та же точка обязана остаться под курсором: 300*2-100 = 500.
            AssertEqual(500, PreviewZoom.Anchor(200, 100, 1.0, 2.0), "точка под курсором осталась на месте");
            AssertEqual(50, PreviewZoom.Anchor(200, 100, 1.0, 0.5), "то же при уменьшении");
            AssertEqual(200, PreviewZoom.Anchor(200, 100, 1.0, 1.0), "без смены масштаба ничего не двигается");
            AssertEqual(0, PreviewZoom.Anchor(0, 100, 1.0, 0.25), "отрицательная прокрутка обрезается нулём");
            AssertEqual(200, PreviewZoom.Anchor(200, 100, 0, 2.0), "нулевой исходный масштаб не ломает расчёт");
        }

        private static void TestPreviewPan()
        {
            var viewport = new System.Drawing.Size(800, 600);
            AssertTrue(PreviewZoom.FitsEntirely(new System.Drawing.Size(800, 600), viewport), "точно по размеру — панорама не нужна");
            AssertTrue(!PreviewZoom.FitsEntirely(new System.Drawing.Size(801, 600), viewport), "шире области — нужна панорама");

            var t = new System.Drawing.Size(8, 8);
            AssertTrue(!PreviewZoom.IsDrag(new System.Drawing.Point(10, 10), new System.Drawing.Point(13, 13), t),
                "дрожание руки при клике перетаскиванием не считается");
            AssertTrue(PreviewZoom.IsDrag(new System.Drawing.Point(10, 10), new System.Drawing.Point(30, 10), t),
                "заметное движение — перетаскивание");
        }

        /// <summary>
        /// Выбор одной стороны двусторонней пачки. Нумерация — пользовательская, с единицы:
        /// нечётные страницы 1, 3, 5 это индексы 0, 2, 4. Перепутать легко, а последствие
        /// заметное — пользователь повернул бы не ту сторону всего документа.
        /// </summary>
        private static void TestSelectEveryOther()
        {
            AssertEqual("0,2,4", string.Join(",", Array.ConvertAll(
                PdfPageGrid.EveryOtherIndices(6, false), Convert.ToString)), "нечётные страницы 1, 3, 5");
            AssertEqual("1,3,5", string.Join(",", Array.ConvertAll(
                PdfPageGrid.EveryOtherIndices(6, true), Convert.ToString)), "чётные страницы 2, 4, 6");
            AssertEqual("0,2", string.Join(",", Array.ConvertAll(
                PdfPageGrid.EveryOtherIndices(3, false), Convert.ToString)), "нечётные при нечётной длине");
            AssertEqual("1", string.Join(",", Array.ConvertAll(
                PdfPageGrid.EveryOtherIndices(3, true), Convert.ToString)), "чётные при нечётной длине");
            AssertEqual(0, PdfPageGrid.EveryOtherIndices(0, true).Length, "пустой документ — пустой выбор");
            AssertEqual(0, PdfPageGrid.EveryOtherIndices(-5, false).Length, "отрицательная длина не роняет");
            AssertEqual(0, PdfPageGrid.EveryOtherIndices(1, true).Length, "одна страница — чётных нет");
        }

        /// <summary>
        /// Добивка до чётного для двусторонней печати: пустая страница ставится после
        /// документа с нечётным числом страниц, но НЕ после последнего — за ним ничего не
        /// печатается, и лишний пустой лист был бы мусором.
        /// </summary>
        private static void TestBlankPagePositions()
        {
            // a: 3 страницы (нечёт) -> пустая после позиции 3; b: 2 (чёт) -> ничего.
            var pages = new List<PdfPageRef>
            {
                PR("a", 0), PR("a", 1), PR("a", 2), PR("b", 0), PR("b", 1)
            };
            AssertEqual("3", string.Join(",", BlankPages.InsertPositions(pages).ConvertAll(Convert.ToString).ToArray()),
                "пустая после нечётного документа");
            AssertTrue(BlankPages.Needed(pages), "добивка нужна");

            // Последний документ не добиваем, даже если он нечётный.
            var tail = new List<PdfPageRef> { PR("a", 0), PR("a", 1), PR("b", 0) };
            AssertEqual(0, BlankPages.InsertPositions(tail).Count, "последний документ не добиваем");
            AssertTrue(!BlankPages.Needed(tail), "добивка не нужна");

            // Три нечётных документа подряд: две вставки, обе по ИСХОДНОЙ нумерации.
            var three = new List<PdfPageRef> { PR("a", 0), PR("b", 0), PR("c", 0) };
            AssertEqual("1,2", string.Join(",", BlankPages.InsertPositions(three).ConvertAll(Convert.ToString).ToArray()),
                "позиции считаются по исходной нумерации");
            AssertEqual(0, BlankPages.InsertPositions(new List<PdfPageRef>()).Count, "пустой набор");
            AssertEqual(0, BlankPages.InsertPositions(null).Count, "null не роняет");
        }

        /// <summary>
        /// ЖИВАЯ проверка добивки: считать позиции мало — надо убедиться, что пустые страницы
        /// действительно попали в файл и того же размера, что соседние, иначе добивочный лист
        /// выпал бы из формата пачки при печати.
        /// </summary>
        private static void TestBlankPageMergeLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerPad_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string a = Path.Combine(dir, "три.pdf"), b = Path.Combine(dir, "две.pdf");
                MakeEmptyPagesPdf(a, 3);
                MakeEmptyPagesPdf(b, 2);
                var order = new List<PdfPageRef>();
                for (int i = 0; i < 3; i++) order.Add(PR(a, i));
                for (int i = 0; i < 2; i++) order.Add(PR(b, i));

                string plain = Path.Combine(dir, "без.pdf");
                PdfMergeService.Merge(order, plain);
                AssertEqual(5, PdfPageCount(plain), "без добивки — пять страниц");

                string padded = Path.Combine(dir, "с.pdf");
                PdfMergeService.Merge(order, padded, null, null, true);
                AssertEqual(6, PdfPageCount(padded), "с добивкой — шесть страниц");

                using (PdfDocument doc = PdfReader.Open(padded, PdfDocumentOpenMode.InformationOnly))
                {
                    // Пустая встала ЧЕТВЁРТОЙ (после трёх страниц первого документа).
                    AssertEqual(doc.Pages[2].Width.Point, doc.Pages[3].Width.Point, "ширина добивочной как у соседней");
                    AssertEqual(doc.Pages[2].Height.Point, doc.Pages[3].Height.Point, "высота добивочной как у соседней");
                }
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        /// <summary>
        /// Шаблон имени НЕОБЯЗАТЕЛЕН: пустой обязан давать ровно прежние имена, иначе
        /// появление новой возможности переименовало бы файлы у всех, кто ей не пользуется.
        /// </summary>
        private static void TestPartNameOptional()
        {
            var v = new NameValues { BaseName = "свод", FileNumber = 3, TotalFiles = 9, CurrentPage = 5 };
            AssertEqual("свод_1-3", PdfSplitService.PartName(null, "свод_1-3", v), "без шаблона — прежнее имя");
            AssertEqual("свод_1-3", PdfSplitService.PartName("", "свод_1-3", v), "пустой шаблон — прежнее имя");
            AssertEqual("свод-003", PdfSplitService.PartName("[BASENAME]-[FILENUMBER###]", "свод_1-3", v),
                "шаблон заменяет имя целиком");
        }

        /// <summary>ЖИВАЯ проверка: с шаблоном части получают заданные имена, без него — прежние.</summary>
        private static void TestSplitTemplateLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerTpl_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = Path.Combine(dir, "исходник.pdf");
                MakeEmptyPagesPdf(src, 6);

                List<string> legacy = PdfSplitService.SplitEveryN(src, 3, dir, "часть");
                AssertEqual("часть", Path.GetFileNameWithoutExtension(legacy[0]).Substring(0, 5),
                    "без шаблона имя начинается как раньше");

                List<string> shaped = PdfSplitService.SplitEveryN(src, 3, dir, "часть", null, null, null,
                    "лист[FILENUMBER##]из[TOTAL_FILES]");
                AssertEqual("лист01из2", Path.GetFileNameWithoutExtension(shaped[0]), "первая часть по шаблону");
                AssertEqual("лист02из2", Path.GetFileNameWithoutExtension(shaped[1]), "вторая часть по шаблону");
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        /// <summary>
        /// Запись «в самого себя» обязана распознаваться до начала работы: приложение не меняет
        /// исходники, а операция ещё и испортила бы файл — источник в этот момент открыт на чтение.
        /// Разный регистр и «.\» в пути — тот же файл, файловая система Windows их не различает.
        /// </summary>
        private static void TestSameFileGuard()
        {
            string dir = Path.GetTempPath();
            string a = Path.Combine(dir, "документ.pdf");
            AssertTrue(OutputFile.IsSameFile(a, a), "тот же путь");
            AssertTrue(OutputFile.IsSameFile(a, Path.Combine(dir, "ДОКУМЕНТ.PDF")), "регистр не различает файлы");
            AssertTrue(OutputFile.IsSameFile(a, Path.Combine(dir, ".", "документ.pdf")), "путь через «.» — тот же файл");
            AssertTrue(!OutputFile.IsSameFile(a, Path.Combine(dir, "другой.pdf")), "разные файлы");
            AssertTrue(!OutputFile.IsSameFile(null, a), "null — не тот же файл");
            AssertTrue(!OutputFile.IsSameFile(a, ""), "пустой путь — не тот же файл");
            AssertTrue(!OutputFile.IsSameFile(a, "|негодный<путь>"), "негодный путь не роняет проверку");
        }

        /// <summary>
        /// Страница обязана вписываться в лист ЦЕЛИКОМ и по центру: растянуть её по краям —
        /// значит срезать поля, где у документов стоят подписи, номера и отметки.
        /// </summary>
        private static void TestPrintFitToPage()
        {
            var sheet = new System.Drawing.Rectangle(0, 0, 1000, 1000);
            // Широкая страница ограничена шириной и центрируется по вертикали.
            System.Drawing.Rectangle wide = PdfPrintService.FitToPage(new System.Drawing.Size(2000, 1000), sheet);
            AssertEqual(1000, wide.Width, "по ширине листа");
            AssertEqual(500, wide.Height, "пропорции сохранены");
            AssertEqual(250, wide.Y, "по центру листа");
            AssertEqual(0, wide.X, "прижата к краю по ширине");
            // Высокая — ограничена высотой и центрируется по горизонтали.
            System.Drawing.Rectangle tall = PdfPrintService.FitToPage(new System.Drawing.Size(1000, 2000), sheet);
            AssertEqual(500, tall.Width, "по высоте листа");
            AssertEqual(250, tall.X, "по центру листа");
            // Мелкая страница увеличивается до листа — на бумаге пустые поля не нужны.
            System.Drawing.Rectangle small = PdfPrintService.FitToPage(new System.Drawing.Size(100, 100), sheet);
            AssertEqual(1000, small.Width, "мелкая страница занимает лист");
            // Вырожденные размеры не роняют печать.
            AssertEqual(sheet, PdfPrintService.FitToPage(new System.Drawing.Size(0, 0), sheet), "нулевая страница");
            AssertEqual(new System.Drawing.Rectangle(0, 0, 0, 0),
                PdfPrintService.FitToPage(new System.Drawing.Size(10, 10), new System.Drawing.Rectangle(0, 0, 0, 0)),
                "нулевой лист");
        }

        /// <summary>
        /// Подписи сжатия называют разрешение, до которого уменьшаются изображения, и число
        /// берётся из того же места, что и аргументы движка, — иначе подпись обещала бы одно,
        /// а получалось бы другое.
        /// </summary>
        private static void TestCompressionLabelsNameDpi()
        {
            // Разрешение называет ровно тот уровень, который его меняет: у остальных числа в
            // подписи быть не должно, иначе «Очень хорошо» обещало бы несуществующие dpi.
            foreach (CompressionLevel level in Enum.GetValues(typeof(CompressionLevel)))
            {
                string label = PdfCompression.Label(level);
                bool namesDpi = label.Contains(PdfCompression.ImageDpi(level).ToString());
                AssertEqual(PdfCompression.Downsamples(level), namesDpi,
                    "разрешение в подписи ровно у пересчитывающих уровней (" + level + "): " + label);
                AssertTrue(!label.Contains("{0}"), "подстановка выполнена, а не показана как есть: " + label);
            }
            // Подпись из списка и подпись уровня — одно и то же (список строится из них).
            string[] labels = PdfCompression.LevelLabels();
            for (int i = 0; i < PdfCompression.LevelLabels().Length; i++)
                AssertEqual(PdfCompression.Label(PdfCompression.LevelAt(i)), labels[i],
                    "подпись позиции " + i + " берётся из уровня");
        }

        /// <summary>
        /// Смена языка при открытом инструменте обязана оставить ГЛАВНЫЙ ЭКРАН рабочим:
        /// он пересоздаётся вместе с остальными, и если после пересборки на нём не окажется
        /// карточек или он окажется закрыт, пользователь останется без входа в инструменты.
        ///
        /// С 1.17.9 экран двухуровневый, поэтому проверяется ещё и РАЗДЕЛ: без его переноса
        /// смена языка выбрасывала бы человека из «Инструментов PDF» на верхний экран — та же
        /// по природе потеря состояния, что и уехавший за край свёрнутый хаб в 1.17.8.
        /// Пересборку зовём напрямую: обычный путь идёт через отложенный вызов и требует
        /// цикла сообщений, а проверяем мы здесь именно её результат.
        /// </summary>
        private static void TestLanguageRebuildKeepsHubUsable()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_lang_rebuild_", delegate
            {
                Lang saved = Loc.Current;
                var ctx = new ShellContext();
                try
                {
                    ctx.OpenTool("split", "Split", delegate(Action back) { return new PdfSplitForm(back); }, HubLevel.Pdf);
                    var before = HubOf(ctx) as StartForm;
                    if (before == null) { failures.Add("главного экрана нет ещё до смены языка"); return; }
                    before.ShowLevel(HubLevel.Pdf); // человек стоит в разделе PDF
                    Loc.Init(saved == Lang.Ru ? Lang.En : Lang.Ru); // язык уже другой, как в момент пересборки
                    typeof(ShellContext).GetMethod("RebuildOpenWindows",
                        System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                        .Invoke(ctx, null);

                    var hub = HubOf(ctx) as StartForm;
                    if (hub == null || hub.IsDisposed)
                        failures.Add("после смены языка главного экрана нет");
                    else
                    {
                        if (hub.Level != HubLevel.Pdf)
                            failures.Add("раздел не сохранился: " + hub.Level);
                        int cards = VisibleCards(hub);
                        if (cards != 5)
                            failures.Add("в разделе PDF карточек " + cards + ", а должно быть 5");
                        if (!hub.Visible)
                            failures.Add("главный экран не показан");
                    }
                }
                catch (Exception ex) { failures.Add(ex.GetType().Name + ": " + ex.Message); }
                finally
                {
                    Loc.Init(saved);
                    try { ctx.Dispose(); } catch { }
                }
            });
            AssertTrue(failures.Count == 0, "пересборка по смене языка: " + string.Join(" | ", failures.ToArray()));
        }

        /// <summary>
        /// Двухуровневый стартовый экран: переходы вперёд и назад показывают ровно свой набор
        /// карточек, кнопка «Назад» появляется только внутри раздела, а Esc возвращает наверх.
        /// Проверяется через настоящие клики по карточкам — иначе потерянный обработчик
        /// (карточка есть, а нажатие ничего не делает) прошёл бы мимо.
        /// </summary>
        private static void TestHubNavigation()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_hub_nav_", delegate
            {
                using (var hub = new StartForm())
                {
                    hub.Show();
                    Check(hub.Level == HubLevel.Main, "стартуем с главного уровня", failures);
                    Check(VisibleCards(hub) == 2, "на главном две карточки разделов", failures);
                    Check(!BackButton(hub).Visible, "на главном «Назад» не показывается", failures);

                    ClickCard(FirstVisibleCard(hub)); // карточка «PDF»
                    Check(hub.Level == HubLevel.Pdf, "клик по разделу PDF открыл его", failures);
                    Check(VisibleCards(hub) == 5, "в разделе PDF пять инструментов", failures);
                    Check(BackButton(hub).Visible, "в разделе есть «Назад»", failures);

                    BackButton(hub).PerformClick();
                    Check(hub.Level == HubLevel.Main, "«Назад» вернул на главный", failures);

                    // Второй раздел: карточек там пока одна, и это нормально.
                    var cards = new List<ChoiceCard>();
                    CollectCards(hub, cards);
                    ClickCard(cards[1]);
                    Check(hub.Level == HubLevel.Other, "клик по «Иному функционалу» открыл его", failures);
                    Check(VisibleCards(hub) == 1, "в ином функционале одна карточка", failures);

                    hub.ShowLevel(HubLevel.Pdf);
                    SendEscape(hub);
                    Check(hub.Level == HubLevel.Main, "Esc возвращает из раздела", failures);
                    hub.Close();
                }
            });
            AssertTrue(failures.Count == 0, "навигация хаба: " + string.Join(" | ", failures.ToArray()));
        }

        /// <summary>
        /// Файлы, брошенные на карточку раздела, ждут выбора инструмента — но НЕ ДОЛЬШЕ. Набор
        /// обязан забыться при уходе из раздела: иначе он «прилипнет», и следующий клик по
        /// карточке откроет инструмент с чужими файлами. Самая вероятная ошибка в этой схеме.
        /// </summary>
        private static void TestHubPendingFilesCleared()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_hub_pending_", delegate
            {
                using (var hub = new StartForm())
                {
                    hub.Show();
                    hub.ShowLevel(HubLevel.Pdf);
                    SetPending(hub, new[] { "a.pdf", "b.pdf" });
                    Check(GetPending(hub) != null, "набор принят", failures);

                    hub.ShowLevel(HubLevel.Main);
                    Check(GetPending(hub) == null, "уход из раздела забывает набор", failures);

                    hub.ShowLevel(HubLevel.Pdf);
                    SetPending(hub, new[] { "a.pdf" });
                    BackButton(hub).PerformClick();
                    Check(GetPending(hub) == null, "«Назад» забывает набор", failures);
                    hub.Close();
                }
            });
            AssertTrue(failures.Count == 0, "придержанные файлы: " + string.Join(" | ", failures.ToArray()));
        }

        /// <summary>
        /// Из хаба обязан открываться КАЖДЫЙ инструмент, и именно свой: перепутанная фабрика
        /// (карточка «Прочие операции», открывающая объединение) — ошибка, которую глазами
        /// ловят в последнюю очередь. Кликаем по всем карточкам обоих разделов и сверяем типы.
        /// </summary>
        private static void TestHubOpensEveryTool()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_hub_tools_", delegate
            {
                var ctx = new ShellContext();
                try
                {
                    var hub = HubOf(ctx) as StartForm;
                    if (hub == null) { failures.Add("хаба нет"); return; }
                    var expected = new List<Type>
                    {
                        // Порядок — как в сетке хаба: верхний ряд операций, нижний — переводы.
                        typeof(PdfMergeForm), typeof(PdfSplitForm), typeof(PdfOpsForm),
                        typeof(OcrForm), typeof(PptxForm)
                    };
                    hub.ShowLevel(HubLevel.Pdf);
                    var cards = new List<ChoiceCard>();
                    CollectCards(hub, cards);
                    if (cards.Count != 5) { failures.Add("карточек в разделе PDF: " + cards.Count); return; }
                    for (int i = 0; i < cards.Count; i++)
                    {
                        ClickCard(cards[i]);
                        if (FindOpenTool(ctx, expected[i]) == null)
                            failures.Add("карточка " + i + " не открыла " + expected[i].Name);
                    }

                    hub.ShowLevel(HubLevel.Other);
                    cards.Clear();
                    CollectCards(hub, cards);
                    if (cards.Count != 1) { failures.Add("карточек в ином функционале: " + cards.Count); return; }
                    ClickCard(cards[0]);
                    if (FindOpenTool(ctx, typeof(MainForm)) == null)
                        failures.Add("карточка не открыла свод Excel");
                }
                catch (Exception ex) { failures.Add(Root(ex).GetType().Name + ": " + Root(ex).Message); }
                finally
                {
                    CloseOpenTools(ctx);
                    try { ctx.Dispose(); } catch { }
                }
            });
            AssertTrue(failures.Count == 0, "открытие инструментов из хаба: " + string.Join(" | ", failures.ToArray()));
        }

        /// <summary>
        /// «Разделение» передаёт ОТКРЫТЫЙ документ в «Прочие операции» одной кнопкой. Путь
        /// длинный и весь из проводки: фабрика окна живёт на стартовом экране, файл едет через
        /// `IFileAcceptor`, а разбор идёт в фоне. Порвётся любое звено — человек будет открывать
        /// один и тот же файл дважды, и заметит это не тест, а он.
        /// </summary>
        private static void TestSplitHandsDocumentToOps()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_ops_bridge_", delegate
            {
                string dir = Path.Combine(Path.GetTempPath(), "iwo_bridge_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(dir);
                string pdf = Path.Combine(dir, "документ.pdf");
                var ctx = new ShellContext();
                try
                {
                    PdfProbe.WriteOnePagePdf(pdf);
                    var hub = HubOf(ctx) as StartForm;
                    if (hub == null) { failures.Add("хаба нет"); return; }
                    hub.ShowLevel(HubLevel.Pdf);
                    var cards = new List<ChoiceCard>();
                    CollectCards(hub, cards);
                    ClickCard(cards[1]); // «Разделение PDF»

                    var split = FindOpenTool(ctx, typeof(PdfSplitForm));
                    if (split == null) { failures.Add("«Разделение» не открылось"); return; }
                    ((IFileAcceptor)split).AcceptFiles(new[] { pdf });
                    if (!WaitFor(delegate { return SourcePathOf(split) != null; }))
                    { failures.Add("документ не загрузился в «Разделение»"); return; }

                    System.Windows.Forms.Button toOps = ButtonWithText(split, Loc.T("split.btn.ops"));
                    if (toOps == null) { failures.Add("кнопки перехода нет"); return; }
                    if (!toOps.Enabled) failures.Add("кнопка перехода недоступна при открытом документе");
                    toOps.PerformClick();

                    var ops = FindOpenTool(ctx, typeof(PdfOpsForm));
                    if (ops == null) { failures.Add("«Прочие операции» не открылись"); return; }
                    if (!WaitFor(delegate { return SourcePathOf(ops) != null; }))
                    { failures.Add("документ не доехал до «Прочих операций»"); return; }
                    if (!string.Equals(SourcePathOf(ops), pdf, StringComparison.OrdinalIgnoreCase))
                        failures.Add("доехал не тот документ: " + SourcePathOf(ops));
                }
                catch (Exception ex) { failures.Add(Root(ex).GetType().Name + ": " + Root(ex).Message); }
                finally
                {
                    CloseOpenTools(ctx);
                    try { ctx.Dispose(); } catch { }
                    try { Directory.Delete(dir, true); } catch { }
                }
            });
            AssertTrue(failures.Count == 0, "передача документа в «Прочие операции»: " +
                string.Join(" | ", failures.ToArray()));
        }

        /// <summary>
        /// «Объединение» отдаёт в «Прочие операции» СОБРАННЫЙ файл — единственный документ,
        /// который оно может назвать своим (источников там много, текущего документа нет).
        /// Проверяем настоящим нажатием пункта меню: обработчик, поле результата и мост — три
        /// звена, и молчаливо порваться может любое.
        /// </summary>
        private static void TestMergeHandsResultToOps()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_merge_ops_", delegate
            {
                string dir = Path.Combine(Path.GetTempPath(), "iwo_merge_ops_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(dir);
                string result = Path.Combine(dir, "собранный.pdf");
                try
                {
                    PdfProbe.WriteOnePagePdf(result);
                    Action back = delegate { };
                    string handed = null;
                    using (var merge = new PdfMergeForm(back))
                    {
                        SetProtected(merge, "OpsBridge", (Action<string>)delegate(string p) { handed = p; });
                        System.Windows.Forms.ToolStripMenuItem item =
                            MenuItemWithText(merge, Loc.T("pdf.menu.ops"));
                        if (item == null) { failures.Add("пункта меню нет"); return; }

                        // Пока ничего не собрано, пункт обязан ОТКРЫТЬ окно, а не объяснять,
                        // почему он ничего не делает: нажатие без последствий читается как
                        // сломанная кнопка. Мост зовётся с пустым путём — «просто открой».
                        bool called = false;
                        SetProtected(merge, "OpsBridge", (Action<string>)delegate(string p)
                        {
                            called = true;
                            handed = p;
                        });
                        item.PerformClick();
                        if (!called)
                            failures.Add("без собранного файла пункт не сделал ничего");
                        if (!string.IsNullOrEmpty(handed))
                            failures.Add("без собранного файла отдан путь: " + handed);

                        // С готовым файлом пункт обязан отдать именно его.
                        SetProtected(merge, "OpsBridge", (Action<string>)delegate(string p) { handed = p; });
                        SetField(merge, "_lastResult", result);
                        item.PerformClick();
                        if (!string.Equals(handed, result, StringComparison.OrdinalIgnoreCase))
                            failures.Add("отдан не тот файл: " + (handed ?? "ничего"));
                    }
                }
                catch (Exception ex) { failures.Add(Root(ex).GetType().Name + ": " + Root(ex).Message); }
                finally { try { Directory.Delete(dir, true); } catch { } }
            });
            AssertTrue(failures.Count == 0, "передача собранного файла: " + string.Join(" | ", failures.ToArray()));
        }

        private static void SetProtected(object target, string property, object value)
        {
            Type t = target.GetType();
            while (t != null)
            {
                System.Reflection.PropertyInfo pi = t.GetProperty(property,
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                if (pi != null) { pi.SetValue(target, value, null); return; }
                t = t.BaseType;
            }
            throw new Exception("нет свойства " + property);
        }

        private static void SetField(object target, string field, object value)
        {
            Type t = target.GetType();
            while (t != null)
            {
                System.Reflection.FieldInfo fi = t.GetField(field,
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                if (fi != null) { fi.SetValue(target, value); return; }
                t = t.BaseType;
            }
            throw new Exception("нет поля " + field);
        }

        private static System.Windows.Forms.ToolStripMenuItem MenuItemWithText(
            System.Windows.Forms.Form form, string text)
        {
            if (form.MainMenuStrip == null)
                return null;
            foreach (System.Windows.Forms.ToolStripItem root in form.MainMenuStrip.Items)
            {
                System.Windows.Forms.ToolStripMenuItem found = FindMenuItem(root, text);
                if (found != null)
                    return found;
            }
            return null;
        }

        private static System.Windows.Forms.ToolStripMenuItem FindMenuItem(
            System.Windows.Forms.ToolStripItem item, string text)
        {
            var mi = item as System.Windows.Forms.ToolStripMenuItem;
            if (mi == null)
                return null;
            if (mi.Text == text)
                return mi;
            foreach (System.Windows.Forms.ToolStripItem sub in mi.DropDownItems)
            {
                System.Windows.Forms.ToolStripMenuItem found = FindMenuItem(sub, text);
                if (found != null)
                    return found;
            }
            return null;
        }

        /// <summary>Подождать условия, прокручивая очередь сообщений (разбор PDF идёт в фоне).</summary>
        private static bool WaitFor(Func<bool> ready)
        {
            for (int i = 0; i < 200; i++)
            {
                if (ready())
                    return true;
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(25);
            }
            return ready();
        }

        private static string SourcePathOf(System.Windows.Forms.Form form)
        {
            return (string)typeof(PdfSingleDocFormBase).GetField("_sourcePath",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .GetValue(form);
        }

        private static System.Windows.Forms.Button ButtonWithText(System.Windows.Forms.Control parent, string text)
        {
            foreach (System.Windows.Forms.Control c in parent.Controls)
            {
                var b = c as System.Windows.Forms.Button;
                if (b != null && b.Text == text)
                    return b;
                System.Windows.Forms.Button inner = ButtonWithText(c, text);
                if (inner != null)
                    return inner;
            }
            return null;
        }

        private static void Check(bool ok, string what, List<string> failures)
        {
            if (!ok)
                failures.Add(what);
        }

        /// <summary>Нажать карточку так же, как это делает мышь (обработчик Click подписан на неё).</summary>
        private static void ClickCard(ChoiceCard card)
        {
            typeof(System.Windows.Forms.Control).GetMethod("OnClick",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .Invoke(card, new object[] { EventArgs.Empty });
        }

        private static void SendEscape(System.Windows.Forms.Form form)
        {
            typeof(System.Windows.Forms.Control).GetMethod("OnKeyDown",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .Invoke(form, new object[]
                {
                    new System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape)
                });
        }

        private static ChoiceCard FirstVisibleCard(System.Windows.Forms.Control parent)
        {
            var cards = new List<ChoiceCard>();
            CollectCards(parent, cards);
            return cards.Count > 0 ? cards[0] : null;
        }

        /// <summary>Видимые карточки текущего уровня, сверху вниз и слева направо (как их видит глаз).</summary>
        private static void CollectCards(System.Windows.Forms.Control parent, List<ChoiceCard> into)
        {
            foreach (System.Windows.Forms.Control c in parent.Controls)
            {
                if (!c.Visible)
                    continue;
                var card = c as ChoiceCard;
                if (card != null)
                    into.Add(card);
                else
                    CollectCards(c, into);
            }
            into.Sort(delegate(ChoiceCard a, ChoiceCard b)
            {
                int byRow = a.Top.CompareTo(b.Top);
                return byRow != 0 ? byRow : a.Left.CompareTo(b.Left);
            });
        }

        private static System.Windows.Forms.Button BackButton(System.Windows.Forms.Control parent)
        {
            foreach (System.Windows.Forms.Control c in parent.Controls)
            {
                var b = c as System.Windows.Forms.Button;
                if (b != null && b.Text == Loc.T("hub.back"))
                    return b;
                System.Windows.Forms.Button inner = BackButton(c);
                if (inner != null)
                    return inner;
            }
            return null;
        }

        private static void SetPending(StartForm hub, string[] files)
        {
            typeof(StartForm).GetMethod("SetPending",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .Invoke(hub, new object[] { files });
        }

        private static string[] GetPending(StartForm hub)
        {
            return (string[])typeof(StartForm).GetField("_pending",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .GetValue(hub);
        }

        private static System.Windows.Forms.Form FindOpenTool(ShellContext ctx, Type type)
        {
            foreach (System.Windows.Forms.Form f in OpenTools(ctx))
                if (type.IsInstanceOfType(f))
                    return f;
            return null;
        }

        private static List<System.Windows.Forms.Form> OpenTools(ShellContext ctx)
        {
            var registry = (ToolRegistry)typeof(ShellContext).GetField("_tools",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .GetValue(ctx);
            return registry.OpenForms();
        }

        private static void CloseOpenTools(ShellContext ctx)
        {
            foreach (System.Windows.Forms.Form f in OpenTools(ctx))
                try { f.Close(); } catch { }
        }

        /// <summary>
        /// Кнопка рисуется вручную, и обе её геометрические величины — радиус скругления и
        /// боковое поле подписи — считаются от размера. Постоянные значения здесь не работают:
        /// кнопки в приложении бывают от 24 до 38 пикселей высотой и от 32 до 230 шириной.
        /// Поле проверяется отдельно, потому что его первая версия (постоянные 10 px) съела
        /// глифы «+» и «−» на квадратных кнопках лупы — они обрезались многоточием.
        /// </summary>
        private static void TestRoundedButtonMetrics()
        {
            // Радиус: пропорция от высоты, но в разумных пределах.
            AssertEqual(6f, RoundedButton.RadiusFor(24), "низкая кнопка");
            AssertEqual(7.5f, RoundedButton.RadiusFor(30), "кнопка панели");
            AssertEqual(9.5f, RoundedButton.RadiusFor(38), "кнопка действия");
            AssertEqual(5f, RoundedButton.RadiusFor(8), "вырожденная высота — нижний предел");
            AssertEqual(10f, RoundedButton.RadiusFor(200), "огромная высота — верхний предел");

            // Поле подписи: на широкой кнопке постоянное, на узкой — доля ширины.
            AssertEqual(10, RoundedButton.TextPadFor(150), "широкая кнопка");
            AssertEqual(10, RoundedButton.TextPadFor(80), "кнопка средней ширины");
            AssertEqual(4, RoundedButton.TextPadFor(32), "квадратная кнопка лупы");
            AssertTrue(RoundedButton.TextPadFor(32) * 2 < 32, "под подпись остаётся место");
            AssertTrue(RoundedButton.TextPadFor(1) >= 0, "вырожденная ширина не даёт отрицательного поля");
        }

        private static void TestCardContentCentered()
        {
            // Блок «значок + заголовок + описание» стоит по центру: поля сверху и снизу
            // равны (расхождение в 1 px — нечётный остаток, делить пиксель не на что).
            foreach (int block in new[] { 60, 113, 160, 197 })
            {
                int top = ChoiceCard.ContentTop(250, block);
                int bottom = 250 - block - top;
                AssertTrue(Math.Abs(top - bottom) <= 1,
                    "блок " + block + " px: сверху " + top + ", снизу " + bottom);
            }

            // Описание во всю карточку не должно уезжать под верхний край.
            AssertTrue(ChoiceCard.ContentTop(250, 240) >= 18, "переполненная карточка прижата к верху");
            AssertTrue(ChoiceCard.ContentTop(250, 400) >= 18, "блок выше карточки не уходит в минус");
        }

        /// <summary>Сколько карточек ВИДНО на текущем уровне хаба (уровни — панели одного окна).</summary>
        private static int VisibleCards(System.Windows.Forms.Control parent)
        {
            int n = 0;
            foreach (System.Windows.Forms.Control c in parent.Controls)
            {
                if (!c.Visible)
                    continue;
                if (c is ChoiceCard)
                    n++;
                else
                    n += VisibleCards(c);
            }
            return n;
        }

        /// <summary>
        /// СВЁРНУТЫЙ главный экран обязан пережить смену языка. У свёрнутого окна Bounds — это
        /// служебные координаты далеко за пределами экранов (-32000, -32000). Пересборка
        /// копировала их на новый хаб, и тот открывался ОБЫЧНЫМ окном за краем рабочего стола:
        /// в панели задач он есть, а на экране его нет и «Главная» будто ничего не делает.
        /// Проверяем и состояние (осталось свёрнутым), и «нормальные» границы (настоящие, на
        /// экране) — по отдельности каждое из них можно было бы удовлетворить и с ошибкой.
        /// </summary>
        private static void TestLanguageRebuildKeepsMinimizedHubOnScreen()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_lang_min_hub_", delegate
            {
                Lang saved = Loc.Current;
                var ctx = new ShellContext();
                try
                {
                    ctx.OpenTool("split", "Split", delegate(Action back) { return new PdfSplitForm(back); }, HubLevel.Pdf);
                    System.Windows.Forms.Form hub = HubOf(ctx);
                    if (hub == null) { failures.Add("главного экрана нет ещё до смены языка"); return; }
                    hub.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                    if (hub.Bounds.X > -30000)
                        failures.Add("окно не свернулось — проверять нечего: " + hub.Bounds);

                    Loc.Init(saved == Lang.Ru ? Lang.En : Lang.Ru);
                    typeof(ShellContext).GetMethod("RebuildOpenWindows",
                        System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                        .Invoke(ctx, null);

                    System.Windows.Forms.Form rebuilt = HubOf(ctx);
                    if (rebuilt == null || rebuilt.IsDisposed)
                    {
                        failures.Add("после смены языка главного экрана нет");
                        return;
                    }
                    if (rebuilt.WindowState != System.Windows.Forms.FormWindowState.Minimized)
                        failures.Add("свёрнутый главный экран развернулся сам: " + rebuilt.WindowState);
                    System.Drawing.Rectangle normal = WindowPlacement.NormalBounds(rebuilt);
                    if (!OnAnyWorkArea(normal))
                        failures.Add("главный экран уехал за пределы экранов: " + normal);
                }
                catch (Exception ex) { failures.Add(ex.GetType().Name + ": " + ex.Message); }
                finally
                {
                    Loc.Init(saved);
                    try { ctx.Dispose(); } catch { }
                }
            });
            AssertTrue(failures.Count == 0, "свёрнутый главный экран: " + string.Join(" | ", failures.ToArray()));
        }

        /// <summary>Текущий главный экран приложения (приватное поле оболочки).</summary>
        private static System.Windows.Forms.Form HubOf(ShellContext ctx)
        {
            return (System.Windows.Forms.Form)typeof(ShellContext)
                .GetField("_hub", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .GetValue(ctx);
        }

        /// <summary>Пересекается ли прямоугольник хоть с одной рабочей областью экранов.</summary>
        private static bool OnAnyWorkArea(System.Drawing.Rectangle r)
        {
            foreach (System.Windows.Forms.Screen s in System.Windows.Forms.Screen.AllScreens)
                if (s.WorkingArea.IntersectsWith(r))
                    return true;
            return false;
        }

        // ---------- Преобразования Ghostscript ----------

        /// <summary>
        /// ЖИВАЯ проверка новых режимов на настоящем движке: аргументы можно сверить и
        /// глазами, а вот что цвет действительно ушёл и что битый файл действительно снова
        /// открылся — только запуском. Ghostscript в сборочной среде ставится раньше тестов,
        /// поэтому проверка выполняется и там, а не только на машине разработчика.
        /// </summary>
        private static void TestConvertLive()
        {
            if (!Ghostscript.Available)
            {
                // Без движка режимы обязаны быть безопасным бездействием, а не ошибкой.
                string tmp = Path.Combine(Path.GetTempPath(), "ExcelMergerConv_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tmp);
                try
                {
                    string p = Path.Combine(tmp, "файл.pdf");
                    MakeColorPdf(p);
                    long before = new FileInfo(p).Length;
                    AssertTrue(!PdfConvert.Apply(p, PdfConvertMode.Grayscale), "без движка — без изменений");
                    AssertEqual(before, new FileInfo(p).Length, "файл не тронут");
                }
                finally { try { Directory.Delete(tmp, true); } catch { } }
                return;
            }

            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerConv_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                // 1. Оттенки серого: на странице ярко-красный прямоугольник.
                string gray = Path.Combine(dir, "цветной.pdf");
                MakeColorPdf(gray);
                AssertTrue(HasColor(gray), "в исходнике есть цвет");
                AssertTrue(PdfConvert.Apply(gray, PdfConvertMode.Grayscale), "перевод в серое применён");
                AssertTrue(PdfCompression.LooksLikePdf(gray), "результат — валидный PDF");
                AssertTrue(!HasColor(gray), "после перевода цветных пикселей не осталось");
                AssertEqual(1, PdfPageCount(gray), "страница на месте");

                // 2. Восстановление: портим таблицу ссылок так, что файл перестаёт открываться.
                string broken = Path.Combine(dir, "битый.pdf");
                MakeColorPdf(broken);
                byte[] bytes = File.ReadAllBytes(broken);
                int at = LastIndexOf(bytes, System.Text.Encoding.ASCII.GetBytes("startxref"));
                AssertTrue(at > 0, "в исходнике есть таблица ссылок");
                var damaged = new List<byte>();
                damaged.AddRange(new List<byte>(bytes).GetRange(0, at));
                damaged.AddRange(System.Text.Encoding.ASCII.GetBytes("startxref\n999999999\n%%EOF\n"));
                File.WriteAllBytes(broken, damaged.ToArray());
                AssertThrowsAny("битый файл не открывается", delegate { PdfPageCount(broken); });

                AssertTrue(PdfConvert.Apply(broken, PdfConvertMode.Repair), "восстановление применено");
                AssertEqual(1, PdfPageCount(broken), "починенный файл снова открывается");

                // 3. Защищённый паролем файл движок прочитать не может, но выходит с НУЛЁМ и
                // оставляет годную по заголовку пустую заглушку в пару килобайт. До 1.17.9 она
                // проходила обе политики замены, и пользователь получал зелёное «Готово» с
                // пустым документом. Проверяем именно исход: преобразование НЕ применилось и
                // файл остался прежним, байт в байт.
                string locked = Path.Combine(dir, "запароленный.pdf");
                MakeProtectedPdf(locked);
                byte[] before = File.ReadAllBytes(locked);
                AssertTrue(!PdfConvert.Apply(locked, PdfConvertMode.Repair),
                    "защищённый файл не считается успешно преобразованным");
                AssertTrue(!PdfCompression.Compress(locked, CompressionLevel.Good),
                    "и сжатым тоже не считается");
                // Сравниваем длины как long: AssertEqual сверяет и тип, а Length у массива int.
                AssertEqual((long)before.Length, new FileInfo(locked).Length, "файл не подменён заглушкой");
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        /// <summary>Одностраничный PDF с паролем на открытие: движок его прочитать не сможет.</summary>
        private static void MakeProtectedPdf(string path)
        {
            using (var doc = new PdfDocument())
            {
                PdfPage page = doc.AddPage();
                using (XGraphics g = XGraphics.FromPdfPage(page))
                    g.DrawString("Проверка", new XFont("Times New Roman", 14), XBrushes.Black,
                        new XPoint(50, 100));
                doc.SecuritySettings.UserPassword = "1234";
                doc.Save(path);
            }
        }

        /// <summary>
        /// Коду возврата Ghostscript верить нельзя: на нечитаемом файле он выходит с нулём и
        /// оставляет заглушку. Признак отказа — «****» в потоке ошибок, и он не срабатывает
        /// на файлах, которые движок штатно чинит (там при -dQUIET поток пуст).
        /// </summary>
        private static void TestEngineSucceeded()
        {
            AssertTrue(GsRewrite.EngineSucceeded(0, ""), "штатная работа: ноль и пустой поток");
            AssertTrue(GsRewrite.EngineSucceeded(0, null), "поток может быть null");
            AssertTrue(!GsRewrite.EngineSucceeded(1, ""), "ненулевой код — отказ");
            AssertTrue(!GsRewrite.EngineSucceeded(-1, "timeout"), "таймаут — отказ");
            AssertTrue(!GsRewrite.EngineSucceeded(0,
                    "GPL Ghostscript 10.07.1:\n   **** This file requires a password for access."),
                "ноль с «****» в потоке — всё равно отказ");
            // Ровно то, что печатает движок на файле, который он чинит: при -dQUIET поток пуст,
            // поэтому «Восстановить повреждённый PDF» не должен ловить ложный отказ.
            AssertTrue(GsRewrite.EngineSucceeded(0, "\n"), "перевод строки не считается сообщением");
        }

        /// <summary>PDF из заданного числа пустых страниц ПО ТОЧНОМУ пути.</summary>
        private static void MakeEmptyPagesPdf(string path, int pages)
        {
            using (var doc = new PdfDocument())
            {
                for (int i = 0; i < pages; i++)
                    doc.AddPage();
                doc.Save(path);
            }
        }

        /// <summary>Одностраничный PDF с заведомо цветным прямоугольником.</summary>
        private static void MakeColorPdf(string path)
        {
            using (var doc = new PdfDocument())
            {
                PdfSharp.Pdf.PdfPage page = doc.AddPage();
                using (XGraphics g = XGraphics.FromPdfPage(page))
                    g.DrawRectangle(new XSolidBrush(XColors.Red), 40, 40, 300, 200);
                doc.Save(path);
            }
        }

        /// <summary>Есть ли на первой странице насыщенный цвет (рендер движком в растр).</summary>
        private static bool HasColor(string pdf)
        {
            string png = pdf + ".probe.png";
            string args = "-sDEVICE=png16m -r36 -dFirstPage=1 -dLastPage=1 -dNOPAUSE -dBATCH -dQUIET -dSAFER" +
                          " -sOutputFile=\"" + png + "\" \"" + pdf + "\"";
            string stderr;
            try
            {
                if (Ghostscript.Run(args, 60000, out stderr) != 0 || !File.Exists(png))
                    return false;
                using (var bmp = new System.Drawing.Bitmap(png))
                    for (int y = 0; y < bmp.Height; y += 2)
                        for (int x = 0; x < bmp.Width; x += 2)
                        {
                            System.Drawing.Color c = bmp.GetPixel(x, y);
                            int max = Math.Max(c.R, Math.Max(c.G, c.B));
                            int min = Math.Min(c.R, Math.Min(c.G, c.B));
                            if (max - min > 24) // насыщенность — признак цвета, а не серого
                                return true;
                        }
                return false;
            }
            finally { try { File.Delete(png); } catch { } }
        }

        private static int LastIndexOf(byte[] haystack, byte[] needle)
        {
            for (int i = haystack.Length - needle.Length; i >= 0; i--)
            {
                bool hit = true;
                for (int j = 0; j < needle.Length && hit; j++)
                    if (haystack[i + j] != needle[j])
                        hit = false;
                if (hit)
                    return i;
            }
            return -1;
        }

        private static void AssertThrowsAny(string what, Action action)
        {
            try { action(); }
            catch { return; }
            throw new Exception(what + ": ожидалось исключение, но его не было");
        }

        private static void TestConvertArguments()
        {
            string gray = PdfConvert.BuildArguments("in.pdf", "out.pdf", PdfConvertMode.Grayscale, null);
            AssertTrue(gray.Contains("-sColorConversionStrategy=Gray"), "перевод цветов в серое");
            // Модель устройства задавать НЕ надо: документация pdfwrite прямо просит не
            // указывать оба переключателя (с 9.11 стратегия выставляет модель сама), и на эту
            // связку заведён баг 693074 — часть изображений остаётся цветной. Пин на то, что
            // ключ не вернётся «для надёжности»: живая проверка ниже показывает, что одной
            // стратегии хватает на всех восьми пробах.
            AssertTrue(!gray.Contains("ProcessColorModel"), "модель устройства отдельно не задаётся");
            string repair = PdfConvert.BuildArguments("in.pdf", "out.pdf", PdfConvertMode.Repair, null);
            AssertTrue(!repair.Contains("ColorConversionStrategy"), "починка цвета не трогает");
            foreach (string args in new[] { gray, repair })
            {
                // Версия вывода 1.4: иначе наш же PdfSharp не откроет результат повторно.
                AssertTrue(args.Contains("-dCompatibilityLevel=1.4"), "версия вывода — 1.4");
                AssertTrue(args.Contains("-dSAFER"), "движок запускается в безопасном режиме");
                AssertTrue(args.Contains("\"in.pdf\"") && args.Contains("-sOutputFile=\"out.pdf\""),
                    "пути в кавычках (пробелы в путях)");
            }
            string bundled = PdfConvert.BuildArguments("in.pdf", "out.pdf", PdfConvertMode.Repair, @"C:\app\gs");
            AssertTrue(bundled.Contains(@"-I ""C:\app\gs\lib"""), "вшитому движку указываются его ресурсы");
        }

        private static void TestConvertShouldReplace()
        {
            // В отличие от сжатия, размер здесь не критерий: серый вариант бывает и больше.
            AssertTrue(PdfConvert.ShouldReplace(1000, 1200), "результат применяется, даже если больше");
            AssertTrue(PdfConvert.ShouldReplace(1000, 10), "результат применяется, если меньше");
            AssertTrue(!PdfConvert.ShouldReplace(1000, 0), "пустой вывод оригинал не заменяет");
            // А у сжатия — строго меньше, иначе в нём нет смысла.
            AssertTrue(!PdfCompression.ShouldReplace(1000, 1200), "сжатие не применяется, если файл вырос");
        }

        // ---------- Что не должно попасть на поток интерфейса ----------

        /// <summary>
        /// Дрейф-гард на заморозку окна. В 1.18.3 дважды получилось так, что работа с диском
        /// оказывалась на потоке интерфейса, и оба раза цена измерялась: разбор документа
        /// ради размеров страницы — 278 мс на 900-страничном файле, а проверка существования
        /// файла на отключённой сетевой шаре — около секунды НА КАЖДЫЙ путь, причём при
        /// запуске, до появления стартового экрана.
        ///
        /// Тест держит оба места асинхронными: обращения к диску обязаны стоять внутри
        /// фонового захода (RunWorker), а не в теле обработчика. Проверяется по исходнику —
        /// поймать это на живом окне можно только с недоступной сетью под рукой.
        /// </summary>
        private static void TestNoDiskWorkOnUiThread()
        {
            string root = RepoRoot();
            if (root == null) { AssertTrue(true, "исходники рядом не найдены — пропускаем"); return; }
            var offenders = new List<string>();

            // Недавние файлы: и чтение истории, и File.Exists — только в фоне.
            string start = File.ReadAllText(Path.Combine(root, "src", "StartForm.cs"));
            int build = start.IndexOf("private void BuildRecentFiles", StringComparison.Ordinal);
            AssertTrue(build > 0, "метод недавних файлов на месте");
            int worker = start.IndexOf("Ui.RunWorker", build, StringComparison.Ordinal);
            int exists = start.IndexOf("File.Exists", build, StringComparison.Ordinal);
            if (worker < 0 || exists < 0 || exists < worker)
                offenders.Add("StartForm: проверка недавних файлов вне фонового захода");

            // Пустая страница: размеры соседа читаются в фоне, а не при нажатии. Метод живёт
            // в общей базе — вставка доступна всюду, где сетка правится, а не в одном окне.
            string ops = File.ReadAllText(Path.Combine(root, "src", "PdfPageOrderFormBase.cs"));
            int add = ops.IndexOf("protected void AddBlankPage", StringComparison.Ordinal);
            AssertTrue(add > 0, "вставка пустой страницы на месте");
            int addWorker = ops.IndexOf("Ui.RunWorker", add, StringComparison.Ordinal);
            int sizeOf = ops.IndexOf("SizeOf(", add, StringComparison.Ordinal);
            if (addWorker < 0 || sizeOf < 0 || sizeOf < addWorker)
                offenders.Add("PdfOpsForm: размеры соседней страницы читаются на потоке интерфейса");

            AssertTrue(offenders.Count == 0, "работа с диском на потоке интерфейса: " +
                string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// Дрейф-гард на зависшее окно: опрос пароля идёт из обработчика, и если он бросит,
        /// окно останется в состоянии загрузки навсегда — кнопки и сетка заблокированы, и
        /// помогает только закрыть инструмент. Оба места, где пароль спрашивается, обязаны
        /// довести дело до конца в любом случае.
        /// </summary>
        private static void TestPasswordPromptCannotHangWindow()
        {
            string root = RepoRoot();
            if (root == null) { AssertTrue(true, "исходники рядом не найдены — пропускаем"); return; }
            foreach (string file in new[] { "PdfOrderedToolFormBase.cs", "PdfSingleDocFormBase.cs" })
            {
                string text = File.ReadAllText(Path.Combine(root, "src", file));
                int ask = text.IndexOf("AskPasswords(", StringComparison.Ordinal);
                AssertTrue(ask > 0, file + ": опрос пароля на месте");
                // Ищем try перед вызовом и catch после него в пределах того же метода.
                int guard = text.LastIndexOf("try", ask, StringComparison.Ordinal);
                int rescue = text.IndexOf("catch", ask, StringComparison.Ordinal);
                AssertTrue(guard > 0 && rescue > ask && rescue - ask < 1200,
                    file + ": опрос пароля не огорожен — сбой оставит окно заблокированным");
            }
        }

        // ---------- Недавние файлы ----------

        /// <summary>
        /// Список быстрого доступа: от новых к старым, без повторов и без исчезнувших файлов.
        /// Половина списка, которая не открывается, хуже отсутствия списка, поэтому проверка
        /// существования — часть правила, а не украшение.
        /// </summary>
        private static void TestRecentFiles()
        {
            var data = new OperationHistory.Data();
            data.Entries.Add(new HistoryEntry { Path = @"C:\нет\старый.pdf" });
            data.Entries.Add(new HistoryEntry { Path = @"C:\есть\первый.pdf" });
            data.Entries.Add(new HistoryEntry { Path = @"C:\есть\второй.pdf" });
            data.Entries.Add(new HistoryEntry { Path = @"C:\есть\первый.pdf" }); // тот же файл ещё раз
            Func<string, bool> exists = delegate(string p) { return p.StartsWith(@"C:\есть"); };

            List<HistoryEntry> recent = OperationHistory.RecentFiles(data, 5, exists);
            AssertEqual(2, recent.Count, "исчезнувший файл в список не попал");
            AssertEqual(@"C:\есть\первый.pdf", recent[0].Path, "самый свежий — первым, и без повтора");
            AssertEqual(@"C:\есть\второй.pdf", recent[1].Path, "затем предыдущий");

            AssertEqual(1, OperationHistory.RecentFiles(data, 1, exists).Count, "предел соблюдается");
            AssertEqual(0, OperationHistory.RecentFiles(data, 0, exists).Count, "нулевой предел — пустой список");
            AssertEqual(0, OperationHistory.RecentFiles(null, 3, exists).Count, "истории нет — списка нет");
            AssertEqual(0, OperationHistory.RecentFiles(data, 3, null).Count, "нечем проверить — не показываем");
            AssertEqual(0, OperationHistory.RecentFiles(new OperationHistory.Data(), 3, exists).Count,
                "пустая история — пустой список");

            // Записи без пути и бросающая проверка не должны ронять стартовый экран.
            var messy = new OperationHistory.Data();
            messy.Entries.Add(null);
            messy.Entries.Add(new HistoryEntry { Path = null });
            messy.Entries.Add(new HistoryEntry { Path = @"C:\есть\целый.pdf" });
            AssertEqual(1, OperationHistory.RecentFiles(messy, 3, exists).Count, "мусорные записи пропущены");
            AssertEqual(0, OperationHistory.RecentFiles(messy, 3,
                delegate { throw new UnauthorizedAccessException("нет доступа"); }).Count,
                "недоступный путь считается исчезнувшим, а не поводом упасть");
        }

        /// <summary>
        /// Значок карточки недавнего файла: сначала по инструменту, которым файл сделан (это
        /// точнее — файл могли переименовать), и лишь потом по расширению.
        /// </summary>
        private static void TestRecentCardGlyph()
        {
            AssertEqual(CardGlyph.PdfSplit, RecentCard.GlyphFor("split.done", "a.pdf"), "разделение");
            AssertEqual(CardGlyph.Excel, RecentCard.GlyphFor("excel.merged", "свод.xlsx"), "свод Excel");
            AssertEqual(CardGlyph.Pptx, RecentCard.GlyphFor("pptx.done", "a.pptx"), "презентация");
            AssertEqual(CardGlyph.Ocr, RecentCard.GlyphFor("word.done", "a.docx"), "PDF → Word");
            AssertEqual(CardGlyph.Tools, RecentCard.GlyphFor("ops.compressed", "a.pdf"), "прочие операции");
            AssertEqual(CardGlyph.Pdf, RecentCard.GlyphFor("pdf.merged", "a.pdf"), "объединение");
            // Операция неизвестна — узнаём по расширению.
            AssertEqual(CardGlyph.Excel, RecentCard.GlyphFor(null, @"C:\путь\книга.XLSX"), "по расширению .xlsx");
            AssertEqual(CardGlyph.Pptx, RecentCard.GlyphFor("", "доклад.pptx"), "по расширению .pptx");
            AssertEqual(CardGlyph.Ocr, RecentCard.GlyphFor("", "письмо.docx"), "по расширению .docx");
            AssertEqual(CardGlyph.Pdf, RecentCard.GlyphFor("", "скан.pdf"), "по расширению .pdf");
            AssertEqual(CardGlyph.Other, RecentCard.GlyphFor("", "заметки.txt"), "чужое расширение — общий значок");
            AssertEqual(CardGlyph.Other, RecentCard.GlyphFor(null, null), "нет ничего — тоже общий, а не падение");
        }

        /// <summary>
        /// Сколько карточек недавних показать в полосе данной ширины: не более трёх и не уже
        /// читаемого. Лучше две целых карточки, чем три огрызка, и лучше ничего, чем ряд
        /// нечитаемых имён.
        /// </summary>
        private static void TestRecentFit()
        {
            const int gap = 10, min = 150;
            AssertEqual(3, StartForm.RecentFit(700, gap, min, 3), "широкая полоса — все три");
            AssertEqual(3, StartForm.RecentFit(480, gap, min, 3), "ровно на три по минимуму");
            AssertEqual(2, StartForm.RecentFit(400, gap, min, 3), "на три не хватило — показываем две");
            AssertEqual(1, StartForm.RecentFit(200, gap, min, 3), "хватило на одну");
            AssertEqual(0, StartForm.RecentFit(100, gap, min, 3), "не хватило ни на одну — ничего");
            AssertEqual(2, StartForm.RecentFit(2000, gap, min, 2), "больше, чем есть записей, не показываем");
            AssertEqual(0, StartForm.RecentFit(0, gap, min, 3), "нулевая полоса");
            AssertEqual(0, StartForm.RecentFit(-50, gap, min, 3), "отрицательная ширина не ломает счёт");
            AssertEqual(0, StartForm.RecentFit(700, gap, min, 0), "записей нет — и карточек нет");
        }

        /// <summary>
        /// ЖИВАЯ проверка стартового экрана С КАРТОЧКАМИ НЕДАВНИХ на минимальном размере:
        /// карточки инструментов и карточки недавних не должны делить одни пиксели. Пустая
        /// история этого не поймает — там полосы недавних просто нет, — поэтому история
        /// заполняется настоящими файлами, а окно затем сжимается до упора.
        /// </summary>
        private static void TestStartScreenWithRecentSurvivesMinimum()
        {
            var offenders = new List<string>();
            string dir = Path.Combine(Path.GetTempPath(), "iwo_recent_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            InIsolatedSettings("iwo_recent_ui_", delegate
            {
                try
                {
                    // Три настоящих файла в истории: список показывает только существующие.
                    for (int i = 0; i < 3; i++)
                    {
                        string file = Path.Combine(dir, "результат" + i + ".pdf");
                        MakeEmptyPagesPdf(file, 1);
                        OperationHistory.Record("pdf.merged", file);
                    }
                    using (var form = new StartForm())
                    {
                        form.Show();
                        Pump(form, 1500);              // карточки строятся в фоне — даём им прийти
                        int cards = CountRecentCards(form);
                        if (cards == 0)
                        {
                            offenders.Add("карточки недавних не появились");
                            return;
                        }
                        foreach (System.Drawing.Size size in new[] { form.MinimumSize, new System.Drawing.Size(1200, 900) })
                        {
                            form.Size = size;
                            form.PerformLayout();
                            Pump(form, 200);
                            CheckControlsDoNotOverlap(form, "стартовый экран " + size.Width + "x" + size.Height, offenders);
                            CheckRecentDoNotCoverCards(form, size, offenders);
                        }
                        form.Close();
                    }
                }
                catch (Exception ex) { offenders.Add("раскладка: " + ex.GetType().Name + " " + ex.Message); }
            });
            try { Directory.Delete(dir, true); } catch { }
            AssertTrue(offenders.Count == 0, "стартовый экран с недавними: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>Сколько карточек недавних сейчас на окне.</summary>
        private static int CountRecentCards(System.Windows.Forms.Form form)
        {
            int n = 0;
            foreach (System.Windows.Forms.Control c in form.Controls)
                if (c.GetType().Name == "RecentCard")
                    n++;
            return n;
        }

        /// <summary>
        /// Карточка недавнего файла не должна перекрывать ни одну карточку инструмента.
        /// Сравниваем в координатах ОКНА: карточки инструментов лежат внутри панели уровня.
        /// </summary>
        private static void CheckRecentDoNotCoverCards(System.Windows.Forms.Form form, System.Drawing.Size size, List<string> offenders)
        {
            var recent = new List<System.Drawing.Rectangle>();
            foreach (System.Windows.Forms.Control c in form.Controls)
                if (c.GetType().Name == "RecentCard" && c.Visible)
                    recent.Add(form.RectangleToClient(c.RectangleToScreen(c.ClientRectangle)));
            foreach (System.Windows.Forms.Control panel in form.Controls)
            {
                if (!(panel is System.Windows.Forms.Panel) || !panel.Visible)
                    continue;
                foreach (System.Windows.Forms.Control card in panel.Controls)
                {
                    if (!card.Visible || card.GetType().Name != "ChoiceCard")
                        continue;
                        System.Drawing.Rectangle box = form.RectangleToClient(card.RectangleToScreen(card.ClientRectangle));
                    foreach (System.Drawing.Rectangle r in recent)
                        if (r.IntersectsWith(box))
                            offenders.Add("недавние наезжают на карточку инструмента при " +
                                size.Width + "x" + size.Height);
                }
            }
        }

        // ---------- Пустая страница ----------

        /// <summary>
        /// Формат вставляемого пустого листа берётся у соседа: лист чужого размера посреди
        /// документа читается как ошибка, а не как намерение. Взять не у чего — A4.
        /// </summary>
        private static void TestBlankSheetSize()
        {
            double w, h;
            BlankPages.SheetSize(new PdfPageInfo { WidthPt = 842, HeightPt = 1191 }, out w, out h);
            AssertEqual(842.0, w, "ширина как у соседней страницы");
            AssertEqual(1191.0, h, "высота как у соседней страницы");
            BlankPages.SheetSize(null, out w, out h);
            AssertEqual(BlankPages.A4WidthPt, w, "соседа нет — ширина A4");
            AssertEqual(BlankPages.A4HeightPt, h, "соседа нет — высота A4");
            BlankPages.SheetSize(new PdfPageInfo { WidthPt = 0, HeightPt = 500 }, out w, out h);
            AssertEqual(BlankPages.A4WidthPt, w, "бессмысленный размер соседа — тоже A4");
            BlankPages.SheetSize(new PdfPageInfo { WidthPt = 400, HeightPt = -1 }, out w, out h);
            AssertEqual(BlankPages.A4HeightPt, h, "отрицательная высота соседа — A4");
        }

        /// <summary>
        /// ЖИВАЯ проверка: пустой лист — настоящая одностраничная обёртка нужного размера,
        /// и он ведёт себя как обычная страница, то есть собирается в документ наравне с прочими.
        /// </summary>
        private static void TestBlankSheetLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerBlank_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string blank = Path.Combine(dir, "пустая.pdf");
                BlankPages.WriteSheet(blank, 842, 1191);   // A3
                AssertEqual(1, PdfPageProbe.PageCount(blank), "обёртка — ровно одна страница");
                List<PdfPageInfo> pages = PdfMergeService.LoadPages(blank);
                AssertEqual(842, (int)Math.Round(pages[0].WidthPt), "ширина листа записана");
                AssertEqual(1191, (int)Math.Round(pages[0].HeightPt), "высота листа записана");

                // Собирается в документ как обычная страница — ради этого обёртка и нужна.
                string doc = Path.Combine(dir, "с пустой.pdf");
                string source = Path.Combine(dir, "исходник.pdf");
                MakeEmptyPagesPdf(source, 2);
                PdfMergeService.Merge(new List<PdfPageRef> {
                    new PdfPageRef { SourcePath = source, PageIndex = 0 },
                    new PdfPageRef { SourcePath = blank, PageIndex = 0 },
                    new PdfPageRef { SourcePath = source, PageIndex = 1 } }, doc);
                AssertEqual(3, PdfPageProbe.PageCount(doc), "пустая страница встала между своими");
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        // ---------- Предпросмотр печати ----------

        /// <summary>
        /// Сколько листов рисует предпросмотр. Ограничение не косметическое: контрол строит
        /// показанные страницы СРАЗУ и держит их растры в памяти, поэтому «показать всё» на
        /// сотне листов — это сотня растров и заморозка окна.
        /// </summary>
        private static void TestPrintPreviewPageLimit()
        {
            AssertEqual(0, PrintPreviewForm.PagesToShow(0), "печатать нечего — и показывать нечего");
            AssertEqual(0, PrintPreviewForm.PagesToShow(-3), "отрицательное число страниц не ломает счёт");
            AssertEqual(1, PrintPreviewForm.PagesToShow(1), "одна страница показывается целиком");
            AssertEqual(PrintPreviewForm.PreviewPages - 1,
                PrintPreviewForm.PagesToShow(PrintPreviewForm.PreviewPages - 1),
                "меньше предела — показываем всё");
            AssertEqual(PrintPreviewForm.PreviewPages, PrintPreviewForm.PagesToShow(PrintPreviewForm.PreviewPages),
                "ровно предел — показываем всё");
            AssertEqual(PrintPreviewForm.PreviewPages, PrintPreviewForm.PagesToShow(500),
                "больше предела — ровно предел, а не всё задание");
        }

        /// <summary>
        /// ЖИВАЯ проверка задания печати: предпросмотр и печать рисуют ОДИН И ТОТ ЖЕ документ,
        /// иначе показанное разошлось бы с напечатанным. Заодно — что задание проходится
        /// дважды: предпросмотр листает его по кругу, и без сброса счётчика второй показ
        /// выдал бы хвост списка вместо начала.
        /// </summary>
        private static void TestPrintDocumentReusable()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerPrint_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = Path.Combine(dir, "печать.pdf");
                MakeEmptyPagesPdf(src, 3);
                var pages = new List<PdfPageRef>();
                for (int i = 0; i < 3; i++)
                    pages.Add(new PdfPageRef { SourcePath = src, PageIndex = i });

                IDisposable renderer;
                using (System.Drawing.Printing.PrintDocument doc =
                    PdfPrintService.CreateDocument(pages, null, null, null, out renderer))
                using (renderer)
                {
                    AssertTrue(doc != null, "задание печати собрано");
                    // На принтер ничего не уходит: считаем листы так же, как предпросмотр.
                    int first = CountPreviewPages(doc);
                    int second = CountPreviewPages(doc);
                    AssertEqual(3, first, "листов столько же, сколько страниц");
                    AssertEqual(first, second, "повторный обход даёт то же самое");
                }

                // Пустой набор — понятная ошибка, а не попытка печатать пустоту.
                AssertThrowsAny("печатать нечего", delegate
                {
                    IDisposable none;
                    PdfPrintService.CreateDocument(new List<PdfPageRef>(), null, null, null, out none);
                });
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        /// <summary>Прогнать задание так же, как предпросмотр, и посчитать листы.</summary>
        private static int CountPreviewPages(System.Drawing.Printing.PrintDocument doc)
        {
            var controller = new System.Drawing.Printing.PreviewPrintController();
            System.Drawing.Printing.PrintController saved = doc.PrintController;
            doc.PrintController = controller;
            try
            {
                doc.Print();
                return controller.GetPreviewPageInfo().Length;
            }
            finally { doc.PrintController = saved; }
        }

        // ---------- Командная строка для PDF ----------

        /// <summary>
        /// Разбор строки запуска. Здесь важна не только удача, но и отказ: сценарий-обёртка
        /// различает «сделано», «не получилось» и «неверно вызвано» по коду возврата, и
        /// опечатка в ключе обязана давать именно третье, а не молча делать не то.
        /// </summary>
        private static void TestPdfCliParse()
        {
            AssertTrue(PdfCli.IsCommand(new[] { "--merge", "o.pdf" }), "команда опознана");
            AssertTrue(PdfCli.IsCommand(new[] { "--TO-IMAGE", "a.pdf", "d" }), "регистр ключа не важен");
            AssertTrue(!PdfCli.IsCommand(new[] { "--cli", "a", "b" }), "режим Excel — не наша команда");
            AssertTrue(!PdfCli.IsCommand(new string[0]), "пустой запуск — не команда");
            AssertTrue(!PdfCli.IsCommand(null), "и null тоже");

            PdfCliCommand merge = PdfCli.Parse(new[] { "--merge", "out.pdf", "a.pdf", "b.pdf" });
            AssertEqual(null, merge.Error, "разбор объединения без ошибок");
            AssertEqual(PdfCliKind.Merge, merge.Kind, "это объединение");
            AssertEqual("out.pdf", merge.Output, "первый путь — результат");
            AssertEqual(2, merge.Inputs.Count, "остальные — исходники");

            PdfCliCommand levelled = PdfCli.Parse(new[] { "--merge", "o.pdf", "a.pdf", "--level", "normal" });
            AssertEqual(CompressionLevel.Small, levelled.Level, "уровень читается по имени");
            AssertEqual(1, levelled.Inputs.Count, "ключ не попал в список файлов");

            PdfCliCommand split = PdfCli.Parse(new[] { "--split", "a.pdf", "dir", "--every", "3" });
            AssertEqual(PdfCliSplitMode.Every, split.SplitMode, "режим — каждые N");
            AssertEqual(3, split.Every, "N прочитано");
            AssertEqual(PdfCliSplitMode.Bookmarks,
                PdfCli.Parse(new[] { "--split", "a.pdf", "dir", "--bookmarks" }).SplitMode, "режим — по закладкам");
            AssertEqual(PdfCliSplitMode.Ranges,
                PdfCli.Parse(new[] { "--split", "a.pdf", "dir", "--ranges", "1-3,5" }).SplitMode, "режим — диапазоны");

            PdfCliCommand images = PdfCli.Parse(new[] { "--to-image", "a.pdf", "dir", "--dpi", "300", "--format", "jpg" });
            AssertEqual(300, images.Dpi, "разрешение прочитано");
            AssertEqual(ImageExportFormat.Jpeg, images.Format, "формат прочитан");
            AssertEqual(ImageExportFormat.Jpeg,
                PdfCli.Parse(new[] { "--to-image", "a.pdf", "d", "--format", "jpeg" }).Format, "jpeg = jpg");

            // Сжатие без уровня бессмысленно — берём средний, а не «ничего не делать».
            AssertEqual(CompressionLevel.Good, PdfCli.Parse(new[] { "--compress", "a.pdf", "b.pdf" }).Level,
                "у сжатия уровень по умолчанию рабочий");

            // Отказы: каждый обязан объяснить, что не так.
            foreach (string[] bad in new[]
            {
                new[] { "--merge" },                                   // нет ни результата, ни входа
                new[] { "--merge", "only-out.pdf" },                   // нет входа
                new[] { "--extract", "a.pdf", "1-3" },                 // нет результата
                new[] { "--split", "a.pdf", "dir" },                   // не сказано, как делить
                new[] { "--split", "a.pdf", "dir", "--every", "0" },   // ноль страниц в части
                new[] { "--split", "a.pdf", "dir", "--every" },        // ключ без значения
                new[] { "--compress", "a.pdf", "b.pdf", "--level", "быстро" }, // нет такого уровня
                new[] { "--to-image", "a.pdf", "d", "--dpi", "-5" },   // отрицательное разрешение
                new[] { "--to-image", "a.pdf", "d", "--format", "tiff" }, // нет такого формата
                new[] { "--merge", "o.pdf", "a.pdf", "--wat" },        // неизвестный ключ
                new[] { "--nonsense", "a" }                            // неизвестная команда
            })
                AssertTrue(!string.IsNullOrEmpty(PdfCli.Parse(bad).Error),
                    "отказ объяснён: " + string.Join(" ", bad));

            // Справка — тоже команда, иначе «--help» открывал бы окно вместо ответа: строка
            // запуска с неизвестным первым словом уводит приложение в обычный запуск с
            // интерфейсом, и человек в консоли не получает ничего. Так и было до 1.18.3.
            foreach (string flag in new[] { "--help", "-h", "/?", "--HELP" })
            {
                AssertTrue(PdfCli.IsCommand(new[] { flag }), flag + " — команда, а не повод открыть окно");
                PdfCliCommand help = PdfCli.Parse(new[] { flag });
                AssertEqual(null, help.Error, flag + " разбирается без ошибок");
                AssertEqual(PdfCliKind.Help, help.Kind, flag + " — это справка");
                AssertEqual(PdfCli.Ok, PdfCli.Execute(help, delegate { }), flag + " — успех, а не отказ");
            }

            AssertTrue(!string.IsNullOrEmpty(PdfCli.Parse(null).Error), "пустой запуск — отказ, а не падение");
            AssertEqual(PdfCli.BadUsage, PdfCli.Execute(PdfCli.Parse(new[] { "--merge" }), delegate { }),
                "неверный вызов — код 2");
            AssertTrue(PdfCli.Usage().Contains("--merge") && PdfCli.Usage().Contains("Exit codes"),
                "справка перечисляет команды и коды возврата");
        }

        private static void TestPdfCliLevels()
        {
            CompressionLevel level;
            AssertTrue(PdfCli.TryParseLevel("none", out level) && level == CompressionLevel.None, "none");
            AssertTrue(PdfCli.TryParseLevel("VeryGood", out level) && level == CompressionLevel.VeryGood, "verygood");
            AssertTrue(PdfCli.TryParseLevel("good", out level) && level == CompressionLevel.Good, "good");
            AssertTrue(PdfCli.TryParseLevel("normal", out level) && level == CompressionLevel.Small, "normal");
            AssertTrue(!PdfCli.TryParseLevel("small", out level), "внутреннее имя уровня наружу не торчит");
            AssertTrue(!PdfCli.TryParseLevel("", out level) && !PdfCli.TryParseLevel(null, out level), "пусто — не уровень");
        }

        /// <summary>
        /// ЖИВОЙ прогон команд: они обязаны делать ровно то же, что кнопки, потому что зовут
        /// те же сервисы. Проверяем сквозь: собрать, извлечь, разрезать, вынуть текст.
        /// </summary>
        private static void TestPdfCliLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerCli_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = Path.Combine(dir, "исходник.pdf");
                MakeEmptyPagesPdf(src, 4);
                var log = new List<string>();
                Action<string> say = delegate(string line) { log.Add(line); };

                string merged = Path.Combine(dir, "объединено.pdf");
                AssertEqual(PdfCli.Ok, PdfCli.Execute(PdfCli.Parse(new[] { "--merge", merged, src, src }), say),
                    "объединение выполнено");
                AssertEqual(8, PdfPageProbe.PageCount(merged), "восемь страниц из двух копий");

                string part = Path.Combine(dir, "часть.pdf");
                AssertEqual(PdfCli.Ok, PdfCli.Execute(PdfCli.Parse(new[] { "--extract", merged, "2-3", part }), say),
                    "извлечение выполнено");
                AssertEqual(2, PdfPageProbe.PageCount(part), "извлечены две страницы");

                string outDir = Path.Combine(dir, "части");
                AssertEqual(PdfCli.Ok, PdfCli.Execute(PdfCli.Parse(new[] { "--split", merged, outDir, "--every", "3" }), say),
                    "разделение выполнено");
                AssertEqual(3, Directory.GetFiles(outDir, "*.pdf").Length, "восемь страниц по три — три файла");

                string txt = Path.Combine(dir, "текст.txt");
                AssertEqual(PdfCli.Ok, PdfCli.Execute(PdfCli.Parse(new[] { "--to-text", merged, txt }), say),
                    "извлечение текста выполнено");
                AssertTrue(File.Exists(txt), "файл текста создан");

                // Исходник не тронут ни одной командой — общее правило приложения.
                AssertEqual(4, PdfPageProbe.PageCount(src), "исходник остался прежним");
                AssertTrue(log.Count >= 4, "каждая команда отчиталась строкой");

                // Несуществующий файл — честный отказ с кодом «не получилось», а не падение.
                AssertEqual(PdfCli.Failed,
                    PdfCli.Execute(PdfCli.Parse(new[] { "--extract", Path.Combine(dir, "нет.pdf"), "1", part }), say),
                    "нет исходника — код 1");
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        // ---------- Пароли к защищённым PDF ----------

        /// <summary>
        /// Реестр паролей: один на процесс, поэтому важно, что тот же файл под другим
        /// написанием пути — это тот же файл, а очистка действительно очищает.
        /// </summary>
        private static void TestPasswordRegistry()
        {
            PdfPasswords.Clear();
            try
            {
                string path = Path.Combine(Path.GetTempPath(), "пример.pdf");
                AssertEqual(null, PdfPasswords.For(path), "пароля нет — null, а не пустая строка");
                PdfPasswords.Remember(path, "тайна");
                AssertEqual("тайна", PdfPasswords.For(path), "запомнили и достали");
                AssertEqual("тайна", PdfPasswords.For(path.ToUpperInvariant()),
                    "регистр пути не делает файл другим");
                AssertEqual("тайна", PdfPasswords.For(Path.Combine(Path.GetTempPath(), ".", "пример.pdf")),
                    "другое написание того же пути — тот же файл");
                PdfPasswords.Remember(path, null);
                AssertEqual(null, PdfPasswords.For(path), "пустой пароль стирает запись");
                PdfPasswords.Remember(path, "снова");
                PdfPasswords.Clear();
                AssertEqual(null, PdfPasswords.For(path), "очистка убирает всё");
                AssertEqual(null, PdfPasswords.For(null), "пустой путь не роняет реестр");
                PdfPasswords.Remember(null, "x");   // не должно бросать
            }
            finally { PdfPasswords.Clear(); }
        }

        /// <summary>
        /// Отличить «нужен пароль» от «файл битый». Ошибиться в эту сторону не страшно —
        /// лишний вопрос человек закроет; страшно обратное: молчаливый отказ на защищённом
        /// файле выглядит поломкой программы. Фразы взяты у самих библиотек.
        /// </summary>
        private static void TestPasswordDetection()
        {
            AssertTrue(PdfPasswords.LooksPasswordProtected(
                new Exception("A password is required to open the PDF document.")),
                "PdfSharp: нужен пароль");
            AssertTrue(PdfPasswords.LooksPasswordProtected(
                new Exception("The specified password is invalid.")),
                "PdfSharp: пароль неверен — тоже повод спросить");
            AssertTrue(PdfPasswords.LooksPasswordProtected(
                new PdfDocumentEncryptedLookalike("документ зашифрован")),
                "PdfPig: тип исключения говорит сам за себя");
            AssertTrue(PdfPasswords.LooksPasswordProtected(
                new Exception("внешняя", new Exception("password required"))),
                "причина во вложенном исключении тоже считается");
            AssertTrue(!PdfPasswords.LooksPasswordProtected(new Exception("Unexpected token in xref table")),
                "битый файл — не повод спрашивать пароль");
            AssertTrue(!PdfPasswords.LooksPasswordProtected(null), "исключения нет — и вопроса нет");
        }

        /// <summary>
        /// Ширина окна пароля считается по самой длинной СТРОКЕ подписи: она многострочная
        /// (имя файла и предупреждение об отказе), и мерка целиком растянула бы окно на весь
        /// экран — ровно та ошибка, из-за которой примечание «бета» когда-то обрезалось.
        /// </summary>
        private static void TestPasswordDialogWidth()
        {
            using (var font = new System.Drawing.Font("Segoe UI", 9f))
            {
                int oneLine = PasswordPromptDialog.WidestLine("Короткая строка", font);
                int twoLines = PasswordPromptDialog.WidestLine("Короткая строка\nЕщё одна короткая", font);
                AssertTrue(twoLines > 0, "ширина считается");
                AssertTrue(twoLines < oneLine * 2, "строки не складываются в одну длину");
                int longer = PasswordPromptDialog.WidestLine(
                    "коротко\nсамая длинная строка подписи в этом окне", font);
                AssertTrue(longer > oneLine, "берётся именно самая длинная строка");
                AssertEqual(PasswordPromptDialog.WidestLine("строка\r\nдругая", font),
                    PasswordPromptDialog.WidestLine("строка\nдругая", font),
                    "перевод строки в двух видах даёт одно и то же");
                AssertEqual(0, PasswordPromptDialog.WidestLine("", font), "пустая подпись — нулевая ширина");
                AssertEqual(0, PasswordPromptDialog.WidestLine("текст", null), "без шрифта мерить нечем");
            }
        }

        /// <summary>Имитация типа PdfPig: опознаём шифрование в том числе по имени типа.</summary>
        private sealed class PdfDocumentEncryptedLookalike : Exception
        {
            public PdfDocumentEncryptedLookalike(string message) : base(message) { }
        }

        /// <summary>
        /// ЖИВАЯ сквозная проверка: защищённый файл без пароля не открывается ни одним из
        /// наших путей, а с паролем — открывается всеми. До 1.18.3 такой документ нельзя было
        /// даже показать в сетке: инструмент отказывал ровно тогда, когда был нужен.
        /// </summary>
        private static void TestPasswordLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerPwd_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            PdfPasswords.Clear();
            try
            {
                string locked = Path.Combine(dir, "под паролем.pdf");
                MakeProtectedPdf(locked);   // пароль «1234»

                // 1. Без пароля — отказ, и он опознаётся именно как «нужен пароль».
                Exception denied = null;
                try { PdfMergeService.LoadPages(locked); }
                catch (Exception ex) { denied = ex; }
                AssertTrue(denied != null, "без пароля файл не открывается");
                AssertTrue(PdfPasswords.LooksPasswordProtected(denied),
                    "и отказ опознан как парольный: " + denied.Message);
                AssertEqual(PdfPageProbe.Unreadable, PdfPageProbe.PageCount(locked),
                    "PdfPig без пароля тоже не читает");

                // 2. С паролем открываются все пути: страницы, счётчик, состав.
                PdfPasswords.Remember(locked, "1234");
                AssertEqual(1, PdfMergeService.LoadPages(locked).Count, "PdfSharp читает с паролем");
                AssertEqual(1, PdfPageProbe.PageCount(locked), "PdfPig читает с паролем");
                AssertTrue(PdfContentProfile.ImageShare(locked) >= 0, "состав документа тоже читается");

                // 3. Полезная работа: страницы защищённого файла собираются в новый документ.
                string result = Path.Combine(dir, "итог.pdf");
                PdfMergeService.Merge(new List<PdfPageRef> {
                    new PdfPageRef { SourcePath = locked, PageIndex = 0 } }, result);
                AssertEqual(1, PdfPageProbe.PageCount(result), "результат собран и открывается без пароля");

                // 4. Неверный пароль — снова отказ, и снова опознанный как парольный.
                PdfPasswords.Remember(locked, "не тот");
                Exception wrong = null;
                try { PdfMergeService.LoadPages(locked); }
                catch (Exception ex) { wrong = ex; }
                AssertTrue(wrong != null && PdfPasswords.LooksPasswordProtected(wrong),
                    "неверный пароль — повод спросить ещё раз");
            }
            finally
            {
                PdfPasswords.Clear();
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        // ---------- Состав документа (почему не сжалось) ----------

        /// <summary>
        /// Когда советовать другой уровень. Случай из жизни: скан на 1,6 МБ, изображения —
        /// 99,8% файла, уровень «Очень хорошо» их не трогает и потому не даёт ничего, а
        /// «Хорошо» снимает 64%. Сказать про такой файл «уже оптимизирован» — обмануть.
        /// </summary>
        private static void TestContentProfileAdvice()
        {
            AssertTrue(PdfContentProfile.ShouldSuggestDownsampling(CompressionLevel.VeryGood, 0.998),
                "документ из картинок и уровень без пересчёта — советуем другой");
            AssertTrue(PdfContentProfile.ShouldSuggestDownsampling(CompressionLevel.None, 0.998),
                "то же и для уровня «без сжатия»");
            AssertTrue(!PdfContentProfile.ShouldSuggestDownsampling(CompressionLevel.Good, 0.998),
                "уровень уже пересчитывает — советовать нечего, файл правда не ужался");
            AssertTrue(!PdfContentProfile.ShouldSuggestDownsampling(CompressionLevel.Small, 0.998),
                "и самый сильный уровень тоже");
            AssertTrue(!PdfContentProfile.ShouldSuggestDownsampling(CompressionLevel.VeryGood, 0.02),
                "текстовый документ: картинок мало, пересчёт не поможет");
            AssertTrue(!PdfContentProfile.ShouldSuggestDownsampling(CompressionLevel.VeryGood, PdfContentProfile.Unknown),
                "долю выяснить не удалось — молчим, а не гадаем");
            // Ровно на пороге советуем: половина файла в картинках — уже повод.
            AssertTrue(PdfContentProfile.ShouldSuggestDownsampling(CompressionLevel.VeryGood, 0.5),
                "половина файла в изображениях — порог сработал");
            AssertTrue(!PdfContentProfile.ShouldSuggestDownsampling(CompressionLevel.VeryGood, 0.49),
                "чуть ниже порога — молчим");
        }

        /// <summary>
        /// ЖИВАЯ проверка доли изображений: документ, собранный из картинок, обязан показать
        /// долю около единицы, а документ из одного текста — около нуля. Без этого правило
        /// «советовать другой уровень» опиралось бы на число, которое никто не измерял.
        /// </summary>
        private static void TestContentProfileLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerShare_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string scan = Path.Combine(dir, "скан.pdf");
                MakeImagePdf(scan, 2);                      // страницы целиком из фотошума
                double scanShare = PdfContentProfile.ImageShare(scan);
                AssertTrue(scanShare > 0.8, "документ из картинок: доля изображений " + scanShare.ToString("F3"));

                string text = Path.Combine(dir, "текст.pdf");
                using (var doc = new PdfDocument())
                    for (int i = 0; i < 3; i++)
                    {
                        PdfSharp.Pdf.PdfPage page = doc.AddPage();
                        using (XGraphics g = XGraphics.FromPdfPage(page))
                            g.DrawString("Обычный текстовый документ без картинок",
                                new XFont("Times New Roman", 12), XBrushes.Black, new XPoint(50, 100));
                        if (i == 2) doc.Save(text);
                    }
                double textShare = PdfContentProfile.ImageShare(text);
                AssertTrue(textShare < 0.1, "текстовый документ: доля изображений " + textShare.ToString("F3"));

                // Сюда же — связка с советом: именно так решает окно «Прочих операций».
                AssertTrue(PdfContentProfile.ShouldSuggestDownsampling(CompressionLevel.VeryGood, scanShare),
                    "скану советуем уровень с пересчётом");
                AssertTrue(!PdfContentProfile.ShouldSuggestDownsampling(CompressionLevel.VeryGood, textShare),
                    "текстовому документу — нет");

                AssertEqual(PdfContentProfile.Unknown, PdfContentProfile.ImageShare(Path.Combine(dir, "нет.pdf")),
                    "несуществующий файл — «выяснить не удалось», а не исключение");
                File.WriteAllText(Path.Combine(dir, "мусор.pdf"), "не PDF");
                AssertEqual(PdfContentProfile.Unknown, PdfContentProfile.ImageShare(Path.Combine(dir, "мусор.pdf")),
                    "мусор вместо PDF — тоже «не удалось»");
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        // ---------- Оглавление (закладки) ----------

        /// <summary>Закладка для проб: заголовок, страница, уровень.</summary>
        private static PdfBookmark Mark(string title, int page, int level)
        {
            return new PdfBookmark { Title = title, PageIndex = page, Level = level, Opened = true };
        }

        /// <summary>
        /// Пересчёт оглавления под новый состав документа — сердце переноса, и вся его
        /// тонкость в том, что делать с закладкой, чья страница не попала в результат.
        /// Выбрасывать вместе с потомками нельзя: удаление одного раздела уносило бы
        /// половину оглавления. Поэтому потомки поднимаются на освободившийся уровень.
        /// </summary>
        private static void TestBookmarksRemap()
        {
            var map = new Dictionary<int, int> { { 0, 0 }, { 2, 1 }, { 5, 2 } };

            // Простой случай: уцелевшие получают новые номера, порядок обхода сохраняется.
            List<PdfBookmark> kept = PdfBookmarks.Remap(
                new List<PdfBookmark> { Mark("A", 0, 0), Mark("B", 2, 0), Mark("C", 5, 0) }, map);
            AssertEqual(3, kept.Count, "уцелели все три");
            AssertEqual("A,0;B,1;C,2", Describe(kept), "номера страниц пересчитаны");

            // Закладка на выпавшую страницу уходит, но её потомок остаётся и поднимается.
            List<PdfBookmark> orphan = PdfBookmarks.Remap(
                new List<PdfBookmark> { Mark("Глава", 1, 0), Mark("Раздел", 2, 1) }, map);
            AssertEqual(1, orphan.Count, "выпавшая закладка не переносится");
            AssertEqual("Раздел,1", Describe(orphan), "а её потомок остаётся");
            AssertEqual(0, orphan[0].Level, "и поднимается на освободившийся уровень");

            // Два выпавших предка подряд — потомок поднимается на оба уровня.
            List<PdfBookmark> deep = PdfBookmarks.Remap(
                new List<PdfBookmark> { Mark("X", 1, 0), Mark("Y", 3, 1), Mark("Z", 5, 2) }, map);
            AssertEqual("Z,2", Describe(deep), "уцелел только самый глубокий");
            AssertEqual(0, deep[0].Level, "и он оказался на верхнем уровне");

            // Выход из ветви выпавшего предка не должен «съедать» уровень у следующей ветви.
            List<PdfBookmark> siblings = PdfBookmarks.Remap(
                new List<PdfBookmark> { Mark("A", 0, 0), Mark("B", 1, 1), Mark("C", 2, 0), Mark("D", 5, 1) }, map);
            AssertEqual("A,0;C,1;D,2", Describe(siblings), "соседняя ветвь уцелела целиком");
            AssertEqual(0, siblings[0].Level, "A на верхнем уровне");
            AssertEqual(0, siblings[1].Level, "C тоже на верхнем — она соседка A, а не потомок B");
            AssertEqual(1, siblings[2].Level, "D осталась вложенной в C");

            // Уровень не прыгает больше чем на ступень: иначе запись построит рваное дерево.
            List<PdfBookmark> jump = PdfBookmarks.Remap(
                new List<PdfBookmark> { Mark("A", 0, 0), Mark("B", 2, 5) }, map);
            AssertEqual(1, jump[1].Level, "вложенность растёт по одной ступени: " + jump[1].Level);

            // Пусто на входе и мусор не должны ронять пересчёт.
            AssertEqual(0, PdfBookmarks.Remap(null, map).Count, "нет закладок — нет и результата");
            AssertEqual(0, PdfBookmarks.Remap(new List<PdfBookmark> { Mark("A", 0, 0) }, null).Count,
                "нет карты — переносить некуда");
            AssertEqual(0, PdfBookmarks.Remap(new List<PdfBookmark> { null }, map).Count,
                "пустая запись пропускается, а не роняет");
        }

        /// <summary>«Заголовок,страница» через точку с запятой — компактная запись для сверки.</summary>
        private static string Describe(IList<PdfBookmark> marks)
        {
            var parts = new List<string>();
            foreach (PdfBookmark m in marks)
                parts.Add(m.Title + "," + m.PageIndex);
            return string.Join(";", parts.ToArray());
        }

        private static void TestBookmarksRangeMap()
        {
            Dictionary<int, int> map = PdfBookmarks.RangeMap(3, 5);
            AssertEqual(3, map.Count, "три страницы диапазона");
            AssertEqual(0, map[3], "первая страница диапазона становится первой");
            AssertEqual(2, map[5], "последняя — последней");
            AssertTrue(!map.ContainsKey(2) && !map.ContainsKey(6), "соседние страницы в карту не попадают");
            AssertEqual(1, PdfBookmarks.RangeMap(4, 4).Count, "диапазон из одной страницы");
            AssertEqual(0, PdfBookmarks.RangeMap(5, 4).Count, "пустой диапазон — пустая карта");
        }

        /// <summary>
        /// ЖИВАЯ проверка на настоящих файлах: до 1.18.3 объединение и разделение переносили
        /// страницы и молча теряли оглавление — файл открывался, выглядел целым, а навигация
        /// исчезала. Проверяем оба пути и то, что закладки ведут на СВОИ страницы после
        /// перестановки, а не просто присутствуют.
        /// </summary>
        private static void TestBookmarksLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerMarks_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = Path.Combine(dir, "с оглавлением.pdf");
                MakeBookmarkedPdf(src);
                AssertEqual("Глава А,0;Раздел А.1,1;Глава Б,2", ReadMarks(src), "исходник: три закладки");

                // 1. Объединение в исходном порядке — оглавление на месте.
                string merged = Path.Combine(dir, "объединено.pdf");
                var order = new List<PdfPageRef>();
                for (int i = 0; i < 4; i++)
                    order.Add(new PdfPageRef { SourcePath = src, PageIndex = i });
                PdfMergeService.Merge(order, merged);
                AssertEqual("Глава А,0;Раздел А.1,1;Глава Б,2", ReadMarks(merged), "объединение сохранило оглавление");

                // 2. Страницы переставлены задом наперёд — закладки обязаны переехать за ними.
                string reversed = Path.Combine(dir, "наоборот.pdf");
                var back = new List<PdfPageRef>();
                for (int i = 3; i >= 0; i--)
                    back.Add(new PdfPageRef { SourcePath = src, PageIndex = i });
                PdfMergeService.Merge(back, reversed);
                AssertEqual("Глава А,3;Раздел А.1,2;Глава Б,1", ReadMarks(reversed),
                    "закладки указывают на новые места своих страниц");

                // 3. Часть страниц выброшена — закладка выпавшей страницы уходит, прочие живут.
                string subset = Path.Combine(dir, "подмножество.pdf");
                PdfSplitService.Extract(src, new List<int> { 0, 2 }, subset);
                AssertEqual("Глава А,0;Глава Б,1", ReadMarks(subset), "осталось оглавление уцелевших страниц");

                // 4. Разделение по диапазонам идёт другим путём в коде — проверяем и его.
                List<string> parts = PdfSplitService.SplitByRanges(src,
                    new List<PageRange> { new PageRange(0, 1), new PageRange(2, 3) }, dir, "часть");
                AssertEqual(2, parts.Count, "две части");
                AssertEqual("Глава А,0;Раздел А.1,1", ReadMarks(parts[0]), "оглавление первой части");
                AssertEqual("Глава Б,0", ReadMarks(parts[1]), "оглавление второй части");

                // 5. Два источника подряд — оглавления идут одно за другим со сдвигом.
                string twice = Path.Combine(dir, "дважды.pdf");
                var both = new List<PdfPageRef>();
                for (int i = 0; i < 4; i++) both.Add(new PdfPageRef { SourcePath = src, PageIndex = i });
                string second = Path.Combine(dir, "второй.pdf");
                MakeBookmarkedPdf(second);
                for (int i = 0; i < 4; i++) both.Add(new PdfPageRef { SourcePath = second, PageIndex = i });
                PdfMergeService.Merge(both, twice);
                AssertEqual("Глава А,0;Раздел А.1,1;Глава Б,2;Глава А,4;Раздел А.1,5;Глава Б,6",
                    ReadMarks(twice), "оглавления обоих файлов, второе со сдвигом");
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        /// <summary>Четыре страницы, две закладки верхнего уровня и одна вложенная.</summary>
        private static void MakeBookmarkedPdf(string path)
        {
            using (var doc = new PdfDocument())
            {
                for (int i = 0; i < 4; i++)
                {
                    PdfSharp.Pdf.PdfPage page = doc.AddPage();
                    using (XGraphics g = XGraphics.FromPdfPage(page))
                        g.DrawString("Страница " + (i + 1), new XFont("Times New Roman", 14),
                            XBrushes.Black, new XPoint(50, 100));
                }
                PdfSharp.Pdf.PdfOutline chapter = doc.Outlines.Add("Глава А", doc.Pages[0], true);
                chapter.Outlines.Add("Раздел А.1", doc.Pages[1], true);
                doc.Outlines.Add("Глава Б", doc.Pages[2], true);
                doc.Save(path);
            }
        }

        /// <summary>Оглавление файла в виде «заголовок,страница» — плоско, в порядке обхода.</summary>
        private static string ReadMarks(string path)
        {
            using (PdfDocument doc = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                return Describe(PdfBookmarks.Read(doc));
        }

        // ---------- Проверка результата прогона Ghostscript ----------

        /// <summary>
        /// Инвариант конвейера: документ обязан остаться документом. Правило одно на все три
        /// режима, и вся его тонкость — в случае «исходник не читается нами самими»
        /// (починка: он битый по условию; экзотика, понятная движку, но не PdfPig). Там
        /// сравнивать не с чем, и строгое сравнение дало бы ложный отказ на ровном месте.
        /// </summary>
        private static void TestPagesKept()
        {
            AssertTrue(PdfPageProbe.PagesKept(10, 10), "столько же страниц — уцелел");
            AssertTrue(!PdfPageProbe.PagesKept(10, 9), "страница пропала — отказ");
            AssertTrue(!PdfPageProbe.PagesKept(10, 11), "страница прибавилась — тоже отказ");
            AssertTrue(!PdfPageProbe.PagesKept(10, 0), "пустой вывод — отказ");
            AssertTrue(!PdfPageProbe.PagesKept(10, PdfPageProbe.Unreadable), "вывод не открылся — отказ");
            // Исходник не прочитан: сверять не с чем, но вывод обязан быть годным.
            AssertTrue(PdfPageProbe.PagesKept(PdfPageProbe.Unreadable, 3), "починка битого — судим по выводу");
            AssertTrue(PdfPageProbe.PagesKept(0, 3), "нулевой счёт исходника считаем «не прочитан»");
            AssertTrue(!PdfPageProbe.PagesKept(PdfPageProbe.Unreadable, 0),
                "и при нечитаемом исходнике пустой вывод — отказ");
            AssertTrue(!PdfPageProbe.PagesKept(PdfPageProbe.Unreadable, PdfPageProbe.Unreadable),
                "оба нечитаемы — отказ");
        }

        /// <summary>Выборка страниц для проверки цвета: все, пока их немного, иначе равномерно по длине.</summary>
        private static void TestColorSamplePages()
        {
            AssertEqual("0,1,2", string.Join(",", ToStrings(PdfColorProbe.SamplePages(3, 8))),
                "страниц меньше лимита — берём все");
            AssertEqual("0,1,2,3,4,5,6,7", string.Join(",", ToStrings(PdfColorProbe.SamplePages(8, 8))),
                "ровно лимит — тоже все");
            int[] many = PdfColorProbe.SamplePages(80, 8);
            AssertEqual(8, many.Length, "страниц больше лимита — ровно лимит проб");
            AssertEqual(0, many[0], "первая страница попадает всегда");
            AssertTrue(many[many.Length - 1] < 80, "индексы не выходят за документ");
            for (int i = 1; i < many.Length; i++)
                AssertTrue(many[i] > many[i - 1], "выборка строго возрастает (без повторных рендеров)");
            AssertEqual(0, PdfColorProbe.SamplePages(0, 8).Length, "пустой документ — нечего смотреть");
            AssertEqual(0, PdfColorProbe.SamplePages(5, 0).Length, "нулевой лимит — пустая выборка");
            AssertEqual(0, PdfColorProbe.SamplePages(PdfPageProbe.Unreadable, 8).Length,
                "нечитаемый документ — пустая выборка, а не падение");
        }

        /// <summary>
        /// Поиск цвета в растре. Порог по ДОЛЕ, а не по одному пикселю: одинокий артефакт
        /// масштабирования не должен отменять операцию, которая на деле удалась.
        /// </summary>
        private static void TestColorHasColor()
        {
            AssertTrue(!PdfColorProbe.HasColor(null), "растра нет — цвет не доказан");
            using (var gray = SolidBitmap(100, 100, System.Drawing.Color.FromArgb(128, 128, 128)))
                AssertTrue(!PdfColorProbe.HasColor(gray), "ровно серый — цвета нет");
            using (var white = SolidBitmap(100, 100, System.Drawing.Color.White))
                AssertTrue(!PdfColorProbe.HasColor(white), "белый лист — цвета нет");
            using (var red = SolidBitmap(100, 100, System.Drawing.Color.Red))
                AssertTrue(PdfColorProbe.HasColor(red), "красный лист — цвет");
            // Ровно на пороге — ещё не цвет (строгое «больше»), на единицу выше — уже цвет.
            using (var edge = SolidBitmap(100, 100,
                System.Drawing.Color.FromArgb(128, 128, 128 - PdfColorProbe.SaturationThreshold)))
                AssertTrue(!PdfColorProbe.HasColor(edge), "насыщенность ровно на пороге — не цвет");
            // Доля: 10000 пикселей, порог 0.5% = 50. Сорок цветных — шум, двести — цвет.
            using (var speck = SolidBitmap(100, 100, System.Drawing.Color.FromArgb(128, 128, 128)))
            {
                PaintPixels(speck, 40, System.Drawing.Color.Red);
                AssertTrue(!PdfColorProbe.HasColor(speck), "единичные цветные точки — не приговор");
            }
            using (var patch = SolidBitmap(100, 100, System.Drawing.Color.FromArgb(128, 128, 128)))
            {
                PaintPixels(patch, 200, System.Drawing.Color.Red);
                AssertTrue(PdfColorProbe.HasColor(patch), "заметная доля цветных точек — цвет");
            }
        }

        private static string[] ToStrings(int[] values)
        {
            var s = new string[values.Length];
            for (int i = 0; i < values.Length; i++)
                s[i] = values[i].ToString();
            return s;
        }

        private static System.Drawing.Bitmap SolidBitmap(int w, int h, System.Drawing.Color color)
        {
            var bmp = new System.Drawing.Bitmap(w, h);
            using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bmp))
            using (var brush = new System.Drawing.SolidBrush(color))
                g.FillRectangle(brush, 0, 0, w, h);
            return bmp;
        }

        /// <summary>Закрасить ровно count первых пикселей растра (построчно) заданным цветом.</summary>
        private static void PaintPixels(System.Drawing.Bitmap bmp, int count, System.Drawing.Color color)
        {
            for (int i = 0; i < count; i++)
                bmp.SetPixel(i % bmp.Width, i / bmp.Width, color);
        }

        /// <summary>
        /// ЖИВАЯ проверка обещания уровня «Очень хорошо»: файл становится меньше, а
        /// изображения остаются В ТОМ ЖЕ разрешении. Смотрим не на слова в подписи, а на
        /// сами картинки внутри документа — сколько в них пикселей до и после. Уровень
        /// «Нормально» тут же служит контролем: он их пересчитывает, и если бы обе ветки
        /// вели себя одинаково, тест бы это увидел.
        /// </summary>
        private static void TestCompressKeepsResolutionLive()
        {
            if (!Ghostscript.Available)
                return;
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerKeep_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string keep = Path.Combine(dir, "чёткость.pdf");
                MakeImagePdf(keep, 2);
                long before = new FileInfo(keep).Length;
                long pixelsBefore = ImagePixels(keep);
                AssertTrue(pixelsBefore > 0, "в исходнике есть изображения");

                AssertTrue(PdfCompression.Compress(keep, CompressionLevel.VeryGood), "«Очень хорошо» применено");
                AssertTrue(new FileInfo(keep).Length < before,
                    "файл стал меньше: " + before + " -> " + new FileInfo(keep).Length);
                AssertEqual(2, PdfPageProbe.PageCount(keep), "страницы на месте");
                AssertEqual(pixelsBefore, ImagePixels(keep), "изображения остались в том же разрешении");

                // Контроль: уровень с пересчётом пиксели действительно уменьшает.
                string shrink = Path.Combine(dir, "пересчёт.pdf");
                MakeImagePdf(shrink, 2);
                AssertTrue(PdfCompression.Compress(shrink, CompressionLevel.Small), "«Нормально» применено");
                AssertTrue(ImagePixels(shrink) < pixelsBefore,
                    "пересчитывающий уровень уменьшает разрешение: " + ImagePixels(shrink) + " < " + pixelsBefore);
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        /// <summary>Сколько всего пикселей в изображениях документа (0 — их нет или файл не открылся).</summary>
        private static long ImagePixels(string path)
        {
            try
            {
                long total = 0;
                using (UglyToad.PdfPig.PdfDocument doc = UglyToad.PdfPig.PdfDocument.Open(path))
                    foreach (UglyToad.PdfPig.Content.Page page in doc.GetPages())
                        foreach (UglyToad.PdfPig.Content.IPdfImage img in page.GetImages())
                            total += (long)img.WidthInSamples * img.HeightInSamples;
                return total;
            }
            catch { return 0; }
        }

        /// <summary>
        /// ЖИВОЙ пин на выбор библиотеки для проверки результата. Страницы считаются PdfPig, и
        /// это не вкусовщина: PdfSharp 1.50 не понимает потоков объектов (PDF 1.5+), а их
        /// пишут по умолчанию и Word, и Acrobat. Считай мы страницы PdfSharp — проверка
        /// отказывала бы на каждом таком файле, то есть сжатие сообщало бы «не удалось» там,
        /// где движок отработал безупречно. Файл для пробы делает сам движок.
        /// Если PdfSharp однажды научится их читать, тест упадёт — и решение можно будет
        /// пересмотреть осознанно, а не случайно.
        /// </summary>
        private static void TestPageProbeLive()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerProbe_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string src = Path.Combine(dir, "исходник.pdf");
                MakeEmptyPagesPdf(src, 3);
                AssertEqual(3, PdfPageProbe.PageCount(src), "обычный файл: страницы посчитаны");
                AssertEqual(PdfPageProbe.Unreadable, PdfPageProbe.PageCount(Path.Combine(dir, "нет.pdf")),
                    "несуществующий файл — «не открылся», а не исключение");
                File.WriteAllText(Path.Combine(dir, "мусор.pdf"), "это вовсе не PDF");
                AssertEqual(PdfPageProbe.Unreadable, PdfPageProbe.PageCount(Path.Combine(dir, "мусор.pdf")),
                    "мусор вместо PDF — «не открылся»");

                if (!Ghostscript.Available)
                    return; // файл с потоками объектов делает движок — без него пина нет

                string modern = Path.Combine(dir, "объектные-потоки.pdf");
                string args = "-sDEVICE=pdfwrite -dCompatibilityLevel=1.5 -dNOPAUSE -dBATCH -dQUIET -dSAFER" +
                              " -sOutputFile=\"" + modern + "\" \"" + src + "\"";
                string stderr;
                Ghostscript.Run(args, 120000, out stderr);
                if (!File.Exists(modern))
                    return;
                // Потоки объектов движок создаёт не всегда: если их в файле нет, пин про
                // PdfSharp бессмысленен, но счёт страниц PdfPig проверить всё равно надо.
                bool hasObjectStreams = IndexOfBytes(File.ReadAllBytes(modern),
                    System.Text.Encoding.ASCII.GetBytes("/ObjStm")) >= 0;
                AssertEqual(3, PdfPageProbe.PageCount(modern), "PdfPig читает файл версии 1.5");
                if (hasObjectStreams)
                    AssertThrowsAny("PdfSharp 1.50 на потоках объектов спотыкается — потому и PdfPig",
                        delegate { PdfPageCount(modern); });
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        /// <summary>
        /// ЖИВАЯ проверка перевода в серое на трудных случаях. Простой цветной прямоугольник
        /// движок переводит и без всяких ухищрений — а вот цвет, спрятанный в JPEG, в модели
        /// CMYK и под прозрачностью, ходит другими путями внутри движка, и раньше связка
        /// «стратегия + модель устройства» проверялась только на самом лёгком из них.
        /// Здесь же проверяется и то, что проба цвета не врёт в обе стороны: на цветном
        /// документе она цвет находит, на переведённом — нет.
        /// </summary>
        private static void TestGrayscaleHardLive()
        {
            if (!Ghostscript.Available)
                return;
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerGray_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string hard = Path.Combine(dir, "трудный.pdf");
                MakeHardColorPdf(hard, dir);
                AssertEqual(3, PdfPageProbe.PageCount(hard), "проба из трёх страниц");
                AssertTrue(HasColor(hard), "на первой странице исходника есть цвет");

                // Без работающего рендера проба цвета молчит по построению (см. PdfColorProbe):
                // тогда её утверждения проверять нечем, а перевод обязан работать как прежде.
                bool canRender = ThumbnailsWork(hard);
                if (canRender)
                    AssertTrue(PdfColorProbe.HasColorPages(hard), "проба находит цвет в исходнике");
                AssertTrue(PdfConvert.Verify(hard, PdfConvertMode.Repair),
                    "починке цвет исходника безразличен");
                if (canRender)
                    AssertTrue(!PdfConvert.Verify(hard, PdfConvertMode.Grayscale),
                        "цветной документ проверку «серое» не проходит");

                AssertTrue(PdfConvert.Apply(hard, PdfConvertMode.Grayscale), "перевод применён");
                AssertEqual(3, PdfPageProbe.PageCount(hard), "все три страницы на месте");
                AssertTrue(!HasColor(hard), "цвета не осталось (страница с JPEG)");
                if (canRender)
                {
                    AssertTrue(!PdfColorProbe.HasColorPages(hard),
                        "цвета не осталось ни на одной из трёх страниц");
                    AssertTrue(PdfConvert.Verify(hard, PdfConvertMode.Grayscale),
                        "переведённый документ проверку проходит");
                }
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        /// <summary>
        /// Потолок высоты растра: ширина уменьшается ровно настолько, чтобы в него уложиться,
        /// и не трогается там, где страница и так низкая. Чистая — под тест.
        /// </summary>
        private static void TestRenderFitWidth()
        {
            // A4 при ширине 140 даёт ~198 px — потолок в 2000 ей не мешает.
            AssertEqual(140, PdfThumbnailRenderer.FitWidth(140, 595, 842, 2000), "обычная страница не ужимается");
            AssertEqual(140, PdfThumbnailRenderer.FitWidth(140, 595, 842, 0), "нулевой потолок ничего не меняет");
            // Умеренно вытянутый лист: потолок достижим, и ширина считается точно под него.
            AssertEqual(100, PdfThumbnailRenderer.FitWidth(140, 100, 2000, 2000),
                "ширина подбирается ровно под потолок высоты");
            // Предельный лист 3×14400: при ширине 140 высота была бы 672000 px. Ужать ровно
            // под потолок тут нельзя — ширина упирается в один пиксель, — но именно это и
            // требуется: растр выходит в 19 КБ вместо 376 МБ.
            int narrow = PdfThumbnailRenderer.FitWidth(140, 3, 14400, 2000);
            AssertEqual(1, narrow, "предельно длинная страница ужимается до минимальной ширины");
            long pixels = (long)(narrow * 14400.0 / 3) * narrow;
            AssertTrue(pixels < 10000L * 1024 / 4, "и растр остаётся копеечным: " + pixels + " пикселей");
            // Ровно на границе — ужимать незачем.
            AssertEqual(100, PdfThumbnailRenderer.FitWidth(100, 100, 2000, 2000), "высота ровно по потолку — как есть");
            // Мусор на входе не должен превращаться в деление на ноль.
            AssertEqual(140, PdfThumbnailRenderer.FitWidth(140, 0, 842, 2000), "нулевая ширина страницы — ширина как есть");
            AssertEqual(140, PdfThumbnailRenderer.FitWidth(140, 595, 0, 2000), "нулевая высота страницы — ширина как есть");
            AssertEqual(0, PdfThumbnailRenderer.FitWidth(0, 595, 842, 2000), "нулевая запрошенная ширина — она же");
        }

        /// <summary>
        /// ЭКСТРЕМАЛЬНАЯ геометрия: PDF разрешает лист до 14400 пунктов при ширине в единицы,
        /// и у пробы цвета высота растра считается от ширины — на такой странице она ушла бы
        /// в сотни тысяч пикселей, то есть в сотни мегабайт под один растр. Проверяем ровно
        /// это: проба обязана отработать, не подвесив и не обвалив процесс.
        /// </summary>
        private static void TestColorProbeExtremePage()
        {
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerLong_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string tall = Path.Combine(dir, "длинная.pdf");
                using (var doc = new PdfDocument())
                {
                    PdfSharp.Pdf.PdfPage page = doc.AddPage();
                    page.Width = 3;       // предельно узкий и предельно длинный лист
                    page.Height = 14400;  // потолок формата
                    using (XGraphics g = XGraphics.FromPdfPage(page))
                        g.DrawRectangle(new XSolidBrush(XColors.Red), 0, 0, 3, 14400);
                    doc.Save(tall);
                }
                long peakBefore = System.Diagnostics.Process.GetCurrentProcess().PeakWorkingSet64;
                var sw = System.Diagnostics.Stopwatch.StartNew();
                bool coloured = PdfColorProbe.HasColorPages(tall); // не должно ни падать, ни виснуть
                sw.Stop();
                long peakAfter = System.Diagnostics.Process.GetCurrentProcess().PeakWorkingSet64;
                Console.WriteLine("      [замер] длинная страница: {0} мс, пик +{1:F1} МБ, цвет={2}",
                    sw.ElapsedMilliseconds, (peakAfter - peakBefore) / 1048576.0, coloured);
                AssertTrue(sw.ElapsedMilliseconds < 20000,
                    "проба на предельной странице укладывается в разумное время: " + sw.ElapsedMilliseconds + " мс");
                // Порог низкий намеренно: без потолка высоты этот лист стоил 400 МБ на
                // сборочной машине, а на машине разработчика движок за него не брался вовсе
                // и не стоил ничего. Разница между машинами — ровно то, ради чего проверка
                // живёт в наборе, а не в голове у автора.
                AssertTrue((peakAfter - peakBefore) < 64L * 1024 * 1024,
                    "проба не съедает десятки мегабайт: +" + (peakAfter - peakBefore) / 1048576 + " МБ");
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        /// <summary>
        /// Подмена файла идёт сразу после того, как его читали (страницы, а для серого ещё и
        /// рендер), а на Windows такое соседство упирается в антивирус: он берётся за файл
        /// сразу после закрытия хэндла, и переименование изредка отказывает. Отсюда и правило
        /// «страницы исходника считаем ДО запуска движка» в GsRewrite.
        ///
        /// Замер, на котором основано решение НЕ добавлять повторные попытки в подмену:
        /// 25 кругов, то есть 50 подмен подряд, — ноль отказов. Пока это так, повторы были бы
        /// кодом на случай, которого нет; если тест однажды упадёт, повод появится.
        /// </summary>
        private static void TestGsReplaceStress()
        {
            if (!Ghostscript.Available)
                return;
            string dir = Path.Combine(Path.GetTempPath(), "ExcelMergerStress_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                // Кругов ровно столько, чтобы систематический отказ проявился, но набор не
                // распухал: полтора десятка подмен ловят его так же, как полсотни.
                const int Rounds = 8;
                int grayFails = 0, compressFails = 0;
                var sw = System.Diagnostics.Stopwatch.StartNew();
                for (int i = 0; i < Rounds; i++)
                {
                    string gray = Path.Combine(dir, "г" + i + ".pdf");
                    MakeColorPdf(gray);
                    if (!PdfConvert.Apply(gray, PdfConvertMode.Grayscale)) grayFails++;
                    string comp = Path.Combine(dir, "с" + i + ".pdf");
                    MakeImagePdf(comp, 1);
                    if (!PdfCompression.Compress(comp, CompressionLevel.Small)) compressFails++;
                }
                sw.Stop();
                Console.WriteLine("      [замер] {0} циклов: отказов серого {1}, отказов сжатия {2}, {3} мс",
                    Rounds, grayFails, compressFails, sw.ElapsedMilliseconds);
                AssertEqual(0, grayFails, "перевод в серое не отказывает на повторах");
                AssertEqual(0, compressFails, "сжатие не отказывает на повторах");
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        }

        /// <summary>Работает ли системный рендер миниатюр на этой машине (WinRT есть не везде).</summary>
        private static bool ThumbnailsWork(string pdf)
        {
            using (var renderer = new PdfThumbnailRenderer())
            using (System.Drawing.Bitmap bmp = renderer.Render(pdf, 0, 80))
                return bmp != null;
        }

        /// <summary>
        /// Три страницы, на каждой цвет своей природы: JPEG (проходит через движок отдельным
        /// путём как DCTDecode), модель CMYK и цвет под прозрачностью (ExtGState).
        /// </summary>
        private static void MakeHardColorPdf(string path, string workDir)
        {
            string jpeg = Path.Combine(workDir, "яркая.jpg");
            using (var bmp = new System.Drawing.Bitmap(120, 90))
            {
                using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bmp))
                {
                    g.Clear(System.Drawing.Color.FromArgb(220, 30, 40));
                    using (var brush = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(20, 90, 220)))
                        g.FillRectangle(brush, 0, 45, 120, 45);
                }
                bmp.Save(jpeg, System.Drawing.Imaging.ImageFormat.Jpeg);
            }

            using (var doc = new PdfDocument())
            {
                // 1. Цвет внутри JPEG.
                using (XGraphics g = XGraphics.FromPdfPage(doc.AddPage()))
                using (XImage img = XImage.FromFile(jpeg))
                    g.DrawImage(img, 40, 40, 300, 200);
                // 2. Цвет в модели CMYK.
                using (XGraphics g = XGraphics.FromPdfPage(doc.AddPage()))
                    g.DrawRectangle(new XSolidBrush(XColor.FromCmyk(0, 1, 1, 0)), 40, 40, 300, 200);
                // 3. Цвет под прозрачностью.
                using (XGraphics g = XGraphics.FromPdfPage(doc.AddPage()))
                    g.DrawRectangle(new XSolidBrush(XColor.FromArgb(128, 20, 60, 220)), 40, 40, 300, 200);
                doc.Save(path);
            }
        }

        /// <summary>Первое вхождение подстроки байтов (или -1).</summary>
        private static int IndexOfBytes(byte[] haystack, byte[] needle)
        {
            for (int i = 0; i + needle.Length <= haystack.Length; i++)
            {
                bool hit = true;
                for (int j = 0; j < needle.Length && hit; j++)
                    if (haystack[i + j] != needle[j])
                        hit = false;
                if (hit)
                    return i;
            }
            return -1;
        }

        // ---------- Простой текст (экспорт в .txt) ----------

        private static OcrParagraph Par(string text, double top, double left)
        {
            var p = new OcrParagraph { TopPt = top, LeftPt = left };
            p.Runs.Add(new OcrRun { Text = text });
            return p;
        }

        private static OcrTableCell Cell(string text)
        {
            var c = new OcrTableCell();
            if (text != null)
                c.Paragraphs.Add(Par(text, 0, 0));
            return c;
        }

        private static OcrTable Grid(double top, double left, params string[][] rows)
        {
            var t = new OcrTable { TopPt = top, LeftPt = left };
            foreach (string[] row in rows)
            {
                var r = new OcrTableRow();
                foreach (string cell in row)
                    r.Cells.Add(Cell(cell));
                t.Rows.Add(r);
                t.ColumnWidthsPt.Clear();
                for (int i = 0; i < row.Length; i++)
                    t.ColumnWidthsPt.Add(10);
            }
            return t;
        }

        /// <summary>
        /// Таблицы обязаны попадать в текст. Разбор УБИРАЕТ слова таблиц из потока абзацев
        /// (они уходят в ячейки), поэтому наивная склейка абзацев теряла бы их целиком —
        /// именно это и проверяем: сборка сводит абзацы и таблицы обратно по вертикали.
        /// </summary>
        private static void TestPlainTextKeepsTables()
        {
            var page = new PdfPageText();
            page.Paragraphs.Add(Par("Заголовок", 800, 50));
            page.Paragraphs.Add(Par("Подпись внизу", 100, 50));
            page.Tables.Add(Grid(500, 50,
                new[] { "Товар", "Цена" },
                new[] { "Болт", "10" }));

            string text = PlainText.Page(page);
            AssertEqual("Заголовок\n\nТовар\tЦена\nБолт\t10\n\nПодпись внизу", text,
                "абзацы и таблица идут сверху вниз, ячейки через табуляцию");
        }

        /// <summary>Ось Y направлена вверх: больший TopPt — выше на странице, значит раньше.</summary>
        private static void TestPlainTextReadingOrder()
        {
            var page = new PdfPageText();
            page.Paragraphs.Add(Par("низ", 100, 50));
            page.Paragraphs.Add(Par("верх", 700, 50));
            page.Paragraphs.Add(Par("середина справа", 400, 300));
            page.Paragraphs.Add(Par("середина слева", 400, 50));
            AssertEqual("верх\n\nсередина слева\n\nсередина справа\n\nниз", PlainText.Page(page),
                "сверху вниз, при равной высоте — слева направо");
            AssertEqual("", PlainText.Page(null), "страницы нет — пустой текст");
            AssertEqual("", PlainText.Page(new PdfPageText()), "пустая страница — пустой текст");
        }

        /// <summary>Накрытые объединением ячейки дают пустое поле, иначе колонки разъезжаются.</summary>
        private static void TestPlainTextMergedCells()
        {
            var t = new OcrTable { TopPt = 500 };
            var row = new OcrTableRow();
            row.Cells.Add(Cell("Итого"));
            var covered = Cell(null);
            covered.Covered = true;
            row.Cells.Add(covered);
            row.Cells.Add(Cell("42"));
            t.Rows.Add(row);
            AssertEqual("Итого\t\t42", PlainText.Table(t), "накрытая ячейка — пустое поле, колонки на месте");
        }

        /// <summary>Страницы разделяются переводом страницы, пустые не съедаются — нумерация не съезжает.</summary>
        private static void TestPlainTextDocument()
        {
            var one = new PdfPageText();
            one.Paragraphs.Add(Par("первая", 700, 50));
            var empty = new PdfPageText();
            var three = new PdfPageText();
            three.Paragraphs.Add(Par("третья", 700, 50));
            AssertEqual("первая\n\f\n\n\f\nтретья",
                PlainText.Document(new List<PdfPageText> { one, empty, three }),
                "пустая страница остаётся на своём месте");
            AssertEqual("", PlainText.Document(null), "нет страниц — пустой текст");
        }

        // ---------- Экспорт в картинки ----------

        private static void TestExportPixelWidth()
        {
            // A4 шириной 595.28 pt: 96 dpi → 794 px, 300 dpi → 2480 px.
            AssertEqual(794, PdfExportService.PixelWidth(595.28, 96), "экранное разрешение");
            AssertEqual(2480, PdfExportService.PixelWidth(595.28, 300), "разрешение печати");
            AssertEqual(1, PdfExportService.PixelWidth(0, 300), "нулевая ширина не даёт нулевую картинку");
            AssertEqual(1, PdfExportService.PixelWidth(595.28, 0), "нулевое разрешение не даёт нулевую картинку");
            AssertEqual(".png", PdfExportService.Extension(ImageExportFormat.Png), "расширение PNG");
            AssertEqual(".jpg", PdfExportService.Extension(ImageExportFormat.Jpeg), "расширение JPEG");
        }

        // ---------- Чередование страниц ----------

        private static PdfPageRef PR(string path, int index)
        {
            return new PdfPageRef { SourcePath = path, PageIndex = index };
        }

        private static string PagesOf(IList<PdfPageRef> pages)
        {
            var parts = new List<string>();
            foreach (PdfPageRef p in pages)
                parts.Add((p.SourcePath ?? "-") + p.PageIndex);
            return string.Join(",", parts.ToArray());
        }

        /// <summary>Порядок разбирается на документы по непрерывным отрезкам одного файла.</summary>
        private static void TestInterleaveRuns()
        {
            var pages = new List<PdfPageRef> { PR("a", 0), PR("a", 1), PR("b", 0), PR("b", 1), PR("b", 2) };
            List<PageRun> runs = PageInterleave.SplitIntoRuns(pages);
            AssertEqual(2, runs.Count, "два документа");
            AssertEqual(0, runs[0].Start, "первый начинается с нуля");
            AssertEqual(2, runs[0].Count, "в первом две страницы");
            AssertEqual(3, runs[1].Count, "во втором три");
            AssertTrue(PageInterleave.CanInterleave(runs), "два документа — чередовать есть что");

            // Тот же файл, добавленный дважды, — это две пачки: пользователь видит их отдельно.
            List<PageRun> twice = PageInterleave.SplitIntoRuns(
                new List<PdfPageRef> { PR("a", 0), PR("b", 0), PR("a", 1) });
            AssertEqual(3, twice.Count, "перемежающиеся страницы дают три отрезка");
            AssertTrue(!PageInterleave.CanInterleave(PageInterleave.SplitIntoRuns(
                new List<PdfPageRef> { PR("a", 0), PR("a", 1) })), "один документ чередовать не с чем");
            AssertTrue(!PageInterleave.CanInterleave(null), "пустой ввод — нечего чередовать");
        }

        /// <summary>
        /// Главный сценарий: односторонний сканер дал лицевые стороны в одном файле, а
        /// оборотные — в другом и в обратном порядке. После чередования должно получиться
        /// 1,2,3,4,5,6 подряд.
        /// </summary>
        private static void TestInterleaveScannerCase()
        {
            // face: стр. 1,3,5 (индексы 0,1,2); back: стр. 6,4,2 (индексы 0,1,2 в обратном порядке)
            var pages = new List<PdfPageRef>
            {
                PR("face", 0), PR("face", 1), PR("face", 2),
                PR("back", 0), PR("back", 1), PR("back", 2)
            };
            List<PageRun> runs = PageInterleave.SplitIntoRuns(pages);
            runs[1].Reverse = true; // оборотная пачка идёт с конца
            List<PdfPageRef> mixed = PageInterleave.Interleave(pages, runs, 1);
            AssertEqual("face0,back2,face1,back1,face2,back0", PagesOf(mixed), "лицевая и оборотная сторона чередуются");
            AssertEqual(pages.Count, mixed.Count, "ни одна страница не потеряна и не задвоена");
        }

        /// <summary>Шаг больше единицы, разная длина пачек и перестановка тех же ссылок.</summary>
        private static void TestInterleavePaceAndTails()
        {
            var pages = new List<PdfPageRef>
            {
                PR("a", 0), PR("a", 1), PR("a", 2), PR("a", 3), PR("b", 0), PR("b", 1)
            };
            List<PageRun> runs = PageInterleave.SplitIntoRuns(pages);
            AssertEqual("a0,a1,b0,b1,a2,a3", PagesOf(PageInterleave.Interleave(pages, runs, 2)), "по две страницы по кругу");
            // Более длинный документ дописывает хвост, когда короткий кончился.
            AssertEqual("a0,b0,a1,b1,a2,a3", PagesOf(PageInterleave.Interleave(pages, runs, 1)), "хвост длинной пачки в конце");
            AssertEqual("a0,b0,a1,b1,a2,a3", PagesOf(PageInterleave.Interleave(pages, runs, 0)), "шаг меньше единицы считается за единицу");

            // Это перестановка: те же самые объекты, а не копии — назначенные повороты уедут со страницами.
            List<PdfPageRef> mixed = PageInterleave.Interleave(pages, runs, 1);
            foreach (PdfPageRef p in mixed)
                AssertTrue(pages.Contains(p), "в результате те же ссылки на страницы");
        }

        // ---------- Шаблон имени выходного файла ----------

        private static void TestNameTemplateBasics()
        {
            var v = new NameValues { BaseName = "отчёт", FileNumber = 2, TotalFiles = 12, CurrentPage = 7 };
            AssertEqual("отчёт_2", NameTemplate.Apply("[BASENAME]_[FILENUMBER]", v), "имя и номер файла");
            AssertEqual("2_из_12", NameTemplate.Apply("[FILENUMBER]_из_[TOTAL_FILES]", v), "номер из всего");
            AssertEqual("стр7", NameTemplate.Apply("стр[CURRENTPAGE]", v), "номер страницы источника");
            AssertEqual("отчёт_2", NameTemplate.Apply(null, v), "пустой шаблон — прежнее поведение");
            AssertEqual("отчёт_2", NameTemplate.Apply("", v), "пустая строка — прежнее поведение");
            AssertEqual("без токенов", NameTemplate.Apply("без токенов", v), "текст без токенов идёт как есть");
        }

        private static void TestNameTemplatePadAndOffset()
        {
            var v = new NameValues { FileNumber = 2, CurrentPage = 3 };
            AssertEqual("002", NameTemplate.Apply("[FILENUMBER###]", v), "решётки дополняют нулями");
            AssertEqual("12", NameTemplate.Apply("[FILENUMBER10]", v), "число — смещение нумерации");
            AssertEqual("012", NameTemplate.Apply("[FILENUMBER###10]", v), "дополнение и смещение вместе");
            AssertEqual("3", NameTemplate.Apply("[CURRENTPAGE#]", v), "одна решётка ничего не дополняет");
        }

        private static void TestNameTemplateUnknownAndUnsafe()
        {
            var v = new NameValues { BaseName = "имя", FileNumber = 1 };
            // Неизвестный токен и одиночные скобки — обычный текст, а не ошибка.
            AssertEqual("[НЕТ]_имя", NameTemplate.Apply("[НЕТ]_[BASENAME]", v), "неизвестный токен остаётся текстом");
            AssertEqual("[незакрытая имя", NameTemplate.Apply("[незакрытая [BASENAME]", v), "незакрытая скобка не ломает разбор");
            // Разделители пути обязаны исчезнуть: шаблон не должен писать за пределы выбранной папки.
            AssertEqual("_.._тайно_имя", NameTemplate.Apply("/../тайно/[BASENAME]", v), "разделители пути обезврежены");
            AssertEqual("имя", NameTemplate.Apply("[BASENAME]...", v), "точки на конце срезаны");
            AssertEqual("___", NameTemplate.Apply(":::", v), "запрещённые символы заменяются, а не выбрасываются");
            // Шаблон, давший пустое имя, откатывается к обычному: файла без имени быть не может.
            AssertEqual("имя_1", NameTemplate.Apply("[BOOKMARK]", v), "нет закладки — откат к шаблону по умолчанию");
            AssertTrue(NameTemplate.Apply("[TIMESTAMP]", v).Length > 0, "без времени шаблон всё равно даёт имя");
        }

        private static void TestNameTemplateUniqueness()
        {
            AssertTrue(NameTemplate.IsUniquePerFile("[BASENAME]_[FILENUMBER]"), "номер файла различает части");
            AssertTrue(NameTemplate.IsUniquePerFile("[CURRENTPAGE###]"), "номер страницы различает части");
            AssertTrue(NameTemplate.IsUniquePerFile(null), "шаблон по умолчанию различает части");
            AssertTrue(!NameTemplate.IsUniquePerFile("[BASENAME]"), "одно имя источника — все части совпадут");
            AssertTrue(!NameTemplate.IsUniquePerFile("постоянное"), "текст без токенов — все части совпадут");
        }

        private static void AssertEqual(object expected, object actual, string what)
        {
            if (!Equals(expected, actual))
                throw new Exception(what + ": ожидалось «" + expected + "», получено «" + actual + "»");
        }

        private static void AssertTrue(bool condition, string what)
        {
            if (!condition)
                throw new Exception(what);
        }

        private static void AssertThrows(string what, Action action)
        {
            try
            {
                action();
            }
            catch (MergeException)
            {
                return; // ожидаемая ошибка ввода
            }
            throw new Exception(what + ": ожидалось исключение, но его не было");
        }

        /// <summary>
        /// Руководство пользователя вшито ресурсом в exe приложения, и его наличие проверяет
        /// самопроверка (<c>--selftest</c>, код 4). Здесь проверяется ОБРАТНОЕ: что та проверка
        /// вообще умеет отвечать «нет». Тестовая сборка руководство не несёт, поэтому
        /// <see cref="UserManual.IsPacked"/> обязана вернуть false — иначе она вернула бы
        /// «всё на месте» и на exe без ресурса, а самопроверка зеленела бы, ничего не проверив.
        /// </summary>
        private static void TestUserManualPackedDetectsAbsence()
        {
            if (UserManual.IsPacked())
                throw new Exception("в тестовой сборке руководства нет, а проверка говорит, что есть");
        }

        /// <summary>
        /// Документ распаковывается РЯДОМ С НАСТРОЙКАМИ и под человеческим именем: имя ресурса
        /// внутри exe латинское (кириллица в именах ресурсов — лишний повод для сюрпризов), а в
        /// заголовке Word пользователь должен увидеть нормальное название.
        /// </summary>
        private static void TestUserManualPath()
        {
            InIsolatedSettings("iwo_manual_", delegate
            {
                string path = UserManual.FilePath;
                if (Path.GetDirectoryName(path) != Path.GetDirectoryName(AppPaths.SettingsFile))
                    throw new Exception("руководство распаковывается не туда, где настройки: " + path);
                if (Path.GetFileName(path) != "Инструкция пользователя.docx")
                    throw new Exception("имя файла руководства не человеческое: " + Path.GetFileName(path));
            });
        }
    }
}
