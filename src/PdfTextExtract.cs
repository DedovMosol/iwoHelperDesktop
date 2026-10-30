using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

namespace ExcelMerger
{
    /// <summary>Растровое изображение страницы (PNG) с рамкой в координатах PDF (ось Y вверх).</summary>
    public class OcrImage
    {
        public byte[] Png;
        public double TopPt;    // верх изображения — для порядка чтения вместе с абзацами
        public double LeftPt;
        public double WidthPt;
        public double HeightPt;
    }

    /// <summary>
    /// Один декодированный source-unit text layer. Это extraction-only модель: наружу не
    /// выходят ни Letter, ни другие объекты PdfPig. Trusted означает подходящую ориентацию
    /// и однозначно декодированное значение; видимость доказывают соседние retained-слова.
    /// </summary>
    internal sealed class PdfSourceTextUnit
    {
        public string Text;
        public bool Trusted;
    }

    /// <summary>Текст одной страницы born-digital PDF (абзацы в порядке чтения).</summary>
    public class PdfPageText
    {
        public int PageIndex;                              // с нуля
        public List<OcrParagraph> Paragraphs = new List<OcrParagraph>();
        public List<OcrImage> Images = new List<OcrImage>();
        public List<OcrTable> Tables = new List<OcrTable>(); // восстановленные таблицы (порядок чтения — по TopPt)
        internal List<PdfLine> Lines = new List<PdfLine>(); // линовка страницы (границы таблиц, подчёркивания)
        public double FirstLineIndentPt;                   // отступ красной строки (pt); 0 — без отступов
        public double WidthPt;
        public double HeightPt;
        // Поля страницы из рамок текста (pt); 0 — на странице не было текста.
        public double LeftMarginPt;
        public double RightMarginPt;
        public double TopMarginPt;
        public double BottomMarginPt;
        // Все слова страницы с рамками (пространство отображения: X вправо, Y вверх) в порядке
        // извлечения. Export-режим Document может заменить слова текстового штампа изображением;
        // Review сохраняет их для семантического сравнения и подсветки.
        internal List<PdfWord> Words = new List<PdfWord>();
        // Декодированный порядок source-unit до GetWords(): только строки и флаги,
        // без ссылок на PdfPig. Review использует его исключительно как положительное
        // доказательство литерального пробела либо доказуемо пустой границы.
        internal List<PdfSourceTextUnit> SourceUnits = new List<PdfSourceTextUnit>();
        // Собственного поворота страницы (/Rotate) здесь нет и не нужно: PdfPig уже применяет его
        // и сдвиг CropBox в начальной матрице, поэтому WidthPt/HeightPt и рамки слов — сразу в
        // пространстве отображения, том же, что показывает системный рендерер (Windows.Data.Pdf).
        // Повторный поворот поверх этого разворачивал подсветку «Сравнения» не туда.
        // Для какого вывода разобрана страница. Писателю слайдов это нужно, чтобы отличить
        // прочерк, СИНТЕЗИРОВАННЫЙ из линии (только Document, см. AddRuleWords), от набранных
        // в источнике подчёркиваний: первые на слайде — шум, вторые — текст бланка.
        // По умолчанию Document — самое осторожное предположение для страниц, собранных в коде.
        internal PageLayoutMode LayoutMode;
        // Страницу разобрать не удалось (RC-11): причина; страница перенесена картинкой во весь
        // лист (или пустой, если рендерить нечем). null — обычная страница.
        internal string ExtractionError;

        /// <summary>Весь текст страницы: абзацы через пустую строку.</summary>
        public string Text
        {
            get
            {
                var parts = new List<string>(Paragraphs.Count);
                foreach (OcrParagraph p in Paragraphs)
                    parts.Add(p.Text);
                return string.Join("\n\n", parts);
            }
        }
    }

    /// <summary>
    /// Извлечение текстового слоя born-digital PDF (PdfPig, Apache 2.0) — без OCR.
    /// Слова с рамками собираются в порядок чтения (<see cref="OcrLayout"/>). Публичные
    /// методы не содержат типов PdfPig в телах: сначала
    /// <see cref="EmbeddedAssemblies.Ensure"/>, затем [NoInlining]-ядро (как в
    /// <see cref="PdfMergeService"/>), иначе JIT падает на резолве вшитой сборки.
    /// </summary>
    public static class PdfTextExtract
    {
        /// <summary>
        /// Текст страниц файла. rotations — пользовательские повороты по индексу исходной
        /// страницы (0/90/180/270 по часовой, null — без поворотов): страница выправляется
        /// ДО анализа макета, боковой текст становится горизонтальным. pageIndexes — какие
        /// страницы (с нуля) разбирать, null — все: две страницы из пятисот не должны стоить
        /// разбора, декодирования картинок и Ghostscript для остальных; результат — только
        /// запрошенные страницы в порядке файла, их PageIndex — исходный. extractImages=false —
        /// картинки не декодируются вовсе (Сравнению и текстовому экспорту они не нужны), а их
        /// рамки по-прежнему работают фильтрами невидимого текста. Битый/зашифрованный файл
        /// или запрет извлечения — <see cref="MergeException"/> с понятным сообщением.
        /// </summary>
        public static List<PdfPageText> Extract(string path, Action<int, int> progress = null, IList<int> rotations = null,
            Func<bool> cancelled = null, PageLayoutMode mode = PageLayoutMode.Document,
            ICollection<int> pageIndexes = null, bool extractImages = true)
        {
            EmbeddedAssemblies.Ensure();
            return ExtractCore(path, progress, rotations, cancelled, mode, pageIndexes, extractImages);
        }

        /// <summary>
        /// Номера страниц (с единицы, по возрастанию) к разбору: все при pageIndexes == null,
        /// иначе только существующие из запрошенных, без повторов. Чистая — под тест.
        /// </summary>
        internal static List<int> PageNumbersToExtract(int pageCount, ICollection<int> pageIndexes)
        {
            var result = new List<int>();
            if (pageIndexes == null)
            {
                for (int n = 1; n <= pageCount; n++)
                    result.Add(n);
                return result;
            }
            var wanted = new SortedSet<int>();
            foreach (int index in pageIndexes)
                if (index >= 0 && index < pageCount)
                    wanted.Add(index + 1);
            result.AddRange(wanted);
            return result;
        }

        /// <summary>
        /// Есть ли хоть на одной странице файла текст (буквы). Дёшево: буквы разбираются вместе
        /// со страницей, картинки НЕ декодируются (15-МБ скан — ~0.1 с против ~3 с и ~18 МБ
        /// полного извлечения). Ошибка чтения — true: пусть полноценное извлечение упадёт со
        /// своим понятным сообщением, проба не должна маскировать причину.
        /// </summary>
        public static bool AnyPageHasText(string path)
        {
            EmbeddedAssemblies.Ensure();
            return AnyPageHasTextCore(path);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static bool AnyPageHasTextCore(string path)
        {
            try
            {
                using (UglyToad.PdfPig.PdfDocument doc = PdfPageProbe.OpenPig(path))
                {
                    foreach (UglyToad.PdfPig.Content.Page page in doc.GetPages())
                        if (page.Letters != null && page.Letters.Count > 0)
                            return true;
                    return false;
                }
            }
            catch { return true; }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<PdfPageText> ExtractCore(string path, Action<int, int> progress, IList<int> rotations,
            Func<bool> cancelled, PageLayoutMode mode, ICollection<int> pageIndexes, bool extractImages)
        {
            try
            {
                using (UglyToad.PdfPig.PdfDocument doc = PdfPageProbe.OpenPig(path))
                {
                    var pages = new List<PdfPageText>();
                    // GetPage(n) разбирает ТОЛЬКО свою страницу: перечисление GetPages() разобрало
                    // бы и все пропускаемые.
                    List<int> numbers = PageNumbersToExtract(doc.NumberOfPages, pageIndexes);
                    int pageCount = numbers.Count;
                    Exception firstFailure = null;
                    foreach (int number in numbers)
                    {
                        Cancellation.ThrowIf(cancelled); // отмена между страницами извлечения
                        int rotation = PageRotation.At(rotations, number - 1);
                        UglyToad.PdfPig.Content.Page page = null;
                        PdfPageText pt;
                        try
                        {
                            page = doc.GetPage(number);
                            pt = ExtractPage(page, path, rotation, mode, extractImages);
                        }
                        catch (Exception ex) when (mode != PageLayoutMode.Review && MergeException.ShouldWrap(ex))
                        {
                            // RC-11: одна сбойная страница не роняет весь файл — она переносится
                            // картинкой, а итог называет её. «Сравнение» (Review) остаётся строгим:
                            // пустая страница там читалась бы как «весь текст удалён».
                            pt = FailedPage(path, number, page, rotation, mode, ex);
                            if (firstFailure == null)
                                firstFailure = ex;
                        }
                        pages.Add(pt);
                        if (progress != null)
                            progress(pages.Count, pageCount);
                    }
                    // Не разобралась НИ одна страница — это уже не «сбойная страница», а битый
                    // файл: прежняя понятная ошибка, а не документ из одних картинок.
                    if (firstFailure != null && pages.TrueForAll(p => p.ExtractionError != null))
                        System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(firstFailure).Throw();
                    return pages;
                }
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                throw new MergeException(string.Format(Loc.T("err.ocr.extractFailed"), Path.GetFileName(path), ex.Message));
            }
        }

        /// <summary>
        /// Страница, которую разобрать не удалось (RC-11): пустая разметка и сама страница
        /// картинкой во весь лист — рендер Ghostscript (нет его — страница остаётся пустой, но
        /// документ пишется). Размер — из PdfPig, если страница открылась, иначе из PdfSharp,
        /// иначе A4. Пользовательский поворот применяется так же, как к обычной странице.
        /// </summary>
        private static PdfPageText FailedPage(string path, int number, UglyToad.PdfPig.Content.Page page,
            int rotation, PageLayoutMode mode, Exception error)
        {
            double w = 595, h = 842;
            if (page != null)
            {
                w = page.Width;
                h = page.Height;
            }
            else
                PdfSharpPageSize(path, number, ref w, ref h);
            var images = new List<OcrImage>();
            OcrImage whole = RenderWholePage(path, number, w, h);
            if (whole != null)
                images.Add(whole);
            PageRotation.RotatePage(new List<PdfWord>(), new List<PdfLine>(), images, rotation, ref w, ref h);
            var pt = new PdfPageText
            {
                PageIndex = number - 1,
                WidthPt = w,
                HeightPt = h,
                Images = images,
                LayoutMode = mode,
                ExtractionError = error.Message
            };
            SetMargins(pt, pt.Words, pt.Images, w, h);
            return pt;
        }

        /// <summary>Размер страницы (pt, с учётом /Rotate) по PdfSharp; не открылся — размер не меняется.</summary>
        private static void PdfSharpPageSize(string path, int number, ref double w, ref double h)
        {
            try
            {
                using (PdfSharp.Pdf.PdfDocument doc = PdfSharp.Pdf.IO.PdfReader.Open(path, PdfPasswords.For(path),
                    PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import))
                {
                    if (number < 1 || number > doc.PageCount)
                        return;
                    PdfSharp.Pdf.PdfPage p = doc.Pages[number - 1];
                    if (p.Width.Point > 0 && p.Height.Point > 0)
                    {
                        w = p.Width.Point;
                        h = p.Height.Point;
                    }
                }
            }
            catch (OutOfMemoryException) { throw; }
            catch { } // и PdfSharp не открыл — остаётся A4: картинка страницы всё равно на месте
        }

        /// <summary>Вся страница картинкой (рамка — весь лист), null — нет Ghostscript или рендер не удался.</summary>
        private static OcrImage RenderWholePage(string path, int number, double w, double h)
        {
            BudgetedBitmap owned = null;
            try
            {
                owned = PageRasterizer.RenderPage(path, number, w, h);
                if (owned == null)
                    return null;
                byte[] png = PageRasterizer.CropRegion(owned.Bitmap, w, h, 0, h, w, h);
                return png == null || png.Length == 0 ? null
                    : new OcrImage { Png = png, LeftPt = 0, TopPt = h, WidthPt = w, HeightPt = h };
            }
            catch (OutOfMemoryException) { throw; }
            catch { return null; }
            finally
            {
                if (owned != null) owned.Dispose();
            }
        }

        /// <summary>
        /// Одна страница источника, разобранная целиком (см. <see cref="ExtractCore"/>). Сбойная
        /// страница бросает исключение; превратить ли её в картинку, решает вызывающий (RC-11).
        /// </summary>
        private static PdfPageText ExtractPage(UglyToad.PdfPig.Content.Page page, string path, int rotation,
            PageLayoutMode mode, bool extractImages)
        {
            // Картинки страницы перечисляются ОДИН раз (декода здесь нет): рамки
            // нужны фильтрам невидимого текста и наложенных картинок, а сами объекты — переносу
            // изображений ниже. Рамки картинок и тёмных векторных заливок — подложки,
            // на которых белый/невидимый текст ЛЕГИТИМЕН (светлые шапки, OCR-слой
            // поверх скана) и скрывать его нельзя.
            List<UglyToad.PdfPig.Content.IPdfImage> pageImages = ListPageImages(page);
            var imageRects = new List<RectPt>(pageImages.Count);
            foreach (UglyToad.PdfPig.Content.IPdfImage img in pageImages)
            {
                RectPt b = RectOf(img.Bounds);
                if (b.Right > b.Left && b.Top > b.Bottom)
                    imageRects.Add(b);
            }
            List<RectPt> darkFills;
            List<PdfLine> lines = ExtractGraphics(page, out darkFills);
            ClipLinesToPage(lines, page.Width, page.Height);
            // Какая ориентация текста станет горизонтальной после поворота страницы
            // пользователем: только её и пускаем в поток (при 0° — прежний фильтр).
            UglyToad.PdfPig.Content.TextOrientation keepOrientation = KeepOrientationFor(rotation);
            List<PdfSourceTextUnit> sourceUnits = SourceUnits(page.Letters,
                keepOrientation);
            Dictionary<UglyToad.PdfPig.Content.Letter, int> sourceOrdinals =
                SourceOrdinals(page.Letters);
            // Разметка структуры страницы, если автор её оставил: по ней видно, какие
            // строки автор считал ОДНИМ абзацем (см. StructureBlocks). Пусто — работаем
            // как раньше, по зазорам.
            Dictionary<UglyToad.PdfPig.Content.Letter, int> blocks = StructureBlocks.Map(page);
            var words = new List<PdfWord>();
            foreach (UglyToad.PdfPig.Content.Word w in page.GetWords())
            {
                RectPt bb = RectOf(w.BoundingBox);
                double size = 0;
                bool bold = false, italic = false;
                int color = 0;
                string family = null;
                string text = w.Text;
                int sourceStart = -1, sourceEnd = -1;
                string sourceText = null;
                bool sourceTrusted = false;
                if (w.Letters != null && w.Letters.Count > 0)
                {
                    UglyToad.PdfPig.Content.Letter first = w.Letters[0];
                    // Повёрнутый текст (вертикальные служебные строки билетов) в поток
                    // не пускаем: построчная сборка рассыпала бы его на посимвольные
                    // «абзацы», вклинивающиеся между обычными строками. При повороте
                    // страницы пользователем горизонтальным станет боковой текст — его
                    // и берём, а прежде-горизонтальный станет боковым и отсеется.
                    if (first.TextOrientation != keepOrientation)
                        continue;
                    // Невидимый текст (белая заливка или режим «не рисовать») ВНЕ
                    // картинок и тёмных плашек — служебные метки («n&#» у
                    // наложенной даты): на белом листе их не видно, в выводе — мусор.
                    // Подложка «спасает» слово, только накрывая его хотя бы
                    // наполовину (OCR-слой скана, белая шапка на цветной плашке);
                    // касание кромки картинки — не подложка.
                    if (IsInvisibleInk(first)
                        && !CoveredByAnyRect(bb, imageRects)
                        && !CoveredByAnyRect(bb, darkFills))
                        continue;
                    bool pseudoBold;
                    text = DedupDoubledGlyphs(w.Letters, text, out pseudoBold);
                    size = first.PointSize;
                    string fn = first.FontName ?? "";
                    // Жирность — из имени шрифта ИЛИ из двойной отрисовки (псевдо-жирный).
                    bold = pseudoBold || fn.IndexOf("Bold", StringComparison.OrdinalIgnoreCase) >= 0;
                    italic = fn.IndexOf("Italic", StringComparison.OrdinalIgnoreCase) >= 0
                          || fn.IndexOf("Oblique", StringComparison.OrdinalIgnoreCase) >= 0;
                    color = ColorArgb(first.Color);
                    family = FontNames.Clean(fn);
                }
                // Значок списка Word из Symbol/Wingdings (U+F0B7 и т.п.) — обычным знаком
                // Юникода (RC-4), в шрифте по умолчанию: в символьном шрифте этого знака
                // нет. «Сравнение» сверяет буквальный текстовый слой — его не трогаем.
                if (mode != PageLayoutMode.Review)
                {
                    string bullet = WordBullets.FromSymbolFont(text, family);
                    if (!ReferenceEquals(bullet, text))
                    {
                        text = bullet;
                        family = null;
                    }
                }
                // Прочие символы приватной зоны юникода (заглушки полей Word, значки вне
                // таблицы WordBullets) не имеют смысла вне своего шрифта — в выводе это «□».
                if (IsPrivateUseOnly(text))
                    continue;
                // Слово, накрытое НИЗКОЙ наложенной картинкой (впечатанные дата и номер
                // поверх пустой формы «____№____»), текстом не выводим: его вид уже
                // несёт картинка, а текст рядом с ней давал бы кашу из двух слоёв.
                if (CoveredByLowImage(bb, imageRects))
                    continue;
                sourceTrusted = TrySourceSpan(w.Letters, sourceOrdinals,
                    sourceUnits, out sourceStart, out sourceEnd, out sourceText) &&
                    string.Equals(sourceText, text, StringComparison.Ordinal);
                // Виртуальная высота прочерков применяется ПОСЛЕ поворота страницы
                // (ApplyUnderscoreHeights): рамка прочерка осмысленна в выправленном
                // пространстве, а при 0° результат прежний.
                double baseX, baseY;
                BaselineOf(w.Letters, out baseX, out baseY);
                int blockId = StructureBlocks.Of(blocks, w.Letters);
                words.Add(new PdfWord
                {
                    Text = text,
                    Left = bb.Left,
                    Right = bb.Right,
                    Bottom = bb.Bottom,
                    Top = bb.Top,
                    BaselineXPt = baseX,
                    BaselineYPt = baseY,
                    BlockId = blockId,
                    SourceStart = sourceTrusted ? sourceStart : -1,
                    SourceEnd = sourceTrusted ? sourceEnd : -1,
                    SourceText = sourceTrusted ? sourceText : null,
                    SourceTrusted = sourceTrusted,
                    FontSizePt = size,
                    Bold = bold,
                    Italic = italic,
                    ColorArgb = color,
                    FontName = family
                });
            }
            AssignHyperlinks(words, page); // рамки ссылок — в исходном пространстве, до поворота
            // Изображения извлекаются ДО поворота: их PNG и фолбэк-кропы Ghostscript
            // живут в исходном пространстве страницы. Поворот ниже развернёт и рамки,
            // и сами пиксели.
            List<OcrImage> images = extractImages
                ? ExtractImages(pageImages, page, path)
                : new List<OcrImage>();
            double pageW = page.Width, pageH = page.Height;
            PageRotation.RotatePage(words, lines, images, rotation, ref pageW, ref pageH);
            ApplyUnderscoreHeights(words);
            // «o» Courier New в начале пункта — значок второго уровня Word (RC-4). После поворота:
            // «начало строки» и «та же строка» осмысленны в выправленном пространстве.
            if (mode != PageLayoutMode.Review)
                WordBullets.MarkCourierBullets(words);
            // Текстовый штамп → переносим картинкой (рендер-кроп), а его слова убираем
            // из потока (иначе штамп задвоится: картинка + те же строки текстом).
            // Печать, нарисованную ТЕКСТОМ, потоковый вывод переносит картинкой: в
            // строку её иначе не поставить. Slide сохраняет текст для редактирования,
            // Review — потому что видимые слова являются семантикой сравнения.
            OcrImage stampImage = mode == PageLayoutMode.Document
                ? ExtractTextStamp(words, page, path, rotation, pageW, pageH)
                : null;
            // Слова таблиц уходят в ячейки; абзацы строятся из ОСТАВШИХСЯ (внетабличных) слов.
            TableDetectResult det = TableDetector.Detect(lines, words, pageW, pageH);
            HashSet<PdfLine> underlines = UnderlineDetector.Mark(det.RemainingWords, lines); // подчёркивания по линовке
            // Одиночные линии, не ставшие ни таблицей, ни подчёркиванием, — прочерки
            // полей («______ №»): переносим текстом-заполнителем, иначе строка
            // полей теряет прочерки и разваливается.
            // Линии-заполнители нужны только потоковому выводу: см. PageLayoutMode.
            if (mode == PageLayoutMode.Document)
                AddRuleWords(det.RemainingWords, det.LoneLines, underlines);
            // Сетки без линовки (чеки, формы «метка … значение») — таблицей без границ.
            GridDetectResult grids = GridDetector.Detect(det.RemainingWords);
            det.Tables.AddRange(grids.Tables);
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(grids.RemainingWords, mode);
            var pt = new PdfPageText
            {
                PageIndex = page.Number - 1, // PdfPig нумерует страницы с 1
                Paragraphs = layout.Paragraphs,
                FirstLineIndentPt = layout.FirstLineIndentPt,
                WidthPt = pageW,
                HeightPt = pageH,
                Lines = lines,
                Tables = det.Tables,
                Words = words,
                SourceUnits = sourceUnits,
                LayoutMode = mode
            };
            pt.Images = images;
            if (stampImage != null)
                pt.Images.Add(stampImage); // штамп — в общий поток изображений (порядок по TopPt)
            SetMargins(pt, words, pt.Images, pageW, pageH);
            return pt;
        }

        private const double UnderscoreMinHeightFactor = 0.35; // виртуальная высота прочерка в долях кегля

        private sealed class LetterReferenceComparer :
            IEqualityComparer<UglyToad.PdfPig.Content.Letter>
        {
            public bool Equals(UglyToad.PdfPig.Content.Letter x,
                UglyToad.PdfPig.Content.Letter y)
            {
                return object.ReferenceEquals(x, y);
            }

            public int GetHashCode(UglyToad.PdfPig.Content.Letter value)
            {
                return value == null ? 0 : RuntimeHelpers.GetHashCode(value);
            }
        }

        /// <summary>
        /// Снять только декодированное значение и пригодность ориентации до GetWords().
        /// Геометрия и предполагаемые зазоры сюда намеренно не входят.
        /// </summary>
        private static List<PdfSourceTextUnit> SourceUnits(
            IReadOnlyList<UglyToad.PdfPig.Content.Letter> letters,
            UglyToad.PdfPig.Content.TextOrientation keepOrientation)
        {
            var result = new List<PdfSourceTextUnit>(letters == null ? 0 : letters.Count);
            if (letters == null)
                return result;
            for (int i = 0; i < letters.Count; i++)
            {
                UglyToad.PdfPig.Content.Letter letter = letters[i];
                string value = letter == null ? null : letter.Value;
                result.Add(new PdfSourceTextUnit
                {
                    Text = value,
                    Trusted = letter != null && value != null && value.Length > 0 &&
                        letter.TextOrientation == keepOrientation
                });
            }
            return result;
        }

        private static Dictionary<UglyToad.PdfPig.Content.Letter, int> SourceOrdinals(
            IReadOnlyList<UglyToad.PdfPig.Content.Letter> letters)
        {
            var result = new Dictionary<UglyToad.PdfPig.Content.Letter, int>(
                new LetterReferenceComparer());
            if (letters == null)
                return result;
            for (int i = 0; i < letters.Count; i++)
            {
                UglyToad.PdfPig.Content.Letter letter = letters[i];
                if (letter == null)
                    continue;
                int previous;
                if (result.TryGetValue(letter, out previous))
                    result[letter] = -1;
                else
                    result.Add(letter, i);
            }
            return result;
        }

        /// <summary>
        /// Слово получает source-span лишь когда его Letter — уникальные соседние unit
        /// исходной страницы, без пробельных/пустых значений и смены ориентации.
        /// Любое расхождение заставляет Review воздержаться от whitespace-семантики.
        /// </summary>
        private static bool TrySourceSpan(
            IReadOnlyList<UglyToad.PdfPig.Content.Letter> letters,
            IDictionary<UglyToad.PdfPig.Content.Letter, int> ordinals,
            IList<PdfSourceTextUnit> units, out int start, out int end,
            out string text)
        {
            start = end = -1;
            text = null;
            if (letters == null || letters.Count == 0 || ordinals == null || units == null)
                return false;
            var value = new System.Text.StringBuilder();
            int previous = -1;
            for (int i = 0; i < letters.Count; i++)
            {
                UglyToad.PdfPig.Content.Letter letter = letters[i];
                int ordinal;
                if (letter == null || !ordinals.TryGetValue(letter, out ordinal) ||
                    ordinal < 0 || ordinal >= units.Count ||
                    (i > 0 && ordinal != previous + 1))
                    return false;
                PdfSourceTextUnit unit = units[ordinal];
                if (unit == null || !unit.Trusted || string.IsNullOrEmpty(unit.Text) ||
                    ContainsWhitespace(unit.Text))
                    return false;
                if (i == 0)
                    start = ordinal;
                previous = ordinal;
                value.Append(unit.Text);
            }
            end = previous;
            text = value.ToString();
            return text.Length > 0;
        }

        private static bool ContainsWhitespace(string value)
        {
            if (string.IsNullOrEmpty(value))
                return false;
            for (int i = 0; i < value.Length; i++)
                if (char.IsWhiteSpace(value[i]))
                    return true;
            return false;
        }

        /// <summary>
        /// Ориентация текста PdfPig, которая станет горизонтальной после поворота страницы
        /// на rotation по часовой: глиф с ориентацией O после поворота видится как O + R,
        /// горизонталь получается при O = 360 − R. Вызывать из ядра (тип PdfPig).
        /// </summary>
        private static UglyToad.PdfPig.Content.TextOrientation KeepOrientationFor(int rotation)
        {
            switch (rotation)
            {
                case 90: return UglyToad.PdfPig.Content.TextOrientation.Rotate270;
                case 180: return UglyToad.PdfPig.Content.TextOrientation.Rotate180;
                case 270: return UglyToad.PdfPig.Content.TextOrientation.Rotate90;
                default: return UglyToad.PdfPig.Content.TextOrientation.Horizontal;
            }
        }

        /// <summary>
        /// Прочерк «____» рисуется у базовой линии рамкой в пару pt: такой вырожденный бокс
        /// не пересекается по вертикали с соседями («№», «от») и отрывался от них в отдельные
        /// строки. Даём прочерку высоту хотя бы трети кегля — как у обычного слова той же
        /// строки. Вызывается ПОСЛЕ поворота страницы (в выправленном пространстве).
        /// Чистая — под тест.
        /// </summary>
        internal static void ApplyUnderscoreHeights(List<PdfWord> words)
        {
            if (words == null)
                return;
            foreach (PdfWord w in words)
                if (w.FontSizePt > 0
                    && w.Top - w.Bottom < UnderscoreMinHeightFactor * w.FontSizePt
                    && IsUnderscoreOnly(w.Text))
                    w.Top = w.Bottom + UnderscoreMinHeightFactor * w.FontSizePt;
        }

        // Прочерк из одиночной линии: пределы толщины/длины и ширина «_» в долях кегля (TNR ≈ 0.5).
        // Толщина до 6 pt: прочерки полей встречаются и «жирным» штрихом (5 pt), а заливки
        // толще этого ExtractLines в линовку уже не пускает.
        private const double RuleMaxThicknessPt = 6;
        private const double RuleMinLenPt = 6;
        private const double UnderscoreAdvanceEm = 0.5;
        private const int RuleMaxChars = 120; // страховка от линейки во всю страницу при крошечном кегле
        // Подъём рамки прочерка над линией: линия прочерка лежит НИЖЕ базовой линии соседних
        // слов («№», «от»), и трети кегля не хватает, чтобы рамки пересеклись и строка собралась.
        private const double RuleAscentFactor = 0.6;

        /// <summary>
        /// Превратить одиночные горизонтальные линии (не таблица и не подчёркивание) в
        /// слова-прочерки «____»: линии-заполнители полей («______ № ______»)
        /// иначе теряются вовсе, и строка полей разваливается. Число «_» подбирается по
        /// длине линии и типичному кеглю слов страницы; рамка прочерка — от линии вверх на треть
        /// кегля, чтобы строка собралась вместе с соседними словами. Чистая — под тест.
        /// </summary>
        internal static void AddRuleWords(List<PdfWord> words, List<PdfLine> loneLines, HashSet<PdfLine> underlineLines)
        {
            if (words == null || loneLines == null || loneLines.Count == 0)
                return;
            var sizes = new List<double>(words.Count);
            foreach (PdfWord w in words)
                if (w.FontSizePt > 0)
                    sizes.Add(w.FontSizePt);
            double size = MathUtil.Median(sizes);
            if (size <= 0) size = 10;

            var added = new List<PdfWord>(); // против задвоения: линию нередко рисуют дважды (туда-обратно)
            foreach (PdfLine line in loneLines)
            {
                if (line.Orientation != LineOrientation.Horizontal
                    || (underlineLines != null && underlineLines.Contains(line))
                    || line.Thickness > RuleMaxThicknessPt
                    || line.Length < RuleMinLenPt)
                    continue;
                bool duplicate = false;
                foreach (PdfWord prev in added)
                {
                    double overlap = Math.Min(prev.Right, line.MaxX) - Math.Max(prev.Left, line.MinX);
                    if (overlap > 0.5 * line.Length && Math.Abs(prev.Bottom - line.Position) <= 1.5)
                    {
                        duplicate = true;
                        break;
                    }
                }
                if (duplicate)
                    continue;
                int n = (int)Math.Round(line.Length / (UnderscoreAdvanceEm * size));
                if (n < 2) n = 2;
                if (n > RuleMaxChars) n = RuleMaxChars;
                var rule = new PdfWord
                {
                    Text = new string('_', n),
                    Left = line.MinX,
                    Right = line.MaxX,
                    Bottom = line.Position,
                    Top = line.Position + RuleAscentFactor * size,
                    FontSizePt = size
                };
                words.Add(rule);
                added.Add(rule);
            }
        }

        /// <summary>Слово целиком из «_» (прочерк-заполнитель поля)?</summary>
        internal static bool IsUnderscoreOnly(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;
            for (int i = 0; i < text.Length; i++)
                if (text[i] != '_')
                    return false;
            return true;
        }

        /// <summary>Все символы слова — из приватной зоны юникода (U+E000–U+F8FF)?</summary>
        internal static bool IsPrivateUseOnly(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;
            bool anyPua = false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (char.IsWhiteSpace(c))
                    continue;
                if (c < 0xE000 || c > 0xF8FF)
                    return false;
                anyPua = true;
            }
            return anyPua;
        }

        /// <summary>
        /// Снять сдвоенные глифы «псевдо-жирного» с букв слова (чистая логика — в
        /// <see cref="GlyphDedup"/>): текст пересобирается из оставшихся букв, pseudoBold —
        /// двойная отрисовка была массовой (слово визуально жирное). Без дублей текст
        /// возвращается как есть. Вызывать из ядра — тело ссылается на типы PdfPig.
        /// </summary>
        private static string DedupDoubledGlyphs(System.Collections.Generic.IReadOnlyList<UglyToad.PdfPig.Content.Letter> letters,
            string text, out bool pseudoBold)
        {
            pseudoBold = false;
            if (letters.Count < 2)
                return text;
            var glyphs = new GlyphInfo[letters.Count];
            for (int i = 0; i < letters.Count; i++)
            {
                UglyToad.PdfPig.Content.Letter l = letters[i];
                UglyToad.PdfPig.Core.PdfRectangle r = l.GlyphRectangle;
                glyphs[i] = new GlyphInfo
                {
                    Value = l.Value,
                    CenterX = (r.Left + r.Right) / 2,
                    CenterY = (r.Bottom + r.Top) / 2,
                    SizePt = l.PointSize
                };
            }
            int dropped;
            List<int> keep = GlyphDedup.Keep(glyphs, out dropped);
            if (dropped == 0)
                return text;
            var sb = new System.Text.StringBuilder(keep.Count);
            foreach (int k in keep)
                sb.Append(letters[k].Value);
            pseudoBold = GlyphDedup.LooksBold(keep.Count, dropped);
            return sb.ToString();
        }

        /// <summary>
        /// Начало базовой линии слова — МЕДИАНА по буквам. Над- и подстрочные знаки стоят на
        /// своих линиях, и одна такая буква увела бы и среднее, и крайнее значение, а медиана
        /// остаётся на линии основного текста. Нули — букв нет (слово синтезировано).
        /// </summary>
        private static void BaselineOf(IReadOnlyList<UglyToad.PdfPig.Content.Letter> letters,
            out double x, out double y)
        {
            x = 0;
            y = 0;
            if (letters == null || letters.Count == 0)
                return;
            var ys = new List<double>(letters.Count);
            var xs = new List<double>(letters.Count);
            for (int i = 0; i < letters.Count; i++)
            {
                UglyToad.PdfPig.Core.PdfPoint p = letters[i].StartBaseLine;
                ys.Add(p.Y);
                xs.Add(p.X);
            }
            ys.Sort();
            xs.Sort();
            y = ys[ys.Count / 2];
            x = xs[xs.Count / 2];
        }

        /// <summary>
        /// Пометить слова гиперссылками: если центр слова попадает в рамку ссылки страницы —
        /// проставить её URI. Сбой получения ссылок не срывает извлечение. Вызывать из ядра.
        /// </summary>
        private static void AssignHyperlinks(List<PdfWord> words, UglyToad.PdfPig.Content.Page page)
        {
            System.Collections.Generic.IReadOnlyList<UglyToad.PdfPig.Content.Hyperlink> links;
            try { links = page.GetHyperlinks(); }
            catch { return; }
            if (links == null || links.Count == 0)
                return;
            // Рамки и адреса — один раз на страницу, а не на каждое слово. Опасный или
            // относительный адрес выбрасывается здесь же: слово остаётся обычным текстом.
            var rects = new List<RectPt>(links.Count);
            var uris = new List<string>(links.Count);
            foreach (UglyToad.PdfPig.Content.Hyperlink h in links)
            {
                try
                {
                    string uri = SafeLinkUri(h.Uri);
                    if (uri == null)
                        continue;
                    rects.Add(RectOf(h.Bounds));
                    uris.Add(uri);
                }
                catch { } // битая аннотация — без ссылки, остальные целы
            }
            if (uris.Count == 0)
                return;
            foreach (PdfWord w in words)
            {
                double mx = (w.Left + w.Right) / 2, my = (w.Top + w.Bottom) / 2;
                for (int i = 0; i < rects.Count; i++)
                {
                    RectPt b = rects[i];
                    if (mx >= b.Left && mx <= b.Right && my >= b.Bottom && my <= b.Top)
                    {
                        w.Uri = uris[i];
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Адрес ссылки PDF, годный для внешней связи .docx/.pptx, или null. Адрес из файла —
        /// чужие данные: javascript:, file:, UNC-путь «\\host\share» или обработчик «ms-…:»
        /// стали бы в документе кликабельной внешней связью — запуском кода, утечкой
        /// NTLM-хэша на чужой сервер, открытием локального файла. Пропускаем только
        /// абсолютные http/https/mailto и отдаём их в каноничной экранированной форме
        /// (AbsoluteUri): писатели кладут адрес в связь как есть. Чистая — под тест.
        /// </summary>
        internal static string SafeLinkUri(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return null;
            Uri uri;
            if (!Uri.TryCreate(raw.Trim(), UriKind.Absolute, out uri))
                return null;
            if (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps &&
                uri.Scheme != Uri.UriSchemeMailto)
                return null;
            return uri.AbsoluteUri;
        }

        /// <summary>
        /// Растровые изображения страницы в PNG с рамками. Недекодируемые форматы (JBIG2,
        /// CCITT, CMYK-JPEG и т.п.) пропускаются; сбой одной картинки не срывает остальные и
        /// текст. Вызывать из ядра (после Ensure) — тело ссылается на типы PdfPig.
        /// </summary>
        private static List<OcrImage> ExtractImages(List<UglyToad.PdfPig.Content.IPdfImage> images,
            UglyToad.PdfPig.Content.Page page, string pdfPath)
        {
            var result = new List<OcrImage>();
            BudgetedBitmap pageRasterOwned = null;
            System.Drawing.Bitmap pageRaster = null; // ленивый рендер страницы — только если понадобится фолбэк
            bool rasterTried = false;
            try
            {
                foreach (UglyToad.PdfPig.Content.IPdfImage img in images)
                {
                    try
                    {
                        UglyToad.PdfPig.Core.PdfRectangle bounds = img.Bounds;
                        RectPt b = RectOf(bounds);
                        double width = b.Right - b.Left, height = b.Top - b.Bottom;
                        if (width <= 0 || height <= 0)
                            continue;
                        // Прямой декод отдаёт пиксели в СОБСТВЕННОЙ ориентации картинки, а на
                        // странице (/Rotate 90, повёрнутая матрица) она стоит повёрнутой — разворачиваем
                        // так, как её видит читатель. Кроп Ghostscript уже повёрнут рендером.
                        int turn = ImageQuarterTurn(bounds.BottomLeft.X, bounds.BottomLeft.Y,
                            bounds.BottomRight.X, bounds.BottomRight.Y, bounds.TopLeft.X, bounds.TopLeft.Y);

                        // Картинка с мягкой маской (SMask/Mask) — прозрачность: PdfPig её игнорирует
                        // и заливает прозрачные области ЧЁРНЫМ (в Word получается чёрный фон у логотипа/
                        // печати). Такую переносим рендер-кропом страницы Ghostscript-ом — он композитит
                        // маску на белое, как она выглядит на странице. Прямой декод для неё пропускаем.
                        bool masked = HasSoftMask(img);
                        byte[] png = masked ? null : PageRotation.RotatePng(EncodePng(img), turn);
                        // Декодер не справился (null) или выдал одноцветный мусор (напр. штрих-код
                        // чёрным ящиком) — фолбэк: рендерим страницу Ghostscript-ом и вырезаем картинку
                        // по её рамке. Так переносится любая картинка, как она выглядит.
                        if (png == null || png.Length == 0 || RasterUtil.IsSolidColor(png))
                        {
                            if (!rasterTried)
                            {
                                pageRasterOwned = PageRasterizer.RenderPage(pdfPath,
                                    page.Number, page.Width, page.Height);
                                pageRaster = pageRasterOwned == null
                                    ? null : pageRasterOwned.Bitmap;
                                rasterTried = true;
                            }
                            byte[] cropped = PageRasterizer.CropRegion(pageRaster, page.Width, page.Height,
                                b.Left, b.Top, width, height);
                            if (cropped != null && !RasterUtil.IsSolidColor(cropped))
                                png = cropped;
                            else if (masked && (png == null || png.Length == 0))
                                png = PageRotation.RotatePng(EncodePng(img), turn); // нет Ghostscript — маскированную берём как есть (может быть чёрный фон), лучше чем терять
                            if (png == null || png.Length == 0 || RasterUtil.IsSolidColor(png))
                                continue; // и фолбэк не помог (нет GS / область пустая) — пропускаем
                        }

                        result.Add(new OcrImage
                        {
                            Png = png,
                            LeftPt = b.Left,
                            TopPt = b.Top,
                            WidthPt = width,
                            HeightPt = height
                        });
                    }
                    catch (OutOfMemoryException) { throw; }
                    catch { } // одна битая картинка не ломает остальные
                }
            }
            finally
            {
                if (pageRasterOwned != null)
                    pageRasterOwned.Dispose();
            }
            return result;
        }

        /// <summary>
        /// Есть ли у картинки мягкая маска прозрачности (ключи SMask/Mask в словаре). PdfPig
        /// такую маску не применяет — прозрачные области выходят чёрными. Сбой доступа к словарю —
        /// считаем, что маски нет (обычный путь). Вызывать из ядра — тело ссылается на типы PdfPig.
        /// </summary>
        private static bool HasSoftMask(UglyToad.PdfPig.Content.IPdfImage img)
        {
            try
            {
                // Словарь картинки и его ключи — типы вендоренной сборки Tokens (не в ссылках проекта),
                // поэтому через dynamic; ключ — NameToken, сравниваем по строковому виду («SMask»/«Mask»).
                dynamic dict = img.ImageDictionary;
                if (dict == null)
                    return false;
                foreach (dynamic kv in dict.Data)
                {
                    string k = kv.Key.ToString();
                    if (k == "SMask" || k == "Mask")
                        return true;
                }
            }
            catch { }
            return false;
        }

        private const double StampCropPadPt = 6.0; // отступ кропа, чтобы захватить рамку печати вокруг текста

        /// <summary>
        /// Распознать текстовый штамп и вернуть его как изображение (рендер-кроп области
        /// Ghostscript-ом), убрав слова штампа из <paramref name="words"/>. Работает в ТЕКУЩЕМ
        /// пространстве страницы (после пользовательского поворота: pageW/pageH — возможно,
        /// поменянные местами); кроп растеризатора выполняется в исходном пространстве через
        /// обратный маппинг, пиксели затем доворачиваются. Если штамп не найден или картинку
        /// получить не удалось (нет Ghostscript, пустая область) — возвращает null и НИЧЕГО не
        /// убирает: текст остаётся текстом (прежнее поведение, без регрессии). Вызывать из ядра
        /// (после Ensure) — тело ссылается на типы PdfPig.
        /// </summary>
        private static OcrImage ExtractTextStamp(List<PdfWord> words, UglyToad.PdfPig.Content.Page page,
            string pdfPath, int rotation, double pageW, double pageH)
        {
            StampRegion stamp = StampDetector.Detect(words, pageW, pageH);
            if (stamp == null)
                return null;

            double left = stamp.Left - StampCropPadPt, top = stamp.Top + StampCropPadPt;
            double width = stamp.Width + 2 * StampCropPadPt, height = stamp.Height + 2 * StampCropPadPt;
            BudgetedBitmap rasterOwned = null;
            System.Drawing.Bitmap raster = null;
            try
            {
                rasterOwned = PageRasterizer.RenderPage(pdfPath, page.Number,
                    page.Width, page.Height);
                raster = rasterOwned == null ? null : rasterOwned.Bitmap;
                if (raster == null)
                    return null; // нет Ghostscript — штамп остаётся текстом
                // Область кропа — в исходном (неповёрнутом) пространстве страницы: рендер
                // Ghostscript ничего не знает о пользовательском повороте.
                double cropLeft = left, cropTop = top, cropW = width, cropH = height;
                if (rotation != 0)
                {
                    double ol, ob, or2, ot;
                    PageRotation.MapBox(left, top - height, left + width, top,
                        PageRotation.Inverse(rotation), pageW, pageH, out ol, out ob, out or2, out ot);
                    cropLeft = ol;
                    cropTop = ot;
                    cropW = or2 - ol;
                    cropH = ot - ob;
                }
                byte[] png = PageRasterizer.CropRegion(raster, page.Width, page.Height, cropLeft, cropTop, cropW, cropH);
                if (png == null || png.Length == 0)
                    return null; // рендер/кроп не удался — не переносим, текст остаётся текстом
                png = PageRotation.RotatePng(png, rotation); // пиксели — в выправленное пространство
                // Проверку на одноцветность НЕ делаем: штамп — это белый фон с тонким текстом и рамкой,
                // разреженная выборка приняла бы его за пустой; наличие текста уже доказано StampDetector.

                var remove = new HashSet<PdfWord>(stamp.Words);
                words.RemoveAll(delegate(PdfWord w) { return remove.Contains(w); }); // снять текст печати
                return new OcrImage { Png = png, LeftPt = left, TopPt = top, WidthPt = width, HeightPt = height };
            }
            catch (OutOfMemoryException) { throw; }
            catch { return null; } // сбой рендера/кропа — штамп остаётся текстом
            finally
            {
                if (rasterOwned != null) rasterOwned.Dispose();
            }
        }

        /// <summary>
        /// PNG для изображения PDF: сначала штатный TryGetPng; если он не смог (частый случай —
        /// DCTDecode/JPEG, для которого PdfPig не строит PNG), декодируем сырые байты потока через
        /// GDI (System.Drawing понимает JPEG/PNG/GIF/BMP) и пересохраняем в PNG. Возвращает null,
        /// если ни один путь не дал картинку (напр. сырьё — недекодируемый формат).
        /// </summary>
        private static byte[] EncodePng(UglyToad.PdfPig.Content.IPdfImage img)
        {
            try
            {
                byte[] direct;
                if (img.TryGetPng(out direct) && direct != null && direct.Length > 0)
                    return direct;
            }
            catch { } // TryGetPng может кинуть на экзотическом формате — уходим в фолбэк

            try
            {
                byte[] raw = img.RawBytes.ToArray(); // ReadOnlyMemory<byte> -> массив (напр. JPEG)
                if (raw.Length == 0)
                    return null;
                using (var ms = new MemoryStream(raw))
                using (System.Drawing.Image bmp = System.Drawing.Image.FromStream(ms))
                using (var outMs = new MemoryStream())
                {
                    bmp.Save(outMs, System.Drawing.Imaging.ImageFormat.Png);
                    return outMs.ToArray();
                }
            }
            catch { return null; } // сырьё не декодируется в растр — пропускаем картинку
        }

        // Классификация отрезка как линовки: отклонение от оси (pt) и минимальная длина.
        private const double LineAxisTol = 1.0; // |dy| для горизонтали / |dx| для вертикали
        private const double MinLineLen = 2.0;  // короче — графический шум, не линовка
        private const double ThinFillMax = 2.5; // залитая полоска тоньше этого — это линия, не фон
        // Тёмная векторная плашка: заливка темнее этого порога (сумма RGB) и крупнее пары em² —
        // подложка, на которой белый текст легитимен (тёмные шапки таблиц).
        private const double DarkFillMaxRgbSum = 1.5;
        private const double DarkFillMinSidePt = 6.0;

        /// <summary>Прямоугольник в координатах PDF (ось Y вверх) — рамки картинок/плашек.</summary>
        internal struct RectPt
        {
            public double Left;
            public double Bottom;
            public double Right;
            public double Top;
        }

        /// <summary>
        /// Рамка PdfPig → нормализованный прямоугольник по min/max ЧЕТЫРЁХ углов. PdfPig 0.1.9
        /// на странице с /Rotate 90/270 (и у объектов под повёрнутой матрицей) отдаёт рамки
        /// картинок, ссылок и боковых слов перевёрнутыми: Left больше Right или Bottom больше
        /// Top, а Width/Height — в собственной ориентации объекта. С такой рамкой перекрытия
        /// выходили отрицательными, центр слова в ссылку не попадал, картинка вставала не туда.
        /// </summary>
        private static RectPt RectOf(UglyToad.PdfPig.Core.PdfRectangle b)
        {
            return NormalizedRect(b.BottomLeft.X, b.BottomLeft.Y, b.BottomRight.X, b.BottomRight.Y,
                b.TopLeft.X, b.TopLeft.Y, b.TopRight.X, b.TopRight.Y);
        }

        /// <summary>Прямоугольник, накрывающий четыре точки (min/max по осям). Чистая — под тест.</summary>
        internal static RectPt NormalizedRect(double x1, double y1, double x2, double y2,
            double x3, double y3, double x4, double y4)
        {
            return new RectPt
            {
                Left = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4)),
                Right = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4)),
                Bottom = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4)),
                Top = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4))
            };
        }

        /// <summary>
        /// На сколько градусов по часовой повернуть декодированные пиксели картинки, чтобы она
        /// легла так, как стоит на странице. Рамка картинки у PdfPig — образ единичного квадрата:
        /// нижний-левый угол — точка (0,0) картинки, нижний-правый — (1,0); направление между
        /// ними — ось X картинки на странице. Ось смотрит вправо — 0, вниз — 90, влево — 180,
        /// вверх — 270. Зеркальная картинка (верхний-левый угол по другую сторону оси X, частый
        /// «переворот по вертикали» генераторов) — 0, как и прежде: поворот отражение не
        /// исправит, а лишь заменит одну неправду другой. Чистая — под тест.
        /// </summary>
        internal static int ImageQuarterTurn(double bottomLeftX, double bottomLeftY,
            double bottomRightX, double bottomRightY, double topLeftX, double topLeftY)
        {
            double dx = bottomRightX - bottomLeftX, dy = bottomRightY - bottomLeftY;
            if (double.IsNaN(dx) || double.IsNaN(dy))
                return 0;
            // Ориентация пары осей картинки (X — к нижнему-правому, Y — к верхнему-левому):
            // отрицательная — отражение.
            double cross = dx * (topLeftY - bottomLeftY) - dy * (topLeftX - bottomLeftX);
            if (!(cross > 0))
                return 0;
            if (Math.Abs(dx) >= Math.Abs(dy))
                return dx >= 0 ? 0 : 180;
            return dy < 0 ? 90 : 270;
        }

        /// <summary>Картинки страницы одним списком (без декодирования). Сбой перечисления — пусто.</summary>
        private static List<UglyToad.PdfPig.Content.IPdfImage> ListPageImages(UglyToad.PdfPig.Content.Page page)
        {
            var images = new List<UglyToad.PdfPig.Content.IPdfImage>();
            try
            {
                foreach (UglyToad.PdfPig.Content.IPdfImage img in page.GetImages())
                    images.Add(img);
            }
            catch { }
            return images;
        }

        /// <summary>
        /// Невидимая отрисовка буквы: белая заливка (на белом листе не видна) или текстовый
        /// режим «не рисовать» (Neither — OCR-слой, служебные метки). Вызывать из ядра.
        /// </summary>
        private static bool IsInvisibleInk(UglyToad.PdfPig.Content.Letter letter)
        {
            try
            {
                if (letter.RenderingMode == UglyToad.PdfPig.Core.TextRenderingMode.Neither)
                    return true;
                if (letter.Color == null)
                    return false;
                var rgb = letter.Color.ToRGBValues();
                return rgb.r >= 0.95 && rgb.g >= 0.95 && rgb.b >= 0.95;
            }
            catch { return false; } // не разобрали цвет/режим — считаем видимым (не теряем текст)
        }

        /// <summary>
        /// Накрыт ли прямоугольник слова каким-либо из rects хотя бы наполовину по ОБЕИМ осям.
        /// Слово «на подложке» — когда подложка реально под ним, а не касается кромкой.
        /// Чистая — под тест.
        /// </summary>
        internal static bool CoveredByAnyRect(RectPt word, List<RectPt> rects)
        {
            double wWidth = word.Right - word.Left, wHeight = word.Top - word.Bottom;
            foreach (RectPt r in rects)
            {
                double xOverlap = Math.Min(word.Right, r.Right) - Math.Max(word.Left, r.Left);
                double yOverlap = Math.Min(word.Top, r.Top) - Math.Max(word.Bottom, r.Bottom);
                bool xOk = wWidth <= 0 ? xOverlap >= 0 : xOverlap >= 0.5 * wWidth;
                bool yOk = wHeight <= 0 ? yOverlap >= 0 : yOverlap >= 0.5 * wHeight;
                if (xOk && yOk)
                    return true;
            }
            return false;
        }

        // «Наложение»: низкая картинка (впечатанные дата/номер) поверх строки-формы.
        private const double LowImageMaxHeightPt = 24.0;
        private const double CoverShare = 0.5; // картинка накрывает не меньше этой доли слова по обеим осям

        /// <summary>
        /// Накрыто ли слово низкой наложенной картинкой: картинка не выше
        /// <see cref="LowImageMaxHeightPt"/> и перекрывает слово не меньше чем на половину по
        /// ширине и высоте. Высокие картинки (схема с номерами мест, штампы, логотипы) под
        /// правило не попадают — текст поверх них легитимен. Чистая — под тест.
        /// </summary>
        internal static bool CoveredByLowImage(RectPt word, List<RectPt> images)
        {
            double wWidth = word.Right - word.Left, wHeight = word.Top - word.Bottom;
            if (wWidth <= 0)
                return false;
            foreach (RectPt img in images)
            {
                if (img.Top - img.Bottom > LowImageMaxHeightPt)
                    continue;
                double xOverlap = Math.Min(word.Right, img.Right) - Math.Max(word.Left, img.Left);
                if (xOverlap < CoverShare * wWidth)
                    continue;
                double yOverlap = Math.Min(word.Top, img.Top) - Math.Max(word.Bottom, img.Bottom);
                if (wHeight <= 0 ? yOverlap >= 0 : yOverlap >= CoverShare * wHeight)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Линовка страницы (границы таблиц, подчёркивания) из векторной графики; попутно —
        /// рамки ТЁМНЫХ заливок (подложки белого текста). Берём только строго
        /// горизонтальные/вертикальные штрихи и рёбра прямоугольников; диагонали, кривые
        /// и крупные заливки (фон ячеек) в линовку не идут. Сбой одной фигуры не срывает
        /// остальные и текст. Вызывать из ядра (после Ensure) — тело ссылается на типы PdfPig.
        /// </summary>
        private static List<PdfLine> ExtractGraphics(UglyToad.PdfPig.Content.Page page, out List<RectPt> darkFills)
        {
            var result = new List<PdfLine>();
            darkFills = new List<RectPt>();
            System.Collections.Generic.IReadOnlyList<UglyToad.PdfPig.Graphics.PdfPath> paths;
            try { paths = page.ExperimentalAccess.Paths; }
            catch { return result; }
            if (paths == null)
                return result;
            foreach (UglyToad.PdfPig.Graphics.PdfPath path in paths)
            {
                try
                {
                    if (path == null || path.IsClipping)
                        continue; // невидимая обтравка — не ink
                    bool stroked = path.IsStroked;
                    bool filled = path.IsFilled;
                    if (!stroked && !filled)
                        continue;
                    double thickness = path.LineWidth;
                    for (int i = 0; i < path.Count; i++)
                    {
                        UglyToad.PdfPig.Core.PdfSubpath sub = path[i];
                        if (sub == null)
                            continue;
                        if (stroked && sub.IsDrawnAsRectangle)
                            AddRectangleEdges(result, sub, thickness);
                        else if (stroked)
                            AddLineCommands(result, sub, thickness);
                        else if (sub.IsDrawnAsRectangle) // filled, не stroked
                        {
                            AddThinFilledRect(result, sub);
                            AddDarkFill(darkFills, path, sub);
                        }
                    }
                }
                catch { } // одна битая фигура не ломает остальные
            }
            return result;
        }

        /// <summary>Крупная ТЁМНАЯ заливка → подложка (на ней белый текст легитимен).</summary>
        private static void AddDarkFill(List<RectPt> darkFills, UglyToad.PdfPig.Graphics.PdfPath path, UglyToad.PdfPig.Core.PdfSubpath sub)
        {
            try
            {
                UglyToad.PdfPig.Core.PdfRectangle? r = sub.GetDrawnRectangle();
                if (r == null)
                    return;
                UglyToad.PdfPig.Core.PdfRectangle rect = r.Value;
                if (rect.Width < DarkFillMinSidePt || rect.Height < DarkFillMinSidePt)
                    return; // мелочь/линии — не подложка
                var fill = path.FillColor;
                if (fill == null)
                    return;
                var rgb = fill.ToRGBValues();
                if (rgb.r + rgb.g + rgb.b <= DarkFillMaxRgbSum)
                    darkFills.Add(RectOf(rect));
            }
            catch { }
        }

        /// <summary>Рёбра нарисованного прямоугольника → 4 линии (границы ячейки/рамки).</summary>
        private static void AddRectangleEdges(List<PdfLine> result, UglyToad.PdfPig.Core.PdfSubpath sub, double thickness)
        {
            UglyToad.PdfPig.Core.PdfRectangle? r = sub.GetDrawnRectangle();
            if (r == null)
                return;
            UglyToad.PdfPig.Core.PdfRectangle rect = r.Value;
            AddSegment(result, rect.Left, rect.Top, rect.Right, rect.Top, thickness);
            AddSegment(result, rect.Left, rect.Bottom, rect.Right, rect.Bottom, thickness);
            AddSegment(result, rect.Left, rect.Bottom, rect.Left, rect.Top, thickness);
            AddSegment(result, rect.Right, rect.Bottom, rect.Right, rect.Top, thickness);
        }

        /// <summary>Прямые команды Line подпути → линовка (диагонали отсеются в AddSegment).</summary>
        private static void AddLineCommands(List<PdfLine> result, UglyToad.PdfPig.Core.PdfSubpath sub, double thickness)
        {
            System.Collections.Generic.IReadOnlyList<UglyToad.PdfPig.Core.PdfSubpath.IPathCommand> cmds = sub.Commands;
            if (cmds == null)
                return;
            foreach (UglyToad.PdfPig.Core.PdfSubpath.IPathCommand c in cmds)
            {
                var line = c as UglyToad.PdfPig.Core.PdfSubpath.Line;
                if (line != null)
                    AddSegment(result, line.From.X, line.From.Y, line.To.X, line.To.Y, thickness);
            }
        }

        /// <summary>Тонкая залитая полоска (линия, нарисованная заливкой) → одна линия по длинной оси.</summary>
        private static void AddThinFilledRect(List<PdfLine> result, UglyToad.PdfPig.Core.PdfSubpath sub)
        {
            UglyToad.PdfPig.Core.PdfRectangle? r = sub.GetDrawnRectangle();
            if (r == null)
                return;
            UglyToad.PdfPig.Core.PdfRectangle rect = r.Value;
            double w = rect.Width, h = rect.Height;
            if (Math.Min(w, h) > ThinFillMax)
                return; // крупная заливка — фон ячейки, а не линия
            double midX = (rect.Left + rect.Right) / 2, midY = (rect.Bottom + rect.Top) / 2;
            if (w >= h)
                AddSegment(result, rect.Left, midY, rect.Right, midY, Math.Min(w, h));
            else
                AddSegment(result, midX, rect.Bottom, midX, rect.Top, Math.Min(w, h));
        }

        /// <summary>
        /// Обрезать линовку по окрестности листа (лист плюс ещё по его размеру с каждой
        /// стороны — запас на смещённый MediaBox). Отрезок далеко за листом невидим, но
        /// стоит дорого: поток в сотню байт «0 0 m 1000000000 0 l S» давал линию на миллиард
        /// пунктов, и детекторы раскладывали её на десятки миллионов клеток — зависание и
        /// нехватка памяти. Линия целиком вне окрестности или NaN — выбрасывается, торчащая —
        /// укорачивается. Чистая — под тест.
        /// </summary>
        internal static void ClipLinesToPage(List<PdfLine> lines, double pageW, double pageH)
        {
            if (lines == null || !(pageW > 0) || !(pageH > 0))
                return;
            double minX = -pageW, maxX = 2 * pageW, minY = -pageH, maxY = 2 * pageH;
            lines.RemoveAll(delegate(PdfLine l)
            {
                bool horizontal = l.Orientation == LineOrientation.Horizontal;
                double pos = l.Position;
                if (horizontal ? !(pos >= minY && pos <= maxY) : !(pos >= minX && pos <= maxX))
                    return true;
                if (horizontal)
                {
                    l.X1 = ClampTo(l.X1, minX, maxX);
                    l.X2 = ClampTo(l.X2, minX, maxX);
                }
                else
                {
                    l.Y1 = ClampTo(l.Y1, minY, maxY);
                    l.Y2 = ClampTo(l.Y2, minY, maxY);
                }
                return !(l.Length >= MinLineLen); // NaN-координаты тоже уходят здесь
            });
        }

        private static double ClampTo(double v, double lo, double hi)
        {
            return v < lo ? lo : (v > hi ? hi : v);
        }

        /// <summary>Классифицировать отрезок и добавить, если это строго H/V-линовка достаточной длины.</summary>
        private static void AddSegment(List<PdfLine> result, double x1, double y1, double x2, double y2, double thickness)
        {
            double dx = Math.Abs(x2 - x1), dy = Math.Abs(y2 - y1);
            if (dy <= LineAxisTol && dx > dy && dx >= MinLineLen)
                result.Add(new PdfLine { Orientation = LineOrientation.Horizontal, X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Thickness = thickness });
            else if (dx <= LineAxisTol && dy > dx && dy >= MinLineLen)
                result.Add(new PdfLine { Orientation = LineOrientation.Vertical, X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Thickness = thickness });
            // иначе диагональ/кривая — не линовка
        }

        // Кап нижнего поля: поле снизу не позиционирует контент (в отличие от верхнего), а лишь
        // ограничивает заполнение листа. «Честное» огромное поле (у счёта текст кончается на
        // середине листа → 336 pt) вместе с inline-вставкой наложенных печатей/подписей
        // выталкивало документ на лишнюю страницу; кап безопасен — раньше листа разрыв не случится.
        private const double BottomMarginMaxPt = 90;

        /// <summary>
        /// Поля страницы из рамок СОДЕРЖИМОГО — слов И изображений (pt, ось Y вверх). Только по
        /// словам поля врут в обе стороны: логотип НАД первым словом «выталкивается» из верхнего
        /// поля и сдвигает весь вывод вниз на свою высоту, а штамп/подписи НИЖЕ последней
        /// строки раздували нижнее поле до полустраницы. Пустая страница — поля 0. Чистая — под тест.
        /// </summary>
        internal static void SetMargins(PdfPageText pt, List<PdfWord> words, List<OcrImage> images, double pageW, double pageH)
        {
            double minL = double.MaxValue, maxR = double.MinValue, minB = double.MaxValue, maxT = double.MinValue;
            foreach (PdfWord w in words)
            {
                if (w.Left < minL) minL = w.Left;
                if (w.Right > maxR) maxR = w.Right;
                if (w.Bottom < minB) minB = w.Bottom;
                if (w.Top > maxT) maxT = w.Top;
            }
            if (images != null)
                foreach (OcrImage img in images)
                {
                    double right = img.LeftPt + img.WidthPt, bottom = img.TopPt - img.HeightPt;
                    if (img.LeftPt < minL) minL = img.LeftPt;
                    if (right > maxR) maxR = right;
                    if (bottom < minB) minB = bottom;
                    if (img.TopPt > maxT) maxT = img.TopPt;
                }
            if (maxR < minL)
                return; // ни слов, ни картинок — поля 0
            pt.LeftMarginPt = minL;
            pt.RightMarginPt = pageW - maxR;
            pt.TopMarginPt = pageH - maxT;
            pt.BottomMarginPt = Math.Min(minB, BottomMarginMaxPt);
        }

        /// <summary>
        /// Цвет буквы PdfPig → 0xRRGGBB; null/чёрный → 0. Текст, залитый узором или градиентом
        /// (PatternColor), одного цвета не имеет, и PdfPig 0.1.9 на ToRGBValues для него бросает
        /// InvalidOperationException — раньше это роняло извлечение ВСЕГО файла. Такой текст
        /// и любой неразобранный цвет пишем чёрным: потерять оттенок лучше, чем документ.
        /// </summary>
        private static int ColorArgb(UglyToad.PdfPig.Graphics.Colors.IColor color)
        {
            if (color == null || color is UglyToad.PdfPig.Graphics.Colors.PatternColor)
                return 0;
            try
            {
                var rgb = color.ToRGBValues();
                int r = Clamp255(rgb.r), g = Clamp255(rgb.g), b = Clamp255(rgb.b);
                return (r << 16) | (g << 8) | b;
            }
            catch (OutOfMemoryException) { throw; }
            catch { return 0; }
        }

        private static int Clamp255(double v)
        {
            int n = (int)Math.Round(v * 255.0);
            return n < 0 ? 0 : (n > 255 ? 255 : n);
        }
    }
}
