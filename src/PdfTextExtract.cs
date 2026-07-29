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
        /// Текст всех страниц. rotations — пользовательские повороты по индексу исходной
        /// страницы (0/90/180/270 по часовой, null — без поворотов): страница выправляется
        /// ДО анализа макета, боковой текст становится горизонтальным. Битый/зашифрованный
        /// файл или запрет извлечения — <see cref="MergeException"/> с понятным сообщением.
        /// </summary>
        public static List<PdfPageText> Extract(string path, Action<int, int> progress = null, IList<int> rotations = null,
            Func<bool> cancelled = null, PageLayoutMode mode = PageLayoutMode.Document)
        {
            EmbeddedAssemblies.Ensure();
            return ExtractCore(path, progress, rotations, cancelled, mode);
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
                using (UglyToad.PdfPig.PdfDocument doc = UglyToad.PdfPig.PdfDocument.Open(path))
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
            Func<bool> cancelled, PageLayoutMode mode)
        {
            try
            {
                using (UglyToad.PdfPig.PdfDocument doc = UglyToad.PdfPig.PdfDocument.Open(path))
                {
                    var pages = new List<PdfPageText>();
                    int pageCount = doc.NumberOfPages;
                    foreach (UglyToad.PdfPig.Content.Page page in doc.GetPages())
                    {
                        Cancellation.ThrowIf(cancelled); // отмена между страницами извлечения
                        int rotation = PageRotation.At(rotations, page.Number - 1);
                        // Картинки страницы перечисляются ОДИН раз (декода здесь нет): рамки
                        // нужны фильтрам невидимого текста и наложенных картинок, а сами объекты — переносу
                        // изображений ниже. Рамки картинок и тёмных векторных заливок — подложки,
                        // на которых белый/невидимый текст ЛЕГИТИМЕН (светлые шапки, OCR-слой
                        // поверх скана) и скрывать его нельзя.
                        List<UglyToad.PdfPig.Content.IPdfImage> pageImages = ListPageImages(page);
                        var imageRects = new List<RectPt>(pageImages.Count);
                        foreach (UglyToad.PdfPig.Content.IPdfImage img in pageImages)
                        {
                            UglyToad.PdfPig.Core.PdfRectangle b = img.Bounds;
                            if (b.Width > 0 && b.Height > 0)
                                imageRects.Add(RectOf(b));
                        }
                        List<RectPt> darkFills;
                        List<PdfLine> lines = ExtractGraphics(page, out darkFills);
                        // Какая ориентация текста станет горизонтальной после поворота страницы
                        // пользователем: только её и пускаем в поток (при 0° — прежний фильтр).
                        UglyToad.PdfPig.Content.TextOrientation keepOrientation = KeepOrientationFor(rotation);
                        // Разметка структуры страницы, если автор её оставил: по ней видно, какие
                        // строки автор считал ОДНИМ абзацем (см. StructureBlocks). Пусто — работаем
                        // как раньше, по зазорам.
                        Dictionary<UglyToad.PdfPig.Content.Letter, int> blocks = StructureBlocks.Map(page);
                        var words = new List<PdfWord>();
                        foreach (UglyToad.PdfPig.Content.Word w in page.GetWords())
                        {
                            UglyToad.PdfPig.Core.PdfRectangle bb = w.BoundingBox;
                            double size = 0;
                            bool bold = false, italic = false;
                            int color = 0;
                            string family = null;
                            string text = w.Text;
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
                                    && !CoveredByAnyRect(RectOf(bb), imageRects)
                                    && !CoveredByAnyRect(RectOf(bb), darkFills))
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
                            // Символы приватной зоны юникода (заглушки полей Word, маркеры Symbol)
                            // не имеют смысла вне своего шрифта — в выводе это «□». Пропускаем.
                            if (IsPrivateUseOnly(text))
                                continue;
                            // Слово, накрытое НИЗКОЙ наложенной картинкой (впечатанные дата и номер
                            // поверх пустой формы «____№____»), текстом не выводим: его вид уже
                            // несёт картинка, а текст рядом с ней давал бы кашу из двух слоёв.
                            if (CoveredByLowImage(RectOf(bb), imageRects))
                                continue;
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
                        List<OcrImage> images = ExtractImages(pageImages, page, path);
                        double pageW = page.Width, pageH = page.Height;
                        PageRotation.RotatePage(words, lines, images, rotation, ref pageW, ref pageH);
                        ApplyUnderscoreHeights(words);
                        // Текстовый штамп → переносим картинкой (рендер-кроп), а его слова убираем
                        // из потока (иначе штамп задвоится: картинка + те же строки текстом).
                        // Печать, нарисованную ТЕКСТОМ, потоковый вывод переносит картинкой: в
                        // строку её иначе не поставить. Слайду это не нужно и вредно — текст
                        // печати должен остаться текстом, а её кольцо и линии придут подложкой.
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
                            Tables = det.Tables
                        };
                        pt.Images = images;
                        if (stampImage != null)
                            pt.Images.Add(stampImage); // штамп — в общий поток изображений (порядок по TopPt)
                        SetMargins(pt, words, pt.Images, pageW, pageH);
                        pages.Add(pt);
                        if (progress != null)
                            progress(pages.Count, pageCount);
                    }
                    return pages;
                }
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                throw new MergeException(string.Format(Loc.T("err.ocr.extractFailed"), Path.GetFileName(path), ex.Message));
            }
        }

        private const double UnderscoreMinHeightFactor = 0.35; // виртуальная высота прочерка в долях кегля

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
            foreach (PdfWord w in words)
            {
                double mx = (w.Left + w.Right) / 2, my = (w.Top + w.Bottom) / 2;
                foreach (UglyToad.PdfPig.Content.Hyperlink h in links)
                {
                    if (string.IsNullOrEmpty(h.Uri))
                        continue;
                    UglyToad.PdfPig.Core.PdfRectangle b = h.Bounds;
                    if (mx >= b.Left && mx <= b.Right && my >= b.Bottom && my <= b.Top)
                    {
                        w.Uri = h.Uri;
                        break;
                    }
                }
            }
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
            System.Drawing.Bitmap pageRaster = null; // ленивый рендер страницы — только если понадобится фолбэк
            bool rasterTried = false;
            try
            {
                foreach (UglyToad.PdfPig.Content.IPdfImage img in images)
                {
                    try
                    {
                        UglyToad.PdfPig.Core.PdfRectangle b = img.Bounds;
                        if (b.Width <= 0 || b.Height <= 0)
                            continue;

                        // Картинка с мягкой маской (SMask/Mask) — прозрачность: PdfPig её игнорирует
                        // и заливает прозрачные области ЧЁРНЫМ (в Word получается чёрный фон у логотипа/
                        // печати). Такую переносим рендер-кропом страницы Ghostscript-ом — он композитит
                        // маску на белое, как она выглядит на странице. Прямой декод для неё пропускаем.
                        bool masked = HasSoftMask(img);
                        byte[] png = masked ? null : EncodePng(img);
                        // Декодер не справился (null) или выдал одноцветный мусор (напр. штрих-код
                        // чёрным ящиком) — фолбэк: рендерим страницу Ghostscript-ом и вырезаем картинку
                        // по её рамке. Так переносится любая картинка, как она выглядит.
                        if (png == null || png.Length == 0 || RasterUtil.IsSolidColor(png))
                        {
                            if (!rasterTried)
                            {
                                pageRaster = PageRasterizer.RenderPage(pdfPath, page.Number);
                                rasterTried = true;
                            }
                            byte[] cropped = PageRasterizer.CropRegion(pageRaster, page.Width, page.Height,
                                b.Left, b.Top, b.Width, b.Height);
                            if (cropped != null && !RasterUtil.IsSolidColor(cropped))
                                png = cropped;
                            else if (masked && (png == null || png.Length == 0))
                                png = EncodePng(img); // нет Ghostscript — маскированную берём как есть (может быть чёрный фон), лучше чем терять
                            if (png == null || png.Length == 0 || RasterUtil.IsSolidColor(png))
                                continue; // и фолбэк не помог (нет GS / область пустая) — пропускаем
                        }

                        result.Add(new OcrImage
                        {
                            Png = png,
                            LeftPt = b.Left,
                            TopPt = b.Top,
                            WidthPt = b.Width,
                            HeightPt = b.Height
                        });
                    }
                    catch { } // одна битая картинка не ломает остальные
                }
            }
            finally
            {
                if (pageRaster != null)
                    pageRaster.Dispose();
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
            System.Drawing.Bitmap raster = null;
            try
            {
                raster = PageRasterizer.RenderPage(pdfPath, page.Number);
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
            catch { return null; } // сбой рендера/кропа — штамп остаётся текстом
            finally { if (raster != null) raster.Dispose(); }
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

        private static RectPt RectOf(UglyToad.PdfPig.Core.PdfRectangle b)
        {
            return new RectPt { Left = b.Left, Bottom = b.Bottom, Right = b.Right, Top = b.Top };
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

        /// <summary>Цвет буквы PdfPig → 0xRRGGBB; null/чёрный → 0.</summary>
        private static int ColorArgb(UglyToad.PdfPig.Graphics.Colors.IColor color)
        {
            if (color == null)
                return 0;
            var rgb = color.ToRGBValues();
            int r = Clamp255(rgb.r), g = Clamp255(rgb.g), b = Clamp255(rgb.b);
            return (r << 16) | (g << 8) | b;
        }

        private static int Clamp255(double v)
        {
            int n = (int)Math.Round(v * 255.0);
            return n < 0 ? 0 : (n > 255 ? 255 : n);
        }
    }
}
