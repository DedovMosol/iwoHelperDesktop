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
        /// Текст всех страниц. Битый/зашифрованный файл или запрет извлечения —
        /// <see cref="MergeException"/> с понятным сообщением.
        /// </summary>
        public static List<PdfPageText> Extract(string path, Action<int, int> progress = null)
        {
            EmbeddedAssemblies.Ensure();
            return ExtractCore(path, progress);
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
        private static List<PdfPageText> ExtractCore(string path, Action<int, int> progress)
        {
            try
            {
                using (UglyToad.PdfPig.PdfDocument doc = UglyToad.PdfPig.PdfDocument.Open(path))
                {
                    var pages = new List<PdfPageText>();
                    int pageCount = doc.NumberOfPages;
                    foreach (UglyToad.PdfPig.Content.Page page in doc.GetPages())
                    {
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
                                // «абзацы», вклинивающиеся между обычными строками.
                                if (first.TextOrientation != UglyToad.PdfPig.Content.TextOrientation.Horizontal)
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
                            double top = bb.Top;
                            // Прочерк «____» рисуется у базовой линии рамкой в пару pt: такой
                            // вырожденный бокс не пересекается по вертикали с соседями («№», «от»)
                            // и отрывался от них в отдельные строки. Даём прочерку высоту хотя бы
                            // трети кегля — как у обычного слова той же строки.
                            if (size > 0 && top - bb.Bottom < UnderscoreMinHeightFactor * size && IsUnderscoreOnly(text))
                                top = bb.Bottom + UnderscoreMinHeightFactor * size;
                            words.Add(new PdfWord
                            {
                                Text = text,
                                Left = bb.Left,
                                Right = bb.Right,
                                Bottom = bb.Bottom,
                                Top = top,
                                FontSizePt = size,
                                Bold = bold,
                                Italic = italic,
                                ColorArgb = color,
                                FontName = family
                            });
                        }
                        AssignHyperlinks(words, page);
                        // Текстовый штамп ЭП → переносим картинкой (рендер-кроп), а его слова убираем
                        // из потока (иначе печать задвоится: картинка + те же строки текстом).
                        OcrImage stampImage = ExtractTextStamp(words, page, path);
                        List<PdfLine> lines = ExtractLines(page);
                        // Слова таблиц уходят в ячейки; абзацы строятся из ОСТАВШИХСЯ (внетабличных) слов.
                        TableDetectResult det = TableDetector.Detect(lines, words, page.Width, page.Height);
                        HashSet<PdfLine> underlines = UnderlineDetector.Mark(det.RemainingWords, lines); // подчёркивания по линовке
                        // Одиночные линии, не ставшие ни таблицей, ни подчёркиванием, — прочерки
                        // реквизитов («______ №»): переносим текстом-заполнителем, иначе строка
                        // реквизитов теряет прочерки и разваливается.
                        AddRuleWords(det.RemainingWords, det.LoneLines, underlines);
                        // Сетки без линовки (чеки, формы «метка … значение») — таблицей без границ.
                        GridDetectResult grids = GridDetector.Detect(det.RemainingWords);
                        det.Tables.AddRange(grids.Tables);
                        OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(grids.RemainingWords);
                        var pt = new PdfPageText
                        {
                            PageIndex = page.Number - 1, // PdfPig нумерует страницы с 1
                            Paragraphs = layout.Paragraphs,
                            FirstLineIndentPt = layout.FirstLineIndentPt,
                            WidthPt = page.Width,
                            HeightPt = page.Height,
                            Lines = lines,
                            Tables = det.Tables
                        };
                        pt.Images = ExtractImages(page, path);
                        if (stampImage != null)
                            pt.Images.Add(stampImage); // штамп ЭП — в общий поток изображений (порядок по TopPt)
                        SetMargins(pt, words, pt.Images, page.Width, page.Height);
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

        // Прочерк из одиночной линии: пределы толщины/длины и ширина «_» в долях кегля (TNR ≈ 0.5).
        // Толщина до 6 pt: прочерки реквизитов встречаются и «жирным» штрихом (5 pt), а заливки
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
        /// слова-прочерки «____»: линии-заполнители реквизитов («______ № ______», «На № … от …»)
        /// иначе теряются вовсе, и строка реквизитов разваливается. Число «_» подбирается по
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

        /// <summary>Слово целиком из «_» (прочерк-заполнитель реквизитов)?</summary>
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
        private static List<OcrImage> ExtractImages(UglyToad.PdfPig.Content.Page page, string pdfPath)
        {
            var result = new List<OcrImage>();
            IEnumerable<UglyToad.PdfPig.Content.IPdfImage> images;
            try { images = page.GetImages(); }
            catch { return result; }

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
                        if (png == null || png.Length == 0 || IsSolidColor(png))
                        {
                            if (!rasterTried)
                            {
                                pageRaster = PageRasterizer.RenderPage(pdfPath, page.Number);
                                rasterTried = true;
                            }
                            byte[] cropped = PageRasterizer.CropRegion(pageRaster, page.Width, page.Height,
                                b.Left, b.Top, b.Width, b.Height);
                            if (cropped != null && !IsSolidColor(cropped))
                                png = cropped;
                            else if (masked && (png == null || png.Length == 0))
                                png = EncodePng(img); // нет Ghostscript — маскированную берём как есть (может быть чёрный фон), лучше чем терять
                            if (png == null || png.Length == 0 || IsSolidColor(png))
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
        /// Распознать текстовый штамп ЭП и вернуть его как изображение (рендер-кроп области
        /// Ghostscript-ом), убрав слова штампа из <paramref name="words"/>. Если штамп не найден
        /// или картинку получить не удалось (нет Ghostscript, пустая область) — возвращает null и
        /// НИЧЕГО не убирает: текст остаётся текстом (прежнее поведение, без регрессии). Вызывать
        /// из ядра (после Ensure) — тело ссылается на типы PdfPig.
        /// </summary>
        private static OcrImage ExtractTextStamp(List<PdfWord> words, UglyToad.PdfPig.Content.Page page, string pdfPath)
        {
            StampRegion stamp = StampDetector.Detect(words, page.Width, page.Height);
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
                byte[] png = PageRasterizer.CropRegion(raster, page.Width, page.Height, left, top, width, height);
                if (png == null || png.Length == 0)
                    return null; // рендер/кроп не удался — не переносим, текст остаётся текстом
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

        /// <summary>
        /// Одноцветный ли растр (все опрошенные пиксели одного цвета). Признак сбоя декодера
        /// PDF-картинки (напр. штрих-код выходит сплошным чёрным прямоугольником) — такое
        /// изображение бессмысленно, его лучше не переносить. Сбой чтения — не считаем одноцветным.
        /// </summary>
        private static bool IsSolidColor(byte[] png)
        {
            try
            {
                using (var ms = new MemoryStream(png))
                using (var bmp = new System.Drawing.Bitmap(ms))
                {
                    int w = bmp.Width, h = bmp.Height;
                    if (w < 2 || h < 2)
                        return true; // 1-пиксельный растр смысла не несёт — считаем вырожденным
                    int stepX = Math.Max(1, w / 16), stepY = Math.Max(1, h / 16);
                    int first = bmp.GetPixel(0, 0).ToArgb();
                    for (int y = 0; y < h; y += stepY)
                        for (int x = 0; x < w; x += stepX)
                            if (bmp.GetPixel(x, y).ToArgb() != first)
                                return false; // нашёлся другой цвет — не одноцветный
                    return true;
                }
            }
            catch { return false; } // не декодировали — не считаем одноцветным (пусть проходит)
        }

        // Классификация отрезка как линовки: отклонение от оси (pt) и минимальная длина.
        private const double LineAxisTol = 1.0; // |dy| для горизонтали / |dx| для вертикали
        private const double MinLineLen = 2.0;  // короче — графический шум, не линовка
        private const double ThinFillMax = 2.5; // залитая полоска тоньше этого — это линия, не фон

        /// <summary>
        /// Линовка страницы (границы таблиц, подчёркивания) из векторной графики. Берём только
        /// строго горизонтальные/вертикальные штрихи и рёбра прямоугольников; диагонали, кривые
        /// и крупные заливки (фон ячеек) пропускаем. Сбой одной фигуры не срывает остальные и
        /// текст. Вызывать из ядра (после Ensure) — тело ссылается на типы PdfPig.
        /// </summary>
        private static List<PdfLine> ExtractLines(UglyToad.PdfPig.Content.Page page)
        {
            var result = new List<PdfLine>();
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
                            AddThinFilledRect(result, sub);
                    }
                }
                catch { } // одна битая фигура не ломает остальные
            }
            return result;
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
        /// словам поля врут в обе стороны: герб НАД первым словом «выталкивается» из верхнего
        /// поля и сдвигает весь вывод вниз на свою высоту, а печать/подписи НИЖЕ последней
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
