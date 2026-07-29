using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Запись разобранных страниц PDF в .pptx. Формат собирается своим кодом (ZIP + XML), а не
    /// через PowerPoint: инструментом пользуются как раз тогда, когда PowerPoint под рукой нет,
    /// и требовать его установки ради складывания текста в слайды несоразмерно. Побочная выгода —
    /// проверяемость: результат разбирается тестом без единого COM-вызова.
    ///
    /// Слайд собирается из ФИГУР с абсолютными координатами, и это делает задачу проще потоковой
    /// вёрстки Word: интервалы, красную строку и выключку воспроизводить нечем и незачем — у
    /// каждой надписи есть своя рамка из PDF. Взамен появляется своя забота: рамка обязана быть
    /// чуть шире измеренной, иначе последнее слово перепрыгнет на вторую строку (см.
    /// <see cref="WidthSlackPt"/>).
    /// </summary>
    internal static class PptxWriter
    {
        /// <summary>Имя приложения в свойствах документа.</summary>
        private const string ApplicationName = "iwo Helper Desktop";

        /// <summary>
        /// Первый свободный номер фигуры на слайде: единицу занимает узел-группа дерева фигур.
        /// </summary>
        private const int FirstShapeId = 2;

        /// <summary>
        /// Запас ширины надписи сверх измеренной рамки. Ширина текста в PowerPoint считается по
        /// его собственным метрикам и на доли процента расходится с тем, что намерил PDF; без
        /// запаса последнее слово абзаца регулярно уезжает на новую строку, и абзац «распухает»
        /// вниз. Доля от ширины плюс минимум в пунктах — короткой строке процента мало.
        /// </summary>
        private const double WidthSlackFraction = 0.04;
        private const double WidthSlackMinPt = 2;

        /// <summary>Высота вырожденной рамки: строка кегля N занимает примерно столько.</summary>
        private const double LineHeightFactor = 1.25;

        /// <summary>
        /// Насколько верх БУКВ ниже верха рамки надписи, в долях кегля. PowerPoint отводит
        /// строке место по метрикам шрифта и ставит буквы не вплотную к верхнему краю, а ниже
        /// на подъёмную часть строки. Рамка же берётся у нас по ЧЕРНИЛАМ исходника, поэтому
        /// без поправки текст на слайде садится ниже, чем стоял, — и подчёркивания с линовкой,
        /// пришедшие подложкой, начинают пересекать буквы.
        ///
        /// Значение измерено: одна строка в известном месте прогонялась через настоящий
        /// PowerPoint и сравнивалась с оригиналом по первому пикселю чернил — Arial 24 дал
        /// 0,260 кегля, Arial 12 — 0,300, Times 24 — 0,330, Times 10 — 0,384, Calibri 18 —
        /// 0,307. Точное значение зависит от шрифта (высота прописных к подъёму разная), и
        /// одной константой его не выразить; взята середина, остаточная ошибка — меньше
        /// пункта против прежних четырёх-восьми.
        /// </summary>
        private const double AscentGapFactor = 0.29;

        /// <summary>
        /// Во сколько кеглей укладывается ОДНА строка. Всё, что выше, — многострочный абзац.
        /// Порог с запасом над <see cref="LineHeightFactor"/>: рамка строки бывает выше кегля
        /// из-за выносных элементов, а две строки дают уже вдвое больше.
        /// </summary>
        private const double SingleLineFactor = 1.6;

        /// <summary>
        /// Пишет .pptx из разобранных страниц. progress — «сделано/всего» по слайдам, может быть
        /// null. cancelled — кооперативная отмена между слайдами (файла на диске к этому моменту
        /// ещё нет). onCommitting — перед записью файла: дальше отмена уже ничего не изменит.
        /// whenUtc попадает в свойства документа.
        /// </summary>
        public static void Write(IList<PdfPageText> pages, string path, DateTime whenUtc,
            Action<int, int> progress = null, Func<bool> cancelled = null, Action onCommitting = null,
            IList<Background> backgrounds = null)
        {
            if (pages == null)
                throw new ArgumentNullException("pages");
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("нет пути результата", "path");

            OoxmlPackage package = BuildPackage(pages, whenUtc, progress, cancelled, backgrounds);
            if (onCommitting != null)
                onCommitting();
            package.Save(path);
        }

        /// <summary>
        /// Собрать пакет целиком в памяти: постоянный каркас + по слайду на страницу.
        /// Отдельно от записи — так пакет проверяется тестом, не создавая файла.
        /// </summary>
        internal static OoxmlPackage BuildPackage(IList<PdfPageText> pages, DateTime whenUtc,
            Action<int, int> progress = null, Func<bool> cancelled = null, IList<Background> backgrounds = null)
        {
            double slideWidthPt, slideHeightPt;
            PptxGeometry.SlideSizePt(pages, out slideWidthPt, out slideHeightPt);

            var package = new OoxmlPackage();
            int count = pages.Count;

            // Каркас, одинаковый для любой колоды.
            package.AddXml("_rels/.rels", PptxParts.RootRels());
            package.AddXml(PptxParts.CorePropsPart, PptxParts.CoreProps(ApplicationName, whenUtc), PptxParts.CtCoreProps);
            package.AddXml(PptxParts.AppPropsPart, PptxParts.AppProps(count, ApplicationName), PptxParts.CtAppProps);
            package.AddXml(PptxParts.PresentationPart,
                PptxParts.Presentation(count, PptxGeometry.Emu(slideWidthPt), PptxGeometry.Emu(slideHeightPt)),
                PptxParts.CtPresentation);
            package.AddXml("ppt/_rels/presentation.xml.rels", PptxParts.PresentationRels(count));
            package.AddXml(PptxParts.SlideMasterPart, PptxParts.SlideMaster(), PptxParts.CtSlideMaster);
            package.AddXml("ppt/slideMasters/_rels/slideMaster1.xml.rels", PptxParts.SlideMasterRels());
            package.AddXml(PptxParts.SlideLayoutPart, PptxParts.SlideLayout(), PptxParts.CtSlideLayout);
            package.AddXml("ppt/slideLayouts/_rels/slideLayout1.xml.rels", PptxParts.SlideLayoutRels());
            package.AddXml(PptxParts.ThemePart, PptxParts.Theme(), PptxParts.CtTheme);
            package.AddXml(PptxParts.TableStylesPart, PptxParts.TableStyles(), PptxParts.CtTableStyles);

            using (var media = new MediaLibrary(package))
                for (int i = 0; i < count; i++)
                {
                    Cancellation.ThrowIf(cancelled); // отмена между слайдами: файла ещё нет
                    Background background = backgrounds != null && i < backgrounds.Count ? backgrounds[i] : null;
                    var slide = new SlideBuilder(pages[i], slideWidthPt, slideHeightPt, media, background);
                    package.AddXml(PptxParts.SlidePart(i), slide.Build(), PptxParts.CtSlide);
                    package.AddXml("ppt/slides/_rels/slide" + PptxParts.Num(i + 1) + ".xml.rels",
                        PptxParts.SlideRels(slide.Rels));
                    if (progress != null)
                        progress(i + 1, count);
                }
            return package;
        }

        /// <summary>
        /// Хранилище картинок пакета: одинаковые байты кладутся ОДИН раз. Логотип, повторённый
        /// на сотне слайдов, — это одна часть и сотня ссылок на неё, иначе файл распухает
        /// кратно числу слайдов.
        /// </summary>
        private sealed class MediaLibrary : IDisposable
        {
            private readonly OoxmlPackage _package;
            private readonly Dictionary<string, string> _byHash = new Dictionary<string, string>(StringComparer.Ordinal);
            // Один считатель отпечатков на весь пакет: картинок бывают сотни, и заводить
            // криптопровайдер на каждую — лишняя работа и лишние дескрипторы.
            private readonly System.Security.Cryptography.SHA1 _sha = System.Security.Cryptography.SHA1.Create();
            private int _next = 1;

            public MediaLibrary(OoxmlPackage package)
            {
                _package = package;
            }

            /// <summary>
            /// Имя части для этих байтов (добавляя её при первой встрече). null — пустая
            /// картинка. jpeg говорит, чем закодированы байты: расширение и тип части обязаны
            /// соответствовать содержимому, иначе читатель откажется открывать файл.
            /// </summary>
            public string Add(byte[] data, bool jpeg = false)
            {
                if (data == null || data.Length == 0)
                    return null;
                // Отпечаток здесь не про безопасность, а про равенство байтов: нужен признак
                // «эту картинку уже клали». Длина в ключе — дешёвая страховка от совпадения.
                string hash = data.Length.ToString(CultureInfo.InvariantCulture) + ":"
                    + Convert.ToBase64String(_sha.ComputeHash(data));
                string part;
                if (_byHash.TryGetValue(hash, out part))
                    return part;
                part = "ppt/media/image" + _next.ToString(CultureInfo.InvariantCulture) + (jpeg ? ".jpeg" : ".png");
                _next++;
                _package.AddBinary(part, data, jpeg ? PptxParts.CtJpeg : PptxParts.CtPng);
                _byHash[hash] = part;
                return part;
            }

            public void Dispose()
            {
                _sha.Dispose();
            }
        }

        /// <summary>
        /// Сборка ОДНОГО слайда. Разметка фигур и список связей слайда строятся вместе: ссылка
        /// r:id в разметке и запись о ней в связях обязаны сходиться, а сойтись они могут только
        /// если их выдаёт одно и то же место.
        /// </summary>
        private sealed class SlideBuilder
        {
            private readonly PdfPageText _page;
            private readonly SlideFit _fit;
            private readonly MediaLibrary _media;
            private readonly Background _background;
            private bool _backgroundPlaced; // подложка действительно попала на слайд
            private readonly List<string> _rels = new List<string>();
            private readonly Dictionary<string, string> _relIdByTarget = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private readonly StringBuilder _shapes = new StringBuilder(4096);
            private int _shapeId = FirstShapeId;
            private int _relId = 2; // rId1 занят макетом слайда

            public SlideBuilder(PdfPageText page, double slideWidthPt, double slideHeightPt, MediaLibrary media,
                Background background)
            {
                _page = page ?? new PdfPageText();
                _media = media;
                _background = background;
                _fit = PptxGeometry.Fit(_page.WidthPt, _page.HeightPt, slideWidthPt, slideHeightPt);
            }

            /// <summary>Связи этого слайда (кроме макета, который добавляет <see cref="PptxParts.SlideRels"/>).</summary>
            public List<string> Rels { get { return _rels; } }

            public string Build()
            {
                string background = AppendBackground(); // подложка — ПОД содержимым, поэтому первой
                foreach (PageBlocks.Block block in Blocks())
                {
                    if (block.Paragraph != null)
                        AppendParagraph(block.Paragraph);
                    else if (block.Image != null)
                    {
                        // Подложка — это вся страница БЕЗ ТЕКСТА, то есть картинки в ней уже
                        // нарисованы. Класть их поверх второй раз значит удвоить вес файла и
                        // обмануть человека: удалив картинку на слайде, он увидит её же в
                        // подложке. Без подложки (нет Ghostscript, пустая страница, исчерпан
                        // бюджет) картинки, наоборот, единственный источник изображения.
                        if (!_backgroundPlaced)
                            AppendImage(block.Image);
                    }
                    else if (block.Table != null)
                        AppendTable(block.Table);
                }

                var sb = new StringBuilder(_shapes.Length + 1024);
                sb.Append(PptxParts.XmlHeader);
                sb.Append("<p:sld").Append(PptxParts.NsPml).Append('>');
                sb.Append("<p:cSld>");
                sb.Append(background); // по схеме фон идёт СТРОГО перед деревом фигур
                sb.Append("<p:spTree>");
                sb.Append("<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>");
                sb.Append("<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>"
                    + "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>");
                sb.Append(_shapes);
                sb.Append("</p:spTree></p:cSld>");
                sb.Append("<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>");
                sb.Append("</p:sld>");
                return sb.ToString();
            }

            /// <summary>
            /// Блоки страницы в порядке чтения. Полосы «бок о бок» раскрываются обратно в
            /// одиночные блоки: они придуманы для ПОТОКОВОЙ вёрстки (там колонки приходится
            /// изображать таблицей), а фигуре с абсолютными координатами соседи не мешают.
            /// Порядок при этом сохраняется — он и задаёт обход слайда с клавиатуры.
            /// </summary>
            private List<PageBlocks.Block> Blocks()
            {
                var result = new List<PageBlocks.Block>();
                foreach (PageBlocks.PageItem item in PageBlocks.OrderedItems(_page))
                {
                    if (!item.IsBand)
                    {
                        result.Add(item.Single);
                        continue;
                    }
                    foreach (List<PageBlocks.Block> column in item.Columns)
                        result.AddRange(column);
                }
                return result;
            }

            // ---------- подложка ----------

            /// <summary>
            /// Подложка страницы. Когда размер страницы совпал с размером слайда (обычный
            /// случай), она становится ФОНОМ слайда: фон нельзя случайно выделить и сдвинуть,
            /// правя текст. Когда страницу пришлось вписывать, фон не годится — он растянулся
            /// бы на весь слайд и исказил пропорции; тогда подложка кладётся картинкой в свою
            /// рамку и запирается от перемещения, изменения размера и выделения.
            /// Возвращает разметку фона (может быть пустой) и, если надо, добавляет фигуру.
            /// </summary>
            private string AppendBackground()
            {
                if (_background == null || _background.Data == null || _background.Data.Length == 0)
                    return string.Empty;
                string part = _media.Add(_background.Data, _background.IsJpeg);
                if (part == null)
                    return string.Empty;
                string relId = ImageRel(part);
                _backgroundPlaced = true;

                if (Math.Abs(_fit.Scale - 1) < 1e-9)
                    return "<p:bg><p:bgPr><a:blipFill><a:blip r:embed=\"" + relId + "\"/>"
                        + "<a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>";

                var sb = new StringBuilder(512);
                sb.Append("<p:pic><p:nvPicPr><p:cNvPr id=\"").Append(PptxParts.Num(_shapeId))
                  .Append("\" name=\"Background ").Append(PptxParts.Num(_shapeId)).Append("\"/>");
                sb.Append("<p:cNvPicPr><a:picLocks noChangeAspect=\"1\" noMove=\"1\" noResize=\"1\" noSelect=\"1\"/>"
                    + "</p:cNvPicPr><p:nvPr/></p:nvPicPr>");
                sb.Append("<p:blipFill><a:blip r:embed=\"").Append(relId)
                  .Append("\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>");
                sb.Append("<p:spPr>");
                AppendFrame(sb, 0, _page.HeightPt, _page.WidthPt, _page.HeightPt);
                sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>");
                _shapes.Append(sb);
                _shapeId++;
                return string.Empty;
            }

            // ---------- надпись ----------

            private void AppendParagraph(OcrParagraph paragraph)
            {
                string body = Runs(paragraph);
                if (body.Length == 0)
                    return; // пустой абзац фигуры не заслуживает: пустая рамка мешает править слайд

                double leftPt = paragraph.LeftPt;
                double widthPt = paragraph.RightPt - paragraph.LeftPt;
                double heightPt = paragraph.TopPt - paragraph.BottomPt;
                double fontPt = FirstFontSizePt(paragraph);
                // Абзац, занимавший в источнике ОДНУ строку, переноситься не должен: он уже
                // помещался, а метрики PowerPoint чуть шире измеренных, и любой перенос здесь —
                // это разрыв там, где в оригинале разрыва не было (подпись «апрель» уезжает
                // второй строкой). Многострочному абзацу перенос, наоборот, нужен — его ширина
                // и есть та, по которой он вёрстан.
                bool singleLine = IsSingleLine(heightPt, fontPt);
                widthPt = SlackWidthPt(leftPt, widthPt, _page.WidthPt);
                if (heightPt <= 0)
                    heightPt = fontPt * LineHeightFactor; // вырожденная рамка — строка своего кегля
                // Рамку поднимаем на подъёмную часть строки и на неё же увеличиваем — иначе
                // текст сядет ниже, чем стоял в источнике (см. AscentGapFactor).
                double ascentGap = AscentGapPt(fontPt);
                double boxTopPt = paragraph.TopPt + ascentGap;
                heightPt += ascentGap;

                var sb = new StringBuilder(body.Length + 512);
                sb.Append("<p:sp><p:nvSpPr><p:cNvPr id=\"").Append(PptxParts.Num(_shapeId))
                  .Append("\" name=\"TextBox ").Append(PptxParts.Num(_shapeId)).Append("\"/>");
                sb.Append("<p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>");
                sb.Append("<p:spPr>");
                AppendFrame(sb, leftPt, boxTopPt, widthPt, heightPt);
                sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>");
                // Нулевые поля надписи обязательны: по умолчанию PowerPoint отступает 0,1" слева
                // и 0,05" сверху, и весь текст уезжает относительно измеренной рамки. Автоподбор
                // выключен — он менял бы кегль, а кегль у нас измеренный.
                sb.Append("<p:txBody><a:bodyPr wrap=\"").Append(singleLine ? "none" : "square")
                    .Append("\" lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\""
                    + " anchor=\"t\"><a:noAutofit/></a:bodyPr><a:lstStyle/>");
                // У однострочной надписи рамка обнимает сам текст, поэтому её левый край и есть
                // точное место строки — выравнивание здесь только сдвинуло бы её на запас ширины.
                OcrAlignment align = singleLine ? OcrAlignment.Left : paragraph.Alignment;
                sb.Append("<a:p>").Append(ParagraphProps(paragraph, align)).Append(body).Append("</a:p>");
                sb.Append("</p:txBody></p:sp>");
                _shapes.Append(sb);
                _shapeId++;
            }

            /// <summary>
            /// Свойства абзаца: выравнивание и отказ от маркера. Маркер списка НЕ снимается с
            /// текста и не заменяется родным маркером PowerPoint намеренно: каждый абзац лежит
            /// в СВОЕЙ надписи, а родная нумерация считается внутри одной надписи — все пункты
            /// стали бы «1.». Литеральный маркер сохраняет и вид, и номер.
            /// </summary>
            /// <summary>
            /// Свойства абзаца. Межстрочный интервал не задаём вовсе: строка источника — это
            /// отдельная надпись, и внутри неё второй строки не бывает (см. PageLayoutMode).
            /// </summary>
            private static string ParagraphProps(OcrParagraph paragraph, OcrAlignment alignment)
            {
                string align;
                switch (alignment)
                {
                    case OcrAlignment.Center: align = "ctr"; break;
                    case OcrAlignment.Justify: align = "just"; break;
                    default: align = "l"; break;
                }
                string bullet = paragraph.ListKind == ListKind.None ? string.Empty : "<a:buNone/>";
                return "<a:pPr algn=\"" + align + "\">" + bullet + "</a:pPr>";
            }

            private string Runs(OcrParagraph paragraph)
            {
                if (paragraph.Runs == null || paragraph.Runs.Count == 0)
                    return string.Empty;
                var sb = new StringBuilder(256);
                bool anyVisible = false;
                foreach (OcrRun run in paragraph.Runs)
                {
                    if (IsFillerRule(run.Text))
                        continue;
                    string text = XmlText.Sanitize(run.Text);
                    if (text.Length == 0)
                        continue;
                    if (text.Trim().Length > 0)
                        anyVisible = true;
                    sb.Append("<a:r>").Append(RunProps(run, text)).Append("<a:t>")
                      .Append(XmlText.Escape(text)).Append("</a:t></a:r>");
                }
                return anyVisible ? sb.ToString() : string.Empty; // только пробелы — это не текст
            }

            /// <summary>
            /// Свойства рана. Порядок дочерних элементов задан схемой: заливка, затем шрифты
            /// (latin, ea, cs), затем ссылка. Перестановка ломает файл так же надёжно, как
            /// отсутствующий элемент.
            /// </summary>
            private string RunProps(OcrRun run, string text)
            {
                var sb = new StringBuilder(192);
                sb.Append("<a:rPr lang=\"").Append(FontResolver.HasCyrillic(text) ? "ru-RU" : "en-US").Append('"');
                sb.Append(" sz=\"").Append(PptxParts.Num(_fit.FontHundredths(run.FontSizePt))).Append('"');
                if (run.Bold)
                    sb.Append(" b=\"1\"");
                if (run.Italic)
                    sb.Append(" i=\"1\"");
                if (run.Underline)
                    sb.Append(" u=\"sng\"");
                // Надстрочный и подстрочный задаются смещением базовой линии в тысячных долях
                // процента — теми же числами, что показывает сам PowerPoint (30% и −25%).
                if (run.Super)
                    sb.Append(" baseline=\"30000\"");
                else if (run.Sub)
                    sb.Append(" baseline=\"-25000\"");
                sb.Append('>');

                sb.Append("<a:solidFill><a:srgbClr val=\"").Append(Rgb(run.ColorArgb)).Append("\"/></a:solidFill>");
                string font = FontResolver.Resolve(run.FontName, text);
                string escaped = XmlText.Encode(font);
                sb.Append("<a:latin typeface=\"").Append(escaped).Append("\"/>");
                sb.Append("<a:ea typeface=\"").Append(escaped).Append("\"/>");
                sb.Append("<a:cs typeface=\"").Append(escaped).Append("\"/>");
                string relId = HyperlinkRel(run.Uri);
                if (relId != null)
                    sb.Append("<a:hlinkClick r:id=\"").Append(relId).Append("\"/>");
                sb.Append("</a:rPr>");
                return sb.ToString();
            }

            /// <summary>Кегль первой значимой строки абзаца — им меряется высота вырожденной рамки.</summary>
            private static double FirstFontSizePt(OcrParagraph paragraph)
            {
                if (paragraph.Runs != null)
                    foreach (OcrRun run in paragraph.Runs)
                        if (run.FontSizePt > 0)
                            return FontResolver.ClampSizePt(run.FontSizePt);
                return FontResolver.DefaultFontSize;
            }

            // ---------- таблица ----------

            /// <summary>
            /// Таблица — не набор надписей, а настоящая таблица слайда: её можно править,
            /// добавлять строки и сортировать. Ячейки, накрытые объединением, обязаны
            /// ПРИСУТСТВОВАТЬ в разметке с пометкой о накрытии: в каждой строке ровно столько
            /// ячеек, сколько колонок в сетке, иначе PowerPoint объявляет файл повреждённым.
            /// </summary>
            private void AppendTable(OcrTable table)
            {
                int cols = table.ColumnCount;
                int rows = table.Rows == null ? 0 : table.Rows.Count;
                if (cols <= 0 || rows <= 0)
                    return;

                CellMerge[,] merge = MergeGrid(table);

                double widthPt = table.RightPt - table.LeftPt;
                double heightPt = table.TopPt - table.BottomPt;
                double rowHeightPt = rows > 0 && heightPt > 0 ? heightPt / rows : FontResolver.DefaultFontSize;

                var sb = new StringBuilder(2048);
                sb.Append("<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"").Append(PptxParts.Num(_shapeId))
                  .Append("\" name=\"Table ").Append(PptxParts.Num(_shapeId)).Append("\"/>");
                sb.Append("<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"1\"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>");
                sb.Append("<p:xfrm><a:off x=\"").Append(PptxParts.Num(_fit.X(table.LeftPt)))
                  .Append("\" y=\"").Append(PptxParts.Num(_fit.Y(table.TopPt))).Append("\"/>");
                long cx = _fit.Size(widthPt), cy = _fit.Size(heightPt);
                sb.Append("<a:ext cx=\"").Append(PptxParts.Num(cx > 0 ? cx : 1))
                  .Append("\" cy=\"").Append(PptxParts.Num(cy > 0 ? cy : 1)).Append("\"/></p:xfrm>");
                sb.Append("<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">");
                // Полосы и выделенная первая строка отключены: раскраска «через строку» —
                // решение оформителя, а мы переносим то, что есть в источнике.
                sb.Append("<a:tbl><a:tblPr firstRow=\"0\" bandRow=\"0\"/><a:tblGrid>");
                for (int c = 0; c < cols; c++)
                    sb.Append("<a:gridCol w=\"").Append(PptxParts.Num(ColumnWidthEmu(table, c, widthPt, cols))).Append("\"/>");
                sb.Append("</a:tblGrid>");

                for (int r = 0; r < rows; r++)
                {
                    sb.Append("<a:tr h=\"").Append(PptxParts.Num(_fit.Size(rowHeightPt))).Append("\">");
                    OcrTableRow row = table.Rows[r];
                    for (int c = 0; c < cols; c++)
                    {
                        // Строки и ячейки проверяем на null так же, как MergeGrid: таблица
                        // приходит от детекторов, и её полнота — их инвариант, а не наш.
                        OcrTableCell cell = row != null && row.Cells != null && c < row.Cells.Count
                            ? row.Cells[c] : null;
                        CellMerge m = merge[r, c];
                        sb.Append("<a:tc");
                        if (m.GridSpan > 1)
                            sb.Append(" gridSpan=\"").Append(PptxParts.Num(m.GridSpan)).Append('"');
                        if (m.RowSpan > 1)
                            sb.Append(" rowSpan=\"").Append(PptxParts.Num(m.RowSpan)).Append('"');
                        if (m.HMerge)
                            sb.Append(" hMerge=\"1\"");
                        if (m.VMerge)
                            sb.Append(" vMerge=\"1\"");
                        sb.Append('>');
                        sb.Append("<a:txBody><a:bodyPr/><a:lstStyle/>");
                        sb.Append(CellParagraphs(cell));
                        sb.Append("</a:txBody>");
                        sb.Append(CellProps(table.Borderless));
                        sb.Append("</a:tc>");
                    }
                    sb.Append("</a:tr>");
                }
                sb.Append("</a:tbl></a:graphicData></a:graphic></p:graphicFrame>");
                _shapes.Append(sb);
                _shapeId++;
            }

            /// <summary>Ширина колонки в EMU; неизвестная — равная доля ширины таблицы.</summary>
            private long ColumnWidthEmu(OcrTable table, int column, double tableWidthPt, int columns)
            {
                double pt = table.ColumnWidthsPt != null && column < table.ColumnWidthsPt.Count
                    ? table.ColumnWidthsPt[column]
                    : (columns > 0 ? tableWidthPt / columns : tableWidthPt);
                long emu = _fit.Size(pt);
                return emu > 0 ? emu : 1;
            }

            /// <summary>
            /// Абзацы ячейки. Пустая ячейка всё равно получает один пустой абзац: текстовое
            /// тело без абзаца разметка запрещает. Выключка по ширине в ячейке заменяется
            /// выравниванием влево — растягивать три слова на всю ширину колонки незачем.
            /// </summary>
            private string CellParagraphs(OcrTableCell cell)
            {
                var sb = new StringBuilder(128);
                if (cell != null && cell.Paragraphs != null)
                    foreach (OcrParagraph paragraph in cell.Paragraphs)
                    {
                        string body = Runs(paragraph);
                        if (body.Length == 0)
                            continue;
                        OcrAlignment align = paragraph.Alignment == OcrAlignment.Justify
                            ? OcrAlignment.Left
                            : paragraph.Alignment;
                        sb.Append("<a:p>").Append(ParagraphProps(paragraph, align)).Append(body).Append("</a:p>");
                    }
                if (sb.Length == 0)
                    sb.Append("<a:p/>");
                return sb.ToString();
            }

            /// <summary>
            /// Свойства ячейки: границы. Безлиновочная сетка (восстановленная по расположению
            /// текста, а не по линиям) рисуется БЕЗ границ — иначе на слайде появилась бы
            /// таблица, которой в источнике не было.
            /// </summary>
            private static string CellProps(bool borderless)
            {
                if (borderless)
                    return "<a:tcPr><a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR>"
                        + "<a:lnT><a:noFill/></a:lnT><a:lnB><a:noFill/></a:lnB></a:tcPr>";
                const string line = "<a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>";
                return "<a:tcPr>"
                    + "<a:lnL w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">" + line + "</a:lnL>"
                    + "<a:lnR w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">" + line + "</a:lnR>"
                    + "<a:lnT w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">" + line + "</a:lnT>"
                    + "<a:lnB w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">" + line + "</a:lnB>"
                    + "</a:tcPr>";
            }

            // ---------- картинка ----------

            private void AppendImage(OcrImage image)
            {
                // Картинки приходят из разбора в PNG. Для фотографии это втрое тяжелее JPEG,
                // а колода из полусотни слайдов на фотографиях распухала на десятки мегабайт.
                bool jpeg;
                string part = _media.Add(RasterUtil.PreferSmaller(image.Png, out jpeg), jpeg);
                if (part == null)
                    return;
                string relId = ImageRel(part);
                var sb = new StringBuilder(512);
                sb.Append("<p:pic><p:nvPicPr><p:cNvPr id=\"").Append(PptxParts.Num(_shapeId))
                  .Append("\" name=\"Picture ").Append(PptxParts.Num(_shapeId)).Append("\"/>");
                sb.Append("<p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>");
                sb.Append("<p:blipFill><a:blip r:embed=\"").Append(relId)
                  .Append("\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>");
                sb.Append("<p:spPr>");
                AppendFrame(sb, image.LeftPt, image.TopPt, image.WidthPt, image.HeightPt);
                sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>");
                _shapes.Append(sb);
                _shapeId++;
            }

            // ---------- общее ----------

            /// <summary>
            /// Рамка фигуры: положение и размер в EMU. Нулевой размер разметка запрещает,
            /// поэтому вырожденная сторона получает минимум в один EMU — фигура остаётся
            /// невидимой, но файл корректен.
            /// </summary>
            private void AppendFrame(StringBuilder sb, double leftPt, double topPt, double widthPt, double heightPt)
            {
                long cx = _fit.Size(widthPt);
                long cy = _fit.Size(heightPt);
                sb.Append("<a:xfrm><a:off x=\"").Append(PptxParts.Num(_fit.X(leftPt)))
                  .Append("\" y=\"").Append(PptxParts.Num(_fit.Y(topPt))).Append("\"/>");
                sb.Append("<a:ext cx=\"").Append(PptxParts.Num(cx > 0 ? cx : 1))
                  .Append("\" cy=\"").Append(PptxParts.Num(cy > 0 ? cy : 1)).Append("\"/></a:xfrm>");
            }

            /// <summary>
            /// Связь на картинку пакета (одна на часть, сколько бы фигур её ни звало). Адрес
            /// связи — относительный от папки слайда, поэтому из имени части берётся только
            /// имя файла.
            /// </summary>
            private string ImageRel(string mediaPart)
            {
                int slash = mediaPart.LastIndexOf('/');
                string file = slash >= 0 ? mediaPart.Substring(slash + 1) : mediaPart;
                return Rel(mediaPart, PptxParts.RelImage, "../media/" + file, false);
            }

            /// <summary>Связь на внешний адрес; null — ссылки нет или она пустая.</summary>
            private string HyperlinkRel(string uri)
            {
                if (string.IsNullOrEmpty(uri))
                    return null;
                string trimmed = uri.Trim();
                return trimmed.Length == 0 ? null : Rel("uri:" + trimmed, PptxParts.RelHyperlink, trimmed, true);
            }

            /// <summary>Связь по ключу: повторный вызов с тем же ключом отдаёт прежний идентификатор.</summary>
            private string Rel(string key, string type, string target, bool external)
            {
                string id;
                if (_relIdByTarget.TryGetValue(key, out id))
                    return id;
                id = "rId" + _relId.ToString(CultureInfo.InvariantCulture);
                _relId++;
                _rels.Add(PptxParts.Relationship(id, type, target, external));
                _relIdByTarget[key] = id;
                return id;
            }
        }

        /// <summary>Подъёмная часть строки в пунктах (см. <see cref="AscentGapFactor"/>). Чистая — под тест.</summary>
        internal static double AscentGapPt(double fontPt)
        {
            double font = fontPt > 0 ? fontPt : FontResolver.DefaultFontSize;
            return AscentGapFactor * font;
        }

        internal static bool IsSingleLine(double heightPt, double fontPt)
        {
            if (heightPt <= 0)
                return true; // вырожденная рамка — точно не многострочная
            double font = fontPt > 0 ? fontPt : FontResolver.DefaultFontSize;
            return heightPt <= SingleLineFactor * font;
        }

        /// <summary>
        /// Ширина надписи с запасом (см. <see cref="WidthSlackFraction"/>), не выходящая за
        /// правый край страницы. Чистая — под тест.
        /// </summary>
        internal static double SlackWidthPt(double leftPt, double widthPt, double pageWidthPt)
        {
            if (widthPt <= 0)
                widthPt = pageWidthPt > leftPt ? pageWidthPt - leftPt : pageWidthPt;
            double slack = widthPt * WidthSlackFraction;
            if (slack < WidthSlackMinPt)
                slack = WidthSlackMinPt;
            double result = widthPt + slack;
            double room = pageWidthPt - leftPt;
            if (room > 0 && result > room)
                result = room;
            return result;
        }

        /// <summary>Как ячейка участвует в объединении: пролёты владельца и пометки накрытия.</summary>
        internal struct CellMerge
        {
            public bool HMerge;   // накрыта объединением СЛЕВА (в своей строке)
            public bool VMerge;   // накрыта объединением СВЕРХУ
            public int GridSpan;  // на сколько колонок тянется объединение (1 — не тянется)
            public int RowSpan;   // на сколько строк тянется объединение
        }

        /// <summary>
        /// Разметка объединений таблицы: для каждой ячейки — пролёты и пометки накрытия.
        /// Правило разметки: ячейка-владелец несёт пролёты; накрытая справа в той же строке —
        /// hMerge и ПРОЛЁТ ПО СТРОКАМ владельца; накрытая снизу — vMerge и пролёт по колонкам;
        /// накрытая по диагонали — обе пометки. Ячейки, накрытые объединением, из разметки НЕ
        /// исчезают: в строке обязано быть ровно столько ячеек, сколько колонок в сетке, иначе
        /// PowerPoint объявит файл повреждённым. Чистая — под тест.
        /// </summary>
        internal static CellMerge[,] MergeGrid(OcrTable table)
        {
            int rows = table == null || table.Rows == null ? 0 : table.Rows.Count;
            int cols = table == null ? 0 : table.ColumnCount;
            var grid = new CellMerge[rows, cols];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                {
                    grid[r, c].GridSpan = 1;
                    grid[r, c].RowSpan = 1;
                }

            for (int r = 0; r < rows; r++)
            {
                OcrTableRow row = table.Rows[r];
                for (int c = 0; c < cols; c++)
                {
                    OcrTableCell cell = row != null && row.Cells != null && c < row.Cells.Count ? row.Cells[c] : null;
                    if (cell == null || cell.Covered)
                        continue;
                    int cs = cell.ColSpan > 1 ? cell.ColSpan : 1;
                    int rs = cell.RowSpan > 1 ? cell.RowSpan : 1;
                    if (cs == 1 && rs == 1)
                        continue;
                    for (int dr = 0; dr < rs && r + dr < rows; dr++)
                        for (int dc = 0; dc < cs && c + dc < cols; dc++)
                            grid[r + dr, c + dc] = new CellMerge
                            {
                                HMerge = dc > 0,
                                VMerge = dr > 0,
                                // Пролёт по колонкам объявляет ЛЕВАЯ ячейка своей полосы,
                                // по строкам — ВЕРХНЯЯ. Остальные несут только пометку.
                                GridSpan = dc > 0 ? 1 : cs,
                                RowSpan = dr > 0 ? 1 : rs
                            };
                }
            }
            return grid;
        }

        /// <summary>
        /// Ран-заполнитель, который на слайде не нужен. Разбор превращает одиночную
        /// горизонтальную линию в слово из «_» (<see cref="PdfTextExtract.AddRuleWords"/>) —
        /// это подпорка ПОТОКОВОЙ вёрстки: в Word линию по координатам не поставить, и
        /// прочерк бланка иначе пропал бы. У слайда координаты есть, а сама линия приходит
        /// подложкой страницы, поэтому здесь такой ран — только шум: сетка диаграммы
        /// превращалась бы в частокол подчёркиваний. Чистая — под тест.
        /// </summary>
        internal static bool IsFillerRule(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;
            // Пробелы внутри значения не мешают: соседние прочерки склеиваются в один ран
            // через пробел («____ ____»), и без этой чистки правило их пропускало бы.
            var sb = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
                if (!char.IsWhiteSpace(text[i]))
                    sb.Append(text[i]);
            return sb.Length >= 2 && PdfTextExtract.IsUnderscoreOnly(sb.ToString());
        }

        /// <summary>Цвет 0xRRGGBB → шесть шестнадцатеричных цифр разметки. Чистая — под тест.</summary>
        internal static string Rgb(int colorArgb)
        {
            return (colorArgb & 0xFFFFFF).ToString("X6", CultureInfo.InvariantCulture);
        }
    }
}
