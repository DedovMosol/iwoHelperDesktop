using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Постоянные части презентации и генераторы тех, что зависят от числа слайдов и размера
    /// листа. Всё, что не меняется от документа к документу (тема, образец, макет «пустой
    /// слайд»), лежит здесь строками: собирать это построителем XML значило бы писать код
    /// ради текста, который никогда не менялся.
    ///
    /// Порядок дочерних элементов в OOXML — часть схемы, а не оформление: `notesSz` обязан
    /// идти после `sldSz`, `p:bg` — перед `p:spTree`, и перестановка ломает файл так же
    /// надёжно, как отсутствие элемента. Поэтому шаблоны правятся целиком и проверяются
    /// валидатором, а не собираются из кусков по месту.
    /// </summary>
    internal static class PptxParts
    {
        // ---- типы содержимого частей ----
        public const string CtPresentation = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
        public const string CtSlide = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
        public const string CtSlideMaster = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
        public const string CtSlideLayout = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
        public const string CtTheme = "application/vnd.openxmlformats-officedocument.theme+xml";
        public const string CtTableStyles = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";
        public const string CtCoreProps = "application/vnd.openxmlformats-package.core-properties+xml";
        public const string CtAppProps = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        public const string CtPng = "image/png";
        public const string CtJpeg = "image/jpeg";

        // ---- типы связей ----
        private const string RelOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string RelCoreProps = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        private const string RelAppProps = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        public const string RelSlideMaster = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
        public const string RelSlide = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
        public const string RelSlideLayout = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
        public const string RelTheme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
        public const string RelTableStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles";
        public const string RelImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        public const string RelHyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

        // ---- имена частей ----
        public const string PresentationPart = "ppt/presentation.xml";
        public const string SlideMasterPart = "ppt/slideMasters/slideMaster1.xml";
        public const string SlideLayoutPart = "ppt/slideLayouts/slideLayout1.xml";
        public const string ThemePart = "ppt/theme/theme1.xml";
        public const string TableStylesPart = "ppt/tableStyles.xml";
        public const string CorePropsPart = "docProps/core.xml";
        public const string AppPropsPart = "docProps/app.xml";

        public const string XmlHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";

        /// <summary>Пространства имён презентации — в корневом элементе каждой части.</summary>
        public const string NsPml = " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
            + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
            + " xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"";

        /// <summary>Пустое дерево фигур: узел-группа с нулевой рамкой, как его пишет PowerPoint.</summary>
        public const string EmptySpTree =
            "<p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
            + "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>"
            + "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr></p:spTree>";

        /// <summary>Один элемент связи. Внешние ссылки (гиперссылки) требуют TargetMode.</summary>
        public static string Relationship(string id, string type, string target, bool external = false)
        {
            var sb = new StringBuilder(160);
            sb.Append("<Relationship Id=\"").Append(id)
              .Append("\" Type=\"").Append(type)
              .Append("\" Target=\"").Append(XmlText.Encode(target)).Append('"');
            if (external)
                sb.Append(" TargetMode=\"External\"");
            sb.Append("/>");
            return sb.ToString();
        }

        /// <summary>Файл связей из готовых элементов <see cref="Relationship"/>.</summary>
        public static string Relationships(IEnumerable<string> items)
        {
            var sb = new StringBuilder(512);
            sb.Append(XmlHeader);
            sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (string item in items)
                sb.Append(item);
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        /// <summary>Корневые связи пакета: сама презентация и свойства документа.</summary>
        public static string RootRels()
        {
            return Relationships(new[]
            {
                Relationship("rId1", RelOfficeDocument, PresentationPart),
                Relationship("rId2", RelCoreProps, CorePropsPart),
                Relationship("rId3", RelAppProps, AppPropsPart)
            });
        }

        /// <summary>
        /// presentation.xml: образец, список слайдов, размер слайда и размер страницы заметок.
        /// slideCount — сколько слайдов в колоде; размеры — в EMU. Идентификаторы слайдов
        /// начинаются с 256 (меньшие схема запрещает), идентификатор образца — с 2147483648.
        /// </summary>
        public static string Presentation(int slideCount, long slideCx, long slideCy)
        {
            var sb = new StringBuilder(512);
            sb.Append(XmlHeader);
            sb.Append("<p:presentation").Append(NsPml).Append(" saveSubsetFonts=\"1\">");
            sb.Append("<p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst>");
            sb.Append("<p:sldIdLst>");
            for (int i = 0; i < slideCount; i++)
                sb.Append("<p:sldId id=\"").Append(256 + i)
                  .Append("\" r:id=\"").Append(SlideRelId(i)).Append("\"/>");
            sb.Append("</p:sldIdLst>");
            sb.Append("<p:sldSz cx=\"").Append(Num(slideCx)).Append("\" cy=\"").Append(Num(slideCy)).Append("\"/>");
            // Размер страницы заметок обязателен по схеме; книжный лист под слайдом — как у PowerPoint.
            sb.Append("<p:notesSz cx=\"6858000\" cy=\"9144000\"/>");
            sb.Append("</p:presentation>");
            return sb.ToString();
        }

        /// <summary>Идентификатор связи слайда (0-based индекс) в presentation.xml.rels.</summary>
        public static string SlideRelId(int slideIndex)
        {
            return "rId" + (slideIndex + 2).ToString(CultureInfo.InvariantCulture); // rId1 занят образцом
        }

        /// <summary>Имя части слайда (0-based индекс).</summary>
        public static string SlidePart(int slideIndex)
        {
            return "ppt/slides/slide" + (slideIndex + 1).ToString(CultureInfo.InvariantCulture) + ".xml";
        }

        /// <summary>Связи презентации: образец, все слайды по порядку, стили таблиц.</summary>
        public static string PresentationRels(int slideCount)
        {
            var items = new List<string>(slideCount + 2);
            items.Add(Relationship("rId1", RelSlideMaster, "slideMasters/slideMaster1.xml"));
            for (int i = 0; i < slideCount; i++)
                items.Add(Relationship(SlideRelId(i), RelSlide, "slides/slide" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".xml"));
            items.Add(Relationship("rId" + (slideCount + 2).ToString(CultureInfo.InvariantCulture), RelTableStyles, "tableStyles.xml"));
            return Relationships(items);
        }

        /// <summary>Образец слайдов: пустое дерево фигур, карта цветов, единственный макет.</summary>
        public static string SlideMaster()
        {
            var sb = new StringBuilder(1024);
            sb.Append(XmlHeader);
            sb.Append("<p:sldMaster").Append(NsPml).Append('>');
            sb.Append("<p:cSld>");
            sb.Append("<p:bg><p:bgPr><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>");
            sb.Append(EmptySpTree);
            sb.Append("</p:cSld>");
            sb.Append("<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\""
                + " accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\""
                + " hlink=\"hlink\" folHlink=\"folHlink\"/>");
            sb.Append("<p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rId1\"/></p:sldLayoutIdLst>");
            sb.Append("</p:sldMaster>");
            return sb.ToString();
        }

        /// <summary>Связи образца: макет и тема.</summary>
        public static string SlideMasterRels()
        {
            return Relationships(new[]
            {
                Relationship("rId1", RelSlideLayout, "../slideLayouts/slideLayout1.xml"),
                Relationship("rId2", RelTheme, "../theme/theme1.xml")
            });
        }

        /// <summary>
        /// Единственный макет — «пустой слайд». Заполнителей нет намеренно: содержимое
        /// приходит из PDF готовыми рамками, а рамка-заполнитель навязала бы им своё место
        /// и свой кегль.
        /// </summary>
        public static string SlideLayout()
        {
            var sb = new StringBuilder(512);
            sb.Append(XmlHeader);
            sb.Append("<p:sldLayout").Append(NsPml).Append(" type=\"blank\" preserve=\"1\">");
            sb.Append("<p:cSld name=\"Blank\">").Append(EmptySpTree).Append("</p:cSld>");
            sb.Append("<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>");
            sb.Append("</p:sldLayout>");
            return sb.ToString();
        }

        /// <summary>Связи макета: свой образец.</summary>
        public static string SlideLayoutRels()
        {
            return Relationships(new[] { Relationship("rId1", RelSlideMaster, "../slideMasters/slideMaster1.xml") });
        }

        /// <summary>Связи слайда: макет + переданные (картинки, гиперссылки).</summary>
        public static string SlideRels(IEnumerable<string> extra)
        {
            var items = new List<string> { Relationship("rId1", RelSlideLayout, "../slideLayouts/slideLayout1.xml") };
            if (extra != null)
                items.AddRange(extra);
            return Relationships(items);
        }

        /// <summary>Пустая таблица стилей таблиц: часть обязательна по связи, содержимое — нет.</summary>
        public static string TableStyles()
        {
            return XmlHeader
                + "<a:tblStyleLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
                + " def=\"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}\"/>";
        }

        /// <summary>
        /// Свойства документа. Порядок элементов задан схемой, а не вкусом: created, creator,
        /// modified, title. whenUtc передаётся снаружи — иначе один и тот же вход давал бы
        /// разные файлы, и тесты не могли бы сравнивать результат.
        /// </summary>
        public static string CoreProps(string title, DateTime whenUtc)
        {
            string stamp = whenUtc.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
            var sb = new StringBuilder(512);
            sb.Append(XmlHeader);
            sb.Append("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\""
                + " xmlns:dc=\"http://purl.org/dc/elements/1.1/\""
                + " xmlns:dcterms=\"http://purl.org/dc/terms/\""
                + " xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\""
                + " xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            sb.Append("<dcterms:created xsi:type=\"dcterms:W3CDTF\">").Append(stamp).Append("</dcterms:created>");
            sb.Append("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">").Append(stamp).Append("</dcterms:modified>");
            sb.Append("<dc:title>").Append(XmlText.Encode(title)).Append("</dc:title>");
            sb.Append("</cp:coreProperties>");
            return sb.ToString();
        }

        /// <summary>Расширенные свойства: число слайдов и имя приложения (порядок — по схеме).</summary>
        public static string AppProps(int slideCount, string application)
        {
            var sb = new StringBuilder(384);
            sb.Append(XmlHeader);
            sb.Append("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\""
                + " xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">");
            sb.Append("<Slides>").Append(slideCount.ToString(CultureInfo.InvariantCulture)).Append("</Slides>");
            sb.Append("<Application>").Append(XmlText.Encode(application)).Append("</Application>");
            sb.Append("</Properties>");
            return sb.ToString();
        }

        /// <summary>Целое в разметку — всегда инвариантной культурой (иначе разделитель поедет).</summary>
        public static string Num(long value)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>Целое в разметку.</summary>
        public static string Num(int value)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Тема оформления. Обязательна: образец слайдов без темы недействителен. Содержимое —
        /// стандартная офисная схема цветов и шрифтов; на вид документа она не влияет, потому
        /// что и цвет, и шрифт каждого рана задаются явно.
        /// </summary>
        public static string Theme()
        {
            var sb = new StringBuilder(4096);
            sb.Append(XmlHeader);
            sb.Append("<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office\">");
            sb.Append("<a:themeElements>");
            sb.Append("<a:clrScheme name=\"Office\">");
            sb.Append("<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>");
            sb.Append("<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>");
            sb.Append("<a:dk2><a:srgbClr val=\"44546A\"/></a:dk2>");
            sb.Append("<a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2>");
            sb.Append("<a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1>");
            sb.Append("<a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2>");
            sb.Append("<a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3>");
            sb.Append("<a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>");
            sb.Append("<a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5>");
            sb.Append("<a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6>");
            sb.Append("<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink>");
            sb.Append("<a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>");
            sb.Append("</a:clrScheme>");
            sb.Append("<a:fontScheme name=\"Office\">");
            sb.Append("<a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>");
            sb.Append("<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>");
            sb.Append("</a:fontScheme>");
            sb.Append("<a:fmtScheme name=\"Office\">");
            // По три стиля в каждом списке — минимум, которого требует схема.
            sb.Append("<a:fillStyleLst>");
            sb.Append("<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>");
            sb.Append("<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>");
            sb.Append("<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>");
            sb.Append("</a:fillStyleLst>");
            sb.Append("<a:lnStyleLst>");
            for (int i = 0; i < 3; i++)
                sb.Append("<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">"
                    + "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
                    + "<a:prstDash val=\"solid\"/></a:ln>");
            sb.Append("</a:lnStyleLst>");
            sb.Append("<a:effectStyleLst>");
            for (int i = 0; i < 3; i++)
                sb.Append("<a:effectStyle><a:effectLst/></a:effectStyle>");
            sb.Append("</a:effectStyleLst>");
            sb.Append("<a:bgFillStyleLst>");
            for (int i = 0; i < 3; i++)
                sb.Append("<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>");
            sb.Append("</a:bgFillStyleLst>");
            sb.Append("</a:fmtScheme>");
            sb.Append("</a:themeElements>");
            sb.Append("</a:theme>");
            return sb.ToString();
        }
    }
}
