using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Аудит E: извлечение текста PDF и писатели «PDF → Word/PowerPoint/текст». Живые тесты
    /// собирают PDF вручную (сырые потоки), чтобы воспроизвести ровно то, что ломалось:
    /// заливку узором, /Rotate 90, ссылки с опасными адресами, линии в миллиард пунктов.
    /// </summary>
    internal static partial class UnitTests
    {
        // ---------- сырой PDF ----------

        /// <summary>
        /// Минимальный PDF из готовых тел объектов (objects[0] — объект 1, каталог). Таблица
        /// xref считается честно: «починка» битой таблицы маскировала бы проверяемое.
        /// Только ASCII — двоичные потоки идут через ASCIIHex.
        /// </summary>
        private static void WriteRawPdf(string path, IList<string> objects)
        {
            var sb = new StringBuilder("%PDF-1.4\n");
            var offsets = new List<int>();
            for (int i = 0; i < objects.Count; i++)
            {
                offsets.Add(Encoding.ASCII.GetByteCount(sb.ToString()));
                sb.Append(i + 1).Append(" 0 obj\n").Append(objects[i]).Append("\nendobj\n");
            }
            int xref = Encoding.ASCII.GetByteCount(sb.ToString());
            sb.Append("xref\n0 ").Append(objects.Count + 1).Append('\n');
            sb.Append("0000000000 65535 f \n");
            foreach (int offset in offsets)
                sb.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
            sb.Append("trailer\n<< /Size ").Append(objects.Count + 1).Append(" /Root 1 0 R >>\nstartxref\n")
              .Append(xref).Append("\n%%EOF\n");
            File.WriteAllBytes(path, Encoding.ASCII.GetBytes(sb.ToString()));
        }

        /// <summary>Тело объекта-потока с честной длиной.</summary>
        private static string RawStream(string dict, string content)
        {
            return "<< " + dict + " /Length " + Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture)
                + " >>\nstream\n" + content + "\nendstream";
        }

        /// <summary>
        /// Одностраничный PDF 400×600: объекты 1–5 — каталог, страницы, страница, Helvetica (/F1),
        /// поток; extraObjects — с номера 6.
        /// </summary>
        private static void WriteRawPage(string path, string content, string extraResources = "",
            string pageExtras = "", params string[] extraObjects)
        {
            var objects = new List<string>
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 600] " + pageExtras
                    + " /Resources << /Font << /F1 4 0 R >> " + extraResources + " >> /Contents 5 0 R >>",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
                RawStream("", content)
            };
            objects.AddRange(extraObjects);
            WriteRawPdf(path, objects);
        }

        // RGB-картинка 2×1: левый пиксель красный, правый синий.
        private static readonly string RawRedBlueImage = RawStream(
            "/Type /XObject /Subtype /Image /Width 2 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /ASCIIHexDecode",
            "FF00000000FF>");

        /// <summary>Трёхстраничный PDF: на каждой странице своя строка и картинка.</summary>
        private static void WriteRawThreePages(string path)
        {
            var objects = new List<string>
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R] /Count 3 >>"
            };
            for (int i = 0; i < 3; i++)
                objects.Add("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 600] /Resources << /Font << /F1 6 0 R >> "
                    + "/XObject << /Im1 10 0 R >> >> /Contents " + (7 + i) + " 0 R >>");
            objects.Add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
            string[] texts = { "Alpha first page", "Bravo second page", "Charlie third page" };
            foreach (string t in texts)
                objects.Add(RawStream("", "BT /F1 14 Tf 50 500 Td (" + t + ") Tj ET q 60 0 0 30 50 300 cm /Im1 Do Q"));
            objects.Add(RawRedBlueImage);
            WriteRawPdf(path, objects);
        }

        private static List<string> WordTexts(PdfPageText page)
        {
            var result = new List<string>();
            foreach (PdfWord w in page.Words)
                result.Add(w.Text);
            return result;
        }

        private static string WordSignature(PdfPageText page)
        {
            var sb = new StringBuilder();
            foreach (PdfWord w in page.Words)
                sb.Append(w.Text).Append('@').Append(Math.Round(w.Left, 2)).Append(',')
                  .Append(Math.Round(w.Bottom, 2)).Append(';');
            return sb.ToString();
        }

        // ---------- 1. цвет узором ----------

        [Test("Аудит E: текст с заливкой узором (Pattern) извлекается, а не роняет файл")]
        private static void AuditEPatternColorText()
        {
            using (var dir = new TempDir("iwo_auditE_pattern_"))
            {
                string path = dir.File("pattern.pdf");
                // Градиентная заливка текста: «/Pattern cs /P1 scn» — цвет буквы у PdfPig
                // становится PatternColor, у которого ToRGBValues бросает исключение.
                WriteRawPage(path,
                    "/Pattern cs /P1 scn BT /F1 24 Tf 50 500 Td (Gradient words) Tj ET "
                    + "0 g BT /F1 12 Tf 50 400 Td (Plain line) Tj ET",
                    "/Pattern << /P1 6 0 R >>", "",
                    "<< /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 400 0] "
                    + "/Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> "
                    + "/Extend [true true] >> >>");
                List<PdfPageText> pages = PdfTextExtract.Extract(path);
                AssertEq(1, pages.Count, "страница извлечена");
                string all = string.Join(" ", WordTexts(pages[0]).ToArray());
                AssertTrue(all.Contains("Gradient") && all.Contains("words"), "слова с узором на месте: " + all);
                AssertTrue(all.Contains("Plain"), "обычная строка на месте: " + all);
                foreach (PdfWord w in pages[0].Words)
                    if (w.Text == "Gradient")
                        AssertEq(0, w.ColorArgb, "текст узором пишется чёрным");
                // Тот же файл — и для Сравнения (свой режим, без картинок).
                List<PdfPageText> review = PdfTextExtract.Extract(path, null, null, null,
                    PageLayoutMode.Review, null, false);
                AssertTrue(PlainText.Page(review[0]).Contains("Gradient"), "Review-извлечение тоже не падает");
            }
        }

        // ---------- 2. адреса ссылок ----------

        [Test("Аудит E: ссылки PDF — только http/https/mailto, в каноничной форме")]
        private static void AuditESafeLinkUri()
        {
            AssertEq("https://example.org/a?b=1&c=2", PdfTextExtract.SafeLinkUri("https://example.org/a?b=1&c=2"), "https как есть");
            AssertEq("http://example.org/", PdfTextExtract.SafeLinkUri("  HTTP://Example.org  "), "схема и хост нормализованы, пробелы сняты");
            AssertEq("mailto:user@example.org", PdfTextExtract.SafeLinkUri("mailto:user@example.org"), "почта разрешена");
            AssertEq("https://example.org/a%20b", PdfTextExtract.SafeLinkUri("https://example.org/a b"), "пробел экранирован");
            AssertEq<string>(null, PdfTextExtract.SafeLinkUri("javascript:alert(1)"), "javascript: запрещён");
            AssertEq<string>(null, PdfTextExtract.SafeLinkUri("file:///C:/Windows/System32/calc.exe"), "file: запрещён");
            AssertEq<string>(null, PdfTextExtract.SafeLinkUri("\\\\evil-host\\share\\doc"), "UNC-путь (утечка NTLM) запрещён");
            AssertEq<string>(null, PdfTextExtract.SafeLinkUri("ms-settings:privacy"), "обработчики ms-* запрещены");
            AssertEq<string>(null, PdfTextExtract.SafeLinkUri("www.example.org"), "относительный адрес — не ссылка");
            AssertEq<string>(null, PdfTextExtract.SafeLinkUri(""), "пусто");
            AssertEq<string>(null, PdfTextExtract.SafeLinkUri(null), "null");
        }

        [Test("Аудит E: опасная ссылка PDF не становится внешней связью, безопасная — становится")]
        private static void AuditEUnsafeLinkLive()
        {
            using (var dir = new TempDir("iwo_auditE_links_"))
            {
                string path = dir.File("links.pdf");
                WriteRawPage(path,
                    "BT /F1 14 Tf 50 500 Td (Safe link) Tj ET BT /F1 14 Tf 50 400 Td (Evil link) Tj ET",
                    "", "/Annots [6 0 R 7 0 R]",
                    "<< /Type /Annot /Subtype /Link /Rect [40 490 200 520] /A << /S /URI /URI (https://example.org/ok) >> >>",
                    "<< /Type /Annot /Subtype /Link /Rect [40 390 200 420] /A << /S /URI /URI (javascript:alert\\(1\\)) >> >>");
                List<PdfPageText> pages = PdfTextExtract.Extract(path);
                foreach (PdfWord w in pages[0].Words)
                {
                    if (w.Text == "Safe")
                        AssertEq("https://example.org/ok", w.Uri, "безопасная ссылка проставлена");
                    if (w.Text == "Evil")
                        AssertEq<string>(null, w.Uri, "javascript: выброшен — слово стало обычным текстом");
                }
            }
        }

        // ---------- 3. /Rotate 90 ----------

        [Test("Аудит E: рамки PdfPig нормализуются по четырём углам, поворот картинки — по её оси")]
        private static void AuditENormalizedRects()
        {
            // Рамка картинки на /Rotate 90 так, как её отдаёт PdfPig 0.1.9 (проверено на файле):
            // BL=(300,350) BR=(300,250) TL=(350,350) TR=(350,250).
            PdfTextExtract.RectPt r = PdfTextExtract.NormalizedRect(300, 350, 300, 250, 350, 350, 350, 250);
            AssertTrue(r.Left == 300 && r.Right == 350 && r.Bottom == 250 && r.Top == 350,
                "min/max по углам: " + r.Left + "," + r.Bottom + "," + r.Right + "," + r.Top);
            AssertEq(90, PdfTextExtract.ImageQuarterTurn(300, 350, 300, 250, 350, 350), "ось X картинки смотрит вниз — 90°");
            AssertEq(0, PdfTextExtract.ImageQuarterTurn(50, 300, 150, 300, 50, 350), "обычная картинка — 0°");
            AssertEq(180, PdfTextExtract.ImageQuarterTurn(150, 300, 50, 300, 150, 250), "ось влево — 180°");
            AssertEq(270, PdfTextExtract.ImageQuarterTurn(50, 300, 50, 400, 0, 300), "ось вверх — 270°");
            AssertEq(0, PdfTextExtract.ImageQuarterTurn(150, 300, 50, 300, 150, 350), "зеркало по горизонтали — как прежде, 0°");
            AssertEq(0, PdfTextExtract.ImageQuarterTurn(50, 350, 150, 350, 50, 300), "зеркало по вертикали — 0°");
            AssertEq(0, PdfTextExtract.ImageQuarterTurn(double.NaN, 0, 1, 0, 0, 1), "мусор — без поворота");

            // Перевёрнутая рамка больше не даёт отрицательных перекрытий.
            var word = new PdfTextExtract.RectPt { Left = 310, Right = 340, Bottom = 260, Top = 270 };
            AssertTrue(PdfTextExtract.CoveredByLowImage(word, new List<PdfTextExtract.RectPt>
            {
                PdfTextExtract.NormalizedRect(300, 275, 300, 255, 350, 275, 350, 255)
            }), "низкая картинка из перевёрнутой рамки накрывает слово");
        }

        [Test("Аудит E: страница /Rotate 90 — ссылки находят слова, картинка на месте и повёрнута")]
        private static void AuditERotatedPageLive()
        {
            using (var dir = new TempDir("iwo_auditE_rot_"))
            {
                string path = dir.File("rot90.pdf");
                // Текст повёрнут в пространстве страницы против часовой — на экране (/Rotate 90)
                // он горизонтален, как у сканера, сохранившего альбомный лист книжным.
                WriteRawPage(path,
                    "BT /F1 20 Tf 0 1 -1 0 100 50 Tm (Hello World) Tj ET q 100 0 0 50 200 300 cm /Im1 Do Q",
                    "/XObject << /Im1 6 0 R >>", "/Rotate 90 /Annots [7 0 R]",
                    RawRedBlueImage,
                    "<< /Type /Annot /Subtype /Link /Rect [70 40 110 180] /A << /S /URI /URI (https://example.org/x) >> >>");
                List<PdfPageText> pages = PdfTextExtract.Extract(path);
                PdfPageText page = pages[0];
                AssertTrue(Math.Abs(page.WidthPt - 600) < 1 && Math.Abs(page.HeightPt - 400) < 1,
                    "страница в пространстве отображения: " + page.WidthPt + "×" + page.HeightPt);
                int linked = 0;
                foreach (PdfWord w in page.Words)
                {
                    AssertTrue(w.Left <= w.Right && w.Bottom <= w.Top, "рамка слова нормализована: " + w.Text);
                    if (w.Uri == "https://example.org/x")
                        linked++;
                }
                AssertTrue(page.Text.Contains("Hello") && page.Text.Contains("World"), "текст на месте: " + page.Text);
                AssertEq(2, linked, "ссылка нашла оба слова");

                AssertEq(1, page.Images.Count, "картинка перенесена");
                OcrImage img = page.Images[0];
                AssertTrue(Math.Abs(img.LeftPt - 300) < 0.5 && Math.Abs(img.TopPt - 200) < 0.5
                    && Math.Abs(img.WidthPt - 50) < 0.5 && Math.Abs(img.HeightPt - 100) < 0.5,
                    "рамка картинки — как на экране: " + img.LeftPt + "," + img.TopPt + " " + img.WidthPt + "×" + img.HeightPt);
                using (var ms = new MemoryStream(img.Png))
                using (var bmp = new System.Drawing.Bitmap(ms))
                {
                    AssertTrue(bmp.Height > bmp.Width, "пиксели повёрнуты: " + bmp.Width + "×" + bmp.Height);
                    System.Drawing.Color top = bmp.GetPixel(bmp.Width / 2, 0);
                    System.Drawing.Color bottom = bmp.GetPixel(bmp.Width / 2, bmp.Height - 1);
                    AssertTrue(top.R > 200 && top.B < 60, "сверху — левый (красный) пиксель: " + top);
                    AssertTrue(bottom.B > 200 && bottom.R < 60, "снизу — правый (синий) пиксель: " + bottom);
                }
            }
        }

        [Test("Аудит E: зеркальная картинка не поворачивается на 180°, повёрнутая матрицей — поворачивается")]
        private static void AuditEMirroredImageLive()
        {
            // 2×2: верхний ряд красный|синий, нижний зелёный|белый (первый ряд данных — верх).
            string image = RawStream(
                "/Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /ASCIIHexDecode",
                "FF00000000FF00FF00FFFFFF>");
            Func<string, System.Drawing.Color> topLeftOf = delegate(string matrix)
            {
                using (var dir = new TempDir("iwo_auditE_mirror_"))
                {
                    string path = dir.File("m.pdf");
                    WriteRawPage(path, "BT /F1 12 Tf 50 550 Td (Text) Tj ET q " + matrix + " cm /Im1 Do Q",
                        "/XObject << /Im1 6 0 R >>", "", image);
                    PdfPageText page = PdfTextExtract.Extract(path)[0];
                    AssertEq(1, page.Images.Count, "картинка перенесена: " + matrix);
                    using (var ms = new MemoryStream(page.Images[0].Png))
                    using (var bmp = new System.Drawing.Bitmap(ms))
                        return bmp.GetPixel(0, 0);
                }
            };
            System.Drawing.Color plain = topLeftOf("100 0 0 100 100 100");
            AssertTrue(plain.R > 200 && plain.G < 60 && plain.B < 60, "обычная: слева сверху красный: " + plain);
            // Зеркало по горизонтали: поворот на 180° дал бы белый угол, прежний перенос — красный.
            System.Drawing.Color mirrored = topLeftOf("-100 0 0 100 300 100");
            AssertTrue(!(mirrored.R > 200 && mirrored.G > 200 && mirrored.B > 200), "зеркало не повёрнуто на 180°: " + mirrored);
            // Настоящий поворот на 180° матрицей: слева сверху — белый (бывший правый нижний).
            System.Drawing.Color turned = topLeftOf("-100 0 0 -100 300 300");
            AssertTrue(turned.R > 200 && turned.G > 200 && turned.B > 200, "поворот 180° развёрнут: " + turned);
        }

        // ---------- 4. только нужные страницы ----------

        [Test("Аудит E: номера страниц к разбору — все или только запрошенные")]
        private static void AuditEPageNumbersToExtract()
        {
            AssertEq("1,2,3", string.Join(",", PdfTextExtract.PageNumbersToExtract(3, null)), "null — все");
            AssertEq("1,3", string.Join(",", PdfTextExtract.PageNumbersToExtract(3, new[] { 2, 0, 2, 7, -1 })),
                "по возрастанию, без повторов и несуществующих");
            AssertEq("", string.Join(",", PdfTextExtract.PageNumbersToExtract(3, new int[0])), "пустой набор — ничего");
        }

        [Test("Аудит E: Extract разбирает только запрошенные страницы, без картинок — по флагу")]
        private static void AuditEExtractSelectedPagesLive()
        {
            using (var dir = new TempDir("iwo_auditE_pages_"))
            {
                string path = dir.File("three.pdf");
                WriteRawThreePages(path);
                List<PdfPageText> full = PdfTextExtract.Extract(path);
                AssertEq(3, full.Count, "полное извлечение");

                int calls = 0, lastTotal = -1;
                List<PdfPageText> one = PdfTextExtract.Extract(path,
                    delegate(int done, int total) { calls++; lastTotal = total; }, null, null,
                    PageLayoutMode.Document, new[] { 2 });
                AssertEq(1, one.Count, "одна страница из трёх");
                AssertEq(2, one[0].PageIndex, "исходный индекс сохранён");
                AssertEq(1, calls, "разобрана ровно одна страница");
                AssertEq(1, lastTotal, "прогресс считает только запрошенные");
                AssertEq(full[2].Text, one[0].Text, "текст тот же, что при полном разборе");
                AssertEq(WordSignature(full[2]), WordSignature(one[0]), "слова и рамки те же");
                AssertEq(full[2].Images.Count, one[0].Images.Count, "картинки те же");
                AssertTrue(full[2].Images.Count == 1, "в полном разборе картинка есть");

                List<PdfPageText> noImages = PdfTextExtract.Extract(path, null, null, null,
                    PageLayoutMode.Review, new[] { 0 }, false);
                AssertEq(0, noImages[0].Images.Count, "extractImages:false — картинки не декодируются");
                AssertTrue(noImages[0].Text.Contains("Alpha"), "текст на месте: " + noImages[0].Text);
            }
        }

        [Test("Аудит E: PdfPageExtraction берёт из файла только выбранные страницы, порядок и дубли целы")]
        private static void AuditEPageExtractionLoadLive()
        {
            using (var dir = new TempDir("iwo_auditE_load_"))
            {
                string path = dir.File("three.pdf");
                WriteRawThreePages(path);
                var order = new List<PdfPageRef>
                {
                    new PdfPageRef { SourcePath = path, PageIndex = 2 },
                    new PdfPageRef { SourcePath = path, PageIndex = 0 },
                    new PdfPageRef { SourcePath = path.ToUpperInvariant(), PageIndex = 2 }
                };
                Dictionary<string, HashSet<int>> needed = PdfPageExtraction.NeededPages(order);
                AssertEq(1, needed.Count, "регистр пути не важен — один источник");
                AssertEq(2, needed[path].Count, "две разные страницы");

                int withText;
                List<PdfPageText> pages = PdfPageExtraction.Load(order, null, null, out withText);
                AssertEq(3, pages.Count, "по странице на ссылку");
                AssertTrue(pages[0].Text.Contains("Charlie") && pages[1].Text.Contains("Alpha"), "порядок ссылок");
                AssertTrue(ReferenceEquals(pages[0], pages[2]), "дубль делит одно извлечение");
                AssertEq(3, withText, "страниц с текстом");

                var gone = new List<PdfPageRef> { new PdfPageRef { SourcePath = path, PageIndex = 5 } };
                AssertThrows("несуществующая страница", delegate { PdfPageExtraction.Load(gone, null, null, out withText); });
            }
        }

        // ---------- 5, 6, 13. PowerPoint ----------

        [Test("Аудит E: pptx — отступ и шаг строк масштабируются вписыванием и клампятся")]
        private static void AuditEPptxScaledIndentAndPitch()
        {
            AssertEq(PptxGeometry.MaxTextIndentEmu, PptxGeometry.TextIndentEmu(1e7), "отступ — не больше ST_TextIndent");
            AssertEq(-PptxGeometry.MaxTextIndentEmu, PptxGeometry.TextIndentEmu(-1e7), "и не меньше");
            AssertEq(PptxGeometry.MaxSpacingHundredths, PptxGeometry.SpacingHundredths(1e6), "шаг — не больше spcPts");
            AssertEq(0, PptxGeometry.SpacingHundredths(-5), "отрицательный шаг — нет шага");
            AssertEq(0, PptxGeometry.SpacingHundredths(double.NaN), "NaN — нет шага");
            AssertEq(2050, PptxGeometry.SpacingHundredths(20.5), "обычный шаг в сотых пункта");

            // Две страницы 600×800 задают слайд, третья 1200×1600 вписывается с масштабом 0,5.
            var pages = new List<PdfPageText> { PptxPage(600, 800), PptxPage(600, 800), PptxPage(1200, 1600) };
            var par = new OcrParagraph
            {
                LeftPt = 200, TopPt = 1400, RightPt = 800, BottomPt = 1360,
                Alignment = OcrAlignment.Left, LineCount = 2, MinLeftPt = 200
            };
            par.LineTopsPt.Add(1400); par.LineTopsPt.Add(1380);
            par.LineBaselinesPt.Add(1392); par.LineBaselinesPt.Add(1372);
            par.LineLeftsPt.Add(200); par.LineLeftsPt.Add(280);
            par.Runs.Add(new OcrRun { Text = "первая", FontSizePt = 12 });
            par.Runs.Add(new OcrRun { Text = "вторая", FontSizePt = 12, StartsLine = true });
            pages[2].Paragraphs.Add(par);
            XDocument slide = PptxSlideXml(PptxWriter.BuildPackage(pages, PptxStamp), 3);
            var props = new List<XElement>(slide.Descendants(PptxA + "pPr"));
            AssertEq(2, props.Count, "две строки");
            AssertEq(PptxGeometry.Emu(40).ToString(CultureInfo.InvariantCulture), (string)props[1].Attribute("indent"),
                "отступ 80 pt при масштабе 0,5 — 40 pt");
            XElement spc = PptxFirst(props[1].Descendants(PptxA + "spcPts"));
            AssertEq("1000", (string)spc.Attribute("val"), "шаг 20 pt при масштабе 0,5 — 10 pt");

            // Мусорная геометрия не выходит за пределы разметки (иначе «файл повреждён»).
            PdfPageText wild = PptxPage(600, 800);
            var huge = new OcrParagraph
            {
                LeftPt = 0, TopPt = 700, RightPt = 500, BottomPt = 650, Alignment = OcrAlignment.Left,
                LineCount = 2, MinLeftPt = 0
            };
            huge.LineTopsPt.Add(700); huge.LineTopsPt.Add(-1e6);
            huge.LineBaselinesPt.Add(692); huge.LineBaselinesPt.Add(-1e6);
            huge.LineLeftsPt.Add(0); huge.LineLeftsPt.Add(1e7);
            huge.Runs.Add(new OcrRun { Text = "a", FontSizePt = 12 });
            huge.Runs.Add(new OcrRun { Text = "b", FontSizePt = 12, StartsLine = true });
            wild.Paragraphs.Add(huge);
            XDocument wildSlide = PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { wild }, PptxStamp), 1);
            var wildProps = new List<XElement>(wildSlide.Descendants(PptxA + "pPr"));
            AssertEq(PptxGeometry.MaxTextIndentEmu.ToString(CultureInfo.InvariantCulture),
                (string)wildProps[1].Attribute("indent"), "отступ прижат к пределу");
            AssertEq(PptxGeometry.MaxSpacingHundredths.ToString(CultureInfo.InvariantCulture),
                (string)PptxFirst(wildProps[1].Descendants(PptxA + "spcPts")).Attribute("val"), "шаг прижат к пределу");
            AssertPptxValid(new List<PdfPageText> { wild });
        }

        [Test("Аудит E: pptx — подчёркивание атрибутом только без подложки")]
        private static void AuditEPptxUnderlineWithBackground()
        {
            Func<bool, XElement> runProps = delegate(bool withBackground)
            {
                PdfPageText page = PptxPage(960, 540);
                var par = new OcrParagraph { LeftPt = 40, TopPt = 500, RightPt = 300, BottomPt = 480 };
                par.Runs.Add(new OcrRun { Text = "подчёркнуто", FontSizePt = 12, Underline = true });
                page.Paragraphs.Add(par);
                List<Background> backgrounds = withBackground
                    ? new List<Background> { new Background { Data = PptxTinyPng() } } : null;
                XDocument slide = PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { page },
                    PptxStamp, null, null, backgrounds), 1);
                return PptxFirst(slide.Descendants(PptxA + "rPr"));
            };
            AssertEq("sng", (string)runProps(false).Attribute("u"), "без подложки линию рисует сам текст");
            AssertEq<string>(null, (string)runProps(true).Attribute("u"),
                "подложка уже несёт линию — второй черты нет");
        }

        [Test("Аудит E: pptx — ссылки, разные регистром, — разные связи; опасные — без связи")]
        private static void AuditEPptxHyperlinkCaseAndSafety()
        {
            PdfPageText page = PptxPage(600, 800);
            var par = new OcrParagraph { LeftPt = 10, TopPt = 700, RightPt = 500, BottomPt = 680 };
            par.Runs.Add(new OcrRun { Text = "A ", FontSizePt = 12, Uri = "https://example.org/Doc?id=A" });
            par.Runs.Add(new OcrRun { Text = "a ", FontSizePt = 12, Uri = "https://example.org/doc?id=a" });
            par.Runs.Add(new OcrRun { Text = "js ", FontSizePt = 12, Uri = "javascript:alert(1)" });
            par.Runs.Add(new OcrRun { Text = "unc", FontSizePt = 12, Uri = "\\\\evil-host\\share" });
            page.Paragraphs.Add(par);
            OoxmlPackage package = PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp);
            string rels = PptxPartText(package, "ppt/slides/_rels/slide1.xml.rels");
            int hyperlinks = 0, at = 0;
            while ((at = rels.IndexOf("/hyperlink\"", at, StringComparison.Ordinal) + 1) > 0)
                hyperlinks++;
            AssertEq(2, hyperlinks, "регистр пути различает ссылки: " + rels);
            AssertTrue(rels.IndexOf("javascript", StringComparison.OrdinalIgnoreCase) < 0, "javascript: не попал в связи");
            AssertTrue(rels.IndexOf("evil-host", StringComparison.OrdinalIgnoreCase) < 0, "UNC не попал в связи");
            XDocument slide = PptxSlideXml(package, 1);
            AssertEq(2, new List<XElement>(slide.Descendants(PptxA + "hlinkClick")).Count, "кликабельны только безопасные");
            AssertEq("A a js unc", PptxAllText(slide), "текст опасных ссылок сохранён");
        }

        [Test("Аудит E: pptx — набранные подчёркивания слайда сохраняются, прочерки потока — нет")]
        private static void AuditEPptxTypedUnderscores()
        {
            Func<PageLayoutMode, string> textOf = delegate(PageLayoutMode mode)
            {
                PdfPageText page = PptxPage(600, 800);
                page.LayoutMode = mode;
                var par = new OcrParagraph { LeftPt = 40, TopPt = 700, RightPt = 300, BottomPt = 690 };
                par.Runs.Add(new OcrRun { Text = "Ф.И.О. ", FontSizePt = 10 });
                par.Runs.Add(new OcrRun { Text = "__________", FontSizePt = 10 });
                page.Paragraphs.Add(par);
                return PptxAllText(PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp), 1));
            };
            AssertEq("Ф.И.О. __________", textOf(PageLayoutMode.Slide), "разбор для слайда прочерков не синтезирует — это поле бланка");
            AssertEq("Ф.И.О. ", textOf(PageLayoutMode.Document), "синтезированный прочерк потока — шум");
        }

        [Test("Аудит E: кегль прижимается к пределам, а не сбрасывается в 12 pt")]
        private static void AuditEFontSizeSaturates()
        {
            AssertEq(72.0, FontResolver.ClampSizePt(80), "документ: 80 pt — к пределу 72, а не 12");
            AssertEq(5.0, FontResolver.ClampSizePt(4.5), "4,5 pt — к нижнему пределу 5");
            AssertEq(10.5, FontResolver.ClampSizePt(10.5), "обычный кегль не тронут");
            AssertEq(FontResolver.DefaultFontSize, FontResolver.ClampSizePt(double.NaN), "NaN — умолчание");
            AssertEq(FontResolver.DefaultFontSize, FontResolver.ClampSizePt(0), "ноль — умолчание");
            AssertEq(FontResolver.DefaultFontSize, FontResolver.ClampSizePt(-3), "отрицательный — умолчание");
            AssertEq(72.0, FontResolver.ClampSizePt(double.PositiveInfinity), "бесконечность — к пределу");
            AssertEq(80.0, FontResolver.ClampSizePt(80, FontResolver.MaxSlideFontSize), "слайд: 80 pt как есть");
            AssertTrue(FontResolver.MaxFontSize <= 1638, "предел документа в пределах Word");

            PdfPageText page = PptxPage(960, 540);
            page.Paragraphs.Add(PptxPar("Заголовок", 40, 500, 900, 400, 80));
            XElement rPr = PptxFirst(PptxSlideXml(PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp), 1)
                .Descendants(PptxA + "rPr"));
            AssertEq("8000", (string)rPr.Attribute("sz"), "заголовок 80 pt на слайде — 80 pt");
        }

        // ---------- 7, 8. таблицы ----------

        [Test("Аудит E: линия в миллиард пунктов не вешает поиск таблиц")]
        private static void AuditEHugeLineTableDetector()
        {
            var lines = new List<PdfLine>
            {
                HLine(100, 0, 100), HLine(50, 0, 100), HLine(0, 0, 100),
                VLine(0, 0, 100), VLine(50, 0, 100), VLine(100, 0, 100),
                HLine(300, 0, 1e9), HLine(-5e8, -1e9, 1e9), VLine(1e9, -1e9, 1e9)
            };
            var words = new List<PdfWord> { Word("A", 25, 75), Word("B", 75, 75), Word("C", 25, 25), Word("D", 75, 25) };
            Stopwatch sw = Stopwatch.StartNew();
            TableDetectResult res = TableDetector.Detect(lines, words, 600, 800);
            sw.Stop();
            AssertTrue(sw.ElapsedMilliseconds < 3000, "поиск закончился быстро: " + sw.ElapsedMilliseconds + " мс");
            AssertEq(1, res.Tables.Count, "настоящая таблица найдена");
            AssertEq(0, res.RemainingWords.Count, "слова в ячейках");
        }

        [Test("Аудит E: линовка обрезается по окрестности листа при сборе")]
        private static void AuditEClipLinesToPage()
        {
            var lines = new List<PdfLine>
            {
                HLine(100, -1e9, 1e9),             // торчит в обе стороны — укорачивается
                HLine(1e9, 0, 100),                // целиком далеко над листом — прочь
                VLine(50, double.NaN, 100),        // NaN — прочь
                VLine(50, 10, 200),                // обычная — как есть
                HLine(-900, 0, 100)                // ниже листа больше, чем на его высоту, — прочь
            };
            PdfTextExtract.ClipLinesToPage(lines, 600, 800);
            AssertEq(2, lines.Count, "остались две линии");
            AssertTrue(lines[0].MinX == -600 && lines[0].MaxX == 1200, "длинная прижата к окрестности: "
                + lines[0].MinX + ".." + lines[0].MaxX);
            AssertTrue(lines[1].MinY == 10 && lines[1].MaxY == 200, "обычная не тронута");
        }

        [Test("Аудит E: слова вложенной рамочной таблицы не задваиваются во внешней")]
        private static void AuditENestedTableWordsOnce()
        {
            var lines = new List<PdfLine>
            {
                // Внешняя 2×2: X 0..200, Y 0..200.
                HLine(200, 0, 200), HLine(100, 0, 200), HLine(0, 0, 200),
                VLine(0, 0, 200), VLine(100, 0, 200), VLine(200, 0, 200),
                // Вложенная 2×2 в ячейке (0,0), не касается её границ.
                HLine(190, 10, 90), HLine(150, 10, 90), HLine(110, 10, 90),
                VLine(10, 110, 190), VLine(50, 110, 190), VLine(90, 110, 190)
            };
            var words = new List<PdfWord>
            {
                Word("a", 30, 170), Word("b", 70, 170), Word("c", 30, 130), Word("d", 70, 130),
                Word("X", 150, 150), Word("Y", 50, 50), Word("Z", 150, 50)
            };
            TableDetectResult res = TableDetector.Detect(lines, words, 600, 800);
            AssertEq(2, res.Tables.Count, "две таблицы");
            var seen = new Dictionary<string, int>();
            foreach (OcrTable t in res.Tables)
                foreach (OcrTableRow row in t.Rows)
                    foreach (OcrTableCell cell in row.Cells)
                        foreach (OcrParagraph p in cell.Paragraphs)
                            foreach (string token in p.Text.Split(' '))
                                seen[token] = (seen.ContainsKey(token) ? seen[token] : 0) + 1;
            foreach (string w in new[] { "a", "b", "c", "d", "X", "Y", "Z" })
                AssertEq(1, seen.ContainsKey(w) ? seen[w] : 0, "слово «" + w + "» ровно в одной ячейке");
            OcrTable inner = res.Tables[0].RightPt < res.Tables[1].RightPt ? res.Tables[0] : res.Tables[1];
            AssertEq("a", TableCellText(inner, 0, 0), "буквы — во вложенной таблице");
        }

        [Test("Аудит E: объединения ячеек не перекрываются")]
        private static void AuditENoOverlappingSpans()
        {
            // Правая колонка без внутренней горизонтали — ячейка (0,1) на две строки; в нижней
            // строке нет вертикали между колонками — ячейка (1,0) тянулась бы на (1,1), уже занятую.
            var lines = new List<PdfLine>
            {
                HLine(100, 0, 100), HLine(0, 0, 100), HLine(50, 0, 50),
                VLine(0, 0, 100), VLine(100, 0, 100), VLine(50, 50, 100)
            };
            TableDetectResult res = TableDetector.Detect(lines, new List<PdfWord>(), 600, 800);
            AssertEq(1, res.Tables.Count, "таблица найдена");
            OcrTable t = res.Tables[0];
            var owners = new int[t.Rows.Count, t.ColumnCount];
            for (int r = 0; r < t.Rows.Count; r++)
                for (int c = 0; c < t.ColumnCount; c++)
                {
                    OcrTableCell cell = t.Rows[r].Cells[c];
                    if (cell.Covered)
                        continue;
                    for (int dr = 0; dr < Math.Max(1, cell.RowSpan); dr++)
                        for (int dc = 0; dc < Math.Max(1, cell.ColSpan); dc++)
                            owners[r + dr, c + dc]++;
                }
            for (int r = 0; r < t.Rows.Count; r++)
                for (int c = 0; c < t.ColumnCount; c++)
                    AssertEq(1, owners[r, c], "позицию (" + r + "," + c + ") накрывает ровно один владелец");
            AssertEq(2, t.Rows[0].Cells[1].RowSpan, "правая ячейка — на две строки");
            AssertEq(1, t.Rows[1].Cells[0].ColSpan, "нижняя левая не залезает под чужой пролёт");
        }

        // ---------- 9. порядок чтения текста ----------

        [Test("Аудит E: текст двух колонок идёт колонками, а не вперемешку")]
        private static void AuditEPlainTextColumns()
        {
            PdfPageText page = PptxPage(600, 800);
            page.Paragraphs.Add(PptxPar("L1", 50, 700, 250, 650));
            page.Paragraphs.Add(PptxPar("R1", 300, 700, 500, 650));
            page.Paragraphs.Add(PptxPar("L2", 50, 640, 250, 590));
            page.Paragraphs.Add(PptxPar("R2", 300, 640, 500, 590));
            AssertEq("L1\n\nL2\n\nR1\n\nR2", PlainText.Page(page), "левая колонка целиком раньше правой");

            PdfPageText single = PptxPage(600, 800);
            single.Paragraphs.Add(PptxPar("second", 50, 600, 500, 580));
            single.Paragraphs.Add(PptxPar("first", 50, 700, 500, 680));
            single.Tables.Add(PptxTable(1, 2, 50, 650, 500, 620));
            single.Tables[0].Rows[0].Cells[0].Paragraphs.Add(PptxPar("cell", 60, 645, 100, 630));
            AssertEq("first\n\ncell\n\nsecond", PlainText.Page(single), "одна колонка — сверху вниз, таблица на своём месте");
        }

        // ---------- 10. повороты дублей ----------

        [Test("Аудит E: подложка дубля повёрнута так же, как его текст")]
        private static void AuditESlotRotations()
        {
            var order = new List<PdfPageRef>
            {
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = 0 },
                new PdfPageRef { SourcePath = "A.pdf", PageIndex = 0, Rotation = 90 },
                new PdfPageRef { SourcePath = "a.pdf", PageIndex = 1, Rotation = 90 },
                new PdfPageRef { SourcePath = "B.pdf", PageIndex = 0, Rotation = 180 },
                new PdfPageRef { SourcePath = "B.pdf", PageIndex = 0 }
            };
            AssertEq("0,0,90,180,180", string.Join(",", Array.ConvertAll(PageBackgrounds.SlotRotations(order),
                delegate(int v) { return v.ToString(CultureInfo.InvariantCulture); })),
                "первый экземпляр решает и для текста, и для подложки");
            AssertEq(0, PageBackgrounds.SlotRotations(null).Length, "пустой порядок");
            AssertEq(0, PdfPageExtraction.EffectiveRotation(null, order[1]), "без карты поворотов — 0");
        }

        // ---------- 11. буквица ----------

        [Test("Аудит E: высокая буквица не сливает строки абзаца")]
        private static void AuditEDropCapLines()
        {
            var words = new List<PdfWord>
            {
                W("Lorem", 40, 100, 30, 10), W("ipsum", 75, 100, 30, 10),
                W("dolor", 40, 88, 30, 10), W("sit", 75, 88, 20, 10),
                W("amet", 40, 76, 30, 10), W("elit", 75, 76, 20, 10),
                W("B", 10, 76, 25, 34) // буквица на три строки
            };
            List<OcrLayout.Line> lines = OcrLayout.ToLines(words);
            AssertEq(3, lines.Count, "три строки, а не одна");
            foreach (OcrLayout.Line line in lines)
            {
                var texts = new List<string>();
                foreach (PdfWord w in line.Words)
                    texts.Add(w.Text);
                string joined = string.Join(" ", texts.ToArray());
                AssertTrue(!(joined.Contains("Lorem") && joined.Contains("dolor"))
                    && !(joined.Contains("dolor") && joined.Contains("amet")), "строки не перемешаны: " + joined);
            }
            string text = string.Join(" ", OcrLayout.ToParagraphs(words).ToArray());
            AssertTrue(text.IndexOf("Lorem ipsum", StringComparison.Ordinal) >= 0
                && text.IndexOf("dolor sit", StringComparison.Ordinal) > text.IndexOf("ipsum", StringComparison.Ordinal)
                && text.IndexOf("amet elit", StringComparison.Ordinal) > text.IndexOf("sit", StringComparison.Ordinal),
                "порядок чтения строк сохранён: " + text);

            // Обычная строка с надстрочным знаком по-прежнему одна строка.
            var sup = new List<PdfWord> { W("E", 0, 100, 10, 10), W("2", 11, 106, 5, 6), W("mc", 20, 100, 20, 7) };
            AssertEq(1, OcrLayout.ToLines(sup).Count, "надстрочный знак — в своей строке");
        }

        // ---------- 12. тире в начале строки ----------

        [Test("Аудит E: перенос фразы на строку с тире — не пункт списка")]
        private static void AuditEDashContinuation()
        {
            var words = new List<PdfWord>
            {
                W("Он", 0, 100, 20, 10), W("долго", 25, 100, 50, 10), W("молчал", 80, 100, 60, 10),
                W("и", 145, 100, 10, 10), W("наконец", 160, 100, 70, 10), W("сказал", 235, 100, 65, 10),
                W("—", 0, 88, 10, 10), W("и", 15, 88, 10, 10), W("вышел", 30, 88, 50, 10)
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEq(1, layout.Paragraphs.Count, "строка дошла до поля и не кончилась точкой — это перенос");
            AssertEq(ListKind.None, layout.Paragraphs[0].ListKind, "и не список");
            AssertEq("Он долго молчал и наконец сказал — и вышел", layout.Paragraphs[0].Text, "текст склеен");
        }

        [Test("Аудит E: настоящий список на тире по-прежнему список")]
        private static void AuditEDashListKept()
        {
            var words = new List<PdfWord>
            {
                W("Требования:", 0, 100, 90, 10),
                W("—", 0, 88, 10, 10), W("первый", 15, 88, 60, 10), W("пункт;", 80, 88, 50, 10),
                W("—", 0, 76, 10, 10), W("второй", 15, 76, 60, 10), W("пункт.", 80, 76, 50, 10)
            };
            OcrLayout.OcrPageLayout layout = OcrLayout.Analyze(words);
            AssertEq(3, layout.Paragraphs.Count, "вводная строка и два пункта");
            AssertEq(ListKind.Bulleted, layout.Paragraphs[1].ListKind, "первый пункт");
            AssertEq(ListKind.Bulleted, layout.Paragraphs[2].ListKind, "второй пункт");

            // Длинный пункт до правого поля, но кончается «;» — следующий тире-пункт отдельно.
            var longItems = new List<PdfWord>
            {
                W("—", 0, 100, 10, 10), W("первый", 15, 100, 60, 10), W("очень", 80, 100, 50, 10),
                W("длинный", 135, 100, 70, 10), W("пункт", 210, 100, 50, 10), W("списка;", 265, 100, 35, 10),
                W("—", 0, 88, 10, 10), W("второй", 15, 88, 60, 10), W("пункт.", 80, 88, 50, 10)
            };
            OcrLayout.OcrPageLayout list = OcrLayout.Analyze(longItems);
            AssertEq(2, list.Paragraphs.Count, "«;» в конце полной строки — граница пункта");
            AssertEq(ListKind.Bulleted, list.Paragraphs[1].ListKind, "второй пункт — пункт");
        }
    }
}
