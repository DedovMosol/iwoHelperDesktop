using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Этап 1 плана распознавания: исправления движка. Каждое правило проверено на настоящих
    /// PDF (PdfSharp) через тот же разбор, что у конвертеров, и каждое — мутантом.
    /// </summary>
    internal static partial class UnitTests
    {
        /// <summary>
        /// Страница A4 (595 × 842 pt) — Я. ист PdfSharp по умолчанию берётся из региона системы:
        /// в en-US это Letter 612 × 792 (проба 30.10.2026), а раннер CI работает в en-US. ез явного
        /// размера геометрия тестов (поля, CropBox, «за краем листа») зависела бы от машины.
        /// </summary>
        private static PdfSharp.Pdf.PdfPage A4(PdfSharp.Pdf.PdfDocument pdf)
        {
            PdfSharp.Pdf.PdfPage page = pdf.AddPage();
            page.Width = 595;
            page.Height = 842;
            return page;
        }

        [Test("Аудит: тесты распознавания не зависят от региона системы (en-US, как на раннере CI)")]
        private static void AuditRecognitionTestsRegionIndependent()
        {
            System.Globalization.CultureInfo saved = System.Threading.Thread.CurrentThread.CurrentCulture;
            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                System.Globalization.CultureInfo.CurrentCulture.ClearCachedData();
                using (var pdf = new PdfSharp.Pdf.PdfDocument())
                {
                    PdfSharp.Pdf.PdfPage plain = pdf.AddPage(); // нарочно без размера — воспроизвести ловушку
                    // Ловушка воспроизводится: без явного размера лист — Letter.
                    AssertTrue(Math.Abs(plain.Height.Point - 792) < 1, "en-US: AddPage() даёт Letter: " + plain.Height.Point);
                    PdfSharp.Pdf.PdfPage a4 = A4(pdf);
                    AssertTrue(Math.Abs(a4.Width.Point - 595) < 1 && Math.Abs(a4.Height.Point - 842) < 1, "A4() — A4 в любом регионе");
                }
                // Геометрические тесты распознавания под en-US: RC-19 до правки здесь падал.
                AuditRc19OffPage();
                AuditRc15RightAlignment();
                AuditRc1WhiteOnFills();
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = saved;
                System.Globalization.CultureInfo.CurrentCulture.ClearCachedData();
            }
            // Страницы тестов — только через A4(): AddPage() в этом файле разрешён ровно в двух
            // местах — в самой A4() (размер задаётся сразу) и выше, где ловушка воспроизводится
            // нарочно. Третий вызов — новый тест, рисующий лист региона машины.
            string root = RepoRoot();
            AssertTrue(root != null, "корень репозитория найден");
            string source = File.ReadAllText(Path.Combine(root, "tests", "AuditTests.Recognition.cs"));
            int calls = System.Text.RegularExpressions.Regex.Matches(source, @"\.AddPage\(\)").Count;
            AssertEq(2, calls, "AddPage() только в A4() и в воспроизведении ловушки");
        }

        // ---------- RC-1: белый текст на светлой цветной заливке ----------

        [Test("Аудит RC-1: контраст с белым по WCAG — эталонные значения")]
        private static void AuditRc1Contrast()
        {
            Func<int, double> of = rgb => PdfTextExtract.ContrastWithWhite(((rgb >> 16) & 0xFF) / 255.0, ((rgb >> 8) & 0xFF) / 255.0, (rgb & 0xFF) / 255.0);
            AssertTrue(Math.Abs(of(0x000000) - 21) < 0.01, "чёрный — 21:1");
            AssertTrue(Math.Abs(of(0xFFFFFF) - 1) < 0.01, "белый — 1:1");
            // Справочные значения WCAG-калькуляторов (WebAIM): серый #767676 — 4,54:1.
            AssertTrue(Math.Abs(of(0x767676) - 4.54) < 0.01, "#767676 — 4,54:1: " + of(0x767676));
            foreach (int accent in new[] { 0x5B9BD5, 0xED7D31, 0xA5A5A5, 0x70AD47, 0x4472C4 })
                AssertTrue(of(accent) >= PdfTextExtract.BackdropMinContrast, "акцент Word #" + accent.ToString("X6") + " — подложка: " + of(accent));
            foreach (int pale in new[] { 0xFFFF00, 0xD9D9D9, 0xF2F2F2 })
                AssertTrue(of(pale) < PdfTextExtract.BackdropMinContrast, "#" + pale.ToString("X6") + " — белое не видно: " + of(pale));
        }

        [Test("Аудит RC-1: белый текст виден на акцентной, обведённой и скруглённой плашке, не виден на жёлтой")]
        private static void AuditRc1WhiteOnFills()
        {
            EmbeddedAssemblies.Ensure();
            using (var dir = new TempDir("iwo_rc1_"))
            {
                string path = dir.File("fills.pdf");
                var font = new PdfSharp.Drawing.XFont("Arial", 12, PdfSharp.Drawing.XFontStyle.Bold,
                    new PdfSharp.Drawing.XPdfFontOptions(PdfSharp.Pdf.PdfFontEncoding.Unicode));
                var cases = new[]
                {
                    new { Word = "Синяя", Fill = 0x5B9BD5, Kind = "plain" },
                    new { Word = "Оранжевая", Fill = 0xED7D31, Kind = "stroked" },
                    new { Word = "Серая", Fill = 0xA5A5A5, Kind = "rounded" },
                    new { Word = "Жёлтая", Fill = 0xFFFF00, Kind = "plain" }
                };
                using (var pdf = new PdfSharp.Pdf.PdfDocument())
                {
                    PdfSharp.Pdf.PdfPage page = A4(pdf);
                    using (PdfSharp.Drawing.XGraphics g = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
                    {
                        g.DrawString("Обычный текст страницы", font, PdfSharp.Drawing.XBrushes.Black, 72, 80);
                        double y = 120;
                        foreach (var c in cases)
                        {
                            var brush = new PdfSharp.Drawing.XSolidBrush(PdfSharp.Drawing.XColor.FromArgb((c.Fill >> 16) & 0xFF, (c.Fill >> 8) & 0xFF, c.Fill & 0xFF));
                            if (c.Kind == "stroked")
                                g.DrawRectangle(PdfSharp.Drawing.XPens.Black, brush, 66, y - 16, 300, 24);
                            else if (c.Kind == "rounded")
                                g.DrawRoundedRectangle(brush, 66, y - 16, 300, 24, 10, 10);
                            else
                                g.DrawRectangle(brush, 66, y - 16, 300, 24);
                            g.DrawString(c.Word + " плашка", font, PdfSharp.Drawing.XBrushes.White, 72, y);
                            y += 40;
                        }
                    }
                    pdf.Save(path);
                }
                foreach (PageLayoutMode mode in new[] { PageLayoutMode.Document, PageLayoutMode.Slide })
                {
                    string text = PdfTextExtract.Extract(path, null, null, null, mode)[0].Text;
                    foreach (var c in cases)
                        AssertTrue(text.Contains(c.Word) == (c.Fill != 0xFFFF00),
                            mode + ": «" + c.Word + "» " + (c.Fill != 0xFFFF00 ? "виден" : "не виден (белое на жёлтом)") + ": " + text);
                }
            }
        }

        // ---------- RC-15: выключка вправо ----------

        [Test("Аудит RC-15: дата и подпись справа — вправо; красная строка, выключка и глубокий пункт — нет")]
        private static void AuditRc15RightAlignment()
        {
            EmbeddedAssemblies.Ensure();
            using (var dir = new TempDir("iwo_rc15_"))
            {
                string path = dir.File("right.pdf");
                var font = new PdfSharp.Drawing.XFont("Arial", 11, PdfSharp.Drawing.XFontStyle.Regular,
                    new PdfSharp.Drawing.XPdfFontOptions(PdfSharp.Pdf.PdfFontEncoding.Unicode));
                const double left = 72, right = 523;
                using (var pdf = new PdfSharp.Pdf.PdfDocument())
                {
                    PdfSharp.Pdf.PdfPage page = A4(pdf);
                    using (PdfSharp.Drawing.XGraphics g = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
                    {
                        Action<string, double> atRight = (s, y) => g.DrawString(s, font, PdfSharp.Drawing.XBrushes.Black, right - g.MeasureString(s, font).Width, y);
                        atRight("Москва, 30 октября 2026 г.", 90);
                        // Полная строка с красной строкой 36 pt, дотянутая до поля выключкой.
                        string[] words = "Настоящим сообщаем что работы по договору выполнены полностью и в срок".Split(' ');
                        double natural = words.Sum(w => g.MeasureString(w, font).Width);
                        double gap = (right - left - 36 - natural) / (words.Length - 1), x = left + 36;
                        foreach (string w in words) { g.DrawString(w, font, PdfSharp.Drawing.XBrushes.Black, x, 130); x += g.MeasureString(w, font).Width + gap; }
                        // Вторая строка абзаца — у левого поля: она задаёт поле колонки. Без неё
                        // все строки начаты с отступа, и поле колонки уехало бы к красной строке.
                        g.DrawString("без замечаний.", font, PdfSharp.Drawing.XBrushes.Black, left, 145);
                        atRight("Директор И. И. Иванов", 200);
                        // Пункт глубокого уровня, дотянутый до поля: список, а не текст вправо.
                        atRight("1. Пункт на глубоком уровне вложенности", 240);
                    }
                    pdf.Save(path);
                }
                List<OcrParagraph> paras = PdfTextExtract.Extract(path)[0].Paragraphs;
                string all = string.Join(" | ", paras.Select(p => p.Alignment + ":" + p.Text).ToArray());
                Func<string, OcrParagraph> find = s => paras.First(p => p.Text.Contains(s));
                AssertEq(OcrAlignment.Right, find("Москва").Alignment, "дата справа: " + all);
                AssertEq(OcrAlignment.Right, find("Директор").Alignment, "подпись справа: " + all);
                AssertTrue(find("Настоящим").Alignment != OcrAlignment.Right, "красная строка — не вправо: " + all);
                OcrParagraph item = find("Пункт");
                AssertTrue(item.ListKind == ListKind.Numbered && item.Alignment == OcrAlignment.Left, "глубокий пункт — пункт влево: " + all);

                // PowerPoint: в ячейке и в рамке без левых краёв строк — «r».
                OoxmlPackage package = PptxWriter.BuildPackage(PdfTextExtract.Extract(path, null, null, null, PageLayoutMode.Slide), PptxStamp);
                string pptx = dir.File("right.pptx");
                package.Save(pptx);
                AssertTrue(PptxValidationProblems(pptx).Count == 0, "пакет валиден");
            }
        }

        [Test("Аудит RC-15: строки вправо — правило по каждой строке, отступ Word справа — до края текста")]
        private static void AuditRc15Rules()
        {
            Func<double, double, OcrLayout.Line> line = (l, r) =>
            {
                var ln = new OcrLayout.Line { MidY = 100 };
                ln.Words.Add(new PdfWord { Text = "x", Left = l, Right = r, Bottom = 95, Top = 105 });
                return ln;
            };
            Func<OcrLayout.Line[], bool> right = g => OcrLayout.IsRightAligned(g.ToList(), 0, 400, 10, 400);
            AssertTrue(right(new[] { line(250, 400) }), "строка справа");
            AssertTrue(right(new[] { line(250, 396) }), "до поля в пределах 0,5 кегля");
            AssertTrue(!right(new[] { line(250, 394) }), "не дотянулась до поля");
            AssertTrue(!right(new[] { line(100, 400) }), "начата в зоне отступов (≤ 25 % ширины)");
            AssertTrue(right(new[] { line(250, 400), line(300, 400) }), "две строки вправо");
            AssertTrue(!right(new[] { line(250, 400), line(0, 400) }), "вторая строка у левого поля — выключка, не вправо");
            AssertTrue(!OcrLayout.IsRightAligned(new List<OcrLayout.Line>(), 0, 400, 10, 400), "пусто");

            AssertEq(0.0, WordDocxWriter.RightEdgeIndent(523, 523), "у поля — без отступа");
            AssertEq(0.0, WordDocxWriter.RightEdgeIndent(520, 523), "3 pt — шум");
            AssertEq(100.0, WordDocxWriter.RightEdgeIndent(423, 523), "край текста левее поля — отступ справа");
        }

        // ---------- RC-19: невидимое за краем страницы ----------

        [Test("Аудит RC-19: текст и метки обреза за CropBox не попадают в вывод")]
        private static void AuditRc19OffPage()
        {
            EmbeddedAssemblies.Ensure();
            using (var dir = new TempDir("iwo_rc19_"))
            {
                string raw = dir.File("raw.pdf"), path = dir.File("bleed.pdf");
                var font = new PdfSharp.Drawing.XFont("Arial", 11);
                using (var pdf = new PdfSharp.Pdf.PdfDocument())
                {
                    // A4 явно: на Letter (en-US, раннер CI) CropBox 36…806 обрезался бы MediaBox 792,
                    // и надпись «на поле» оказалась бы в видимой части — так тест и упал на CI.
                    PdfSharp.Pdf.PdfPage page = A4(pdf); // видимая часть — 36…559 × 36…806
                    using (PdfSharp.Drawing.XGraphics g = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
                    {
                        g.DrawString("Видимый текст страницы", font, PdfSharp.Drawing.XBrushes.Black, 72, 100);
                        g.DrawString("Служебная надпись на поле под обрез", font, PdfSharp.Drawing.XBrushes.Black, 72, 20);
                        g.DrawString("Справа", font, PdfSharp.Drawing.XBrushes.Black, 565, 400);
                        var pen = new PdfSharp.Drawing.XPen(PdfSharp.Drawing.XColors.Black, 0.5);
                        // Метки обреза в каждом углу: по горизонтали и по вертикали, за CropBox.
                        foreach (double x in new[] { 36.0, 559.0 })
                            foreach (double y in new[] { 36.0, 806.0 })
                            {
                                double hx = x < 300 ? 4 : 565, vy = y < 400 ? 4 : 812;
                                g.DrawLine(pen, hx, y, hx + 26, y);
                                g.DrawLine(pen, x, vy, x, vy + 26);
                            }
                    }
                    pdf.Save(raw);
                }
                using (PdfSharp.Pdf.PdfDocument pdf = PdfSharp.Pdf.IO.PdfReader.Open(raw, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify))
                {
                    pdf.Pages[0].CropBox = new PdfSharp.Pdf.PdfRectangle(new PdfSharp.Drawing.XPoint(36, 36), new PdfSharp.Drawing.XPoint(559, 806));
                    pdf.Save(path);
                }
                foreach (PageLayoutMode mode in new[] { PageLayoutMode.Document, PageLayoutMode.Slide, PageLayoutMode.Review })
                {
                    PdfPageText page = PdfTextExtract.Extract(path, null, null, null, mode)[0];
                    string words = string.Join(" ", page.Words.Select(w => w.Text).ToArray());
                    AssertTrue(words.Contains("Видимый"), mode + ": видимый текст на месте: " + words);
                    AssertTrue(!words.Contains("Служебная") && !words.Contains("Справа"), mode + ": текст за CropBox отсечён: " + words);
                    // Прочерки из одиночных линий (AddRuleWords) попадают в абзацы, а не в Words.
                    AssertTrue(!page.Text.Contains("_"), mode + ": метки обреза — не прочерки: " + page.Text);
                    AssertEq(0, page.Lines.Count, mode + ": линий на видимой странице нет: " + page.Lines.Count);
                }
            }
        }

        private static PdfLine Rc19Line(bool horizontal, double pos, double from, double to)
        {
            return horizontal
                ? new PdfLine { Orientation = LineOrientation.Horizontal, X1 = from, X2 = to, Y1 = pos, Y2 = pos }
                : new PdfLine { Orientation = LineOrientation.Vertical, X1 = pos, X2 = pos, Y1 = from, Y2 = to };
        }

        [Test("Аудит RC-19: границы видимости слова и линии")]
        private static void AuditRc19Rules()
        {
            Func<double, double, bool> word = (x, y) => PdfTextExtract.CenterOnPage(new PdfTextExtract.RectPt { Left = x - 5, Right = x + 5, Bottom = y - 3, Top = y + 3 }, 600, 800);
            AssertTrue(word(300, 400) && word(0, 400) && word(600, 800), "внутри и на кромке");
            AssertTrue(!word(-1, 400) && !word(601, 400) && !word(300, -1) && !word(300, 801), "центр за краем");
            var lines = new List<PdfLine>
            {
                Rc19Line(true, 0, 0, 600),      // кромка листа — видна
                Rc19Line(true, 100, -50, 10),   // заходит на лист — видна
                Rc19Line(true, -10, 0, 600),    // под листом
                Rc19Line(true, 100, -40, -5),   // левее листа
                Rc19Line(false, 610, 0, 800),   // правее листа
                Rc19Line(false, 300, 805, 830), // выше листа
                Rc19Line(false, 600.3, 0, 800)  // в допуске кромки — видна
            };
            PdfTextExtract.DropLinesOffPage(lines, 600, 800);
            AssertEq("0,100,600.3", string.Join(",", lines.Select(l => l.Position.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToArray()),
                "остались только видимые линии");
        }

        // ---------- RC-22: синтезированный курсив ----------

        [Test("Аудит RC-22: наклон рамки буквы не зависит от поворота и масштаба")]
        private static void AuditRc22SlantGeometry()
        {
            AssertTrue(Math.Abs(PdfTextExtract.SlantOf(0, 0, 10, 0, 0, 8)) < 1e-9, "прямая буква — 0");
            AssertTrue(Math.Abs(PdfTextExtract.SlantOf(0, 0, 10, 0, 2.8, 8) - 0.35) < 1e-9, "наклон 0,35");
            // Та же буква, повёрнутая на 90° против часовой: строка идёт вверх, «верх» — влево.
            AssertTrue(Math.Abs(PdfTextExtract.SlantOf(0, 0, 0, 10, -8, 2.8) - 0.35) < 1e-9, "поворот не меняет наклон");
            AssertTrue(Math.Abs(PdfTextExtract.SlantOf(0, 0, 30, 0, 8.4, 24) - 0.35) < 1e-9, "масштаб не меняет наклон");
            AssertTrue(PdfTextExtract.SlantOf(0, 0, 10, 0, -2.8, 8) < 0, "обратный наклон — отрицательный");
            AssertEq(0.0, PdfTextExtract.SlantOf(5, 5, 5, 5, 5, 9), "вырожденная рамка — 0");
            AssertEq(0.0, PdfTextExtract.SlantOf(0, 0, 10, 0, 3, 0), "нулевая высота — 0");
            var slanted = new double[] { 0, 0, 10, 0, 2.8, 8 };
            AssertTrue(Math.Abs(PdfTextExtract.WordSlant(new List<double[]> { new double[] { 4, 4, 4, 4, 4, 4 }, slanted }) - 0.35) < 1e-9,
                "вырожденная первая буква не прячет наклон слова");
            AssertTrue(Math.Abs(PdfTextExtract.WordSlant(new List<double[]> { new double[] { 0, 0, 0.005, 0, 2, 8 }, slanted }) - 0.35) < 1e-9,
                "буква нулевой ширины тоже пропускается");
            AssertEq(0.0, PdfTextExtract.WordSlant(new List<double[]>()), "без букв — 0");
        }

        [Test("Аудит RC-22: синтезированный курсив Tahoma и Verdana — курсив с настоящим кеглем; прямой и настоящий курсив не тронуты")]
        private static void AuditRc22SyntheticItalic()
        {
            EmbeddedAssemblies.Ensure();
            using (var dir = new TempDir("iwo_rc22_"))
            {
                string path = dir.File("italic.pdf");
                var opt = new PdfSharp.Drawing.XPdfFontOptions(PdfSharp.Pdf.PdfFontEncoding.Unicode);
                var cases = new[]
                {
                    new { Word = "Прямой", Font = "Tahoma", Style = PdfSharp.Drawing.XFontStyle.Regular, Size = 10.5, Italic = false, Bold = false },
                    new { Word = "Наклонный", Font = "Tahoma", Style = PdfSharp.Drawing.XFontStyle.Italic, Size = 10.5, Italic = true, Bold = false },
                    new { Word = "Широкий", Font = "Verdana", Style = PdfSharp.Drawing.XFontStyle.Italic, Size = 14.0, Italic = true, Bold = false },
                    new { Word = "Жирнокурсив", Font = "Tahoma", Style = PdfSharp.Drawing.XFontStyle.BoldItalic, Size = 9.0, Italic = true, Bold = true },
                    new { Word = "Настоящий", Font = "Arial", Style = PdfSharp.Drawing.XFontStyle.Italic, Size = 12.0, Italic = true, Bold = false }
                };
                using (var pdf = new PdfSharp.Pdf.PdfDocument())
                {
                    PdfSharp.Pdf.PdfPage page = A4(pdf);
                    using (PdfSharp.Drawing.XGraphics g = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
                    {
                        double y = 100;
                        foreach (var c in cases)
                        {
                            g.DrawString(c.Word, new PdfSharp.Drawing.XFont(c.Font, c.Size, c.Style, opt), PdfSharp.Drawing.XBrushes.Black, 72, y);
                            y += 40;
                        }
                    }
                    pdf.Save(path);
                }
                List<PdfWord> words = PdfTextExtract.Extract(path)[0].Words;
                foreach (var c in cases)
                {
                    PdfWord w = words.First(x => x.Text == c.Word);
                    AssertEq(c.Italic, w.Italic, c.Word + ": курсив");
                    AssertEq(c.Bold, w.Bold, c.Word + ": полужирный");
                    AssertTrue(Math.Abs(w.FontSizePt - c.Size) < 0.05, c.Word + ": кегль " + w.FontSizePt + " вместо " + c.Size);
                }
            }
        }

        // ---------- RC-4: значки списков Word из Symbol, Wingdings и Courier New ----------

        [Test("Аудит RC-4: значок Symbol/Wingdings — знак Юникода, прочие символы шрифтов не трогаются")]
        private static void AuditRc4SymbolGlyphs()
        {
            AssertEq("\u2022", WordBullets.FromSymbolFont("\uF0B7", "Symbol"), "Symbol 0xB7 — •");
            AssertEq("\u25AA", WordBullets.FromSymbolFont("\uF0A7", "Wingdings"), "Wingdings 0xA7 — ▪");
            AssertEq("\u2756", WordBullets.FromSymbolFont("\uF076", "Wingdings"), "Wingdings 0x76 — ❖");
            AssertEq("\u2B9A", WordBullets.FromSymbolFont("\uF0D8", "Wingdings"), "Wingdings 0xD8 — ⮚");
            AssertEq("\u2714", WordBullets.FromSymbolFont("\uF0FC", "Wingdings"), "Wingdings 0xFC — ✔");
            AssertEq("\uF0A7", WordBullets.FromSymbolFont("\uF0A7", "Wingdings 2"), "Wingdings 2 — другая раскладка");
            AssertEq("\uF0B7", WordBullets.FromSymbolFont("\uF0B7", "Segoe UI Symbol"), "не символьный шрифт Word");
            AssertEq("\uF0E0", WordBullets.FromSymbolFont("\uF0E0", "Wingdings"), "значка нет в таблице — не угадываем");
            AssertEq("\uF0B7\uF0B7", WordBullets.FromSymbolFont("\uF0B7\uF0B7", "Symbol"), "только одиночный значок");
            AssertEq("\uF0B7", WordBullets.FromSymbolFont("\uF0B7", null), "шрифт неизвестен");
            foreach (char g in WordBullets.Glyphs)
                AssertEq(ListKind.Bulleted, ListMarker.Detect(g + " пункт").Kind, "«" + g + "» — маркер списка");
        }

        private static PdfWord Rc4Word(string text, string font, double left, double width, double baseline = 700)
        {
            return new PdfWord { Text = text, FontName = font, FontSizePt = 12, Left = left, Right = left + width, Bottom = baseline - 3, Top = baseline + 9 };
        }

        [Test("Аудит RC-4: «o» Courier New в начале пункта — значок второго уровня, в коде и во фразе — буква")]
        private static void AuditRc4CourierBullet()
        {
            var bullet = new List<PdfWord> { Rc4Word("o", "Courier New", 90, 7), Rc4Word("Пункт", "Arial", 108, 40) };
            AssertEq(1, WordBullets.MarkCourierBullets(bullet), "значок пункта");
            AssertEq("\u25E6", bullet[0].Text, "стал «◦»");
            AssertTrue(bullet[0].FontName == null, "в шрифте по умолчанию");
            var code = new List<PdfWord> { Rc4Word("o", "Courier New", 90, 7), Rc4Word("=", "Courier New", 104, 7) };
            AssertEq(0, WordBullets.MarkCourierBullets(code), "листинг кода: сосед того же Courier");
            var inside = new List<PdfWord> { Rc4Word("слово", "Courier New", 40, 40), Rc4Word("o", "Courier New", 87, 7), Rc4Word("Пункт", "Arial", 108, 40) };
            AssertEq(0, WordBullets.MarkCourierBullets(inside), "не в начале строки");
            var arial = new List<PdfWord> { Rc4Word("o", "Arial", 90, 7), Rc4Word("Пункт", "Times New Roman", 108, 40) };
            AssertEq(0, WordBullets.MarkCourierBullets(arial), "«o» не Courier — буква");
            var far = new List<PdfWord> { Rc4Word("o", "Courier New", 90, 7), Rc4Word("Пункт", "Arial", 160, 40) };
            AssertEq(0, WordBullets.MarkCourierBullets(far), "текст дальше четырёх кеглей — не тот же пункт");
            var other = new List<PdfWord> { Rc4Word("o", "Courier New", 90, 7), Rc4Word("Пункт", "Arial", 108, 40, 680) };
            AssertEq(0, WordBullets.MarkCourierBullets(other), "сосед на другой строке");
        }

        /// <summary>PDF: пункты со значками Word — Symbol, Courier New «o», Wingdings — и текстом Arial.</summary>
        private static string Rc4Pdf(string dir)
        {
            string path = Path.Combine(dir, "bullets.pdf");
            var unicode = new PdfSharp.Drawing.XPdfFontOptions(PdfSharp.Pdf.PdfFontEncoding.Unicode);
            var text = new PdfSharp.Drawing.XFont("Arial", 12, PdfSharp.Drawing.XFontStyle.Regular, unicode);
            using (var pdf = new PdfSharp.Pdf.PdfDocument())
            {
                PdfSharp.Pdf.PdfPage page = A4(pdf);
                using (PdfSharp.Drawing.XGraphics g = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
                {
                    var items = new[] { new { M = "\uF0B7", F = "Symbol", X = 72.0 }, new { M = "o", F = "Courier New", X = 96.0 }, new { M = "\uF0A7", F = "Wingdings", X = 120.0 } };
                    // Вводная строка во всю ширину задаёт правое поле: без неё три коротких
                    // пункта со ступенчатым отступом выглядят центрированными (это RC-3).
                    g.DrawString("Вводный абзац перед списком задаёт ширину колонки текста документа целиком",
                        text, PdfSharp.Drawing.XBrushes.Black, 72, 80);
                    double y = 100;
                    foreach (var it in items)
                    {
                        var mf = new PdfSharp.Drawing.XFont(it.F, 12, PdfSharp.Drawing.XFontStyle.Regular, unicode);
                        g.DrawString(it.M, mf, PdfSharp.Drawing.XBrushes.Black, it.X, y);
                        g.DrawString("Пункт списка уровня", text, PdfSharp.Drawing.XBrushes.Black, it.X + 18, y);
                        y += 20;
                    }
                }
                pdf.Save(path);
            }
            return path;
        }

        [Test("Аудит RC-4: пункты со значками Word распознаны списком; «Сравнение» читает слой буквально")]
        private static void AuditRc4ListsRecognized()
        {
            EmbeddedAssemblies.Ensure();
            using (var dir = new TempDir("iwo_rc4_"))
            {
                string path = Rc4Pdf(dir.Path);
                List<OcrParagraph> paras = PdfTextExtract.Extract(path)[0].Paragraphs;
                string all = string.Join(" | ", paras.Select(p => p.ListKind + p.ListLevel.ToString() + ":" + p.Text).ToArray());
                AssertEq(4, paras.Count, "вводный абзац и три пункта: " + all);
                AssertEq(ListKind.None, paras[0].ListKind, "вводный абзац — не пункт");
                paras.RemoveAt(0);
                AssertTrue(paras.All(p => p.ListKind == ListKind.Bulleted), "все — маркированный список: " + all);
                AssertEq("0,1,2", string.Join(",", paras.Select(p => p.ListLevel.ToString()).ToArray()), "уровни по отступу: " + all);
                AssertTrue(paras[0].Text.StartsWith("\u2022") && paras[1].Text.StartsWith("\u25E6") && paras[2].Text.StartsWith("\u25AA"),
                    "маркеры — знаки Юникода: " + all);
                List<PdfWord> literal = PdfTextExtract.Extract(path, null, null, null, PageLayoutMode.Review)[0].Words;
                AssertTrue(literal.Any(w => w.Text == "o") && !literal.Any(w => w.Text == "\u2022" || w.Text == "\u25E6"),
                    "Сравнение: текстовый слой без подмен: " + string.Join(" ", literal.Select(w => w.Text).ToArray()));
            }
        }

        // ---------- RC-11: одна сбойная страница не роняет конвертацию ----------

        /// <summary>
        /// PDF из трёх страниц с текстом «Page N»; у страниц из broken ресурсы заменены числом.
        /// PdfPig на такой странице бросает исключение при разборе содержимого (проба: «Cannot
        /// get item from empty stack»), остальные страницы читаются. Вторая страница альбомная —
        /// по ней видно, что размер сбойной страницы берётся настоящий, а не A4.
        /// </summary>
        private static string Rc11Pdf(string dir, string name, params int[] broken)
        {
            return Rc11Pdf(dir, name, false, broken);
        }

        /// <summary>blankFirst — первая страница без текста (только рамка), как картинка-скан.</summary>
        private static string Rc11Pdf(string dir, string name, bool blankFirst, params int[] broken)
        {
            string path = Path.Combine(dir, name);
            using (var pdf = new PdfSharp.Pdf.PdfDocument())
            {
                var font = new PdfSharp.Drawing.XFont("Arial", 14);
                for (int i = 0; i < 3; i++)
                {
                    PdfSharp.Pdf.PdfPage page = A4(pdf);
                    if (i == 1) { page.Width = 842; page.Height = 595; }
                    using (PdfSharp.Drawing.XGraphics g = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
                    {
                        if (i == 0 && blankFirst)
                            g.DrawRectangle(PdfSharp.Drawing.XPens.Black, 72, 72, 200, 100);
                        else
                            g.DrawString("Page " + (i + 1), font, PdfSharp.Drawing.XBrushes.Black, 72, 100);
                    }
                }
                foreach (int b in broken)
                    pdf.Pages[b].Elements.SetValue("/Resources", new PdfSharp.Pdf.PdfInteger(5));
                pdf.Save(path);
            }
            return path;
        }

        [Test("Аудит RC-11: сбойная страница переносится картинкой во весь лист, остальные разобраны")]
        private static void AuditRc11BrokenPageDegrades()
        {
            if (!Ghostscript.Available)
                Skip("Ghostscript не найден");
            EmbeddedAssemblies.Ensure();
            using (var dir = new TempDir("iwo_rc11_"))
            {
                string path = Rc11Pdf(dir.Path, "broken.pdf", 1);
                foreach (PageLayoutMode mode in new[] { PageLayoutMode.Document, PageLayoutMode.Slide })
                {
                    List<PdfPageText> pages = PdfTextExtract.Extract(path, null, null, null, mode);
                    AssertEq(3, pages.Count, mode + ": все три страницы на месте");
                    AssertTrue(pages[0].ExtractionError == null && pages[0].Text.Contains("Page 1"), mode + ": страница 1 разобрана");
                    AssertTrue(pages[2].ExtractionError == null && pages[2].Text.Contains("Page 3"), mode + ": страница 3 разобрана");
                    PdfPageText bad = pages[1];
                    AssertTrue(bad.ExtractionError != null, mode + ": страница 2 помечена как сбойная");
                    AssertEq(1, bad.PageIndex, mode + ": номер сбойной страницы исходный");
                    AssertTrue(Math.Abs(bad.WidthPt - 842) < 1 && Math.Abs(bad.HeightPt - 595) < 1,
                        mode + ": размер настоящий (альбом), а не A4: " + bad.WidthPt + "×" + bad.HeightPt);
                    AssertEq(1, bad.Images.Count, mode + ": страница перенесена одной картинкой");
                    OcrImage img = bad.Images[0];
                    AssertTrue(img.Png != null && img.Png.Length > 0 && img.LeftPt == 0
                        && Math.Abs(img.TopPt - 595) < 1 && Math.Abs(img.WidthPt - 842) < 1 && Math.Abs(img.HeightPt - 595) < 1,
                        mode + ": картинка во весь лист");
                    AssertEq(0, bad.Paragraphs.Count, mode + ": текста у сбойной страницы нет");
                }
                // Писатель слайдов принимает такую страницу: пакет валиден, на слайде 2 — картинка.
                List<PdfPageText> slides = PdfTextExtract.Extract(path, null, null, null, PageLayoutMode.Slide);
                OoxmlPackage package = PptxWriter.BuildPackage(slides, PptxStamp);
                string pptx = dir.File("broken.pptx");
                package.Save(pptx);
                List<string> problems = PptxValidationProblems(pptx);
                AssertTrue(problems.Count == 0, "пакет валиден: " + string.Join("; ", problems.ToArray()));
            }
        }

        [Test("Аудит RC-11: размер сбойной страницы — видимый (CropBox), а не MediaBox")]
        private static void AuditRc11CropBoxSize()
        {
            if (!Ghostscript.Available)
                Skip("Ghostscript не найден");
            EmbeddedAssemblies.Ensure();
            using (var dir = new TempDir("iwo_rc11_crop_"))
            {
                string path = Rc11Pdf(dir.Path, "crop.pdf", 1);
                // Обрезать ВТОРУЮ (альбомную, сбойную) страницу до 500×300: так PdfPig меряет
                // видимую часть исправной страницы, и Ghostscript (-dUseCropBox) рисует её же.
                string cropped = dir.File("cropped.pdf");
                using (PdfSharp.Pdf.PdfDocument pdf = PdfSharp.Pdf.IO.PdfReader.Open(path, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify))
                {
                    pdf.Pages[1].CropBox = new PdfSharp.Pdf.PdfRectangle(new PdfSharp.Drawing.XPoint(100, 100), new PdfSharp.Drawing.XPoint(600, 400));
                    pdf.Save(cropped);
                }
                PdfPageText bad = PdfTextExtract.Extract(cropped)[1];
                AssertTrue(bad.ExtractionError != null, "сбойная");
                AssertTrue(Math.Abs(bad.WidthPt - 500) < 1 && Math.Abs(bad.HeightPt - 300) < 1,
                    "размер — CropBox 500×300: " + bad.WidthPt + "×" + bad.HeightPt);
                OcrImage img = bad.Images[0];
                AssertTrue(Math.Abs(img.WidthPt - bad.WidthPt) < 0.01 && Math.Abs(img.HeightPt - bad.HeightPt) < 0.01,
                    "картинка — ровно лист");
            }
        }

        [Test("Аудит RC-11: без Ghostscript сбойная страница пустая, и итог говорит это, а не «картинкой»")]
        private static void AuditRc11NoGhostscript()
        {
            EmbeddedAssemblies.Ensure();
            Ghostscript.ResetResolutionForTests();
            Ghostscript.ResolveForTests = delegate { return null; };
            try
            {
                using (var dir = new TempDir("iwo_rc11_nogs_"))
                {
                    string path = Rc11Pdf(dir.Path, "broken.pdf", 1);
                    List<PdfPageText> pages = PdfTextExtract.Extract(path);
                    AssertEq(3, pages.Count, "документ всё равно разобран");
                    AssertTrue(pages[1].ExtractionError != null && pages[1].Images.Count == 0, "страница пустая");
                    AssertTrue(Math.Abs(pages[1].WidthPt - 842) < 1, "размер — подсказка PdfSharp: " + pages[1].WidthPt);
                    AssertEq(1, PdfPageExtraction.FailedBlankCount(pages), "пустых — одна");
                }
            }
            finally
            {
                Ghostscript.ResetResolutionForTests();
            }
            Lang before = Loc.Current;
            try
            {
                foreach (Lang lang in new[] { Lang.Ru, Lang.En })
                {
                    Loc.Init(lang);
                    var blank = new ConvertResult { Pages = 3, PagesWithText = 2, FailedPages = new List<int> { 2 }, FailedBlank = 1 };
                    string status = PdfConvertFormBase.DoneStatus(blank, "Done {0}.");
                    AssertTrue(status.Contains("Ghostscript") && !status.Contains(string.Format(Loc.T("convert.status.failedPages"), 1, "2")),
                        lang + ": пустая страница не названа «картинкой»: " + status);
                }
            }
            finally
            {
                Loc.Init(before);
            }
        }

        [Test("Аудит RC-11: сбойная страница — картинка на слайде и в .docx (живой Word)")]
        private static void AuditRc11WritersCarryPicture()
        {
            if (!Ghostscript.Available)
                Skip("Ghostscript не найден");
            EmbeddedAssemblies.Ensure();
            using (var dir = new TempDir("iwo_rc11_out_"))
            {
                string path = Rc11Pdf(dir.Path, "broken.pdf", 1);
                OoxmlPackage package = PptxWriter.BuildPackage(PdfTextExtract.Extract(path, null, null, null, PageLayoutMode.Slide), PptxStamp);
                string slide;
                using (var ms = new MemoryStream(package.ToArray()))
                using (var zip = new System.IO.Compression.ZipArchive(ms))
                using (var reader = new StreamReader(zip.GetEntry("ppt/slides/slide2.xml").Open()))
                    slide = reader.ReadToEnd();
                AssertTrue(slide.Contains("<p:pic>"), "на слайде 2 — картинка страницы");

                AuditBRequireWord();
                string docx = dir.File("broken.docx");
                List<PdfPageText> pages = PdfTextExtract.Extract(path);
                AuditBOfficeSta(180000, delegate { WordDocxWriter.Write(pages, docx); });
                System.Xml.XmlDocument xml = AuditBDocumentXml(docx);
                var ns = new System.Xml.XmlNamespaceManager(xml.NameTable);
                ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                string text = xml.DocumentElement.InnerText;
                AssertTrue(text.Contains("Page 1") && text.Contains("Page 3"), "исправные страницы — текстом: " + text);
                AssertEq(1, xml.SelectNodes("//w:drawing", ns).Count, "одна картинка — сбойная страница");
                AssertEq(3, xml.SelectNodes("//w:sectPr", ns).Count, "три раздела — три страницы");
            }
        }

        [Test("Аудит RC-11: поворот пользователя применяется к сбойной странице")]
        private static void AuditRc11RotationApplies()
        {
            if (!Ghostscript.Available)
                Skip("Ghostscript не найден");
            EmbeddedAssemblies.Ensure();
            using (var dir = new TempDir("iwo_rc11_rot_"))
            {
                string path = Rc11Pdf(dir.Path, "broken.pdf", 1);
                List<PdfPageText> pages = PdfTextExtract.Extract(path, null, new[] { 0, 90, 0 }, null, PageLayoutMode.Document);
                PdfPageText bad = pages[1];
                AssertTrue(bad.ExtractionError != null, "сбойная");
                AssertTrue(Math.Abs(bad.WidthPt - 595) < 1 && Math.Abs(bad.HeightPt - 842) < 1,
                    "альбом, повёрнутый на 90°, стал книжным: " + bad.WidthPt + "×" + bad.HeightPt);
                OcrImage img = bad.Images[0];
                AssertTrue(Math.Abs(img.WidthPt - 595) < 1 && Math.Abs(img.HeightPt - 842) < 1, "картинка повёрнута вместе с листом");
                using (var ms = new MemoryStream(img.Png))
                using (var bmp = new System.Drawing.Bitmap(ms))
                    AssertTrue(bmp.Height > bmp.Width, "пиксели повёрнуты: " + bmp.Width + "×" + bmp.Height);
            }
        }

        [Test("Аудит RC-11: целиком битый файл и «Сравнение» по-прежнему дают понятную ошибку")]
        private static void AuditRc11StrictWhereNeeded()
        {
            EmbeddedAssemblies.Ensure();
            using (var dir = new TempDir("iwo_rc11_strict_"))
            {
                string all = Rc11Pdf(dir.Path, "all.pdf", 0, 1, 2);
                AssertThrowsExactly<MergeException>("все страницы сбойные — это битый файл", () => PdfTextExtract.Extract(all));
                string one = Rc11Pdf(dir.Path, "one.pdf", 1);
                AssertThrowsExactly<MergeException>("Сравнение: пустая страница читалась бы как «весь текст удалён»",
                    () => PdfTextExtract.Extract(one, null, null, null, PageLayoutMode.Review));
            }
        }

        [Test("Аудит RC-11: страница без текста и сбойная — не «скан», конвертация идёт")]
        private static void AuditRc11FailedIsNotScan()
        {
            EmbeddedAssemblies.Ensure();
            using (var dir = new TempDir("iwo_rc11_scan_"))
            {
                string path = Rc11Pdf(dir.Path, "mixed.pdf", true, 1);
                var order = new List<PdfPageRef>
                {
                    new PdfPageRef { SourcePath = path, PageIndex = 0 },
                    new PdfPageRef { SourcePath = path, PageIndex = 1 }
                };
                int withText;
                List<PdfPageText> pages = PdfPageExtraction.Load(order, null, null, out withText);
                AssertEq(0, withText, "текста на выбранных страницах нет");
                AssertEq("2", string.Join(",", PdfPageExtraction.FailedPositions(pages)),
                    "вторая — сбойная, и она не повод сказать «это скан»");
            }
        }

        [Test("Аудит RC-11: итог конвертации называет сбойные страницы и не считает их ещё раз «без текста»")]
        private static void AuditRc11DoneStatus()
        {
            Lang before = Loc.Current;
            try
            {
                foreach (Lang lang in new[] { Lang.Ru, Lang.En })
                {
                    Loc.Init(lang);
                    const string fmt = "Done {0}.";
                    var result = new ConvertResult { Pages = 5, PagesWithText = 3, FailedPages = new List<int> { 2, 4 } };
                    string status = PdfConvertFormBase.DoneStatus(result, fmt);
                    AssertTrue(status.Contains("(2, 4)"), lang + ": названы номера страниц: " + status);
                    AssertTrue(!status.Contains(string.Format(Loc.T("convert.status.noTextPages"), 2)),
                        lang + ": сбойные страницы не повторены как «без текста»: " + status);
                    var many = new ConvertResult { Pages = 30, PagesWithText = 18, FailedPages = Enumerable.Range(1, 12).ToList() };
                    string longStatus = PdfConvertFormBase.DoneStatus(many, fmt);
                    AssertTrue(longStatus.Contains("10, …") && !longStatus.Contains("11"), lang + ": длинный список обрезан: " + longStatus);
                }
            }
            finally
            {
                Loc.Init(before);
            }
            var pages = new List<PdfPageText> { new PdfPageText(), new PdfPageText { ExtractionError = "x" }, new PdfPageText() };
            AssertEq("2", string.Join(",", PdfPageExtraction.FailedPositions(pages)), "номера — с 1, в порядке результата");
        }
    }
}
