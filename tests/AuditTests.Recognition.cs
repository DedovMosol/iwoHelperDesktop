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
                PdfSharp.Pdf.PdfPage page = pdf.AddPage();
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
                    PdfSharp.Pdf.PdfPage page = pdf.AddPage();
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
