using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ExcelMerger;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Регрессии этапа 0 аудита 1.18.5 (audit.md). Каждый тест назван по ID находки и
    /// воспроизводит её на production-коде: до правки он падал, после — проходит.
    /// </summary>
    internal static partial class UnitTests
    {
        // ---------- RS-1: поворот страницы в «Сравнении» ----------

        /// <summary>
        /// Матрица текста, при которой слово «Hello» на экране читается горизонтально при данном
        /// /Rotate (MediaBox 400×600). Для 90° — так же, как в живом тесте извлечения.
        /// </summary>
        private static string Stage0RotatedText(int rotate, string word)
        {
            string tm;
            switch (rotate)
            {
                case 90: tm = "0 1 -1 0 100 50"; break;
                case 180: tm = "-1 0 0 -1 300 100"; break;
                case 270: tm = "0 -1 1 0 300 550"; break;
                default: tm = "1 0 0 1 100 500"; break;
            }
            return "BT /F1 20 Tf " + tm + " Tm (" + word + ") Tj ET";
        }

        /// <summary>Рамка тёмных пикселей растра (пиксели, Y вниз); Empty — чернил нет.</summary>
        private static Rectangle Stage0InkBounds(Bitmap bitmap)
        {
            var rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
            BitmapData data = bitmap.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                int minX = int.MaxValue, minY = int.MaxValue, maxX = -1, maxY = -1;
                var row = new byte[bitmap.Width * 4];
                for (int y = 0; y < bitmap.Height; y++)
                {
                    Marshal.Copy(IntPtr.Add(data.Scan0, y * data.Stride), row, 0, row.Length);
                    for (int x = 0; x < bitmap.Width; x++)
                    {
                        int sum = row[x * 4] + row[x * 4 + 1] + row[x * 4 + 2];
                        if (sum >= 300)
                            continue;
                        if (x < minX) minX = x;
                        if (x > maxX) maxX = x;
                        if (y < minY) minY = y;
                        if (y > maxY) maxY = y;
                    }
                }
                return maxX < 0 ? Rectangle.Empty : Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
        }

        [Test("Аудит RS-1: рамка слова «Сравнения» лежит на чернилах рендера при /Rotate 0/90/180/270 и CropBox")]
        private static void AuditRs1ReviewBoxesMatchRenderedInk()
        {
            var failures = new List<string>();
            using (var dir = new TempDir("iwo_stage0_rs1_"))
            using (var renderer = new PdfThumbnailRenderer())
            {
                var cases = new List<KeyValuePair<string, string[]>>();
                foreach (int rotate in new[] { 0, 90, 180, 270 })
                    cases.Add(new KeyValuePair<string, string[]>("rot" + rotate, new[]
                    {
                        Stage0RotatedText(rotate, "Hello"), rotate == 0 ? "" : "/Rotate " + rotate
                    }));
                // CropBox со сдвинутым началом: координаты отсчитываются от его угла.
                cases.Add(new KeyValuePair<string, string[]>("crop", new[]
                {
                    "BT /F1 20 Tf 1 0 0 1 120 300 Tm (Hello) Tj ET", "/CropBox [50 200 350 450]"
                }));
                cases.Add(new KeyValuePair<string, string[]>("crop-rot90", new[]
                {
                    "BT /F1 20 Tf 0 1 -1 0 200 260 Tm (Hello) Tj ET", "/CropBox [50 200 350 450] /Rotate 90"
                }));
                foreach (KeyValuePair<string, string[]> c in cases)
                {
                    string path = dir.File(c.Key + ".pdf");
                    WriteRawPage(path, c.Value[0], "", c.Value[1]);
                    PdfReviewDocument doc = PdfReviewService.Load(path, null, null, PdfReviewLimits.Default());
                    PdfReviewPage page = doc.Pages[0];
                    if (page.Words.Count != 1)
                    {
                        failures.Add(c.Key + ": слов " + page.Words.Count);
                        continue;
                    }
                    using (BudgetedBitmap rendered = renderer.RenderOwned(path, 0, 400, 0, RasterBudget.PreviewPixels))
                    {
                        if (rendered == null)
                            Skip("системный рендерер PDF недоступен");
                        Bitmap bitmap = rendered.Bitmap;
                        double aspectView = page.ViewHeightPt / page.ViewWidthPt;
                        double aspectBitmap = (double)bitmap.Height / bitmap.Width;
                        if (Math.Abs(aspectView - aspectBitmap) > 0.02)
                            failures.Add(c.Key + ": пропорции вида " + page.ViewWidthPt + "×" + page.ViewHeightPt +
                                " не совпадают с растром " + bitmap.Width + "×" + bitmap.Height);
                        RectangleF box = PdfReviewGeometry.ToPixelRect(page.Words[0].Box,
                            page.ViewWidthPt, page.ViewHeightPt, bitmap.Width, bitmap.Height);
                        Rectangle ink = Stage0InkBounds(bitmap);
                        if (ink.IsEmpty)
                        {
                            failures.Add(c.Key + ": на растре нет чернил");
                            continue;
                        }
                        RectangleF grown = box;
                        grown.Inflate(4, 4);
                        var inkCenter = new PointF(ink.Left + ink.Width / 2f, ink.Top + ink.Height / 2f);
                        if (!grown.Contains(inkCenter) || !ink.IntersectsWith(Rectangle.Round(box)))
                            failures.Add(c.Key + ": рамка " + box + " мимо чернил " + ink);
                    }
                }
            }
            AssertTrue(failures.Count == 0, "рамки «Сравнения» на повёрнутых страницах: " +
                string.Join(" | ", failures.ToArray()));
        }

        [Test("Аудит RS-1: замена слова на странице /Rotate 90 остаётся правкой после растрового уточнения")]
        private static void AuditRs1RotatedReplacementSurvives()
        {
            using (var dir = new TempDir("iwo_stage0_rs1cmp_"))
            {
                string left = dir.File("left.pdf");
                string right = dir.File("right.pdf");
                const string common = "BT /F1 14 Tf 0 1 -1 0 60 50 Tm (Common words stay here) Tj ET ";
                WriteRawPage(left, common + Stage0RotatedText(90, "Alpha"), "", "/Rotate 90");
                WriteRawPage(right, common + Stage0RotatedText(90, "Omega"), "", "/Rotate 90");
                PdfReviewResult result = PdfReviewService.Compare(left, right);
                AssertEq(1, result.Stats.DeletedWords, "удалённое слово на повёрнутой странице");
                AssertEq(1, result.Stats.InsertedWords, "добавленное слово на повёрнутой странице");
                foreach (PdfReviewWordOp op in result.Operations)
                    AssertTrue(op.MatchKind != PdfReviewMatchKind.RasterEquivalent,
                        "разные слова не объявлены растрово равными");
            }
        }

        // ---------- RS-2: отсутствие чернил — не доказательство равенства ----------

        [Test("Аудит RS-2: пустые растры с обеих сторон не снимают замену слова")]
        private static void AuditRs2NoInkIsNotEquivalence()
        {
            PdfReviewPage leftPage = ReviewVisualPage(0, ReviewVisualWord("Alpha", 20, 30, 80, 70));
            PdfReviewPage rightPage = ReviewVisualPage(0, ReviewVisualWord("Omega", 20, 30, 80, 70));
            using (Bitmap leftBitmap = ReviewVisualBitmap(null))
            using (Bitmap rightBitmap = ReviewVisualBitmap(null))
            {
                bool refined;
                PdfReviewResult result = RefineVisualResult(leftPage, rightPage, leftBitmap, rightBitmap, out refined);
                AssertTrue(!refined, "без чернил уточнение воздерживается");
                AssertEq(1, result.Stats.DeletedWords, "Delete сохранён");
                AssertEq(1, result.Stats.InsertedWords, "Insert сохранён");
            }
            // Контроль: одинаковые чернила в рамке слова — равенство доказано и правка снята.
            using (Bitmap leftBitmap = ReviewVisualBitmap(g => g.FillRectangle(Brushes.Black, 20, 30, 60, 40)))
            using (Bitmap rightBitmap = ReviewVisualBitmap(g => g.FillRectangle(Brushes.Black, 20, 30, 60, 40)))
            {
                bool refined;
                PdfReviewResult result = RefineVisualResult(leftPage, rightPage, leftBitmap, rightBitmap, out refined);
                AssertTrue(refined, "одинаковые чернила — доказательство");
                AssertEq(0, result.Stats.DeletedWords, "правка снята по доказательству");
            }
        }

        // ---------- GR-2: нехватка памяти в потоке миниатюр ----------

        [Test("Аудит GR-2: нехватка памяти при рендере миниатюры не роняет процесс, рендер продолжается")]
        private static void AuditGr2ThumbWorkerSurvivesOutOfMemory()
        {
            int calls = 0, delivered = 0;
            RunSta(delegate
            {
                using (var host = new Form())
                {
                    var grid = new PdfPageGrid { Dock = DockStyle.Fill };
                    grid.RenderThumbForTests = delegate(string path, int index, int width)
                    {
                        int call = Interlocked.Increment(ref calls);
                        if (call <= 3)
                            throw new OutOfMemoryException("проба: нехватка памяти");
                        Interlocked.Increment(ref delivered);
                        return AuditDBitmap(40, 60);
                    };
                    host.Controls.Add(grid);
                    host.Show();
                    grid.SetPages(AuditDFakePages("oom.pdf", 6));
                    WaitFor(delegate { return Volatile.Read(ref delivered) > 0; }, 15000);
                    Pump(host, 200);
                }
            });
            AssertTrue(calls > 3, "после отказов рендер вызывался снова: вызовов " + calls);
            AssertTrue(delivered > 0, "после нехватки памяти миниатюры снова отрисовываются");
        }

        // ---------- PS-1: закладки с любыми видами цели ----------

        /// <summary>
        /// Двухстраничный PDF с оглавлением. outlineItems — словари закладок (объекты 7, 8, …),
        /// связанные /First /Next /Parent самим вызывающим.
        /// </summary>
        private static void Stage0WriteOutlinedPdf(string path, string catalogExtra, params string[] outlineItems)
        {
            var objects = new List<string>
            {
                "<< /Type /Catalog /Pages 2 0 R /Outlines 6 0 R " + catalogExtra + " >>",
                "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << >> /Contents 5 0 R >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << >> /Contents 5 0 R >>",
                RawStream("", "0 0 m 10 10 l S"),
                "<< /Type /Outlines /First 7 0 R /Count " + outlineItems.Length + " >>"
            };
            objects.AddRange(outlineItems);
            WriteRawPdf(path, objects);
        }

        private static string Stage0BookmarkSig(IList<PdfBookmark> marks)
        {
            var parts = new List<string>();
            foreach (PdfBookmark m in marks)
                parts.Add(m.Title + "@" + m.PageIndex + "/L" + m.Level + (m.Opened ? "+" : ""));
            return string.Join(", ", parts.ToArray());
        }

        private static List<PdfBookmark> Stage0ReadBookmarks(string path)
        {
            using (PdfDocument doc = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                return PdfBookmarks.Read(doc);
        }

        [Test("Аудит PS-1: закладки читаются при любом виде цели — именованной, числовой, через /Dests")]
        private static void AuditPs1BookmarksAllDestinationKinds()
        {
            using (var dir = new TempDir("iwo_stage0_ps1_"))
            {
                string path = dir.File("latex.pdf");
                Stage0WriteOutlinedPdf(path,
                    "/Names << /Dests 13 0 R >> /Dests 14 0 R",
                    // 7: hyperref — /GoTo с именованным назначением-строкой (дерево имён)
                    "<< /Title (Intro) /Parent 6 0 R /Next 8 0 R /A << /S /GoTo /D (sec.1) >> /First 12 0 R /Last 12 0 R /Count 1 >>",
                    // 8: /Dest-строка
                    "<< /Title (Strings) /Parent 6 0 R /Prev 7 0 R /Next 9 0 R /Dest (sec.2) >>",
                    // 9: /Dest-имя через словарь каталога /Dests
                    "<< /Title (Names) /Parent 6 0 R /Prev 8 0 R /Next 10 0 R /Dest /legacy >>",
                    // 10: номер страницы числом
                    "<< /Title (Number) /Parent 6 0 R /Prev 9 0 R /Next 11 0 R /Dest [1 /Fit] >>",
                    // 11: папка без цели и переход в другой файл — пропускаются
                    "<< /Title (Remote) /Parent 6 0 R /Prev 10 0 R /A << /S /GoToR /F (other.pdf) /D [0 /Fit] >> >>",
                    // 12: ребёнок первой закладки — обычный массив
                    "<< /Title (Child) /Parent 7 0 R /Dest [4 0 R /XYZ 0 0 0] >>",
                    // 13: дерево имён с промежуточным узлом и словарём назначения /D
                    "<< /Kids [15 0 R] >>",
                    // 14: словарь /Dests каталога (PDF 1.1)
                    "<< /legacy [4 0 R /Fit] >>",
                    // 15: лист дерева имён
                    "<< /Limits [(sec.1) (sec.2)] /Names [(sec.1) [3 0 R /Fit] (sec.2) << /D [4 0 R /Fit] >>] >>");
                List<PdfBookmark> marks = Stage0ReadBookmarks(path);
                AssertEqual("Intro@0/L0+, Child@1/L1, Strings@1/L0, Names@1/L0, Number@1/L0",
                    Stage0BookmarkSig(marks), "все виды цели разрешены, /GoToR пропущен");

                // Объединение с таким файлом раньше падало целиком на чтении оглавления.
                string merged = dir.File("merged.pdf");
                PdfMergeService.Merge(new List<PdfPageRef>
                {
                    new PdfPageRef { SourcePath = path, PageIndex = 0 },
                    new PdfPageRef { SourcePath = path, PageIndex = 1 }
                }, merged);
                AssertEqual("Intro@0/L0+, Child@1/L1, Strings@1/L0, Names@1/L0, Number@1/L0",
                    Stage0BookmarkSig(Stage0ReadBookmarks(merged)), "оглавление перенесено в результат");
            }
        }

        private static PdfDictionary Stage0Dict(PdfItem item)
        {
            PdfReference reference = item as PdfReference;
            return (reference != null ? reference.Value : item) as PdfDictionary;
        }

        private static void Stage0RawCounts(PdfDictionary parent, List<string> into)
        {
            for (PdfDictionary item = Stage0Dict(parent.Elements["/First"]); item != null;
                item = Stage0Dict(item.Elements["/Next"]))
            {
                PdfItem count = item.Elements["/Count"];
                into.Add(item.Elements.GetString("/Title") + "=" + (count == null ? "-" : count.ToString()));
                Stage0RawCounts(item, into);
            }
        }

        [Test("Аудит PS-1: /Count пишется по ISO 32000-1 — раскрытая глава с листьями остаётся раскрытой")]
        private static void AuditPs1BookmarkCountsFollowIso()
        {
            using (var dir = new TempDir("iwo_stage0_ps1c_"))
            {
                string path = dir.File("counts.pdf");
                using (var doc = new PdfDocument())
                {
                    for (int i = 0; i < 3; i++)
                        doc.AddPage();
                    PdfBookmarks.Write(doc, new List<PdfBookmark>
                    {
                        new PdfBookmark { Title = "A", PageIndex = 0, Level = 0, Opened = true },
                        new PdfBookmark { Title = "A1", PageIndex = 1, Level = 1 },
                        new PdfBookmark { Title = "A1a", PageIndex = 1, Level = 2 },
                        new PdfBookmark { Title = "B", PageIndex = 2, Level = 0 },
                        new PdfBookmark { Title = "B1", PageIndex = 2, Level = 1, Opened = true },
                        new PdfBookmark { Title = "B1a", PageIndex = 2, Level = 2 }
                    });
                    doc.Save(path);
                }
                AssertEqual("A@0/L0+, A1@1/L1, A1a@1/L2, B@2/L0, B1@2/L1+, B1a@2/L2",
                    Stage0BookmarkSig(Stage0ReadBookmarks(path)), "раскрытость пережила запись и чтение");
                using (PdfDocument doc = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                {
                    PdfDictionary root = Stage0Dict(doc.Internals.Catalog.Elements["/Outlines"]);
                    var counts = new List<string> { "root=" + root.Elements["/Count"] };
                    Stage0RawCounts(root, counts);
                    // Видно: A, A1 (A раскрыта), B. Свёрнутая A1 прячет 1, свёрнутая B — 2 (B1 и,
                    // раз B1 раскрыта, B1a). У листьев /Count нет.
                    AssertEqual("root=3, A=1, A1=-1, A1a=-, B=-2, B1=1, B1a=-",
                        string.Join(", ", counts.ToArray()), "/Count каждой закладки");
                }
            }
        }

        [Test("Аудит PS-1: тысячи вложенных «папок» без цели не переполняют стек")]
        private static void AuditPs1BookmarksDeepFoldersDoNotOverflow()
        {
            using (var dir = new TempDir("iwo_stage0_ps1d_"))
            {
                // Цепочка из 5000 вложенных закладок без цели (объекты 7…), у последней цель есть.
                // Уровень выпавших не растёт, поэтому ограничивать обязана глубина рекурсии.
                const int count = 5000;
                var items = new string[count];
                for (int i = 0; i < count; i++)
                {
                    int self = 7 + i;
                    string parent = (i == 0 ? 6 : self - 1) + " 0 R";
                    items[i] = i == count - 1
                        ? "<< /Title (Leaf) /Parent " + parent + " /Dest [3 0 R /Fit] >>"
                        : "<< /Title (Folder) /Parent " + parent + " /First " + (self + 1) + " 0 R /Last "
                            + (self + 1) + " 0 R /Count 1 >>";
                }
                string path = dir.File("deep.pdf");
                Stage0WriteOutlinedPdf(path, "", items);
                List<PdfBookmark> marks = null;
                Exception failure = null;
                // Малый стек делает переполнение до исправления неизбежным, а не делом случая.
                var worker = new Thread(delegate()
                {
                    try { marks = Stage0ReadBookmarks(path); }
                    catch (Exception ex) { failure = ex; }
                }, 256 * 1024);
                worker.IsBackground = true;
                worker.Start();
                AssertTrue(worker.Join(30000), "обход закончился");
                AssertTrue(failure == null, "чтение не бросило: " + failure);
                AssertEqual("", Stage0BookmarkSig(marks), "закладка глубже предела отброшена, остальное — папки без цели");
            }
        }

        [Test("Аудит PS-1: закладка части, дописанная к готовому оглавлению, не портит /Count корня")]
        private static void AuditPs1AppendTopLevelKeepsRootCount()
        {
            using (var dir = new TempDir("iwo_stage0_ps1e_"))
            {
                string path = dir.File("part.pdf");
                using (var doc = new PdfDocument())
                {
                    doc.AddPage();
                    doc.AddPage();
                    PdfBookmarks.Write(doc, new List<PdfBookmark>
                    {
                        new PdfBookmark { Title = "A", PageIndex = 0, Level = 0, Opened = true },
                        new PdfBookmark { Title = "A1", PageIndex = 1, Level = 1 }
                    });
                    doc.Save(path);
                }
                // Так разделение по закладкам подписывает часть, чья глава выпала из сетки.
                using (PdfDocument doc = PdfReader.Open(path, PdfDocumentOpenMode.Modify))
                {
                    PdfBookmarks.AppendTopLevel(doc, "Part", 1);
                    doc.Save(path);
                }
                AssertEqual("A@0/L0+, A1@1/L1, Part@1/L0", Stage0BookmarkSig(Stage0ReadBookmarks(path)),
                    "старое оглавление на месте, подпись части — в конце");
                using (PdfDocument doc = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                {
                    PdfDictionary root = Stage0Dict(doc.Internals.Catalog.Elements["/Outlines"]);
                    var counts = new List<string> { "root=" + root.Elements["/Count"] };
                    Stage0RawCounts(root, counts);
                    AssertEqual("root=3, A=1, A1=-, Part=-", string.Join(", ", counts.ToArray()),
                        "видно A, A1 и Part; PDFsharp не затёр /Count корня своим счётчиком");
                }
            }
        }

        [Test("Аудит PS-1: цикл /Next и испорченная цель не вешают и не роняют чтение закладок")]
        private static void AuditPs1BookmarksSurviveBrokenTrees()
        {
            using (var dir = new TempDir("iwo_stage0_ps1b_"))
            {
                string path = dir.File("broken.pdf");
                Stage0WriteOutlinedPdf(path, "",
                    "<< /Title (One) /Parent 6 0 R /Next 8 0 R /Dest [3 0 R /Fit] >>",
                    // цель — строка, которой нет нигде; /Next замкнут на первую закладку
                    "<< /Title (Missing) /Parent 6 0 R /Next 7 0 R /Dest (nowhere) /First 9 0 R >>",
                    // ребёнок выпавшей закладки поднимается на её уровень
                    "<< /Title (Orphan) /Parent 8 0 R /Dest [4 0 R /Fit] >>");
                List<PdfBookmark> marks = null;
                var worker = new Thread(delegate() { marks = Stage0ReadBookmarks(path); });
                worker.IsBackground = true;
                worker.Start();
                AssertTrue(worker.Join(10000), "обход закончился (цикл /Next не бесконечен)");
                AssertEqual("One@0/L0, Orphan@1/L0", Stage0BookmarkSig(marks),
                    "битая закладка выпала, её ребёнок поднят, цикл оборван");
            }
        }

        // ---------- RW-1: поле пути в высокой контрастности ----------

        [Test("Аудит RW-1: поля ввода «Сравнения» берут системные цвета в высокой контрастности")]
        private static void AuditRw1InputColorsFollowHighContrast()
        {
            using (var box = new TextBox())
            {
                PdfReviewForm.ApplyInputColors(box, true);
                AssertEq(SystemColors.Window, box.BackColor, "фон — системный Window");
                AssertEq(SystemColors.WindowText, box.ForeColor, "текст — системный WindowText");
                PdfReviewForm.ApplyInputColors(box, false);
                AssertEq(Color.White.ToArgb(), box.BackColor.ToArgb(), "обычная тема — белое поле");
                AssertEq(Theme.TextPrimary.ToArgb(), box.ForeColor.ToArgb(), "обычная тема — основной текст");
            }
            var failures = new List<string>();
            InIsolatedSettings("iwo_stage0_rw1_", delegate
            {
                using (var form = new PdfReviewForm(delegate { }))
                {
                    var boxes = new List<TextBox>();
                    Stage0Collect(form, boxes);
                    if (boxes.Count < 2)
                        failures.Add("полей ввода " + boxes.Count);
                    Color back = SystemInformation.HighContrast ? SystemColors.Window : Color.White;
                    Color fore = SystemInformation.HighContrast ? SystemColors.WindowText : Theme.TextPrimary;
                    foreach (TextBox b in boxes)
                        if (b.BackColor.ToArgb() != back.ToArgb() || b.ForeColor.ToArgb() != fore.ToArgb())
                            failures.Add((b.AccessibleName ?? b.Name) + ": " + b.BackColor + "/" + b.ForeColor);
                }
            });
            AssertTrue(failures.Count == 0, "цвета полей ввода окна: " + string.Join(" | ", failures.ToArray()));
        }

        private static void Stage0Collect<T>(Control parent, List<T> into) where T : Control
        {
            foreach (Control c in parent.Controls)
            {
                var hit = c as T;
                if (hit != null)
                    into.Add(hit);
                Stage0Collect(c, into);
            }
        }

        // ---------- PF-2: «Разделение» → «Прочие операции» с правленой сеткой ----------

        /// <summary>Ответить на ближайший вопрос MessageForm (Да/Нет), когда он откроется.</summary>
        private static System.Windows.Forms.Timer Stage0AnswerNextConfirm(bool yes, Action seen)
        {
            var timer = new System.Windows.Forms.Timer { Interval = 50 };
            timer.Tick += delegate
            {
                foreach (Form f in Application.OpenForms)
                    if (f is MessageForm && f.Visible)
                    {
                        timer.Stop();
                        if (seen != null) seen();
                        f.DialogResult = yes ? DialogResult.Yes : DialogResult.No;
                        return;
                    }
            };
            timer.Start();
            return timer;
        }

        [Test("Аудит PF-2: переход в «Прочие операции» с правленой сеткой сначала спрашивает")]
        private static void AuditPf2SplitAsksBeforeDroppingEdits()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_stage0_pf2_", delegate
            {
                using (var dir = new TempDir("iwo_stage0_pf2_"))
                {
                    string pdf = dir.File("документ.pdf");
                    MakeEmptyPagesPdf(pdf, 2);
                    var bridged = new List<string>();
                    using (var form = new PdfSplitForm(delegate { }, delegate(string p) { bridged.Add(p); }))
                    {
                        form.Show();
                        form.AcceptFiles(new[] { pdf });
                        if (!WaitFor(delegate { return SourcePathOf(form) == pdf && !form.IsBusy; }, 10000))
                        { failures.Add("документ не открылся"); return; }
                        Button toOps = ButtonWithText(form, Loc.T("split.btn.ops"));
                        if (toOps == null) { failures.Add("кнопки перехода нет"); return; }

                        // Нетронутая сетка — переход без вопросов.
                        bool asked = false;
                        using (Stage0AnswerNextConfirm(true, delegate { asked = true; }))
                        {
                            toOps.PerformClick();
                            Pump(form, 300);
                        }
                        if (asked) failures.Add("нетронутая сетка: вопрос не нужен");
                        if (bridged.Count != 1) failures.Add("нетронутая сетка: переходов " + bridged.Count);

                        // Правка сетки: страница удалена.
                        var order = (PdfPageOrder)AuditDField(form, "_order");
                        order.RemoveAt(new[] { 1 });
                        asked = false;
                        using (Stage0AnswerNextConfirm(false, delegate { asked = true; }))
                            toOps.PerformClick();
                        if (!asked) failures.Add("правленая сетка: вопроса не было");
                        if (bridged.Count != 1) failures.Add("ответ «Нет» всё равно открыл «Прочие операции»");

                        asked = false;
                        using (Stage0AnswerNextConfirm(true, delegate { asked = true; }))
                            toOps.PerformClick();
                        if (!asked) failures.Add("повторный переход: вопроса не было");
                        if (bridged.Count != 2 || !string.Equals(bridged[bridged.Count - 1], pdf, StringComparison.OrdinalIgnoreCase))
                            failures.Add("ответ «Да» не передал исходный файл");
                        form.Close();
                    }
                }
            });
            AssertTrue(failures.Count == 0, "мост в «Прочие операции»: " + string.Join(" | ", failures.ToArray()));
        }

        // ---------- SH-1: публикация результата не зависит от сброса на диск ----------

        [Test("Аудит SH-1: сброс на диск перед публикацией не мешает, пока временный файл открыт другим процессом")]
        private static void AuditSh1CommitSurvivesUnflushableTemp()
        {
            using (var dir = new TempDir("iwo_stage0_sh1_"))
            {
                // Новый файл и замена существующего — два пути Commit.
                foreach (bool replace in new[] { false, true })
                {
                    string target = dir.File(replace ? "replace.txt" : "new.txt");
                    if (replace)
                        File.WriteAllText(target, "old");
                    using (var output = new AtomicOutput(target))
                    {
                        File.WriteAllText(output.TempPath, "fresh");
                        // Другой процесс (антивирус) держит temp открытым на чтение и не делит
                        // запись: открыть файл для сброса нельзя (IOException), а переименовать
                        // можно (общий Delete). Публикация обязана пройти без сброса.
                        using (new FileStream(output.TempPath, FileMode.Open, FileAccess.Read,
                            FileShare.Read | FileShare.Delete))
                            output.Commit();
                    }
                    AssertEqual("fresh", File.ReadAllText(target), (replace ? "замена" : "новый") + ": содержимое опубликовано");
                    AssertEq(1, Directory.GetFiles(dir.Path, Path.GetFileName(target) + "*").Length +
                        Directory.GetFiles(dir.Path, Path.GetFileNameWithoutExtension(target) + ".iwo-*").Length,
                        (replace ? "замена" : "новый") + ": временных файлов не осталось");
                }

                // Temp «только чтение»: открыть для сброса нельзя (UnauthorizedAccessException),
                // переименование при этом законно.
                string readOnlyTarget = dir.File("readonly.txt");
                using (var output = new AtomicOutput(readOnlyTarget))
                {
                    File.WriteAllText(output.TempPath, "fresh");
                    File.SetAttributes(output.TempPath, FileAttributes.ReadOnly);
                    output.Commit();
                }
                File.SetAttributes(readOnlyTarget, FileAttributes.Normal); // иначе TempDir не уберёт
                AssertEqual("fresh", File.ReadAllText(readOnlyTarget), "только чтение: содержимое опубликовано");
            }
        }

        // ---------- SH-4: уборка производных страниц в %TEMP% ----------

        [Test("Аудит SH-4: уборка при запуске убирает старые рендеры страниц и папки картинок, своё и чужое различает")]
        private static void AuditSh4SweepRemovesPageRendersAndImageDirs()
        {
            using (var tmp = new TempDir("iwo_stage0_sh4_"))
            {
                DateTime now = DateTime.UtcNow;
                string guid = Guid.NewGuid().ToString("N");
                string oldPng = tmp.File("iwo_pg_" + guid + ".png");
                string freshPng = tmp.File("iwo_pg_" + Guid.NewGuid().ToString("N") + ".png");
                string foreignPng = tmp.File("iwo_pg_photo.png");
                string oldImg = tmp.File("iwo_img_" + Guid.NewGuid().ToString("N"));
                string freshImg = tmp.File("iwo_img_" + Guid.NewGuid().ToString("N"));
                string foreignDir = tmp.File("iwo_img_backup");
                foreach (string f in new[] { oldPng, freshPng, foreignPng })
                    File.WriteAllText(f, "x");
                foreach (string d in new[] { oldImg, freshImg, foreignDir })
                {
                    Directory.CreateDirectory(d);
                    File.WriteAllText(Path.Combine(d, "img0.png"), "x");
                }
                File.SetLastWriteTimeUtc(oldPng, now.AddHours(-2));
                File.SetLastWriteTimeUtc(foreignPng, now.AddHours(-2));
                foreach (string d in new[] { oldImg, foreignDir })
                {
                    File.SetLastWriteTimeUtc(Path.Combine(d, "img0.png"), now.AddHours(-2));
                    Directory.SetLastWriteTimeUtc(d, now.AddHours(-2));
                }

                int removed = StartupSweep.SweepTempCopies(tmp.Path, now);

                AssertEq(2, removed, "удалены ровно старый рендер и старая папка картинок");
                AssertTrue(!File.Exists(oldPng) && !Directory.Exists(oldImg), "старые копии убраны");
                AssertTrue(File.Exists(freshPng) && Directory.Exists(freshImg), "свежие (живая работа) не тронуты");
                AssertTrue(File.Exists(foreignPng) && Directory.Exists(foreignDir), "чужие имена не тронуты");
            }
        }

        // ---------- PF-4: окно не убирает обёртки соседнего окна ----------

        [Test("Аудит PF-4: первая вставка листа не стирает обёртки другого окна, открытого дольше суток")]
        private static void AuditPf4WindowDoesNotSweepNeighbourWrappers()
        {
            string neighbour = Path.Combine(Path.GetTempPath(), "iwo_wrap_" + Guid.NewGuid().ToString("N"));
            string wrapper = Path.Combine(neighbour, "blank0000.pdf");
            try
            {
                Directory.CreateDirectory(neighbour);
                File.WriteAllText(wrapper, "x");
                Directory.SetLastWriteTimeUtc(neighbour, DateTime.UtcNow.AddDays(-2));
                InIsolatedSettings("iwo_stage0_pf4_", delegate
                {
                    using (var form = new PdfMergeForm(delegate { }, null))
                    {
                        MethodInfo wrapperDir = typeof(PdfPageOrderFormBase).GetMethod("WrapperDir",
                            BindingFlags.Instance | BindingFlags.NonPublic);
                        var own = (string)wrapperDir.Invoke(form, null);
                        AssertTrue(Directory.Exists(own), "своя папка обёрток создана");
                    }
                });
                AssertTrue(File.Exists(wrapper), "обёртка соседнего окна цела");
            }
            finally
            {
                try { Directory.Delete(neighbour, true); } catch { }
            }
        }

        // ---------- WP-1 / RC-16: растр Ghostscript — это CropBox ----------

        [Test("Аудит WP-1: растр страницы Ghostscript совпадает с CropBox, а не с MediaBox")]
        private static void AuditWp1RasterizerUsesCropBox()
        {
            if (!Ghostscript.Available)
                Skip("Ghostscript не найден");
            using (var dir = new TempDir("iwo_stage0_wp1_"))
            {
                string path = dir.File("crop.pdf");
                // Видимая часть 300×150 (альбомная) на книжном MediaBox 400×600.
                WriteRawPage(path, "0 0 0 rg 60 260 40 40 re f", "", "/CropBox [50 250 350 400]");
                using (BudgetedBitmap page = PageRasterizer.RenderPage(path, 1, 300, 150))
                {
                    AssertTrue(page != null, "Ghostscript отрисовал страницу");
                    Bitmap bitmap = page.Bitmap;
                    AssertTrue(Math.Abs((double)bitmap.Width / bitmap.Height - 2.0) < 0.05,
                        "пропорции растра — CropBox 2:1, а не MediaBox 2:3: " + bitmap.Width + "×" + bitmap.Height);
                    Rectangle ink = Stage0InkBounds(bitmap);
                    // Квадрат 40 pt в 10 pt от левого и нижнего края CropBox.
                    double sx = bitmap.Width / 300.0;
                    AssertTrue(Math.Abs(ink.Left - 10 * sx) <= 3 && Math.Abs(bitmap.Height - ink.Bottom - 10 * sx) <= 3,
                        "квадрат на месте относительно CropBox: " + ink);
                }
            }
        }

        // ---------- PS-2 / PS-3: аргументы pdfwrite ----------

        [Test("Аудит PS-2/PS-3: pdfwrite не крутит страницы, «Очень хорошо» кодирует картинки без потерь")]
        private static void AuditPs2Ps3PdfWriteArguments()
        {
            foreach (CompressionLevel level in new[] { CompressionLevel.VeryGood, CompressionLevel.Good, CompressionLevel.Small })
                AssertTrue(PdfCompression.BuildArguments("in.pdf", "out.pdf", level, null).Contains("-dAutoRotatePages=/None"),
                    "сжатие " + level + ": ориентация страниц не угадывается");
            foreach (PdfConvertMode mode in Enum.GetValues(typeof(PdfConvertMode)))
                AssertTrue(PdfConvert.BuildArguments("in.pdf", "out.pdf", mode, null).Contains("-dAutoRotatePages=/None"),
                    "преобразование " + mode + ": ориентация страниц не угадывается");
            string keep = PdfCompression.BuildArguments("in.pdf", "out.pdf", CompressionLevel.VeryGood, null);
            foreach (string key in new[] { "-dAutoFilterColorImages=false", "-dAutoFilterGrayImages=false",
                "-dColorImageFilter=/FlateEncode", "-dGrayImageFilter=/FlateEncode" })
                AssertTrue(keep.Contains(key), "«Очень хорошо»: " + key);
            string ebook = PdfCompression.BuildArguments("in.pdf", "out.pdf", CompressionLevel.Good, null);
            AssertTrue(!ebook.Contains("FlateEncode"), "пересчитывающий уровень сохраняет выбор фильтра движка");
        }

        [Test("Аудит PS-2 (живой): страница с текстом поперёк листа после pdfwrite стоит как стояла")]
        private static void AuditPs2PdfWriteKeepsOrientationLive()
        {
            if (!Ghostscript.Available)
                Skip("Ghostscript не найден");
            using (var dir = new TempDir("iwo_stage0_ps2_"))
            {
                string source = dir.File("portrait.pdf");
                var text = new StringBuilder();
                for (int i = 0; i < 12; i++)
                    text.Append("BT /F1 14 Tf 0 1 -1 0 ").Append(40 + i * 25).Append(" 60 Tm (Landscape table row ")
                        .Append(i).Append(" with long text) Tj ET ");
                WriteRawPage(source, text.ToString(), "", "");
                foreach (CompressionLevel level in new[] { CompressionLevel.Good, CompressionLevel.Small })
                {
                    string output = dir.File("out-" + level + ".pdf");
                    string stderr;
                    int exit = Ghostscript.Run(PdfCompression.BuildArguments(source, output, level, Ghostscript.BundledRoot),
                        120000, out stderr);
                    AssertEq(0, exit, level + ": движок отработал (" + stderr + ")");
                    using (PdfDocument doc = PdfReader.Open(output, PdfDocumentOpenMode.Import))
                    {
                        PdfPage page = doc.Pages[0];
                        AssertEq(0, page.Rotate, level + ": /Rotate не добавлен");
                        AssertTrue(page.Width.Point < page.Height.Point, level + ": лист остался книжным");
                    }
                }
            }
        }

        // ---------- CI-10: из нескольких Ghostscript берётся самый новый ----------

        [Test("Аудит CI-10: из установленных Ghostscript выбирается самый новый, а не первый по порядку")]
        private static void AuditCi10GhostscriptPicksNewest()
        {
            var versions = new Dictionary<string, Version>(StringComparer.OrdinalIgnoreCase)
            {
                { @"C:\gs\gs9.56.1\bin\gswin64c.exe", new Version(9, 56, 1) },
                { @"C:\gs\gs10.07.1\bin\gswin64c.exe", new Version(10, 7, 1) },
                { @"C:\gs\gs10.08.0\bin\gswin64c.exe", new Version(10, 8, 0) },
                { @"C:\user\gs64x\bin\gswin64c.exe", null }
            };
            var order = new List<string>(versions.Keys);
            string picked = Ghostscript.PickNewest(order, versions.ContainsKey, delegate(string p) { return versions[p]; });
            AssertEqual(@"C:\gs\gs10.08.0\bin\gswin64c.exe", picked, "самый новый из четырёх");
            AssertEqual(@"C:\user\gs64x\bin\gswin64c.exe", Ghostscript.PickNewest(
                new[] { @"C:\user\gs64x\bin\gswin64c.exe" }, versions.ContainsKey, delegate(string p) { return versions[p]; }),
                "единственный без версии всё равно годится");
            AssertEqual(@"C:\gs\gs9.56.1\bin\gswin64c.exe", Ghostscript.PickNewest(
                new[] { @"C:\user\gs64x\bin\gswin64c.exe", @"C:\gs\gs9.56.1\bin\gswin64c.exe" },
                versions.ContainsKey, delegate(string p) { return versions[p]; }), "с версией важнее, чем без неё");
            AssertEqual(null, Ghostscript.PickNewest(order, delegate { return false; }, delegate { return null; }), "ни одного — null");
            AssertEqual(@"C:\gs\gs10.07.1\bin\gswin64c.exe", Ghostscript.PickNewest(order,
                delegate(string p) { return !p.Contains("10.08") && !p.Contains("9.56"); },
                delegate(string p) { if (p.Contains("gs64x")) throw new IOException("нет доступа"); return versions[p]; }),
                "сбой чтения версии исключает только этого кандидата");
            // Порядок уровней: вшитый → установленные → PATH, и PATH не опрашивается, пока есть
            // установленный (его сетевые каталоги могли бы подвесить UI-поток).
            bool pathAsked = false;
            Func<IEnumerable<string>> pathList = delegate
            {
                pathAsked = true;
                return new[] { @"D:\tools\gs\gswin64c.exe" };
            };
            var all = new Dictionary<string, Version>(versions, StringComparer.OrdinalIgnoreCase)
            {
                { @"C:\app\gs\bin\gswin64c.exe", new Version(10, 8, 0) },
                { @"D:\tools\gs\gswin64c.exe", new Version(99, 0, 0) }
            };
            Func<string, Version> versionOf = delegate(string p) { return all[p]; };
            AssertEqual(@"C:\app\gs\bin\gswin64c.exe", Ghostscript.Choose(new[] { @"C:\app\gs\bin\gswin64c.exe" },
                delegate { return order; }, pathList, all.ContainsKey, versionOf), "вшитый — первым");
            AssertEqual(@"C:\gs\gs10.08.0\bin\gswin64c.exe", Ghostscript.Choose(new[] { @"C:\app\none.exe" },
                delegate { return order; }, pathList, all.ContainsKey, versionOf), "без вшитого — самый новый установленный");
            AssertTrue(!pathAsked, "при установленном движке PATH не опрашивается");
            AssertEqual(@"D:\tools\gs\gswin64c.exe", Ghostscript.Choose(new string[0],
                delegate { return new string[0]; }, pathList, all.ContainsKey, versionOf), "PATH — последний шанс");
            AssertTrue(pathAsked, "без установленных PATH опрошен");

            AssertEq(new Version(10, 7, 1), Ghostscript.ParseVersion("10.07.1"), "10.07.1");
            AssertEq(new Version(10, 8, 0), Ghostscript.ParseVersion("10.08.0 (2026-09-08)"), "хвост после версии");
            AssertEq(null, Ghostscript.ParseVersion("x"), "мусор");
            AssertEq(null, Ghostscript.ParseVersion(null), "пусто");
        }
    }
}
