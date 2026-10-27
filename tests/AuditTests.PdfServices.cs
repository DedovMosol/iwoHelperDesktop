using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using ExcelMerger;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
namespace ExcelMerger.Tests
{
    /// <summary>Аудит зоны C: PDF-сервисы, Ghostscript, метаданные, закладки, имена.</summary>
    internal static partial class UnitTests
    {
        // ---------- 1. «%» в пути вывода Ghostscript ----------

        [Test("Аудит C: «%» в -sOutputFile удваивается, шаблон страниц — только в имени")]
        private static void AuditGsOutputPercentEscaped()
        {
            AssertEq(@"C:\Счета 50%%\Contract%%20draft.pdf",
                GsArguments.LiteralOutput(@"C:\Счета 50%\Contract%20draft.pdf"), "буквальный путь");
            AssertEq(@"C:\a\b.pdf", GsArguments.LiteralOutput(@"C:\a\b.pdf"), "без «%» — как есть");
            AssertEq(@"C:\Папка 50%%\bg-%04d.png",
                GsArguments.OutputPattern(@"C:\Папка 50%", "bg-%04d.png"), "папка экранирована, шаблон — нет");

            string compress = PdfCompression.BuildArguments(@"C:\в 50%\a.pdf", @"C:\в 50%\a%20.tmp",
                CompressionLevel.Small, @"C:\gs");
            AssertTrue(compress.Contains("-sOutputFile=\"C:\\в 50%%\\a%%20.tmp\""), "сжатие: " + compress);
            AssertTrue(compress.EndsWith(" \"C:\\в 50%\\a.pdf\""), "вход не экранируется: " + compress);
            AssertTrue(compress.Contains("-I \"C:\\gs\\lib\""), "-I вшитого движка: " + compress);

            string gray = PdfConvert.BuildArguments(@"C:\x\a.pdf", @"C:\x\50%.tmp",
                PdfConvertMode.Grayscale, null);
            AssertTrue(gray.Contains("-sOutputFile=\"C:\\x\\50%%.tmp\""), "серое: " + gray);
            AssertTrue(!gray.Contains("-I "), "без вшитого движка -I нет: " + gray);

            string raster = PageRasterizer.BuildArgs(@"C:\x\a.pdf", 1, 2, 96,
                GsArguments.OutputPattern(@"C:\x 5%", "bg-%04d.png"), true, @"C:\gs");
            AssertTrue(raster.Contains("-sOutputFile=\"C:\\x 5%%\\bg-%04d.png\""), "растр: " + raster);
            AssertTrue(raster.Contains("-I \"C:\\gs\\Resource\\Init\""), "растр: -I: " + raster);
        }

        [Test("Аудит C: сжатие, серое и растр в папке и файле с «%» и «%20d» — без лишних файлов")]
        private static void AuditGsPercentPathsLive()
        {
            if (!Ghostscript.Available)
                Skip("Ghostscript не найден");
            using (var tmp = new TempDir("ExcelMergerAuditC_"))
            {
                string dir = Path.Combine(tmp.Path, "Счета 50% и %20d");
                Directory.CreateDirectory(dir);
                string img = Path.Combine(dir, "Предоплата 50%.pdf");
                MakeImagePdf(img, 2);
                string color = Path.Combine(dir, "Contract%20draft.pdf");
                AuditCColorPages(color, 3);
                string[] before = AuditCNames(dir);

                AssertTrue(PdfCompression.Compress(img, CompressionLevel.Small), "сжатие с «%» применено");
                AssertEq(2, PdfPageCount(img), "сжатие: страницы на месте");
                AssertTrue(PdfConvert.Apply(color, PdfConvertMode.Grayscale), "серое с «%20d» применено");
                AssertEq(3, PdfPageCount(color), "серое: страницы на месте");
                // Короткое ожидание — не на наш код: однажды наблюдался мимолётный
                // «<ИМЯ>.IWO-<guid>.PDF.tmp» в верхнем регистре — так пишет не приложение, а
                // замена файла средствами ОС/антивирус. Файлы на страницу от «%20d» сами не
                // исчезли бы, поэтому проверку это не ослабляет.
                string expected = string.Join("|", before);
                WaitFor(delegate { return expected == string.Join("|", AuditCNames(dir)); });
                AssertEq(expected, string.Join("|", AuditCNames(dir)),
                    "в папке ни лишних, ни пропавших файлов");

                string outDir = Path.Combine(dir, "фон 5%");
                List<string> pages = PageRasterizer.RenderPagesWithoutText(color, 1, 3, 36, outDir, null);
                AssertEq(3, pages.Count, "растр: по файлу на страницу в папке с «%»");
                AssertEq(3, Directory.GetFiles(outDir).Length, "растр: в папке только страницы");
            }
        }

        private static string[] AuditCNames(string dir)
        {
            var names = new List<string>();
            foreach (string f in Directory.GetFiles(dir))
                names.Add(Path.GetFileName(f));
            names.Sort(StringComparer.Ordinal);
            return names.ToArray();
        }

        /// <summary>Несколько страниц с цветной заливкой — многостраничный вход для серого.</summary>
        private static void AuditCColorPages(string path, int pages)
        {
            using (var doc = new PdfDocument())
            {
                for (int i = 0; i < pages; i++)
                {
                    PdfPage page = doc.AddPage();
                    using (XGraphics g = XGraphics.FromPdfPage(page))
                        g.DrawRectangle(new XSolidBrush(XColors.Red), 40, 40, 300, 200);
                }
                doc.Save(path);
            }
        }

        // ---------- 2. Метаданные: XMP и пустые ключи ----------

        [Test("Аудит C: запись свойств удаляет XMP каталога и пустые ключи Info")]
        private static void AuditMetadataDropsXmpAndEmptyKeys()
        {
            using (var tmp = new TempDir("ExcelMergerAuditC_"))
            {
                string src = tmp.File("исходник.pdf");
                AuditCXmpPdf(src, "Ivan Petrov");
                byte[] sourceBytes = File.ReadAllBytes(src);
                AssertTrue(AuditCHasXmp(src), "в исходнике есть XMP");
                string output = tmp.File("результат.pdf");
                PdfMetadataService.Write(src, output, new PdfMetadata
                {
                    Title = "Новый заголовок", Author = "", Subject = null, Keywords = ""
                });
                AssertTrue(!AuditCHasXmp(output), "XMP каталога удалён");
                using (PdfDocument doc = PdfReader.Open(output, PdfDocumentOpenMode.Import))
                {
                    AssertTrue(!doc.Info.Elements.ContainsKey("/Author"), "пустой автор — ключа нет");
                    AssertTrue(!doc.Info.Elements.ContainsKey("/Keywords"), "пустые ключевые слова — ключа нет");
                    AssertTrue(!doc.Info.Elements.ContainsKey("/Subject"), "null-тема — ключа нет");
                    AssertEq("Новый заголовок", doc.Info.Title, "заголовок записан");
                }
                AssertTrue(AuditCIndexOf(File.ReadAllBytes(output), Encoding.ASCII.GetBytes("Ivan Petrov")) < 0,
                    "имени автора в файле не осталось нигде");
                AssertTrue(Convert.ToBase64String(sourceBytes) == Convert.ToBase64String(File.ReadAllBytes(src)),
                    "исходник не изменён");
            }
        }

        [Test("Аудит C: запись свойств поверх исходника запрещена (в т.ч. через жёсткую ссылку)")]
        private static void AuditMetadataRefusesSource()
        {
            using (var tmp = new TempDir("ExcelMergerAuditC_"))
            {
                string src = tmp.File("исходник.pdf");
                AuditCXmpPdf(src, "Ivan Petrov");
                byte[] sourceBytes = File.ReadAllBytes(src);
                var meta = new PdfMetadata { Title = "x" };
                AssertThrowsExactly<MergeException>("тот же путь",
                    delegate { PdfMetadataService.Write(src, src, meta); });
                string link = tmp.File("ссылка.pdf");
                if (AuditCCreateHardLink(link, src, IntPtr.Zero))
                    AssertThrowsExactly<MergeException>("жёсткая ссылка на исходник",
                        delegate { PdfMetadataService.Write(src, link, meta); });
                AssertTrue(Convert.ToBase64String(sourceBytes) == Convert.ToBase64String(File.ReadAllBytes(src)),
                    "исходник не изменён");
            }
        }

        /// <summary>Одностраничный PDF с автором и в Info, и в XMP-потоке каталога.</summary>
        private static void AuditCXmpPdf(string path, string author)
        {
            using (var doc = new PdfDocument())
            {
                doc.AddPage();
                doc.Info.Author = author;
                doc.Info.Title = "Старый";
                doc.Info.Keywords = "k";
                string xmp = "<?xpacket begin=\"\" id=\"W5M0MpCehiHzreSzNTczkc9d\"?>" +
                    "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
                    "<rdf:Description rdf:about=\"\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\">" +
                    "<dc:creator><rdf:Seq><rdf:li>" + author + "</rdf:li></rdf:Seq></dc:creator>" +
                    "</rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end=\"w\"?>";
                var meta = new PdfDictionary(doc);
                meta.Elements["/Type"] = new PdfName("/Metadata");
                meta.Elements["/Subtype"] = new PdfName("/XML");
                meta.CreateStream(Encoding.UTF8.GetBytes(xmp));
                doc.Internals.AddObject(meta);
                doc.Internals.Catalog.Elements["/Metadata"] = meta.Reference;
                doc.Save(path);
            }
        }

        private static bool AuditCHasXmp(string path)
        {
            using (PdfDocument doc = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                return doc.Internals.Catalog.Elements.ContainsKey("/Metadata");
        }

        private static int AuditCIndexOf(byte[] haystack, byte[] needle)
        {
            for (int i = 0; i + needle.Length <= haystack.Length; i++)
            {
                int j = 0;
                while (j < needle.Length && haystack[i + j] == needle[j])
                    j++;
                if (j == needle.Length)
                    return i;
            }
            return -1;
        }

        // ---------- 3. Закладки: дети закладки без цели ----------

        [Test("Аудит C: закладка без цели выпадает, а её дети поднимаются на её уровень")]
        private static void AuditBookmarksPromoteChildren()
        {
            using (var doc = new PdfDocument())
            {
                for (int i = 0; i < 3; i++)
                    doc.AddPage();
                var folder = new PdfOutline { Title = "Папка" }; // «папка» без назначения
                doc.Outlines.Add(folder);
                PdfOutline first = folder.Outlines.Add("Глава 1", doc.Pages[0], true);
                first.Outlines.Add("Раздел 1.1", doc.Pages[1], true);
                folder.Outlines.Add("Глава 2", doc.Pages[2], true);
                doc.Outlines.Add("После", doc.Pages[2], true);
                // Читаем так же, как объединение и разделение: из сохранённого файла.
                string path = Path.Combine(Path.GetTempPath(), "iwo_auditc_marks_" + Guid.NewGuid().ToString("N") + ".pdf");
                List<PdfBookmark> marks;
                try
                {
                    doc.Save(path);
                    using (PdfDocument saved = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                        marks = PdfBookmarks.Read(saved);
                }
                finally { try { File.Delete(path); } catch { } }
                var sb = new StringBuilder();
                foreach (PdfBookmark m in marks)
                    sb.Append(m.Title).Append(':').Append(m.Level).Append(':').Append(m.PageIndex).Append(';');
                AssertEq("Глава 1:0:0;Раздел 1.1:1:1;Глава 2:0:2;После:0:2;", sb.ToString(),
                    "дети «папки» на месте");
            }
        }

        // ---------- 4. Имена устройств Windows ----------

        [Test("Аудит C: имена устройств Windows получают «_» без учёта регистра")]
        private static void AuditSanitizeDeviceNames()
        {
            AssertEq("CON_", NameTemplate.Sanitize("CON"), "CON");
            AssertEq("con_.part", NameTemplate.Sanitize("con.part"), "con.part");
            AssertEq("Nul_", NameTemplate.Sanitize("Nul"), "Nul");
            AssertEq("prn_", NameTemplate.Sanitize("prn."), "хвостовая точка срезается, потом проверка");
            AssertEq("AUX_.a.b", NameTemplate.Sanitize("AUX.a.b"), "часть до первой точки");
            AssertEq("COM1_", NameTemplate.Sanitize("COM1"), "COM1");
            AssertEq("lpt9_.pdf", NameTemplate.Sanitize("lpt9.pdf"), "lpt9");
            AssertEq("COM\u00B9_", NameTemplate.Sanitize("COM\u00B9"), "COM¹");
            AssertEq("COM0", NameTemplate.Sanitize("COM0"), "COM0 — обычное имя");
            AssertEq("CONSOLE", NameTemplate.Sanitize("CONSOLE"), "CONSOLE — обычное имя");
            AssertEq("CON отчёт", NameTemplate.Sanitize("CON отчёт"), "CON с продолжением — обычное имя");
            AssertEq("LPT10", NameTemplate.Sanitize("LPT10"), "LPT10 — обычное имя");
        }

        [Test("Аудит C: разбиение по закладке «CON» создаёт файл")]
        private static void AuditSplitByDeviceBookmark()
        {
            using (var tmp = new TempDir("ExcelMergerAuditC_"))
            {
                string src = tmp.File("исходник.pdf");
                using (var doc = new PdfDocument())
                {
                    doc.AddPage();
                    doc.AddPage();
                    doc.Outlines.Add("CON", doc.Pages[0], true);
                    doc.Outlines.Add("nul", doc.Pages[1], true);
                    doc.Save(src);
                }
                string outDir = Path.Combine(tmp.Path, "out");
                List<string> parts = PdfSplitService.SplitByBookmarks(src, outDir, "исходник",
                    null, null, null, "[BOOKMARK]");
                AssertEq(2, parts.Count, "две части");
                AssertEq("CON_.pdf", Path.GetFileName(parts[0]), "часть «CON»");
                AssertEq("nul_.pdf", Path.GetFileName(parts[1]), "часть «nul»");
                AssertTrue(File.Exists(parts[0]) && File.Exists(parts[1]), "файлы на диске");
            }
        }

        // ---------- 5. Идентичность файлов ----------

        [Test("Аудит C: IsSameFile узнаёт жёсткую ссылку")]
        private static void AuditIsSameFileIdentity()
        {
            using (var tmp = new TempDir("ExcelMergerAuditC_"))
            {
                string src = tmp.File("очень длинное имя исходника.pdf");
                File.WriteAllText(src, "x");
                string other = tmp.File("другой.pdf");
                File.WriteAllText(other, "x");
                AssertTrue(!OutputFile.IsSameFile(src, other), "разные файлы с одинаковым содержимым");
                AssertTrue(!OutputFile.IsSameFile(src, tmp.File("нет такого.pdf")), "несуществующий — не тот же");

                string link = tmp.File("жёсткая ссылка.pdf");
                AssertTrue(AuditCCreateHardLink(link, src, IntPtr.Zero), "жёсткая ссылка создана");
                AssertTrue(OutputFile.IsSameFile(src, link), "жёсткая ссылка — тот же файл");
                AssertTrue(OutputFile.IsSameFile(link, src), "симметрично");
            }
        }

        [Test("Аудит C: IsSameFile узнаёт короткое имя 8.3")]
        private static void AuditIsSameFileShortName()
        {
            using (var tmp = new TempDir("ExcelMergerAuditC_"))
            {
                string src = tmp.File("очень длинное имя исходника.pdf");
                File.WriteAllText(src, "x");
                var shortName = new StringBuilder(1024);
                uint n = AuditCShortPath(src, shortName, (uint)shortName.Capacity);
                if (n == 0 || n >= shortName.Capacity ||
                    string.Equals(Path.GetFileName(shortName.ToString()), Path.GetFileName(src),
                        StringComparison.OrdinalIgnoreCase))
                    SkipMissing("shortnames", "короткие имена 8.3 на этом томе выключены");
                AssertTrue(OutputFile.IsSameFile(src, shortName.ToString()), "имя 8.3 — тот же файл: " + shortName);
            }
        }

        [Test("Аудит C: слияние в жёсткую ссылку на исходник запрещено")]
        private static void AuditMergeRefusesHardLinkToSource()
        {
            using (var tmp = new TempDir("ExcelMergerAuditC_"))
            {
                string src = tmp.File("исходник.pdf");
                MakeEmptyPagesPdf(src, 2);
                byte[] bytes = File.ReadAllBytes(src);
                string link = tmp.File("ссылка.pdf");
                AssertTrue(AuditCCreateHardLink(link, src, IntPtr.Zero), "жёсткая ссылка создана");
                AssertThrowsExactly<MergeException>("слияние в ссылку", delegate
                {
                    PdfMergeService.Merge(new List<PdfPageRef>
                    {
                        new PdfPageRef { SourcePath = src, PageIndex = 0 },
                        new PdfPageRef { SourcePath = src, PageIndex = 1 }
                    }, link);
                });
                AssertTrue(Convert.ToBase64String(bytes) == Convert.ToBase64String(File.ReadAllBytes(src)),
                    "исходник не изменён");
            }
        }

        // ---------- 6. PATH: только полные пути ----------

        [Test("Аудит C: из PATH берутся только полностью заданные каталоги")]
        private static void AuditGhostscriptPathCandidates()
        {
            List<string> list = Ghostscript.PathCandidates(
                @".;bin;C:tools;\gs\bin; ;;..\x;C:\gs\bin; \\srv\share\gs ;D:/gs");
            AssertEq(string.Join("|", new[]
                {
                    @"C:\gs\bin\gswin64c.exe", @"C:\gs\bin\gswin32c.exe",
                    @"\\srv\share\gs\gswin64c.exe", @"\\srv\share\gs\gswin32c.exe",
                    @"D:/gs\gswin64c.exe", @"D:/gs\gswin32c.exe"
                }), string.Join("|", list.ToArray()), "кандидаты");
            AssertEq(0, Ghostscript.PathCandidates(null).Count, "PATH не задан");
            AssertEq(0, Ghostscript.PathCandidates("C:|bad\x01;C:").Count, "мусор не роняет");
        }

        // ---------- 7. Разделение открывает исходник один раз ----------

        [Test("Аудит C: разделение открывает исходник один раз на всю операцию")]
        private static void AuditSplitOpensSourceOnce()
        {
            using (var tmp = new TempDir("ExcelMergerAuditC_"))
            {
                string src = tmp.File("исходник.pdf");
                MakeBookmarkedPdf(src);
                PdfSourceSet.OpensForTests = 0;
                List<string> every = PdfSplitService.SplitEveryN(src, 1, Path.Combine(tmp.Path, "n"), "исходник");
                AssertEq(4, every.Count, "постранично — четыре части");
                AssertEq(1, PdfSourceSet.OpensForTests, "постранично: исходник разобран один раз");
                PdfSourceSet.OpensForTests = 0;
                List<string> byMarks = PdfSplitService.SplitByBookmarks(src, Path.Combine(tmp.Path, "b"), "исходник");
                AssertEq(2, byMarks.Count, "по закладкам — две части");
                AssertEq(1, PdfSourceSet.OpensForTests, "по закладкам: исходник разобран один раз");
                // Кэш оглавления на серию не должен ломать перенос: вложенная закладка на месте.
                string marks = ReadMarks(byMarks[0]);
                AssertTrue(marks.Contains("Глава А") && marks.Contains("Раздел А.1"), "оглавление части: " + marks);
            }
        }

        // ---------- 8. GsRewrite: цель изменилась во время работы движка ----------

        [Test("Аудит C: GsRewrite не подменяет файл, изменённый во время работы движка")]
        private static void AuditGsRewriteDetectsConcurrentChange()
        {
            if (!Ghostscript.Available)
                Skip("Ghostscript не найден");
            using (var tmp = new TempDir("ExcelMergerAuditC_"))
            {
                string path = tmp.File("документ.pdf");
                AuditCColorPages(path, 1);
                string temp = GsRewrite.TempOutput(path);
                string args = PdfConvert.BuildArguments(path, temp, PdfConvertMode.Grayscale, Ghostscript.BundledRoot);
                byte[] newer = null;
                // Хук verify вызывается между работой движка и подменой — ровно окно гонки.
                GsRewriteResult result = GsRewrite.RunDetailed(path, temp, args, 60000,
                    delegate(long a, long b) { return true; },
                    delegate(string produced)
                    {
                        AuditCColorPages(path, 2); // «кто-то другой» пересохранил файл
                        newer = File.ReadAllBytes(path);
                        return true;
                    });
                AssertEq(GsRewriteResult.Failed, result, "подмена отменена");
                AssertTrue(newer != null && Convert.ToBase64String(newer) ==
                    Convert.ToBase64String(File.ReadAllBytes(path)), "файл остался в новом виде");
                AssertTrue(!File.Exists(temp), "временный выход удалён");
            }
        }

        // ---------- 9. Доля изображений: общая картинка считается один раз ----------

        [Test("Аудит C: картинка, общая для многих страниц, считается в доле один раз")]
        private static void AuditImageShareCountsSharedOnce()
        {
            using (var tmp = new TempDir("ExcelMergerAuditC_"))
            {
                string jpg = tmp.File("logo.jpg");
                using (var bmp = new Bitmap(200, 200))
                {
                    var rnd = new Random(7);
                    for (int y = 0; y < 200; y++)
                        for (int x = 0; x < 200; x++)
                            bmp.SetPixel(x, y, Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256)));
                    bmp.Save(jpg, System.Drawing.Imaging.ImageFormat.Jpeg);
                }
                long jpgBytes = new FileInfo(jpg).Length;
                string pdf = tmp.File("бланк.pdf");
                using (var doc = new PdfDocument())
                using (XImage logo = XImage.FromFile(jpg))
                {
                    var font = new XFont("Arial", 9);
                    for (int i = 0; i < 40; i++)
                    {
                        PdfPage page = doc.AddPage();
                        using (XGraphics g = XGraphics.FromPdfPage(page))
                        {
                            g.DrawImage(logo, 20, 20, 60, 60); // один и тот же объект на каждой странице
                            for (int line = 0; line < 60; line++)
                                g.DrawString("Строка " + line + " страницы " + i + " " + Guid.NewGuid().ToString("N"),
                                    font, XBrushes.Black, new XPoint(30, 100 + line * 11));
                        }
                    }
                    doc.Save(pdf);
                }
                long total = new FileInfo(pdf).Length;
                double share = PdfContentProfile.ImageShare(pdf);
                double expected = (double)jpgBytes / total;
                AssertTrue(share > 0, "картинка учтена: " + share);
                AssertTrue(share <= expected * 1.1 + 0.01,
                    "доля ≈ одной картинки: " + share + " при ожидаемой " + expected);
            }
        }

        // ---------- 10. «Файл занят» — только когда занят ----------

        [Test("Аудит C: слияние различает «файл занят» и прочие ошибки записи")]
        private static void AuditMergeBusyVsOtherErrors()
        {
            using (var tmp = new TempDir("ExcelMergerAuditC_"))
            {
                string src = tmp.File("исходник.pdf");
                MakeEmptyPagesPdf(src, 1);
                var order = new List<PdfPageRef> { new PdfPageRef { SourcePath = src, PageIndex = 0 } };
                string busyText = Loc.T("err.pdf.fileBusy");

                string locked = tmp.File("открыт.pdf");
                File.WriteAllText(locked, "x");
                using (new FileStream(locked, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    MergeException busy = AssertThrowsExactly<MergeException>("занятый файл",
                        delegate { PdfMergeService.Merge(order, locked); });
                    AssertEq(busyText, busy.Message, "занятый файл — «занят»");
                }

                string folder = tmp.File("это папка.pdf");
                Directory.CreateDirectory(folder);
                MergeException other = AssertThrowsExactly<MergeException>("путь — папка",
                    delegate { PdfMergeService.Merge(order, folder); });
                AssertTrue(other.Message != busyText, "не «занят»: " + other.Message);

                string missing = Path.Combine(tmp.Path, "нет папки", "итог.pdf");
                MergeException gone = AssertThrowsExactly<MergeException>("нет папки",
                    delegate { PdfMergeService.Merge(order, missing); });
                AssertTrue(gone.Message != busyText, "нет папки — не «занят»: " + gone.Message);

                AssertTrue(OutputFile.IsBusy(new IOException("x", unchecked((int)0x80070020))), "sharing violation");
                AssertTrue(OutputFile.IsBusy(new IOException("x", unchecked((int)0x80070021))), "lock violation");
                AssertTrue(!OutputFile.IsBusy(new IOException("x", unchecked((int)0x80070035))), "нет сети — не «занят»");
                AssertTrue(!OutputFile.IsBusy(new UnauthorizedAccessException()), "нет прав — не «занят»");
            }
        }

        // ---------- 11. Расшифрованные копии растра ----------

        [Test("Аудит C: расшифрованная копия растра — в %TEMP% с префиксом и удаляется")]
        private static void AuditRasterDecryptedCopyCleaned()
        {
            if (!Ghostscript.Available)
                Skip("Ghostscript не найден");
            using (var tmp = new TempDir("ExcelMergerAuditC_"))
            {
                string pdf = tmp.File("защищён.pdf");
                MakeProtectedPdf(pdf);
                PdfPasswords.Remember(pdf, "1234");
                try
                {
                    var before = new HashSet<string>(Directory.GetFiles(Path.GetTempPath(),
                        PageRasterizer.DecryptedPrefix + "*.pdf"), StringComparer.OrdinalIgnoreCase);
                    string outDir = Path.Combine(tmp.Path, "фон");
                    List<string> pages = PageRasterizer.RenderPagesWithoutText(pdf, 1, 1, 36, outDir, null);
                    AssertEq(1, pages.Count, "страница отрисована с паролем");
                    AssertEq(1, Directory.GetFiles(outDir).Length, "в папке вызывающего только страница");
                    using (BudgetedBitmap bmp = PageRasterizer.RenderPage(pdf, 1, 595, 842))
                        AssertTrue(bmp != null, "одиночный рендер с паролем");
                    foreach (string f in Directory.GetFiles(Path.GetTempPath(), PageRasterizer.DecryptedPrefix + "*.pdf"))
                        AssertTrue(before.Contains(f), "расшифрованная копия осталась: " + f);
                }
                finally { PdfPasswords.Clear(); }
            }
        }

        [DllImport("kernel32.dll", EntryPoint = "CreateHardLinkW", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AuditCCreateHardLink(string fileName, string existingFileName, IntPtr securityAttributes);

        [DllImport("kernel32.dll", EntryPoint = "GetShortPathNameW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint AuditCShortPath(string longPath, StringBuilder shortPath, uint bufferLength);
    }
}
