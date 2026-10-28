using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Drawing;
using System.IO.Compression;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using ExcelMerger;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Тесты этапа «1. Корректность данных» плана аудита (audit.md). Каждый тест называет
    /// запись аудита, которую закрывает, и падает на коде до исправления.
    /// </summary>
    internal static partial class UnitTests
    {
        // ---------- PF-1: имена картинок из «Прочих операций» ----------

        [Test("Аудит PF-1: картинки собранного документа называются по документу, а не по временной сборке")]
        private static void AuditPf1ImagesNamedAfterUserDocument()
        {
            using (var dir = new TempDir("iwo_stage1_pf1_"))
            {
                // Так «Прочие операции» называют сборку после перестановки страниц.
                string assembled = dir.File("iwo_ops_" + Guid.NewGuid().ToString("N") + ".pdf");
                File.Copy(MakePagesPdf(dir.Path, 2), assembled);
                string named = dir.File("named");
                List<string> files = PdfExportService.ToImages(assembled, new List<int> { 0, 1 }, named,
                    NameTemplate.Default, ImageExportFormat.Png, 96, null, null, "Отчёт");
                var names = new List<string>();
                foreach (string f in files)
                    names.Add(Path.GetFileName(f));
                AssertEq("Отчёт_1.png|Отчёт_2.png", string.Join("|", names.ToArray()),
                    "имя — от документа человека");

                // Без явного имени — как раньше, по читаемому файлу (CLI).
                string plain = dir.File("plain");
                List<string> legacy = PdfExportService.ToImages(assembled, new List<int> { 0 }, plain,
                    NameTemplate.Default, ImageExportFormat.Png, 96);
                AssertEq(Path.GetFileNameWithoutExtension(assembled) + "_1.png", Path.GetFileName(legacy[0]),
                    "без baseName имя берётся у читаемого файла");
            }
        }

        // ---------- PF-3: отмена сжатия, серого, починки и записи свойств ----------

        /// <summary>«Прочие операции» с прямым входом в пролог неизмеримой операции.</summary>
        private sealed class Stage1UnmeasuredProbe : PdfOpsForm
        {
            public Stage1UnmeasuredProbe() : base(delegate { }) { }
            public void Begin(int units) { BeginUnmeasuredOperation(Loc.T("ops.status.compressing"), units); }
            public void BeginUnknown() { BeginUnmeasuredOperation(Loc.T("ops.status.converting"), UnknownSize); }
            public void End() { EndOperation(); }
        }

        [Test("Аудит PF-3: неизмеримая операция от порога страниц предлагает «Отмену»")]
        private static void AuditPf3UnmeasuredOperationOffersCancel()
        {
            var failures = new List<string>();
            int enough = 1;
            while (!PdfToolFormBase.ShouldOfferCancel(enough) && enough < 1000)
                enough++;
            InIsolatedSettings("iwo_stage1_pf3_", delegate
            {
                using (var f = new Stage1UnmeasuredProbe())
                {
                    f.Show();
                    f.Begin(enough - 1);
                    if (FindVisibleButton(f, Loc.T("common.cancel")) != null)
                        failures.Add("короче порога «Отмена» не нужна");
                    f.End();
                    f.Begin(enough);
                    if (FindVisibleButton(f, Loc.T("common.cancel")) == null)
                        failures.Add("от порога страниц «Отмены» нет");
                    f.End();
                    f.BeginUnknown();
                    if (FindVisibleButton(f, Loc.T("common.cancel")) == null)
                        failures.Add("при неизвестном размере (починка) «Отмены» нет");
                    f.End();
                    f.Close();
                }
            });
            AssertTrue(failures.Count == 0, "отмена неизмеримых операций: " + string.Join(" | ", failures.ToArray()));
        }

        /// <summary>Временные файлы движка рядом с path (path.iwo-gs-&lt;guid&gt;.tmp).</summary>
        private static string[] Stage1EngineLeftovers(string path)
        {
            return Directory.GetFiles(Path.GetDirectoryName(path), Path.GetFileName(path) + ".iwo-gs-*");
        }

        [Test("Аудит PF-3: отмена сжатия и преобразования завершает движок и не трогает файл")]
        private static void AuditPf3EngineCancellationLeavesFileIntact()
        {
            if (!Ghostscript.Available)
                Skip("Ghostscript не найден");
            using (var dir = new TempDir("iwo_stage1_pf3e_"))
            {
                string source = MakePagesPdf(dir.Path, 8);
                byte[] before = File.ReadAllBytes(source);
                var cases = new Dictionary<string, Action<Func<bool>>>
                {
                    { "сжатие", delegate(Func<bool> c) { PdfCompression.CompressDetailedUnpublished(source, CompressionLevel.Small, c); } },
                    { "серое", delegate(Func<bool> c) { PdfConvert.ApplyUnpublished(source, PdfConvertMode.Grayscale, c); } },
                    { "починка", delegate(Func<bool> c) { PdfConvert.ApplyUnpublished(source, PdfConvertMode.Repair, c); } }
                };
                foreach (KeyValuePair<string, Action<Func<bool>>> item in cases)
                {
                    // Первая проверка пропускает (движок стартует), следующая — отмена во время работы.
                    int calls = 0;
                    bool threw = false;
                    try { item.Value(delegate { return ++calls > 1; }); }
                    catch (OperationCanceledException) { threw = true; }
                    AssertTrue(threw, item.Key + ": отмена дошла до вызывающего как OperationCanceledException, а не «сбой»");
                    // Проверка до движка, опрос самим движком (на нём он и прерван), проверка после.
                    AssertTrue(calls >= 3, item.Key + ": отмену опрашивал сам работающий движок (вызовов " + calls + ")");
                    AssertTrue(Convert.ToBase64String(before) == Convert.ToBase64String(File.ReadAllBytes(source)),
                        item.Key + ": файл не изменён");
                    AssertEq(0, Stage1EngineLeftovers(source).Length, item.Key + ": вывод движка удалён");
                }
                // Без отмены те же вызовы работают (иначе «бросило» ничего бы не доказывало).
                AssertTrue(PdfConvert.ApplyUnpublished(source, PdfConvertMode.Repair, delegate { return false; }),
                    "без отмены починка применяется");
            }
        }

        [Test("Аудит PF-3: отменённая запись свойств не создаёт файл")]
        private static void AuditPf3MetadataWriteCancellation()
        {
            using (var dir = new TempDir("iwo_stage1_pf3m_"))
            {
                string source = MakePagesPdf(dir.Path, 1);
                string output = dir.File("свойства.pdf");
                bool threw = false;
                try { PdfMetadataService.Write(source, output, new PdfMetadata { Title = "T" }, delegate { return true; }); }
                catch (OperationCanceledException) { threw = true; }
                AssertTrue(threw, "отмена бросает OperationCanceledException");
                AssertTrue(!File.Exists(output), "результата нет");
                AssertEq(1, Directory.GetFiles(dir.Path).Length, "временных файлов не осталось");
                PdfMetadataService.Write(source, output, new PdfMetadata { Title = "T" }, delegate { return false; });
                AssertTrue(File.Exists(output), "без отмены файл записан");
            }
        }

        // ---------- PS-5: свойства документа под паролем ----------

        /// <summary>PDF с РАЗНЫМИ паролями на открытие и владельца — как их ставит Acrobat.</summary>
        private static void Stage1OwnerProtectedPdf(string path)
        {
            using (var doc = new PdfDocument())
            {
                doc.AddPage();
                doc.AddPage();
                doc.SecuritySettings.UserPassword = "1234";
                doc.SecuritySettings.OwnerPassword = "владелец";
                doc.Save(path);
            }
        }

        [Test("Аудит PS-5: свойства файла под паролем записываются, когда человек знает только пароль на открытие")]
        private static void AuditPs5MetadataOfProtectedDocument()
        {
            using (var dir = new TempDir("iwo_stage1_ps5_"))
            {
                string locked = dir.File("под паролем.pdf");
                Stage1OwnerProtectedPdf(locked);
                PdfPasswords.Clear();
                try
                {
                    PdfPasswords.Remember(locked, "1234");
                    // Так было: прямая правка требует пароля владельца.
                    bool direct = false;
                    try { PdfMetadataService.Write(locked, dir.File("напрямую.pdf"), new PdfMetadata { Title = "T" }); }
                    catch (MergeException) { direct = true; }
                    AssertTrue(direct, "проверка предпосылки: PDFsharp правит такой файл только с паролем владельца");

                    Type assembled = typeof(PdfOpsForm).GetNestedType("AssembledDoc", BindingFlags.NonPublic);
                    MethodInfo whole = assembled.GetMethod("Whole", BindingFlags.Public | BindingFlags.Static);
                    string output = dir.File("со свойствами.pdf");
                    var doc = (IDisposable)whole.Invoke(null, new object[] { locked, AuditDPristine(locked, 2), 2, null, true });
                    try
                    {
                        string readFrom = (string)assembled.GetField("FilePath").GetValue(doc);
                        AssertTrue(!string.Equals(readFrom, locked, StringComparison.OrdinalIgnoreCase),
                            "для записи защищённый документ собирается");
                        PdfMetadataService.Write(readFrom, output, new PdfMetadata { Title = "Отчёт", Author = "" });
                    }
                    finally { doc.Dispose(); }
                    using (PdfDocument result = PdfReader.Open(output, PdfDocumentOpenMode.Import))
                    {
                        AssertEq(2, result.PageCount, "страницы на месте");
                        AssertEq("Отчёт", result.Info.Title, "заголовок записан");
                    }

                    // Чтение (текст) по-прежнему идёт прямо из исходника.
                    var reading = (IDisposable)whole.Invoke(null, new object[] { locked, AuditDPristine(locked, 2), 2, null, false });
                    try
                    {
                        AssertEq(locked, (string)assembled.GetField("FilePath").GetValue(reading),
                            "для чтения нетронутый документ не собирается");
                    }
                    finally { reading.Dispose(); }
                }
                finally { PdfPasswords.Clear(); }
            }
        }

        [Test("Аудит PS-5: отметка PDF/A распознаётся до записи свойств, которая её снимет")]
        private static void AuditPs5DetectsPdfA()
        {
            using (var dir = new TempDir("iwo_stage1_pdfa_"))
            {
                string archival = dir.File("архив.pdf");
                using (var doc = new PdfDocument())
                {
                    doc.AddPage();
                    string xmp = "<?xpacket begin=\"\" id=\"W5M0MpCehiHzreSzNTczkc9d\"?>" +
                        "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
                        "<rdf:Description rdf:about=\"\" xmlns:pdfaid=\"http://www.aiim.org/pdfa/ns/id/\">" +
                        "<pdfaid:part>2</pdfaid:part><pdfaid:conformance>B</pdfaid:conformance>" +
                        "</rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end=\"w\"?>";
                    var meta = new PdfDictionary(doc);
                    meta.Elements["/Type"] = new PdfName("/Metadata");
                    meta.Elements["/Subtype"] = new PdfName("/XML");
                    meta.CreateStream(System.Text.Encoding.UTF8.GetBytes(xmp));
                    doc.Internals.AddObject(meta);
                    doc.Internals.Catalog.Elements["/Metadata"] = meta.Reference;
                    doc.Save(archival);
                }
                string plain = MakePagesPdf(dir.Path, 1);
                AssertTrue(PdfMetadataService.Read(archival).DeclaresPdfA, "файл с pdfaid — архивный");
                AssertTrue(!PdfMetadataService.Read(plain).DeclaresPdfA, "обычный — нет");

                // Признак, разрезанный границей блока чтения, всё равно находится.
                byte[] marker = System.Text.Encoding.ASCII.GetBytes("http://www.aiim.org/pdfa/ns/id/");
                var data = new byte[64 * 1024 + 40];
                Array.Copy(marker, 0, data, 64 * 1024 - 10, marker.Length);
                AssertTrue(PdfMetadataService.DeclaresPdfA(new MemoryStream(data)), "признак на границе блока");
                AssertTrue(!PdfMetadataService.DeclaresPdfA(new MemoryStream(new byte[100])), "пустой поток");

                // Запись свойств снимает отметку — о чём и предупреждает вопрос.
                string written = dir.File("свойства.pdf");
                PdfMetadataService.Write(archival, written, new PdfMetadata { Title = "T" });
                AssertTrue(!PdfMetadataService.Read(written).DeclaresPdfA, "после записи отметки нет");
            }
        }

        // ---------- PS-4: печать с учётом жёстких полей принтера ----------

        [Test("Аудит PS-4: страница печатается по центру листа и целиком в области печати принтера")]
        private static void AuditPs4PrintTargetRespectsHardMargins()
        {
            var a4 = new Rectangle(0, 0, 827, 1169);
            AssertEq(a4, PdfPrintService.PrintTarget(a4, new RectangleF(0, 0, 827, 1169), 0, 0, false),
                "без жёстких полей — весь лист");
            // Поля 20/27 по горизонтали и 16/13 по вертикали: сужаем на большее по каждой оси.
            var printable = new RectangleF(20, 16, 780, 1140);
            AssertEq(new Rectangle(7, 0, 773, 1137), PdfPrintService.PrintTarget(a4, printable, 20, 16, false),
                "принтер: начало координат — угол печатаемой области");
            AssertEq(new Rectangle(27, 16, 773, 1137), PdfPrintService.PrintTarget(a4, printable, 20, 16, true),
                "предпросмотр: начало — угол листа, место на листе то же");
            Rectangle onPaper = PdfPrintService.PrintTarget(a4, printable, 20, 16, false);
            AssertTrue(onPaper.X >= 0 && onPaper.Y >= 0 && onPaper.Right <= 780 && onPaper.Bottom <= 1140,
                "целиком внутри печатаемой области: " + onPaper);
            var landscape = new Rectangle(0, 0, 1169, 827);
            AssertEq(new Rectangle(0, 7, 1137, 773), PdfPrintService.PrintTarget(landscape, printable, 16, 20, false),
                "альбомный лист с «книжной» областью печати от драйвера");
            AssertEq(a4, PdfPrintService.PrintTarget(a4, RectangleF.Empty, 0, 0, false), "нет данных — весь лист");
            AssertEq(a4, PdfPrintService.PrintTarget(a4, new RectangleF(500, 700, 10, 10), 500, 700, false),
                "поля больше листа — весь лист");
        }

        // ---------- XL-2: повтор переносит только упавшие листы ----------

        private static FileResult Stage1Sheet(string path, string sheet, bool ok)
        {
            return new FileResult
            {
                FileName = Path.GetFileName(path), FullPath = path, SourceSheet = sheet, Ok = ok,
                SheetName = ok ? Path.GetFileNameWithoutExtension(path) + " · " + sheet : null,
                Note = ok ? null : "лист не перенесён (тест)"
            };
        }

        [Test("Аудит XL-2: повтор «все листы» берёт каждый файл один раз и только упавшие листы")]
        private static void AuditXl2RetryOnlyFailedSheets()
        {
            const string a = @"C:\in\a.xlsx", b = @"C:\in\b.xlsx", c = @"C:\in\c.xlsx";
            var previous = new MergeResult();
            previous.Files.Add(Stage1Sheet(a, "Янв", true));
            previous.Files.Add(Stage1Sheet(a, "Мар", false));
            previous.Files.Add(Stage1Sheet(a, "Апр", false));
            previous.Files.Add(Stage1Sheet(b, null, false)); // файл не открылся
            previous.Files.Add(Stage1Sheet(c, "Лист1", false));
            previous.Files.Add(Stage1Sheet(c, "Лист2", false)); // ни один лист файла не перенёсся
            AssertEq(a + "|" + b + "|" + c, string.Join("|", MergeRetry.Paths(previous).ToArray()),
                "каждый файл — один раз, порядок прежний");
            Dictionary<string, HashSet<string>> sheets = MergeRetry.Sheets(previous);
            var forA = new List<string>(sheets[a]);
            forA.Sort(StringComparer.Ordinal);
            AssertEq("Апр|Мар", string.Join("|", forA.ToArray()), "у частично перенесённого — только упавшие листы");
            AssertTrue(sheets[a].Contains("мар"), "имена листов без учёта регистра, как в Excel");
            AssertTrue(sheets[b] == null, "не открывшийся файл — целиком");
            AssertTrue(sheets[c] == null, "файл без единого перенесённого листа — целиком");

            var fresh = new List<FileResult> { Stage1Sheet(a, "Мар", true), Stage1Sheet(a, "Апр", true) };
            var partial = new MergeResult();
            partial.Files.Add(Stage1Sheet(a, "Янв", true));
            partial.Files.Add(Stage1Sheet(a, "Мар", false));
            partial.Files.Add(Stage1Sheet(a, "Апр", false));
            MergeResult combined = MergeRetry.Combine(partial, fresh);
            var names = new List<string>();
            foreach (FileResult f in combined.Files)
                names.Add(f.SourceSheet + (f.Ok ? "+" : "-"));
            AssertEq("Янв+,Мар+,Апр+", string.Join(",", names.ToArray()), "без дублей перенесённых листов");
        }

        /// <summary>Книга с листами names (первый — «Лист1» переименовывается), значение в A1 каждого.</summary>
        private static void Stage1SaveMultiSheetBook(object app, string path, params string[] names)
        {
            AuditBSaveBook(app, path, delegate(dynamic ws)
            {
                dynamic book = ws.Parent;
                ws.Name = names[0];
                ws.Range("A1").Value2 = names[0];
                for (int i = 1; i < names.Length; i++)
                {
                    dynamic added = book.Worksheets.Add(Type.Missing, book.Worksheets[book.Worksheets.Count]);
                    added.Name = names[i];
                    added.Range("A1").Value2 = names[i];
                }
            });
        }

        private static List<string> Stage1SheetNames(string path)
        {
            var names = new List<string>();
            AuditBReadBook(path, delegate(dynamic wb)
            {
                int count = (int)wb.Sheets.Count;
                for (int i = 1; i <= count; i++)
                    names.Add((string)wb.Sheets[i].Name);
            });
            return names;
        }

        [Test("Аудит XL-2: повтор после упавшего листа не копирует заново перенесённые (живой Excel)")]
        private static void AuditXl2RetryCopiesOnlyFailedSheetLive()
        {
            AuditBRequireExcel();
            using (var dir = new TempDir("iwo_stage1_xl2_"))
            {
                string source = dir.File("Отчёт.xlsx");
                string digest = dir.File("Свод.xlsx");
                var options = new MergeOptions { AllSheets = true };
                MergeResult retried = null;
                AuditBOfficeSta(180000, delegate
                {
                    AuditBWithExcel(delegate(object app) { Stage1SaveMultiSheetBook(app, source, "Янв", "Фев", "Мар"); });
                    MergeResult first = new MergeService().Merge(new[] { source }, digest, options);
                    AssertEq(3, first.OkCount, "первый прогон перенёс три листа");
                    // Как будто «Мар» не перенёсся: его записи в результате нет как успешной.
                    foreach (FileResult fr in first.Files)
                        if (fr.SourceSheet == "Мар")
                        {
                            fr.Ok = false;
                            fr.SheetName = null;
                            fr.Note = "лист не перенесён (тест)";
                        }
                    first.OkCount = 2;
                    first.SkipCount = 1;
                    retried = new MergeService().RetrySkipped(digest, options, first);
                });
                AssertEq(3, retried.Files.Count, "в итоге три записи, а не пять");
                AssertEq(0, retried.SkipCount, "пропусков нет");
                List<string> sheets = Stage1SheetNames(digest);
                AssertEq(4, sheets.Count, "в своде прибавился один лист (Мар), а не три: " + string.Join(", ", sheets.ToArray()));
            }
        }

        // ---------- XL-4: свой экземпляр Office с чужими файлами не закрывается ----------

        [Test("Аудит XL-4: Excel с книгой пользователя не закрывается без сохранения, пустой — закрывается (живой Excel)")]
        private static void AuditXl4ExcelWithUserBookIsHandedOverLive()
        {
            AuditBRequireExcel();
            AuditBOfficeSta(90000, delegate
            {
                // С «книгой пользователя»: экземпляр остаётся и показывается.
                object app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                int pid = OwnedOfficeProcess.FromWindow((long)((dynamic)app).Hwnd).Id;
                object book = null;
                try
                {
                    ((dynamic)app).Visible = false;
                    book = ((dynamic)app).Workbooks.Add();
                    AssertEq(1, MergeService.QuitOrHandOver(app), "одна чужая книга");
                    AssertTrue((bool)((dynamic)app).Visible, "Excel показан человеку");
                    AssertTrue(AuditBAlive(pid), "процесс не завершён");
                    AssertEq(1, (int)((dynamic)app).Workbooks.Count, "книга на месте");
                }
                finally
                {
                    if (book != null) { try { ((dynamic)book).Close(false); } catch { } ComSafe.Release(book); }
                    try { ((dynamic)app).Quit(); } catch { }
                    ComSafe.Release(app);
                    ComSafe.Collect();
                }
                // Пустой: закрывается.
                object empty = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                int emptyPid = OwnedOfficeProcess.FromWindow((long)((dynamic)empty).Hwnd).Id;
                AssertEq(0, MergeService.QuitOrHandOver(empty), "чужих книг нет");
                ComSafe.Release(empty);
                ComSafe.Collect();
                AssertTrue(WaitFor(delegate { return !AuditBAlive(emptyPid); }, 20000), "пустой экземпляр завершился");
            });
        }

        [Test("Аудит XL-4: Word с документом пользователя не закрывается без сохранения (живой Word)")]
        private static void AuditXl4WordWithUserDocumentIsHandedOverLive()
        {
            AuditBRequireWord();
            AuditBOfficeSta(90000, delegate
            {
                object app = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
                object document = null;
                try
                {
                    ((dynamic)app).Visible = false;
                    document = ((dynamic)app).Documents.Add();
                    AssertEq(1, WordCom.QuitOrHandOver(app), "один чужой документ");
                    AssertTrue((bool)((dynamic)app).Visible, "Word показан человеку");
                    AssertEq(1, (int)((dynamic)app).Documents.Count, "документ на месте");
                }
                finally
                {
                    if (document != null) { try { ((dynamic)document).Close(0); } catch { } ComSafe.Release(document); }
                    try { ((dynamic)app).Quit(0); } catch { }
                    ComSafe.Release(app);
                    ComSafe.Collect();
                }
            });
        }

        // ---------- XL-5: пределы формата .xls ----------

        [Test("Аудит XL-5: данные за пределами .xls дают примечание, лист нового формата в .xls-свод — понятную причину")]
        private static void AuditXl5LegacyGridLimits()
        {
            AssertTrue(MergeService.LegacyGridNote(65536, 256) == null, "ровно предел помещается");
            AssertTrue(MergeService.LegacyGridNote(65537, 1) != null, "строка за пределом");
            AssertTrue(MergeService.LegacyGridNote(1, 257) != null, "столбец за пределом");
            AssertTrue(MergeService.LegacyGridNote(0, 0) == null, "пустой лист");
        }

        [Test("Аудит XL-5: свод в .xls отмечает отброшенные строки и не пытается дописать лист .xlsx (живой Excel)")]
        private static void AuditXl5LegacyGridLimitsLive()
        {
            AuditBRequireExcel();
            using (var dir = new TempDir("iwo_stage1_xl5_"))
            {
                string big = dir.File("Большой.xlsx");
                string small = dir.File("Малый.xls");
                string digest = dir.File("Свод.xls");
                MergeResult merged = null, retried = null;
                AuditBOfficeSta(180000, delegate
                {
                    AuditBWithExcel(delegate(object app)
                    {
                        AuditBSaveBook(app, big, delegate(dynamic ws)
                        {
                            ws.Range("A1").Value2 = 1;
                            ws.Range("A70000").Value2 = 2; // за 65 536-й строкой
                            ws.Range("C1:C3").Interior.Color = 255; // только оформление — не данные
                        });
                        AuditBSaveBook(app, small, delegate(dynamic ws) { ws.Range("A1").Value2 = 3; });
                    });
                    merged = new MergeService().Merge(new[] { big, small }, digest, new MergeOptions());
                    // Повтор: .xlsx в уже существующий .xls-свод (как будто «Большой» не перенёсся).
                    var previous = new MergeResult { OutputPath = digest };
                    previous.Files.Add(new FileResult { FileName = "Большой.xlsx", FullPath = big, Note = "тест" });
                    foreach (FileResult fr in merged.Files)
                        if (fr.FullPath == small)
                            previous.Files.Add(fr);
                    retried = new MergeService().RetrySkipped(digest, new MergeOptions(), previous);
                });
                FileResult bigResult = merged.Files[0];
                AssertTrue(bigResult.Ok, "лист перенесён");
                AssertTrue((bigResult.Note ?? "").Contains("65 536"), "примечание об отброшенном: «" + bigResult.Note + "»");
                AssertTrue(string.IsNullOrEmpty(merged.Files[1].Note), "лист .xls без примечания: «" + merged.Files[1].Note + "»");
                FileResult again = null;
                foreach (FileResult fr in retried.Files)
                    if (fr.FullPath == big)
                        again = fr;
                AssertTrue(again != null && !again.Ok && again.Note == MergeService.BigGridIntoLegacyNote,
                    "понятная причина вместо текста Excel: «" + (again == null ? "нет записи" : again.Note) + "»");
            }
        }

        // ---------- XL-1: динамические массивы при «формулы → значения» ----------

        [Test("Аудит XL-1: разлив динамического массива переносится значениями целиком (живой Excel 2021+)")]
        private static void AuditXl1SpillKeptAsValuesLive()
        {
            AuditBRequireExcel();
            bool spills = false;
            AuditBOfficeSta(60000, delegate
            {
                AuditBWithExcel(delegate(object app)
                {
                    dynamic wb = ((dynamic)app).Workbooks.Add();
                    try
                    {
                        wb.Sheets[1].Range("A1").Formula2 = "=SEQUENCE(3)";
                        spills = (bool)wb.Sheets[1].Range("A1").HasSpill;
                    }
                    catch { spills = false; }
                    finally { object o = wb; try { wb.Close(false); } catch { } ComSafe.Release(o); }
                });
            });
            if (!spills)
                SkipMissing("office", "в этом Excel нет динамических массивов (нужен 2021 или Microsoft 365)");
            using (var dir = new TempDir("iwo_stage1_xl1_"))
            {
                string source = dir.File("Разлив.xlsx");
                string digest = dir.File("Свод.xlsx");
                MergeResult result = null;
                AuditBOfficeSta(180000, delegate
                {
                    AuditBWithExcel(delegate(object app)
                    {
                        AuditBSaveBook(app, source, delegate(dynamic ws) { ws.Range("A1").Formula2 = "=SEQUENCE(3)*10"; });
                    });
                    result = new MergeService().Merge(new[] { source }, digest, new MergeOptions { ValuesOnly = true });
                });
                AssertTrue(result.Files[0].Ok && string.IsNullOrEmpty(result.Files[0].Note), "заменено полностью: «" + result.Files[0].Note + "»");
                string values = null;
                bool formulas = true;
                AuditBReadBook(digest, delegate(dynamic wb)
                {
                    dynamic ws = wb.Sheets[1];
                    values = ws.Range("A1").Value2 + "," + ws.Range("A2").Value2 + "," + ws.Range("A3").Value2;
                    formulas = MergeService.AnyFormula((object)ws.UsedRange.HasFormula);
                });
                AssertEq("10,20,30", values, "весь разлив на месте значениями");
                AssertTrue(!formulas, "формул не осталось");
            }
        }

        // ---------- GR-1: окно докачки миниатюр не больше кэша ----------

        /// <summary>
        /// Тики сетки на LRU заданной ёмкости: видимые [50..60] и окно докачки вокруг них.
        /// Промах — «рендер» (Add). Возвращает число рендеров на третьем тике: у устойчивого
        /// окна оно нулевое, у окна больше кэша — он вытесняет свои же страницы по кругу.
        /// </summary>
        private static int Stage1ThirdTickRenders(int capacity, int windowCapacity, bool touchVisibleLast)
        {
            var cache = new LruCache<string>(capacity, null);
            int renders = 0;
            for (int tick = 1; tick <= 3; tick++)
            {
                if (tick == 3)
                    renders = 0;
                int lo, hi;
                PdfPageGrid.PrefetchWindow(50, 60, 200, 16, windowCapacity, out lo, out hi);
                var order = new List<int>();
                for (int i = 50; i <= 60; i++)
                    order.Add(i);
                for (int d = 1; 50 - d >= lo || 60 + d <= hi; d++)
                {
                    if (50 - d >= lo) order.Add(50 - d);
                    if (60 + d <= hi) order.Add(60 + d);
                }
                string value;
                foreach (int i in order)
                    if (!cache.TryGet("p" + i, out value))
                    {
                        cache.Add("p" + i, "p" + i);
                        renders++;
                    }
                if (touchVisibleLast)
                    for (int i = 50; i <= 60; i++)
                        cache.TryGet("p" + i, out value);
            }
            return renders;
        }

        [Test("Аудит GR-1: окно докачки миниатюр помещается в кэш и не крутится в цикле «отрисовать → вытеснить»")]
        private static void AuditGr1PrefetchWindowFitsCache()
        {
            int lo, hi;
            PdfPageGrid.PrefetchWindow(50, 60, 200, 16, 24, out lo, out hi);
            AssertEq("44,66", lo + "," + hi, "11 видимых + по 6 с каждой стороны + 1 место под добавление = 24");
            PdfPageGrid.PrefetchWindow(0, 10, 200, 16, 24, out lo, out hi);
            AssertEq("0,22", lo + "," + hi, "у начала списка остаток буфера уходит вперёд");
            PdfPageGrid.PrefetchWindow(0, 30, 200, 16, 24, out lo, out hi);
            AssertEq("0,30", lo + "," + hi, "видимое не меньше кэша — без буфера");

            AssertTrue(Stage1ThirdTickRenders(24, 512, false) > 0,
                "проверка предпосылки: прежнее окно (±16) больше кэша на 24 и рендерит по кругу");
            AssertEq(0, Stage1ThirdTickRenders(24, 24, true), "окно по ёмкости устойчиво: повторных рендеров нет");

            var evicted = new List<string>();
            var lru = new LruCache<string>(4, delegate(string v) { evicted.Add(v); });
            foreach (string k in new[] { "a", "b", "c", "d" })
                lru.Add(k, k);
            lru.SetCapacity(2);
            AssertEq("a,b", string.Join(",", evicted.ToArray()), "уменьшение ёмкости вытесняет самые несвежие");
            AssertEq(2, lru.Count, "осталось по новой ёмкости");
            AssertTrue(ThumbZoom.PageCacheCapacity(48L << 20, ThumbZoom.RenderWidthFor(96, ThumbZoom.BaseWidth)) >
                ThumbZoom.PageCacheCapacity(48L << 20, ThumbZoom.RenderWidthFor(96, ThumbZoom.MaxWidth)),
                "на обычном масштабе (x86) кэш держит больше страниц, чем на крупном");
        }

        // ---------- RS-4: слово переноса подсвечивается по строкам ----------

        [Test("Аудит RS-4: слово, склеенное из переноса, подсвечивается рамкой на каждой строке, а не прямоугольником на две")]
        private static void AuditRs4HyphenatedWordBoxesPerLine()
        {
            PdfReviewPage split = ReviewPageFromWords(0,
                W("inter-", 10, 100, 30, 10),
                W("national", 5, 80, 40, 10));
            PdfReviewWord joined = split.Words[0];
            AssertEq("international", joined.Key, "проверка предпосылки: перенос склеен");
            IList<PdfReviewBox> lines = joined.DisplayBoxes;
            AssertEq(2, lines.Count, "две строки — две рамки");
            AssertTrue(BoxSig(lines[0]) == "10;100;40;110" && BoxSig(lines[1]) == "5;80;45;90",
                "рамки строк — как у фрагментов: " + BoxSig(lines[0]) + " | " + BoxSig(lines[1]));
            AssertEq("5;80;45;110", BoxSig(joined.Box), "общая рамка для сравнения не изменилась");
            PdfReviewPage plain = ReviewPageFromWords(0, W("word", 10, 100, 30, 10));
            AssertEq(1, plain.Words[0].DisplayBoxes.Count, "обычное слово — одна рамка");

            var highlight = new PdfReviewHighlight
            {
                ViewWidthPt = 200, ViewHeightPt = 200, Color = Color.Red, EdgeColor = Color.Red,
                ChangeBarSide = PdfReviewChangeBarSide.Left
            };
            foreach (PdfReviewBox box in lines)
            {
                highlight.Boxes.Add(box);
                highlight.Words.Add(joined);
            }
            List<PdfReviewChangeBar> bars = PdfReviewPageView.LayoutChangeBars(highlight, 200, 200, null);
            AssertEq(2, bars.Count, "рейка — у каждой строки своя");
            foreach (PdfReviewChangeBar bar in bars)
                AssertTrue(bar.Band.Height < 20, "полоса рейки высотой в строку, а не в две: " + bar.Band);
        }

        // ---------- RS-3: растровое равенство проверяется в окне знака ----------

        /// <summary>Строка из 30 «знаков»-штрихов в рамке слова; changed — индекс другого знака (−1 — нет).</summary>
        private static Bitmap Stage1GlyphLine(int changed, int shift)
        {
            return ReviewVisualBitmap(delegate(Graphics g)
            {
                for (int i = 0; i < 30; i++)
                {
                    int x = 8 + i * 6 + shift;
                    if (i == changed)
                        g.FillRectangle(Brushes.Black, x - 1, 48, 5, 3); // «—» вместо «|»
                    else
                        g.FillRectangle(Brushes.Black, x, 44, 2, 12);
                }
            });
        }

        [Test("Аудит RS-3: замена одного знака в длинной строке не снимается растровым уточнением")]
        private static void AuditRs3OneGlyphChangeInLongLineStays()
        {
            PdfReviewPage leftPage = ReviewVisualPage(0, ReviewVisualWord("111111111111111111111111111111", 5, 40, 195, 60));
            PdfReviewPage rightPage = ReviewVisualPage(0, ReviewVisualWord("111111111111111-11111111111111", 5, 40, 195, 60));
            bool refined;
            PdfReviewResult result;
            using (Bitmap left = Stage1GlyphLine(-1, 0))
            using (Bitmap right = Stage1GlyphLine(15, 0))
                result = RefineVisualResult(leftPage, rightPage, left, right, out refined);
            AssertTrue(!refined, "другой знак в строке из тридцати — настоящая правка");
            AssertEq(PdfReviewPairStatus.Changed, result.Pairs[0].Status, "правка осталась");

            // Положительный контроль: тот же рисунок, сдвинутый на пиксель, равенство доказывает.
            using (Bitmap left = Stage1GlyphLine(-1, 0))
            using (Bitmap right = Stage1GlyphLine(-1, 1))
                RefineVisualResult(leftPage, rightPage, left, right, out refined);
            AssertTrue(refined, "одинаковый рисунок со сдвигом — равенство, правка снимается");

            // Местная проверка сама по себе.
            var mass = new long[40];
            var miss = new long[40];
            for (int x = 0; x < 40; x++)
                mass[x] = 100;
            AssertTrue(PdfReviewVisualDiff.LocallyEquivalent(mass, miss, 4), "без расхождений — равны");
            miss[20] = 100; // одно окно с четвертью расхождения
            AssertTrue(!PdfReviewVisualDiff.LocallyEquivalent(mass, miss, 4), "расхождение в одном знаке — не равны");
            miss[20] = 10;
            AssertTrue(PdfReviewVisualDiff.LocallyEquivalent(mass, miss, 4), "мелкое сглаживание — равны");
            var sparse = new long[40];
            var noise = new long[40];
            for (int x = 0; x < 40; x++)
                sparse[x] = x < 20 ? 100 : 0;
            noise[30] = 2; // пиксель сглаживания на пустом фоне
            AssertTrue(PdfReviewVisualDiff.LocallyEquivalent(sparse, noise, 4), "пиксель на пустом фоне не проваливает равенство");
            AssertEq(2, PdfReviewVisualDiff.LocalWindow(1), "окно не уже двух столбцов");
            AssertEq(11, PdfReviewVisualDiff.LocalWindow(22), "окно — половина высоты строки");
        }

        [Test("Аудит XL-2: строка файла после повтора начинается с перенесённых листов; лист без имени не повторяется")]
        private static void AuditXl2RetryKeepsMergedSheetsForRows()
        {
            const string a = @"C:\in\a.xlsx", d = @"C:\in\d.xlsx";
            var previous = new MergeResult();
            previous.Files.Add(Stage1Sheet(a, "Янв", true));
            previous.Files.Add(Stage1Sheet(a, "Фев", false));
            previous.Files.Add(Stage1Sheet(d, "Лист1", true));
            previous.Files.Add(Stage1Sheet(d, null, false)); // имя листа не прочиталось
            Dictionary<string, HashSet<string>> plan = MergeRetry.Sheets(previous);
            AssertTrue(plan[d] != null && plan[d].Count == 0,
                "лист без имени в частично перенесённом файле не повторяется — иначе дубли перенесённых");
            AssertEq(a, string.Join("|", MergeRetry.Paths(previous).ToArray()),
                "файл, где повторять нечего, в Excel не открывается");
            var onlyNameless = new MergeResult();
            onlyNameless.Files.Add(Stage1Sheet(d, "Лист1", true));
            onlyNameless.Files.Add(Stage1Sheet(d, null, false));
            AssertEq(0, MergeRetry.Paths(onlyNameless).Count, "повторять нечего — кнопки «Повторить» нет");
            var kept = new List<string>();
            foreach (FileResult fr in MergeRetry.Kept(previous))
                kept.Add(Path.GetFileNameWithoutExtension(fr.FullPath) + ":" + (fr.SourceSheet ?? "?") + (fr.Ok ? "+" : "-"));
            AssertEq("a:Янв+,d:Лист1+,d:?-", string.Join(",", kept.ToArray()),
                "строкам остаются перенесённые листы и неповторяемый пропуск");
            MergeResult combined = MergeRetry.Combine(previous,
                new List<FileResult> { Stage1Sheet(a, "Фев", true) });
            var names = new List<string>();
            foreach (FileResult fr in combined.Files)
                names.Add(Path.GetFileNameWithoutExtension(fr.FullPath) + ":" + (fr.SourceSheet ?? "?") + (fr.Ok ? "+" : "-"));
            AssertEq("a:Янв+,a:Фев+,d:Лист1+,d:?-", string.Join(",", names.ToArray()),
                "неповторённый пропуск в итоге не теряется");
        }

        [Test("Аудит PS-4: предпросмотр узнаётся и через обёртку PrintPreviewControl")]
        private static void AuditPs4PreviewDetection()
        {
            using (var doc = new System.Drawing.Printing.PrintDocument())
            {
                AssertTrue(!PdfPrintService.IsPreview(doc), "обычная печать — не предпросмотр");
                doc.PrintController = new System.Drawing.Printing.PreviewPrintController();
                AssertTrue(PdfPrintService.IsPreview(doc), "PreviewPrintController — предпросмотр");
                // Так подменяет контроллер PrintPreviewControl: обёртка с окном хода печати.
                doc.PrintController = new System.Windows.Forms.PrintControllerWithStatusDialog(
                    new System.Drawing.Printing.PreviewPrintController());
                AssertTrue(PdfPrintService.IsPreview(doc), "обёртка отвечает за вложенный предпросмотр");
                doc.PrintController = new System.Drawing.Printing.StandardPrintController();
                AssertTrue(!PdfPrintService.IsPreview(doc), "StandardPrintController — печать");
            }
        }

        [Test("Аудит SH-10: общий срок обрывает сервер, отдающий тело по байту")]
        private static void AuditSh10DeadlineAbortsTricklingServer()
        {
            var listener = new System.Net.Sockets.TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            int port = ((IPEndPoint)listener.LocalEndpoint).Port;
            const int TrickleBytes = 300;
            int sent = 0;
            var server = new System.Threading.Thread(delegate()
            {
                try
                {
                    using (System.Net.Sockets.TcpClient client = listener.AcceptTcpClient())
                    using (System.Net.Sockets.NetworkStream stream = client.GetStream())
                    {
                        var buffer = new byte[4096];
                        stream.Read(buffer, 0, buffer.Length); // заголовки запроса
                        byte[] head = System.Text.Encoding.ASCII.GetBytes(
                            "HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: 100000\r\n\r\n");
                        stream.Write(head, 0, head.Length);
                        // По байту раз в 200 мс: ни одно чтение не ждёт дольше ReadWriteTimeout,
                        // и только общий срок может закончить запрос.
                        for (int i = 0; i < TrickleBytes; i++)
                        {
                            stream.WriteByte((byte)'x');
                            System.Threading.Interlocked.Increment(ref sent);
                            System.Threading.Thread.Sleep(200);
                        }
                    }
                }
                catch { } // клиент оборвал соединение — так и задумано
            });
            server.IsBackground = true;
            server.Start();
            bool aborted = false;
            int sentWhenReturned;
            try { UpdateChecker.Download("http://127.0.0.1:" + port + "/", "application/json", 1500); }
            catch (WebException) { aborted = true; }
            catch (IOException) { aborted = true; }
            finally
            {
                sentWhenReturned = System.Threading.Volatile.Read(ref sent);
                listener.Stop();
            }
            AssertTrue(aborted, "запрос оборван ошибкой сети");
            // Мера — ход сервера, а не секундомер (TS-5): запрос обязан вернуться, пока сервер ещё
            // отдаёт тело (60 с по байту). Без общего срока он дождался бы последнего байта:
            // ни одно чтение не длится дольше шагового таймаута.
            AssertTrue(sentWhenReturned < TrickleBytes, "запрос вернулся после " + sentWhenReturned +
                " из " + TrickleBytes + " байт: общий срок не оборвал медленный сервер");
            AssertTrue(server.Join(5000), "сервер закончил работу после обрыва соединения");
        }

        // ---------- SH-10: сроки и прокси проверки обновлений ----------

        [Test("Аудит SH-10: запрос обновлений ограничен по времени и на чтении тела, прокси получает учётные данные")]
        private static void AuditSh10UpdateRequestIsBounded()
        {
            var request = (HttpWebRequest)WebRequest.Create("https://example.invalid/"); // сеть не трогается
            UpdateChecker.Configure(request, "application/json");
            AssertEq(UpdateChecker.StepTimeoutMs, request.Timeout, "соединение и заголовки");
            AssertEq(UpdateChecker.StepTimeoutMs, request.ReadWriteTimeout,
                "чтение тела — не 300 с по умолчанию");
            AssertTrue(UpdateChecker.DeadlineMs >= UpdateChecker.StepTimeoutMs, "общий срок не короче шага");
            AssertTrue(request.Proxy == null || request.Proxy.Credentials != null,
                "системный прокси идёт с учётными данными Windows");
            AssertEq("application/json", request.Accept, "Accept");
            AssertTrue(!string.IsNullOrEmpty(request.UserAgent), "User-Agent для GitHub");

            AssertEq("abcdefghij", UpdateChecker.ReadLimited(new StringReader("abcdefghij"), 10), "ровно предел");
            bool threw = false;
            try { UpdateChecker.ReadLimited(new StringReader(new string('x', 20000)), 10000); }
            catch (InvalidDataException) { threw = true; }
            AssertTrue(threw, "ответ больше предела — ошибка, а не чтение без конца");
        }

        // ---------- WP-2: одноцветная небелая страница → фон слайда ----------

        private static void Stage1ValidatePptx(byte[] package, string what)
        {
            string path = Path.Combine(Path.GetTempPath(), "iwo_stage1_" + Guid.NewGuid().ToString("N") + ".pptx");
            try
            {
                File.WriteAllBytes(path, package);
                List<string> problems = PptxValidationProblems(path);
                var consistency = new List<string>();
                AssertOoxmlSelfConsistent(package, consistency);
                problems.AddRange(consistency);
                AssertTrue(problems.Count == 0, what + ": " + string.Join(" | ", problems.ToArray()));
            }
            finally { try { File.Delete(path); } catch { } }
        }

        [Test("Аудит WP-2: одноцветная небелая подложка становится цветом фона слайда")]
        private static void AuditWp2SolidBackgroundKeepsColor()
        {
            using (var navy = new Bitmap(40, 30))
            {
                using (Graphics g = Graphics.FromImage(navy))
                    g.Clear(Color.FromArgb(255, 0x10, 0x20, 0x30));
                int argb;
                AssertTrue(RasterUtil.TrySolidColor(navy, out argb), "сплошной цвет распознан");
                AssertEq(unchecked((int)0xFF102030), argb, "и назван точно");
                navy.SetPixel(39, 29, Color.White);
                AssertTrue(!RasterUtil.TrySolidColor(navy, out argb), "одна точка другого цвета — не сплошной");
                AssertTrue(!RasterUtil.IsSolidColor(navy), "IsSolidColor согласован с TrySolidColor");
            }
            AssertTrue(Background.IsBlank(unchecked((int)0xFFFFFFFF)), "белый — пустота");
            AssertTrue(Background.IsBlank(0x00102030), "прозрачный — пустота");
            AssertTrue(!Background.IsBlank(unchecked((int)0xFF102030)), "тёмно-синий — фон");

            // Слайд размером со страницу и вписанная страница: фон цветом на обоих.
            var pages = new List<PdfPageText> { PptxPage(960, 540), PptxPage(595, 842) };
            foreach (PdfPageText page in pages)
            {
                var par = new OcrParagraph { LeftPt = 40, TopPt = 500, RightPt = 300, BottomPt = 480 };
                par.Runs.Add(new OcrRun { Text = "белый текст", FontSizePt = 12, ColorArgb = 0xFFFFFF, Underline = true });
                page.Paragraphs.Add(par);
            }
            var solid = new Background { SolidArgb = unchecked((int)0xFF102030) };
            OoxmlPackage package = PptxWriter.BuildPackage(pages, PptxStamp, null, null,
                new List<Background> { solid, solid });
            for (int n = 1; n <= 2; n++)
            {
                XDocument slide = PptxSlideXml(package, n);
                XElement color = PptxFirst(slide.Descendants(PptxP + "bg"));
                color = color == null ? null : PptxFirst(color.Descendants(PptxA + "srgbClr"));
                AssertEq("102030", color == null ? null : (string)color.Attribute("val"), "слайд " + n + ": фон цветом страницы");
                AssertEq("sng", (string)PptxFirst(slide.Descendants(PptxA + "rPr")).Attribute("u"),
                    "слайд " + n + ": цвет — не подложка, подчёркивание остаётся");
            }
            byte[] bytes = package.ToArray();
            AssertEq(0, PptxMediaCount(bytes), "цвет не занимает места картинкой");
            Stage1ValidatePptx(bytes, "пакет с цветным фоном");
        }

        // ---------- WP-4: таблица поверх подложки с линовкой ----------

        [Test("Аудит WP-4: таблица слайда поверх подложки — без своей сетки, с настоящими высотами строк и полями")]
        private static void AuditWp4TableOverBackground()
        {
            Func<bool, XDocument> slideOf = delegate(bool withBackground)
            {
                PdfPageText page = PptxPage(600, 800);
                OcrTable table = PptxTable(2, 2, 50, 700, 350, 640);
                table.Rows[0].HeightPt = 10;
                table.Rows[1].HeightPt = 50;
                table.Rows[0].Cells[0].Paragraphs.Add(PptxPar("Ячейка", 54, 698, 150, 692));
                page.Tables.Add(table);
                List<Background> backgrounds = withBackground
                    ? new List<Background> { new Background { Data = PptxTinyPng() } } : null;
                OoxmlPackage package = PptxWriter.BuildPackage(new List<PdfPageText> { page }, PptxStamp, null, null, backgrounds);
                Stage1ValidatePptx(package.ToArray(), withBackground ? "таблица с подложкой" : "таблица без подложки");
                return PptxSlideXml(package, 1);
            };

            XDocument over = slideOf(true);
            foreach (XElement ln in over.Descendants(PptxA + "lnL"))
                AssertTrue(ln.Element(PptxA + "noFill") != null, "поверх линовки подложки своя сетка не рисуется");
            var rows = new List<XElement>(over.Descendants(PptxA + "tr"));
            AssertEq(PptxGeometry.Emu(10).ToString(), (string)rows[0].Attribute("h"), "высота первой строки — по линовке");
            AssertEq(PptxGeometry.Emu(50).ToString(), (string)rows[1].Attribute("h"), "второй — тоже, а не равная доля");
            XElement tcPr = PptxFirst(over.Descendants(PptxA + "tcPr"));
            AssertEq(PptxGeometry.Emu(4).ToString(), (string)tcPr.Attribute("marL"), "левое поле — где стоит текст в ячейке");
            AssertEq(PptxGeometry.Emu(2).ToString(), (string)tcPr.Attribute("marT"), "верхнее — тоже");
            AssertEq("0", (string)tcPr.Attribute("marR"), "правое у текста влево — ноль");
            AssertEq("0", (string)tcPr.Attribute("marB"), "нижнее — ноль");

            XDocument plain = slideOf(false);
            XElement line = PptxFirst(plain.Descendants(PptxA + "lnL"));
            AssertTrue(line != null && line.Element(PptxA + "solidFill") != null, "без подложки линии рисуем сами");

            double left, top;
            PptxWriter.CellInsets(null, 0, 100, 50, 20, out left, out top);
            AssertTrue(left == 0 && top == 0, "пустая ячейка — без полей");
            var far = new OcrTableCell();
            far.Paragraphs.Add(new OcrParagraph { LeftPt = 500, TopPt = -100, Alignment = OcrAlignment.Left });
            PptxWriter.CellInsets(far, 0, 100, 50, 20, out left, out top);
            AssertTrue(left == 25 && top == 10, "поле не больше половины ячейки");
            // «Итого» по центру ячейки 200 pt: поле слева сдвинуло бы центр на 40 pt — его нет.
            var centered = new OcrTableCell();
            centered.Paragraphs.Add(new OcrParagraph { LeftPt = 80, RightPt = 120, TopPt = 98, Alignment = OcrAlignment.Center });
            PptxWriter.CellInsets(centered, 0, 100, 200, 20, out left, out top);
            AssertTrue(left == 0 && top == 2, "по центру — без поля слева: " + left);
        }

        // ---------- WP-5: уровни вложенности списков Word ----------

        private static string Stage1ListSteps(params object[] items)
        {
            var state = new WordDocxWriter.ListState();
            var steps = new List<string>();
            for (int i = 0; i < items.Length; i += 3)
            {
                if (items[i] == null)
                {
                    state.PrevKind = ListKind.None; // обычный абзац между пунктами
                    i -= 2;
                    continue;
                }
                var kind = (ListKind)items[i];
                int number = (int)items[i + 1], level = (int)items[i + 2];
                steps.Add(state.Decide(kind, number, level).ToString());
                state.Record(kind, number, level);
            }
            return string.Join(",", steps.ToArray());
        }

        [Test("Аудит WP-5: продолжение списка учитывает уровень вложенности и номера по уровням")]
        private static void AuditWp5ListLevels()
        {
            ListKind n = ListKind.Numbered, b = ListKind.Bulleted;
            AssertEq("Restart,SetLevel,Inherit,SetLevel",
                Stage1ListSteps(n, 1, 0, n, 1, 1, n, 2, 1, n, 2, 0), "1. 1) 2) 2.: подсписок — уровень глубже, возврат — уровень назад");
            AssertEq("Restart,SetLevel,Inherit,Inherit,SetLevel",
                Stage1ListSteps(n, 1, 0, n, 1, 1, n, 2, 1, n, 3, 1, n, 2, 0), "после «3)» пункт «2.» продолжает верхний уровень, а не начинает новый список");
            AssertEq("Restart,SetLevel,SetLevel",
                Stage1ListSteps(b, 0, 0, b, 0, 1, b, 0, 0), "вложенный маркер получает свой уровень");
            AssertEq("Restart,Continue",
                Stage1ListSteps(n, 1, 0, null, n, 2, 0), "пункт после обычного абзаца продолжает список");
            AssertEq("Restart,Inherit,Inherit,Restart",
                Stage1ListSteps(n, 1, 0, n, 2, 0, n, 3, 0, n, 1, 0), "«1.» после «3.» — новый список");
            AssertEq("Restart,SetLevel,SetLevel,Restart",
                Stage1ListSteps(n, 1, 0, n, 1, 1, n, 2, 0, n, 2, 1), "«2)» после возврата наверх — уровень обнулён, новый подсписок");
        }

        // ---------- WP-5/WP-6 вживую: разметка .docx ----------

        private static OcrParagraph Stage1ListItem(string text, int number, int level, double top)
        {
            var paragraph = new OcrParagraph
            {
                Alignment = OcrAlignment.Left, TopPt = top, BottomPt = top - 14, LeftPt = 56 + level * 20, RightPt = 539,
                ListKind = ListKind.Numbered, ListNumber = number, ListLevel = level,
                ListContentStart = text.IndexOf(' ') + 1
            };
            paragraph.Runs.Add(new OcrRun { Text = text, FontSizePt = 11, FontName = "Arial" });
            return paragraph;
        }

        private static PdfPageText Stage1WordPage(int index, params OcrParagraph[] paragraphs)
        {
            var page = new PdfPageText
            {
                PageIndex = index, WidthPt = 595, HeightPt = 842,
                LeftMarginPt = 56, RightMarginPt = 56, TopMarginPt = 56, BottomMarginPt = 56
            };
            page.Paragraphs.AddRange(paragraphs);
            return page;
        }

        /// <summary>Абзацы document.xml: «ilvl/numId» у пунктов списка, «-» у прочих; «§» — абзац с разрывом раздела.</summary>
        private static List<string> Stage1DocxParagraphs(string docx, out XmlNamespaceManager ns, out XmlDocument xml)
        {
            xml = AuditBDocumentXml(docx);
            ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            var result = new List<string>();
            foreach (XmlNode p in xml.SelectNodes("//w:body/w:p", ns))
            {
                XmlNode ilvl = p.SelectSingleNode("w:pPr/w:numPr/w:ilvl/@w:val", ns);
                XmlNode numId = p.SelectSingleNode("w:pPr/w:numPr/w:numId/@w:val", ns);
                string mark = numId == null ? "-" : (ilvl == null ? "0" : ilvl.Value) + "/" + numId.Value;
                if (p.SelectSingleNode("w:pPr/w:sectPr", ns) != null)
                    mark = "§" + mark;
                result.Add(mark);
            }
            return result;
        }

        [Test("Аудит WP-5: подсписок в .docx — тот же список уровнем глубже (живой Word)")]
        private static void AuditWp5NestedListLive()
        {
            AuditBRequireWord();
            using (var dir = new TempDir("iwo_stage1_wp5_"))
            {
                string path = dir.File("list.docx");
                AuditBOfficeSta(180000, delegate
                {
                    WordDocxWriter.Write(new List<PdfPageText> { Stage1WordPage(0,
                        Stage1ListItem("1. Первый", 1, 0, 780),
                        Stage1ListItem("1) вложенный", 1, 1, 760),
                        Stage1ListItem("2) вложенный", 2, 1, 740),
                        Stage1ListItem("2. Второй", 2, 0, 720)) }, path);
                });
                XmlNamespaceManager ns;
                XmlDocument xml;
                var items = new List<string>();
                foreach (string mark in Stage1DocxParagraphs(path, out ns, out xml))
                    if (mark.TrimStart('§') != "-")
                        items.Add(mark.TrimStart('§'));
                AssertEq(4, items.Count, "четыре пункта: " + string.Join(" ", items.ToArray()));
                var levels = new List<string>();
                foreach (string item in items)
                    levels.Add(item.Substring(0, item.IndexOf('/')));
                AssertEq("0,1,1,0", string.Join(",", levels.ToArray()), "уровни пунктов");
                AssertTrue(items[0].Split('/')[1] == items[3].Split('/')[1],
                    "второй пункт верхнего уровня продолжает тот же список: " + string.Join(" ", items.ToArray()));
            }
        }

        [Test("Аудит WP-6: разрыв раздела после пункта списка не получает номера (живой Word)")]
        private static void AuditWp6SectionBreakAfterListItemLive()
        {
            AuditBRequireWord();
            using (var dir = new TempDir("iwo_stage1_wp6_"))
            {
                string path = dir.File("pages.docx");
                AuditBOfficeSta(180000, delegate
                {
                    WordDocxWriter.Write(new List<PdfPageText>
                    {
                        Stage1WordPage(0, Stage1ListItem("1. пункт", 1, 0, 780)),
                        Stage1WordPage(1, Stage1ListItem("2. пункт", 2, 0, 780))
                    }, path);
                });
                XmlNamespaceManager ns;
                XmlDocument xml;
                List<string> marks = Stage1DocxParagraphs(path, out ns, out xml);
                string all = string.Join(" ", marks.ToArray());
                int numbered = 0, breaks = 0;
                foreach (string mark in marks)
                {
                    if (mark.StartsWith("§", StringComparison.Ordinal))
                    {
                        breaks++;
                        AssertEq("§-", mark, "абзац с разрывом раздела — без номера: " + all);
                    }
                    if (mark.TrimStart('§') != "-")
                        numbered++;
                }
                AssertEq(1, breaks, "один разрыв раздела: " + all);
                AssertEq(2, numbered, "пронумерованы только два пункта: " + all);
            }
        }
    }
}
