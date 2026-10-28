using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Аудит зоны D: PDF-окна, сетка страниц, печать, план разделения. Живые окна — на STA с
    /// изолированными настройками; буфер обмена не трогаем. Каждый тест падал на коде до правки.
    /// </summary>
    internal static partial class UnitTests
    {
        // ---------- 1. Папка обёрток: что предлагать для сохранения и что удалять ----------

        [Test("Аудит D: папка сохранения — первой настоящей страницы, не обёртки")]
        private static void AuditDFirstUserFolderSkipsWrappers()
        {
            using (var root = new TempDir("iwo_auditd_root_"))
            {
                string wrapper = Path.Combine(root.Path, "iwo_wrap_abc", "blank0000.pdf");
                string user = Path.Combine(root.Path, "Документы", "отчёт.pdf");
                // Папка с тем же префиксом, но НЕ прямо в папке временных — это папка человека.
                string lookalike = Path.Combine(root.Path, "Документы", "iwo_wrap_мой", "скан.pdf");

                AssertEq(Path.GetDirectoryName(user),
                    PdfPageOrderFormBase.FirstUserFolder(new[] { wrapper, user }, root.Path),
                    "пустой лист первым: папка берётся у следующей настоящей страницы");
                AssertEq(Path.GetDirectoryName(user),
                    PdfPageOrderFormBase.FirstUserFolder(new[] { null, "", "пустая страница", user }, root.Path),
                    "null, пустая строка и подпись вместо пути пропускаются");
                AssertEq(null, PdfPageOrderFormBase.FirstUserFolder(new[] { wrapper }, root.Path),
                    "одни обёртки — папку выбирает сам диалог");
                AssertEq(null, PdfPageOrderFormBase.FirstUserFolder(null, root.Path), "null-набор");
                AssertEq(Path.GetDirectoryName(lookalike),
                    PdfPageOrderFormBase.FirstUserFolder(new[] { lookalike }, root.Path),
                    "похожее имя вне папки временных — обычная папка человека");
                AssertTrue(PdfPageOrderFormBase.IsWrapperPath(wrapper, root.Path + "\\"),
                    "завершающий разделитель у корня не мешает узнать обёртку");
                AssertTrue(!PdfPageOrderFormBase.IsWrapperPath("a|b?.pdf", root.Path),
                    "недопустимые символы в пути — не обёртка и не исключение");
            }
        }

        [Test("Аудит D: закрытие окна удаляет свои обёртки, но не чужой файл в их папке")]
        private static void AuditDWrapperDirKeepsForeignFile()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_auditd_wrap_", delegate
            {
                using (var docs = new TempDir("iwo_auditd_docs_"))
                {
                    string doc = docs.File("документ.pdf");
                    PdfProbe.WriteOnePagePdf(doc);
                    string wrapDir = null;
                    var form = new PdfMergeForm();
                    try
                    {
                        form.Show();
                        CallInternal(form, "AddBlankPage");
                        if (!WaitFor(delegate { return !form.IsBusy && FindPageGrid(form).Count == 1; }, 10000))
                        {
                            failures.Add("пустой лист не вставился");
                            return;
                        }
                        wrapDir = form.WrapperDirectory;
                        string[] own = Directory.GetFiles(wrapDir);
                        if (own.Length != 1)
                            failures.Add("ожидалась одна своя обёртка, в папке: " + own.Length);

                        // Первая страница — пустой лист, за ним настоящий документ: диалог
                        // сохранения обязан предложить папку документа, а не временную.
                        var order = form.PageOrder;
                        order.AddDocument(doc, 1);
                        string hint = (string)CallInternal(form, "OutputFolderHint", new object[] { null });
                        if (!string.Equals(hint, docs.Path, StringComparison.OrdinalIgnoreCase))
                            failures.Add("предложена папка «" + hint + "» вместо «" + docs.Path + "»");

                        // Человек всё же сохранил результат в папку обёрток.
                        string foreign = Path.Combine(wrapDir, "итог.pdf");
                        File.WriteAllText(foreign, "результат человека");
                        form.Close();
                        form.Dispose();

                        if (!File.Exists(foreign))
                            failures.Add("закрытие окна удалило чужой файл из папки обёрток");
                        foreach (string file in own)
                            if (File.Exists(file))
                                failures.Add("своя обёртка осталась: " + Path.GetFileName(file));
                    }
                    finally
                    {
                        form.Dispose();
                        if (wrapDir != null)
                            try { Directory.Delete(wrapDir, true); } catch { }
                    }
                }
            });
            AssertTrue(failures.Count == 0, "папка обёрток: " + string.Join(" | ", failures.ToArray()));
        }

        [Test("Аудит D: уборка прошлых сеансов трогает только обёртки и пустые папки")]
        private static void AuditDSweepKeepsForeignFiles()
        {
            using (var root = new TempDir("iwo_auditd_sweep_"))
            {
                DateTime now = DateTime.UtcNow;
                string stale = Path.Combine(root.Path, "iwo_wrap_old");
                string staleEmpty = Path.Combine(root.Path, "iwo_wrap_gone");
                string fresh = Path.Combine(root.Path, "iwo_wrap_live");
                foreach (string dir in new[] { stale, staleEmpty, fresh })
                    Directory.CreateDirectory(dir);
                File.WriteAllText(Path.Combine(stale, "blank0000.pdf"), "x");
                File.WriteAllText(Path.Combine(stale, "img0001.pdf"), "x");
                File.WriteAllText(Path.Combine(stale, "итог.pdf"), "результат человека");
                File.WriteAllText(Path.Combine(staleEmpty, "blank0003.pdf"), "x");
                File.WriteAllText(Path.Combine(fresh, "blank0000.pdf"), "x");
                // Время папки ставим ПОСЛЕ файлов: их создание его обновляет.
                Directory.SetLastWriteTimeUtc(stale, now.AddDays(-2));
                Directory.SetLastWriteTimeUtc(staleEmpty, now.AddDays(-2));

                PdfPageOrderFormBase.SweepOldWrapperDirs(root.Path, now);

                AssertTrue(File.Exists(Path.Combine(stale, "итог.pdf")), "чужой файл в старой папке сохранён");
                AssertTrue(!File.Exists(Path.Combine(stale, "blank0000.pdf")) &&
                    !File.Exists(Path.Combine(stale, "img0001.pdf")), "старые обёртки удалены");
                AssertTrue(!Directory.Exists(staleEmpty), "опустевшая старая папка удалена");
                AssertTrue(File.Exists(Path.Combine(fresh, "blank0000.pdf")), "свежий сеанс не тронут");
                PdfPageOrderFormBase.SweepOldWrapperDirs(Path.Combine(root.Path, "нет такой"), now); // не бросает
            }
        }

        [Test("Аудит D: уборка прошлых сеансов узнаёт обёртки по точному имени, а не по началу")]
        private static void AuditDSweepMatchesExactWrapperNames()
        {
            using (var root = new TempDir("iwo_auditd_sweepnames_"))
            {
                string stale = Path.Combine(root.Path, "iwo_wrap_old");
                Directory.CreateDirectory(stale);
                // Обёртки: нынешние (blank/img + номер, номер может перерасти 4 знака) и картинки
                // «Прочих» до 1.18.5 (img + 32 hex из Guid).
                string[] wrappers = { "blank0003.pdf", "img0002.pdf", "blank12345.pdf",
                    "img" + Guid.NewGuid().ToString("N") + ".pdf" };
                // Файлы человека, сохранённые в ту же папку, — с похожим началом имени.
                string[] users = { "IMG_2024 скан.pdf", "Blank form.pdf", "imgres.pdf", "blank.pdf",
                    "img0002 копия.pdf", "blank0003.pdf.pdf", "img123.pdf" };
                foreach (string name in wrappers)
                    File.WriteAllText(Path.Combine(stale, name), "x");
                foreach (string name in users)
                    File.WriteAllText(Path.Combine(stale, name), "результат человека");
                DateTime now = DateTime.UtcNow;
                Directory.SetLastWriteTimeUtc(stale, now.AddDays(-2));

                PdfPageOrderFormBase.SweepOldWrapperDirs(root.Path, now);

                var lost = new List<string>();
                foreach (string name in users)
                    if (!File.Exists(Path.Combine(stale, name)))
                        lost.Add(name);
                var left = new List<string>();
                foreach (string name in wrappers)
                    if (File.Exists(Path.Combine(stale, name)))
                        left.Add(name);
                AssertTrue(lost.Count == 0, "уборка удалила файлы человека: " + string.Join(", ", lost.ToArray()));
                AssertTrue(left.Count == 0, "обёртки остались: " + string.Join(", ", left.ToArray()));
            }
        }

        // ---------- 2. Перетаскивание плитки из ДРУГОЙ сетки ----------

        [Test("Аудит D: плитка чужой сетки не переставляет страницы")]
        private static void AuditDForeignTileIsNotReorder()
        {
            var failures = new List<string>();
            RunSta(delegate
            {
                using (var host = new Form())
                {
                    var a = new PdfPageGrid { AllowReorder = true, Dock = DockStyle.Left, Width = 300 };
                    var b = new PdfPageGrid { AllowReorder = true, Dock = DockStyle.Right, Width = 300 };
                    host.Controls.Add(a);
                    host.Controls.Add(b);
                    a.SetPages(AuditDFakePages("a.pdf", 3));
                    b.SetPages(AuditDFakePages("b.pdf", 3));
                    host.Show();
                    var calls = new List<string>();
                    a.ReorderRequested += delegate(int from, int to) { calls.Add(from + "→" + to); };
                    var listA = a.ListControl;
                    var listB = b.ListControl;

                    AuditDDropItem(a, listA, listB.Items[2], 1);
                    if (calls.Count != 0)
                        failures.Add("плитка соседней сетки принята как перестановка: " + string.Join(",", calls.ToArray()));

                    AuditDDropItem(a, listA, listA.Items[0], 2);
                    if (calls.Count != 1 || calls[0] != "0→2")
                        failures.Add("своя плитка должна переставляться: " + string.Join(",", calls.ToArray()));

                    AssertTrue(!PdfPageGrid.IsOwnItem(null, listA) && !PdfPageGrid.IsOwnItem(listA.Items[0], null) &&
                        !PdfPageGrid.IsOwnItem(new ListViewItem("x"), listA), "null и элемент вне списка — не свои");
                }
            });
            AssertTrue(failures.Count == 0, "дроп плитки: " + string.Join(" | ", failures.ToArray()));
        }

        private static void AuditDDropItem(PdfPageGrid grid, ListView list, ListViewItem item, int markIndex)
        {
            list.InsertionMark.Index = markIndex;
            list.InsertionMark.AppearsAfterItem = false;
            var e = new DragEventArgs(new DataObject(item), 0, 0, 0, DragDropEffects.Move, DragDropEffects.Move);
            CallInternal(grid, "OnListDragDrop", list, e);
        }

        // ---------- 3. План разделения со вставленными листами ----------

        [Test("Аудит D: разделение по диапазонам — вставленный лист идёт за своей страницей")]
        private static void AuditDSplitRangesKeepInsertedSheet()
        {
            var ranges = new List<PageRange> { new PageRange(0, 1), new PageRange(2, 3) };
            // Лист вставлен после стр. 3: его место — во второй части, а не в первой по «номеру» 0.
            List<PdfPageRef> working = AuditDWorking("s0", "s1", "s2", "blank", "s3");
            List<PdfSplitPart> parts = PdfSplitPlan.ByRanges(working, ranges);
            AssertEq(2, parts.Count, "две части");
            AssertEq("s0,s1", AuditDSig(parts[0].Pages), "первая часть без чужого листа");
            AssertEq("s2,blank,s3", AuditDSig(parts[1].Pages), "лист — вслед за стр. 3");

            // Ведущий лист: в первую непустую часть, а не в часть со стр. 1 (её здесь нет).
            parts = PdfSplitPlan.ByRanges(AuditDWorking("blank", "s0", "s1", "s2", "s3"),
                new List<PageRange> { new PageRange(2, 3) });
            AssertEq(1, parts.Count, "одна часть");
            AssertEq("blank,s2,s3", AuditDSig(parts[0].Pages), "ведущий лист — в начало первой части");

            // Лист за страницей, не вошедшей ни в один диапазон, уходит вместе с ней.
            parts = PdfSplitPlan.ByRanges(AuditDWorking("s0", "blank", "s1", "s2"),
                new List<PageRange> { new PageRange(1, 2) }, AuditDSource);
            AssertEq("s1,s2", AuditDSig(parts[0].Pages), "лист за исключённой стр. 1 не попадает в часть");

            AssertEq(0, PdfSplitPlan.ByRanges(null, ranges).Count, "null-набор");
            AssertEq(0, PdfSplitPlan.ByRanges(working, null).Count, "null-диапазоны");
            AssertEq(AuditDSource, PdfSplitPlan.InferSource(AuditDWorking("blank", "blank", "blank", "s0", "s1")),
                "источник — файл с наибольшим числом разных страниц, а не самый частый лист");
            AssertEq(null, PdfSplitPlan.InferSource(null), "null-набор — источника нет");
        }

        [Test("Аудит D: разделение по закладкам — вставленный лист идёт за своей страницей")]
        private static void AuditDSplitBookmarksKeepInsertedSheet()
        {
            var marks = new List<PdfSplitBookmark>
            {
                new PdfSplitBookmark { PageIndex = 0, Title = "A" },
                new PdfSplitBookmark { PageIndex = 2, Title = "B" }
            };
            List<PdfSplitPart> parts = PdfSplitPlan.ByBookmarks(
                AuditDWorking("s0", "s1", "s2", "blank", "s3"), marks, 4);
            AssertEq(2, parts.Count, "две главы");
            AssertEq("s0,s1", AuditDSig(parts[0].Pages), "глава A без чужого листа");
            AssertEq("s2,blank,s3", AuditDSig(parts[1].Pages), "лист — в главе B, за стр. 3");
            AssertEq("B", parts[1].BookmarkTitle, "название главы сохранено");

            parts = PdfSplitPlan.ByBookmarks(AuditDWorking("blank", "s2", "s3"), marks, 4, AuditDSource);
            AssertEq(1, parts.Count, "пустая глава A пропущена");
            AssertEq("blank,s2,s3", AuditDSig(parts[0].Pages), "ведущий лист — в первую живую главу");
        }

        [Test("Аудит D: источник разделения при равенстве — документ, а не вставленный лист")]
        private static void AuditDInferSourcePrefersRealFile()
        {
            var marks = new List<PdfSplitBookmark>
            {
                new PdfSplitBookmark { PageIndex = 0, Title = "A" },
                new PdfSplitBookmark { PageIndex = 5, Title = "B" }
            };
            // От документа осталась стр. 8, перед ней вставлен пустой лист: по одной странице у обоих.
            List<PdfPageRef> working = AuditDWorking("blank", "s7");
            List<PdfSplitPart> parts = PdfSplitPlan.ByBookmarks(working, marks, 10);
            AssertEq(1, parts.Count, "одна глава");
            AssertEq("B", parts[0].BookmarkTitle, "стр. 8 — в главе B, а не в A");
            AssertEq("blank,s7", AuditDSig(parts[0].Pages), "ведущий лист — вместе с ней");

            AssertEq(AuditDSource, PdfSplitPlan.InferSource(AuditDWorking("s7", "blank", "blank")),
                "скопированный лист чаще встречается, но источником не становится");
            var twoFiles = new List<PdfPageRef>
            {
                new PdfPageRef { SourcePath = @"C:\docs\first.pdf", PageIndex = 0 },
                new PdfPageRef { SourcePath = AuditDSource, PageIndex = 0 }
            };
            AssertEq(@"C:\docs\first.pdf", PdfSplitPlan.InferSource(twoFiles),
                "два настоящих файла поровну — по-прежнему первый встреченный");
        }

        private const string AuditDSource = @"C:\docs\source.pdf";

        /// <summary>Вставленный пустой лист: обёртка в папке окна во временных файлах.</summary>
        private static readonly string AuditDBlankPath =
            Path.Combine(Path.GetTempPath(), "iwo_wrap_x", "blank0000.pdf");

        /// <summary>«sN» — страница N источника, «blank» — вставленный лист (своя обёртка, индекс 0).</summary>
        private static List<PdfPageRef> AuditDWorking(params string[] tokens)
        {
            var pages = new List<PdfPageRef>();
            foreach (string token in tokens)
                pages.Add(token == "blank"
                    ? new PdfPageRef { SourcePath = AuditDBlankPath, PageIndex = 0 }
                    : new PdfPageRef { SourcePath = AuditDSource, PageIndex = int.Parse(token.Substring(1)) });
            return pages;
        }

        private static string AuditDSig(IList<PdfPageRef> pages)
        {
            var parts = new List<string>();
            foreach (PdfPageRef page in pages)
                parts.Add(page.SourcePath == AuditDSource ? "s" + page.PageIndex : "blank");
            return string.Join(",", parts.ToArray());
        }

        // ---------- 4. Печать: неотрисованная страница и отмена ----------

        [Test("Аудит D: печать — неотрисованная страница прерывает задание ошибкой")]
        private static void AuditDPrintFailsOnUnrenderedPage()
        {
            int calls = 0;
            var pages = new List<PdfPageRef> { new PdfPageRef { SourcePath = @"C:\docs\договор.pdf", PageIndex = 4 } };
            using (PrintDocument doc = PdfPrintService.CreateJob(pages, null, null, null,
                delegate { calls++; return null; }, true))
            {
                MergeException error = AssertThrowsExactly<MergeException>("пустой рендер",
                    delegate { AuditDPrintPage(doc); });
                AssertTrue(error.Message.Contains("5") && error.Message.Contains("договор.pdf"),
                    "сообщение называет страницу файла и файл: " + error.Message);
            }
            AssertEq(2, calls, "одна повторная попытка перед отказом");

            // Со второй попытки отрисовалась — печатается как обычно.
            calls = 0;
            using (PrintDocument doc = PdfPrintService.CreateJob(pages, null, null, null,
                delegate(PdfPageRef p, int w) { return ++calls == 1 ? null : AuditDBitmap(40, 60); }, true))
            {
                PrintPageEventArgs e = AuditDPrintPage(doc);
                AssertTrue(!e.HasMorePages, "единственный лист — последний");
            }
            AssertEq(2, calls, "повтор дал страницу");

            // Предпросмотр не строгий: исключение из отрисовки окна уронило бы приложение.
            using (PrintDocument doc = PdfPrintService.CreateJob(pages, null, null, null,
                delegate { return null; }, false))
                AuditDPrintPage(doc);
        }

        [Test("Аудит D: печать — отмена между листами прерывает задание, а не «успех»")]
        private static void AuditDPrintCancelThrows()
        {
            var pages = new List<PdfPageRef>
            {
                new PdfPageRef { SourcePath = @"C:\docs\a.pdf", PageIndex = 0 },
                new PdfPageRef { SourcePath = @"C:\docs\a.pdf", PageIndex = 1 }
            };
            var progress = new List<int>();
            using (PrintDocument doc = PdfPrintService.CreateJob(pages, null,
                delegate(int done, int total) { progress.Add(done); }, delegate { return true; },
                delegate { return AuditDBitmap(40, 60); }, true))
                AssertThrowsExactly<OperationCanceledException>("отмена после первого листа",
                    delegate { AuditDPrintPage(doc); });
            AssertEq("1", string.Join(",", progress.ConvertAll(delegate(int i) { return i.ToString(); }).ToArray()),
                "начатый лист допечатан и учтён");

            // Отмена после ПОСЛЕДНЕГО листа уже ничего не отменяет: задание целиком отдано.
            using (PrintDocument doc = PdfPrintService.CreateJob(pages.GetRange(0, 1), null, null,
                delegate { return true; }, delegate { return AuditDBitmap(40, 60); }, true))
                AuditDPrintPage(doc);
        }

        [Test("Аудит D: печать (живое) — ошибка листа выходит из PrintDocument.Print")]
        private static void AuditDPrintExceptionLeavesPrint()
        {
            var pages = new List<PdfPageRef> { new PdfPageRef { SourcePath = @"C:\docs\a.pdf", PageIndex = 0 } };
            using (PrintDocument doc = PdfPrintService.CreateJob(pages, null, null, null,
                delegate { return null; }, true))
            {
                // Контроллер предпросмотра рисует в метафайл: настоящий цикл печати без задания.
                // Путь — тот же, что у печати: ошибка листа не подменяется «отменой».
                AuditDUseLocalPrinter(doc, new PreviewPrintController());
                AssertThrowsExactly<MergeException>("исключение обработчика листа",
                    delegate { PdfPrintService.RunJob(doc); });
            }
        }

        [Test("Аудит D: печать — в ошибке и в очереди принтера страница названа так, как её знает человек")]
        private static void AuditDPrintNamesPageAsUserSeesIt()
        {
            string blank = AuditDBlankPath;
            string image = Path.Combine(Path.GetDirectoryName(AuditDBlankPath), "img0002.pdf");
            var pages = new List<PdfPageRef>
            {
                new PdfPageRef { SourcePath = blank, PageIndex = 0 },
                new PdfPageRef { SourcePath = image, PageIndex = 0 }
            };
            string caption = Loc.T("ops.blankPage.name");
            var names = new List<string> { caption, @"C:\Фото\отпуск.jpg" };
            using (PrintDocument doc = PdfPrintService.CreateJob(pages, null, null, null,
                delegate { return null; }, true, names))
            {
                AssertEq(caption, doc.DocumentName, "имя задания в очереди принтера");
                MergeException error = AssertThrowsExactly<MergeException>("пустой рендер листа",
                    delegate { AuditDPrintPage(doc); });
                AssertTrue(error.Message.Contains(caption) && !error.Message.Contains("blank0000"),
                    "пустой лист назван подписью, а не именем обёртки: " + error.Message);
            }
            using (PrintDocument doc = PdfPrintService.CreateJob(pages.GetRange(1, 1), null, null, null,
                delegate { return null; }, true, names.GetRange(1, 1)))
            {
                MergeException error = AssertThrowsExactly<MergeException>("пустой рендер картинки",
                    delegate { AuditDPrintPage(doc); });
                AssertTrue(error.Message.Contains("отпуск.jpg") && !error.Message.Contains("img0002"),
                    "картинка названа своим именем: " + error.Message);
            }
            // Без имён — как раньше: имя файла страницы.
            using (PrintDocument doc = PdfPrintService.CreateJob(pages, null, null, null,
                delegate { return null; }, false))
                AssertEq("blank0000.pdf", doc.DocumentName, "имена не переданы — имя файла");
        }

        [Test("Аудит D: вставленный пустой лист для человека — «пустая страница», а не имя обёртки")]
        private static void AuditDBlankWrapperUserPath()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_auditd_userpath_", delegate
            {
                foreach (Form form in new Form[] { new PdfMergeForm(), new PdfSplitForm() })
                    using (form)
                    {
                        string blank = (string)CallInternal(form, "UserPathOf", AuditDBlankPath);
                        if (blank != Loc.T("ops.blankPage.name"))
                            failures.Add(form.GetType().Name + ": пустой лист — «" + blank + "»");
                        if ((string)CallInternal(form, "UserPathOf", AuditDSource) != AuditDSource)
                            failures.Add(form.GetType().Name + ": путь документа изменён");
                        if (CallInternal(form, "UserPathOf", new object[] { null }) != null)
                            failures.Add(form.GetType().Name + ": null — не null");
                    }
            });
            AssertTrue(failures.Count == 0, "имя страницы для человека: " + string.Join(" | ", failures.ToArray()));
        }

        // ---------- 4а. «Отмена» самого контроллера печати ----------

        [Test("Аудит D: печать (живое) — «Отмена» окна состояния печати прерывает задание, а не «успех»")]
        private static void AuditDPrintControllerCancelThrows()
        {
            var controller = new AuditDCancelingController { CancelOnEndPage = true };
            int rendered = 0;
            using (PrintDocument doc = AuditDLiveJob(controller, 3, delegate { rendered++; }))
                AssertThrowsExactly<OperationCanceledException>("«Отмена» после первого листа",
                    delegate { PdfPrintService.RunJob(doc); });
            AssertEq(1, rendered, "после «Отмены» листы больше не рисуются");
            AssertTrue(controller.EndPrintCancel == true,
                "контроллер закрыл задание как отменённое (AbortDoc), получено: " + controller.EndPrintCancel);
        }

        [Test("Аудит D: печать (живое) — отказ в начале задания (StartDoc: ERROR_CANCELLED) — отмена")]
        private static void AuditDPrintStartCanceledThrows()
        {
            // Так ведёт себя StandardPrintController, когда в окне «Сохранить как» у печати в
            // файл (Microsoft Print to PDF) нажали «Отмена»: e.Cancel в OnStartPrint, без исключения.
            int rendered = 0;
            using (PrintDocument doc = AuditDLiveJob(new AuditDCancelingController { CancelOnStartPrint = true }, 2,
                delegate { rendered++; }))
                AssertThrowsExactly<OperationCanceledException>("задание не начато",
                    delegate { PdfPrintService.RunJob(doc); });
            AssertEq(0, rendered, "ни одного листа не нарисовано");
        }

        [Test("Аудит D: печать (живое) — «Отмена», нажатая после отправки задания, не отмена")]
        private static void AuditDPrintCancelAfterEndDocIsSuccess()
        {
            // Окно состояния ставит e.Cancel и ПОСЛЕ того, как внутренний контроллер закрыл
            // задание (EndDoc): такое задание уже в очереди принтера, «Отменено» было бы неправдой.
            var controller = new AuditDCancelingController { CancelAfterEndPrint = true };
            int rendered = 0;
            using (PrintDocument doc = AuditDLiveJob(controller, 2, delegate { rendered++; }))
                PdfPrintService.RunJob(doc);
            AssertEq(2, rendered, "напечатаны оба листа");
            AssertTrue(controller.EndPrintCancel == false, "задание закрыто штатно (EndDoc)");
        }

        /// <summary>
        /// Строгое задание из count листов поверх живого контроллера (метафайл, без очереди).
        /// Нет принтера — Skip: базовому контроллеру не из чего взять параметры листа.
        /// </summary>
        private static PrintDocument AuditDLiveJob(PrintController controller, int count, Action onRender)
        {
            var pages = new List<PdfPageRef>();
            for (int i = 0; i < count; i++)
                pages.Add(new PdfPageRef { SourcePath = @"C:\docs\a.pdf", PageIndex = i });
            PrintDocument doc = PdfPrintService.CreateJob(pages, null, null, null,
                delegate { onRender(); return AuditDBitmap(40, 60); }, true);
            try
            {
                AuditDUseLocalPrinter(doc, controller);
                return doc;
            }
            catch
            {
                doc.Dispose();
                throw;
            }
        }

        /// <summary>Принтер для живых тестов печати (лучше локальный PDF: без бумаги) и контроллер.</summary>
        private static void AuditDUseLocalPrinter(PrintDocument doc, PrintController controller)
        {
            doc.PrinterSettings.PrinterName = RequirePrinter().PrinterName; // общая политика (TS-7)
            doc.PrintController = controller;
        }

        /// <summary>
        /// Имитация окна состояния печати поверх контроллера предпросмотра (рисует в метафайл):
        /// ставит e.Cancel там же, где PrintControllerWithStatusDialog после нажатия «Отмена»,
        /// а StandardPrintController — при ERROR_CANCELLED от StartDoc.
        /// </summary>
        private sealed class AuditDCancelingController : PreviewPrintController
        {
            public bool CancelOnStartPrint, CancelOnEndPage, CancelAfterEndPrint;
            public bool? EndPrintCancel; // с каким e.Cancel контроллер закрывал задание

            public override void OnStartPrint(PrintDocument document, PrintEventArgs e)
            {
                base.OnStartPrint(document, e);
                if (CancelOnStartPrint)
                    e.Cancel = true;
            }

            public override void OnEndPage(PrintDocument document, PrintPageEventArgs e)
            {
                base.OnEndPage(document, e);
                if (CancelOnEndPage)
                    e.Cancel = true;
            }

            public override void OnEndPrint(PrintDocument document, PrintEventArgs e)
            {
                EndPrintCancel = e.Cancel;
                base.OnEndPrint(document, e);
                if (CancelAfterEndPrint)
                    e.Cancel = true;
            }
        }

        private static PrintPageEventArgs AuditDPrintPage(PrintDocument doc)
        {
            using (var sheet = new Bitmap(85, 110))
            using (Graphics g = Graphics.FromImage(sheet))
            {
                var bounds = new Rectangle(0, 0, 85, 110);
                var e = new PrintPageEventArgs(g, bounds, bounds, new PageSettings());
                CallInternal(doc, "OnPrintPage", e);
                return e;
            }
        }

        // ---------- 5. Сбой сжатия — не «уже оптимизирован» ----------

        [Test("Аудит D: сбой движка сжатия в статусе — предупреждение, а не успех")]
        private static void AuditDCompressionFailureStatus()
        {
            string failed = Loc.T("common.suffix.compressFailed");
            string notSmaller = Loc.T("common.suffix.notCompressed");

            string merge = PdfMergeForm.DoneStatus(3, GsRewriteResult.Failed, CompressionLevel.Small);
            AssertTrue(merge.StartsWith("⚠") && merge.Contains(failed) && !merge.Contains(notSmaller),
                "объединение, движок упал: " + merge);
            AssertTrue(PdfMergeForm.DoneStatus(3, GsRewriteResult.RejectedByPolicy, CompressionLevel.Small)
                .StartsWith("✓"), "файл не уменьшился — это всё ещё успех");
            AssertEq(null, PdfToolFormBase.CompressedPart(GsRewriteResult.Failed, CompressionLevel.None),
                "сжатие не заказывали — о нём молчим");

            string extract = PdfSplitForm.DoneStatus(false, 1, 4, 0, 1, CompressionLevel.Good);
            AssertTrue(extract.StartsWith("⚠") && extract.Contains(failed), "извлечение, движок упал: " + extract);

            string allFailed = PdfSplitForm.DoneStatus(true, 5, 10, 0, 5, CompressionLevel.Small);
            AssertTrue(allFailed.StartsWith("⚠") && !allFailed.Contains(notSmaller) &&
                allFailed.Contains(string.Format(Loc.T("split.suffix.compressFailed"), 5)),
                "все части не сжаты из-за сбоя: " + allFailed);
            string mixed = PdfSplitForm.DoneStatus(true, 5, 10, 3, 2, CompressionLevel.Small);
            AssertTrue(mixed.StartsWith("⚠") &&
                mixed.Contains(string.Format(Loc.T("split.suffix.compressFailed"), 2)) &&
                mixed.Contains(string.Format(Loc.T("split.suffix.compressed"), 3, PdfCompression.ImageDpi(CompressionLevel.Small))),
                "часть сжата, часть сорвалась — сказано про обе: " + mixed);
            AssertTrue(PdfSplitForm.DoneStatus(true, 5, 10, 0, 0, CompressionLevel.Small).StartsWith("✓"),
                "без сбоев — галочка");
        }

        // ---------- 6. Запароленный нетронутый документ для движка ----------

        [Test("Аудит D: рабочая копия запароленного документа для движка — расшифрована")]
        private static void AuditDWorkingCopyOfProtectedIsDecrypted()
        {
            using (var dir = new TempDir("iwo_auditd_pwd_"))
            {
                string locked = dir.File("под паролем.pdf");
                string plain = dir.File("открытый.pdf");
                MakeProtectedPdf(locked); // пароль «1234»
                PdfProbe.WriteOnePagePdf(plain);
                PdfPasswords.Clear();
                try
                {
                    PdfPasswords.Remember(locked, "1234"); // человек открыл файл, введя пароль
                    string copy = dir.File("копия.pdf");
                    PdfOpsForm.WriteWorkingCopy(locked, AuditDPristine(locked, 1), 1, copy);
                    AssertEq(1, PdfPageProbe.PageCount(copy),
                        "копия читается без пароля — движку не нужен пароль в командной строке");
                    if (Ghostscript.Available)
                        AssertTrue(PdfCompression.CompressDetailedUnpublished(copy, CompressionLevel.Small) !=
                            GsRewriteResult.Failed, "Ghostscript справился с копией");

                    // Без пароля нетронутый файл по-прежнему копируется байт в байт.
                    string plainCopy = dir.File("копия открытого.pdf");
                    PdfOpsForm.WriteWorkingCopy(plain, AuditDPristine(plain, 1), 1, plainCopy);
                    AssertTrue(Convert.ToBase64String(File.ReadAllBytes(plain)) ==
                        Convert.ToBase64String(File.ReadAllBytes(plainCopy)), "незащищённый — точная копия");
                }
                finally { PdfPasswords.Clear(); }
            }
        }

        [Test("Аудит D: справка «Прочих операций» говорит, что результат из файла под паролем — без пароля")]
        private static void AuditDOpsHelpMentionsPasswordLoss()
        {
            Lang saved = Loc.Current;
            try
            {
                Loc.Init(Lang.Ru);
                string ru = Loc.T("ops.help.body");
                Loc.Init(Lang.En);
                string en = Loc.T("ops.help.body");
                AssertTrue(ru.Contains("без пароля") && ru.Contains("«Сжать…»"),
                    "русская справка молчит о снятии пароля");
                AssertTrue(en.Contains("without the password") && en.Contains("“Compress…”"),
                    "английская справка молчит о снятии пароля");
            }
            finally { Loc.Init(saved); }
        }

        // ---------- 7. Отмена после записи в выбранный файл ----------

        [Test("Аудит D: отмена после записи извлечения не удаляет файл человека")]
        private static void AuditDExtractCancelAfterCommitKeepsFile()
        {
            var failures = new List<string>();
            foreach (CompressionLevel level in new[] { CompressionLevel.None, CompressionLevel.Small })
            {
                CompressionLevel current = level;
                InIsolatedSettings("iwo_auditd_extract_", delegate
                {
                    using (var dir = new TempDir("iwo_auditd_extract_"))
                    using (var form = new PdfSplitForm())
                    {
                        string chosen = dir.File("выбранный.pdf");
                        PdfProbe.WriteOnePagePdf(chosen); // уже ЗАМЕНЁН новым результатом
                        form.Show();
                        FieldInfo flag = typeof(PdfToolFormBase).GetField("_cancelRequested",
                            BindingFlags.Instance | BindingFlags.NonPublic);
                        Func<Action<int, int>, List<string>> work = delegate(Action<int, int> pr)
                        {
                            flag.SetValue(form, true); // «Отмена» нажата сразу после записи
                            return new List<string> { chosen };
                        };
                        // openTarget = null: результат не открываем во внешней программе.
                        CallInternal(form, "RunSplit", work, current, null, false, (Action)delegate { }, 1);
                        if (!WaitFor(delegate { return !form.IsBusy; }, 60000))
                            failures.Add(current + ": операция не завершилась");
                        if (!File.Exists(chosen))
                            failures.Add(current + ": отмена после записи удалила выбранный файл");
                    }
                });
            }
            AssertTrue(failures.Count == 0, "извлечение: " + string.Join(" | ", failures.ToArray()));
        }

        // ---------- 8. Колесо над диалогом поверх предпросмотра ----------

        [Test("Аудит D: сообщение окна — своё только для дерева окон формы")]
        private static void AuditDMessageForOwnWindowTree()
        {
            var failures = new List<string>();
            RunSta(delegate
            {
                using (var form = new Form())
                using (var other = new Form())
                {
                    var panel = new Panel { Dock = DockStyle.Fill };
                    var box = new TextBox();
                    panel.Controls.Add(box);
                    form.Controls.Add(panel);
                    form.Show();
                    other.Show();
                    IntPtr popup = AuditDCreatePopup(form.Handle);
                    try
                    {
                        if (!PagePreviewForm.IsMessageFor(form, form.Handle)) failures.Add("сама форма");
                        if (!PagePreviewForm.IsMessageFor(form, box.Handle)) failures.Add("вложенный контрол");
                        if (PagePreviewForm.IsMessageFor(form, other.Handle)) failures.Add("чужая форма");
                        // Нативный диалог печати — всплывающее окно, ВЛАДЕЛЕЦ которого форма.
                        if (PagePreviewForm.IsMessageFor(form, popup))
                            failures.Add("нативное всплывающее окно поверх формы принято за её");
                        if (PagePreviewForm.IsMessageFor(form, IntPtr.Zero)) failures.Add("нулевой hwnd");
                        if (PagePreviewForm.IsMessageFor(null, box.Handle)) failures.Add("null-форма");
                    }
                    finally { AuditDDestroyWindow(popup); }
                }
            });
            AssertTrue(failures.Count == 0, "адресат сообщения: " + string.Join(" | ", failures.ToArray()));
        }

        [Test("Аудит D: колесо над диалогом поверх предпросмотра не листает страницу (живое)")]
        private static void AuditDPreviewWheelIgnoresDialogOnTop()
        {
            var failures = new List<string>();
            string skip = null;
            InIsolatedSettings("iwo_auditd_wheel_", delegate
            {
                using (var dir = new TempDir("iwo_auditd_wheel_"))
                using (var image = new Bitmap(800, 2000))
                {
                    PagePreviewForm form = AuditDNewPreview(dir.File("нет.pdf"));
                    IntPtr popup = IntPtr.Zero;
                    try
                    {
                        form.Show();
                        AuditDJoinPreviewWorker(form);
                        Application.DoEvents();
                        SetInternal(form, "_image", image);
                        // Курсор — над областью просмотра: иначе фильтр отказался бы и без правки.
                        Point cursor = Cursor.Position;
                        form.Location = new Point(cursor.X - form.Width / 2, cursor.Y - form.Height / 2);
                        var viewport = form.ViewportPanel;
                        Func<bool> cursorOver = delegate
                        {
                            return viewport.RectangleToScreen(viewport.ClientRectangle).Contains(Cursor.Position);
                        };
                        // Нет рабочего стола (RDP без сеанса, служба) или мышь двигают во время
                        // прогона — проверять нечего, это не провал правки.
                        const string noCursor = "курсор не над областью просмотра (нет рабочего стола или мышь двигали) — проверка не состоялась";
                        if (!cursorOver())
                        {
                            skip = noCursor;
                            return;
                        }
                        popup = AuditDCreatePopup(form.Handle);
                        bool dialogWheelTaken = AuditDWheel(form, popup);
                        bool pageWheelTaken = AuditDWheel(form, viewport.Handle);
                        if (!cursorOver())
                        {
                            skip = noCursor;
                            return;
                        }
                        if (dialogWheelTaken)
                            failures.Add("колесо над диалогом поверх окна перехвачено предпросмотром");
                        if (!pageWheelTaken)
                            failures.Add("колесо над самой страницей больше не прокручивает");
                    }
                    finally
                    {
                        AuditDDestroyWindow(popup);
                        SetInternal(form, "_image", null); // растр теста освобождает тест
                        form.Close();
                        form.Dispose();
                    }
                }
            });
            if (skip != null)
                SkipMissing("desktop", skip); // нет интерактивного рабочего стола или курсор увели (TS-7)
            AssertTrue(failures.Count == 0, "колесо в предпросмотре: " + string.Join(" | ", failures.ToArray()));
        }

        private static bool AuditDWheel(Form form, IntPtr hwnd)
        {
            Message m = Message.Create(hwnd, 0x020A, (IntPtr)(120 << 16), IntPtr.Zero);
            return ((IMessageFilter)form).PreFilterMessage(ref m);
        }

        private static PagePreviewForm AuditDNewPreview(string missingPdf)
        {
            var page = new PdfPageRef { SourcePath = missingPdf, PageIndex = 0 };
            return (PagePreviewForm)Activator.CreateInstance(typeof(PagePreviewForm),
                BindingFlags.Instance | BindingFlags.NonPublic, null,
                new object[] { page, "стр. 1", new Size(600, 500), null }, null);
        }

        /// <summary>Дождаться фонового рендера окна (несуществующий файл — быстрый отказ).</summary>
        private static void AuditDJoinPreviewWorker(PagePreviewForm form)
        {
            var worker = form.RenderThread;
            if (worker != null && !worker.Join(15000))
                throw new Exception("фоновый рендер предпросмотра не завершился");
        }

        [DllImport("user32.dll", EntryPoint = "CreateWindowExW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr AuditDCreateWindowEx(int exStyle, string className, string windowName,
            int style, int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);

        [DllImport("user32.dll", EntryPoint = "DestroyWindow")]
        private static extern bool AuditDDestroyWindowCore(IntPtr hwnd);

        /// <summary>Нативное всплывающее окно (WS_POPUP) с владельцем owner — как нативный диалог.</summary>
        private static IntPtr AuditDCreatePopup(IntPtr owner)
        {
            const int WsPopup = unchecked((int)0x80000000);
            IntPtr hwnd = AuditDCreateWindowEx(0, "STATIC", "", WsPopup, 0, 0, 10, 10, owner,
                IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            if (hwnd == IntPtr.Zero)
                throw new Exception("CreateWindowEx: ошибка " + Marshal.GetLastWin32Error());
            return hwnd;
        }

        private static void AuditDDestroyWindow(IntPtr hwnd)
        {
            if (hwnd != IntPtr.Zero)
                AuditDDestroyWindowCore(hwnd);
        }

        // ---------- 9. Картинки из дропа вместе с PDF ----------

        [Test("Аудит D: дроп PDF с картинкой в занятое окно — ничего не откладывается")]
        private static void AuditDBusyDropQueuesNothing()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_auditd_busydrop_", delegate
            {
                using (var dir = new TempDir("iwo_auditd_busydrop_"))
                using (var form = new PdfOpsForm(delegate { }))
                {
                    string first = dir.File("первый.pdf"), second = dir.File("второй.pdf"), png = dir.File("фото.png");
                    PdfProbe.WriteOnePagePdf(first);
                    PdfProbe.WriteOnePagePdf(second);
                    AuditDWritePng(png);
                    form.Show();

                    CallInternal(form, "BeginLoad", "занято"); // окно занято чужой работой
                    CallInternal(form, "AcceptDroppedPaths", new object[] { new[] { first, png } });
                    CallInternal(form, "EndLoad");
                    // Следующая, не связанная загрузка не должна «подобрать» отклонённую картинку.
                    CallInternal(form, "LoadSource", second);
                    if (!WaitFor(delegate { return SourcePathOf(form) == second; }, 10000))
                        failures.Add("второй документ не открылся");
                    AuditDWaitIdle(form);
                    int count = FindPageGrid(form).Count;
                    if (count != 1)
                        failures.Add("в наборе страниц: " + count + " (картинка из отклонённого дропа добавилась)");
                }
            });
            AssertTrue(failures.Count == 0, "дроп в занятое окно: " + string.Join(" | ", failures.ToArray()));
        }

        [Test("Аудит D: отменённая загрузка PDF не добавляет брошенные с ним картинки")]
        private static void AuditDCanceledLoadDropsImages()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_auditd_canceldrop_", delegate
            {
                using (var dir = new TempDir("iwo_auditd_canceldrop_"))
                using (var form = new PdfOpsForm(delegate { }))
                {
                    string pdf = dir.File("документ.pdf"), png = dir.File("фото.png");
                    PdfProbe.WriteOnePagePdf(pdf);
                    AuditDWritePng(png);
                    form.Show();

                    CallInternal(form, "AcceptDroppedPaths", new object[] { new[] { pdf, png } });
                    if (!form.IsBusy)
                        failures.Add("разбор документа не начался");
                    SetInternal(form, "_cancelRequested", true); // «Отмена» во время разбора
                    AuditDWaitIdle(form);
                    int count = FindPageGrid(form).Count;
                    if (count != 0)
                        failures.Add("после отмены в наборе страниц: " + count);
                }
            });
            AssertTrue(failures.Count == 0, "отменённый дроп: " + string.Join(" | ", failures.ToArray()));
        }

        private static void AuditDWritePng(string path)
        {
            using (var bmp = new Bitmap(60, 40))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                    g.Clear(Color.SteelBlue);
                bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
            }
        }

        // ---------- 10. Пустые перестановки и история отмены ----------

        [Test("Аудит D: перестановка «на место» не стирает ветку Ctrl+Y")]
        private static void AuditDNoOpMoveKeepsRedo()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_auditd_noop_", delegate
            {
                using (var form = new PdfMergeForm())
                {
                    form.Show();
                    var order = form.PageOrder;
                    order.AddDocument(@"C:\docs\нет.pdf", 4);
                    CallInternal(form, "RefreshGrid");
                    order.Checkpoint();
                    order.Move(0, 3);
                    order.Undo(); // есть что вернуть по Ctrl+Y

                    var noOps = new List<object[]>
                    {
                        new object[] { "OnReorder", 1, 1 }, new object[] { "OnReorder", 1, 2 },
                        new object[] { "OnReorder", 99, 0 }, new object[] { "OnReorder", -1, 2 },
                        new object[] { "OnMoveRange", new[] { 1, 2 }, 1 }, new object[] { "OnMoveRange", new[] { 1, 2 }, 3 },
                        new object[] { "OnMoveRange", new int[0], 0 }, new object[] { "OnMoveRange", new[] { 7, -3 }, 0 }
                    };
                    foreach (object[] call in noOps)
                    {
                        var args = new object[call.Length - 1];
                        Array.Copy(call, 1, args, 0, args.Length);
                        CallInternal(form, (string)call[0], args);
                        if (!order.CanRedo || order.CanUndo)
                        {
                            failures.Add(call[0] + "(" + string.Join(",", Array.ConvertAll(args, AuditDText)) +
                                ") записал пустой шаг отмены");
                            return;
                        }
                    }
                    CallInternal(form, "OnReorder", 0, 2); // настоящая перестановка — шаг есть
                    if (!order.CanUndo)
                        failures.Add("настоящая перестановка не попала в историю");
                }
            });
            AssertTrue(failures.Count == 0, "пустые перестановки: " + string.Join(" | ", failures.ToArray()));

            AssertTrue(PdfPageOrderFormBase.IsNoOpMove(3, 0, 0) && PdfPageOrderFormBase.IsNoOpMove(3, 2, 3) &&
                PdfPageOrderFormBase.IsNoOpMove(3, 0, int.MaxValue) && PdfPageOrderFormBase.IsNoOpMove(0, 0, 0),
                "на себя, сразу за собой, за край, пустой набор — без правки");
            AssertTrue(!PdfPageOrderFormBase.IsNoOpMove(3, 0, 2) && !PdfPageOrderFormBase.IsNoOpMove(3, 2, 0),
                "перенос вперёд и назад — правка");
            AssertTrue(!PdfPageOrderFormBase.IsNoOpRangeMove(5, new[] { 1, 3 }, 2),
                "несмежные страницы, собранные вместе, — правка");
            AssertTrue(PdfPageOrderFormBase.IsNoOpRangeMove(5, null, 0) &&
                PdfPageOrderFormBase.IsNoOpRangeMove(5, new[] { 3, 4, 4 }, int.MaxValue),
                "null-набор и вставка хвоста в конец — без правки");
        }

        private static string AuditDText(object value)
        {
            var array = value as int[];
            return array == null ? Convert.ToString(value) :
                "[" + string.Join(" ", Array.ConvertAll(array, delegate(int i) { return i.ToString(); })) + "]";
        }

        // ---------- 11. Пополнение бюджета перерисовывает плитки ----------

        [Test("Аудит D: после пополнения бюджета плитки-заглушки перерисовываются (живое)")]
        private static void AuditDMemoryReleaseRepaintsTiles()
        {
            var failures = new List<string>();
            string skip = null;
            RunSta(delegate
            {
                using (var dir = new TempDir("iwo_auditd_tiles_"))
                using (var host = new Form { ClientSize = new Size(500, 400) })
                {
                    string pdf = dir.File("страница.pdf");
                    PdfProbe.WriteOnePagePdf(pdf);
                    var grid = new PdfPageGrid { Dock = DockStyle.Fill };
                    host.Controls.Add(grid);
                    grid.SetPages(AuditDPristine(pdf, 1));
                    host.Show();
                    if (!WaitFor(delegate { return grid.CachedTileCount > 0; }, 15000))
                    {
                        skip = "миниатюра не отрисовалась: системный рендер PDF (WinRT) недоступен";
                        return;
                    }

                    PdfMemoryLease big = null, refused;
                    try
                    {
                        // Плитки сняты (их резерв вернулся в бюджет), а весь бюджет тут же выбрали
                        // другие окна: при отрисовке плитка не составляется — остаётся заглушка.
                        CallInternal(grid, "ClearTiles");
                        if (!PdfMemoryBudget.TryAcquire(PdfMemoryBudget.LimitBytes - PdfMemoryBudget.Used, out big))
                        {
                            failures.Add("не удалось занять весь бюджет");
                            return;
                        }
                        if (PdfMemoryBudget.TryAcquire(1, out refused)) // отказ: кто-то ждёт пополнения
                            refused.Dispose();
                        grid.ListControl.Invalidate();
                        Pump(host, 300);
                        if (grid.CachedTileCount != 0)
                        {
                            failures.Add("плитка составилась без бюджета — проверка не состоялась");
                            return;
                        }
                        big.Dispose(); // память вернули — сетка обязана перерисоваться сама
                        if (!WaitFor(delegate { return grid.CachedTileCount > 0; }, 3000))
                            failures.Add("после пополнения бюджета плитка осталась заглушкой");
                    }
                    finally { if (big != null) big.Dispose(); }
                }
            });
            if (skip != null)
                Skip(skip);
            AssertTrue(failures.Count == 0, "перерисовка после пополнения: " + string.Join(" | ", failures.ToArray()));
        }

        // ---------- 12. Одно событие выделения на серию ----------

        [Test("Аудит D: выделение диапазона — одно SelectionChanged с итоговым выделением")]
        private static void AuditDSelectionChangedOncePerSeries()
        {
            var failures = new List<string>();
            RunSta(delegate
            {
                using (var host = new Form())
                {
                    var grid = new PdfPageGrid { Dock = DockStyle.Fill };
                    host.Controls.Add(grid);
                    grid.SetPages(AuditDFakePages("много.pdf", 40));
                    host.Show();
                    Pump(host, 100);
                    var seen = new List<int>();
                    grid.SelectionChanged += delegate { seen.Add(grid.SelectedCount); };

                    grid.SelectRange(0, 40);
                    WaitFor(delegate { return seen.Count > 0; }, 2000);
                    Pump(host, 150);
                    if (seen.Count != 1 || seen[0] != 40)
                        failures.Add("первая серия: событий " + seen.Count + ", выделено при них: " +
                            string.Join(",", seen.ConvertAll(delegate(int i) { return i.ToString(); }).ToArray()));

                    seen.Clear();
                    grid.SelectIndex(5); // следующая серия — снова ровно одно событие
                    WaitFor(delegate { return seen.Count > 0; }, 2000);
                    Pump(host, 150);
                    if (seen.Count != 1 || seen[0] != 1)
                        failures.Add("вторая серия: событий " + seen.Count);
                }
            });
            AssertTrue(failures.Count == 0, "события выделения: " + string.Join(" | ", failures.ToArray()));
        }

        // ---------- 13. Дроп разбирается после выхода из обработчика OLE ----------

        [Test("Аудит D: брошенный файл открывается после выхода из обработчика дропа")]
        private static void AuditDDropHandledAfterDragDropReturns()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_auditd_latedrop_", delegate
            {
                using (var dir = new TempDir("iwo_auditd_latedrop_"))
                using (var form = new PdfSplitForm())
                {
                    string pdf = dir.File("документ.pdf");
                    PdfProbe.WriteOnePagePdf(pdf);
                    form.Show();
                    form.AcceptFiles(new[] { pdf }); // так зовёт обработчик DragDrop окна и карточки хаба
                    if (form.IsBusy || SourcePathOf(form) != null)
                        failures.Add("дроп разобран прямо в обработчике: вопрос «заменить?» держал бы Проводник");
                    if (!WaitFor(delegate { return SourcePathOf(form) == pdf && !form.IsBusy; }, 10000))
                        failures.Add("отложенный дроп так и не открыл документ");
                }
            });
            AssertTrue(failures.Count == 0, "отложенный дроп: " + string.Join(" | ", failures.ToArray()));
        }

        // ---------- 14. Растры «в пути» при закрытии окна ----------

        [Test("Аудит D: рендер миниатюры, не дошедший до закрытой сетки, возвращает бюджет")]
        private static void AuditDGridInFlightRenderReleased()
        {
            long leaked = 0;
            RunSta(delegate
            {
                BudgetedBitmap page = null;
                var host = new Form();
                try
                {
                    var grid = new PdfPageGrid { Dock = DockStyle.Fill };
                    host.Controls.Add(grid);
                    host.Show();
                    long before = PdfMemoryBudget.Used;
                    page = AuditDBitmap(64, 64);
                    CallInternal(grid, "PostPage", "k", page, 64, 0); // как фоновый рендер
                    host.Dispose(); // окно закрыли раньше, чем UI-поток дошёл до ApplyPage
                    for (int i = 0; i < 5; i++)
                        Application.DoEvents();
                    leaked = PdfMemoryBudget.Used - before;
                }
                finally
                {
                    host.Dispose();
                    if (page != null) page.Dispose();
                }
            });
            AssertEq(0L, leaked, "резерв растра после закрытия сетки (байт)");
        }

        [Test("Аудит D: рендер предпросмотра, не дошедший до закрытого окна, возвращает бюджет")]
        private static void AuditDPreviewInFlightRenderReleased()
        {
            long leaked = 0;
            InIsolatedSettings("iwo_auditd_preview_", delegate
            {
                using (var dir = new TempDir("iwo_auditd_preview_"))
                {
                    BudgetedBitmap page = null;
                    PagePreviewForm form = AuditDNewPreview(dir.File("нет.pdf"));
                    try
                    {
                        form.Show();
                        AuditDJoinPreviewWorker(form);
                        long before = PdfMemoryBudget.Used;
                        page = AuditDBitmap(64, 64);
                        CallInternal(form, "Post", page); // как фоновый рендер
                        form.Close(); // раньше, чем UI-поток дошёл до Apply
                        form.Dispose();
                        for (int i = 0; i < 5; i++)
                            Application.DoEvents();
                        leaked = PdfMemoryBudget.Used - before;
                    }
                    finally
                    {
                        form.Dispose();
                        if (page != null) page.Dispose();
                    }
                }
            });
            AssertEq(0L, leaked, "резерв растра после закрытия предпросмотра (байт)");
        }

        // ---------- 15. Свойства документа читаются в фоне ----------

        [Test("Аудит D: «Свойства» читают документ в фоне и отменяются без диалога")]
        private static void AuditDMetadataReadInBackground()
        {
            var failures = new List<string>();
            InIsolatedSettings("iwo_auditd_meta_", delegate
            {
                using (var dir = new TempDir("iwo_auditd_meta_"))
                using (var form = new PdfOpsForm(delegate { }))
                using (var closer = new Timer { Interval = 50 })
                {
                    string pdf = dir.File("документ.pdf");
                    PdfProbe.WriteOnePagePdf(pdf);
                    form.Show();
                    CallInternal(form, "LoadSource", pdf);
                    if (!WaitFor(delegate { return SourcePathOf(form) == pdf && !form.IsBusy; }, 10000))
                    {
                        failures.Add("документ не открылся");
                        return;
                    }
                    // Страховка: модальный диалог свойств закрывается сам — тест не повиснет.
                    bool dialogShown = false;
                    closer.Tick += delegate
                    {
                        var open = new List<Form>();
                        foreach (Form f in Application.OpenForms)
                            if (f is MetadataForm)
                                open.Add(f);
                        foreach (Form f in open)
                        {
                            dialogShown = true;
                            f.Close();
                        }
                    };
                    closer.Start();

                    CallInternal(form, "EditMetadata");
                    bool busy = form.IsBusy;
                    SetInternal(form, "_cancelRequested", true); // «Отмена», пока свойства читаются
                    AuditDWaitIdle(form);
                    if (!busy)
                        failures.Add("свойства прочитаны на потоке интерфейса — окно не было занято");
                    if (dialogShown)
                        failures.Add("после отмены всё равно открылся диалог свойств");
                }
            });
            AssertTrue(failures.Count == 0, "чтение свойств: " + string.Join(" | ", failures.ToArray()));
        }

        // ---------- общие помощники зоны D ----------

        /// <summary>Нетронутый набор: все страницы файла по порядку.</summary>
        private static List<PdfPageRef> AuditDPristine(string path, int count)
        {
            var pages = new List<PdfPageRef>();
            for (int i = 0; i < count; i++)
                pages.Add(new PdfPageRef { SourcePath = path, PageIndex = i });
            return pages;
        }

        /// <summary>Страницы несуществующего файла: сетке хватает, рендер молча даёт заглушки.</summary>
        private static List<PdfPageRef> AuditDFakePages(string name, int count)
        {
            return AuditDPristine(Path.Combine(Path.GetTempPath(), "iwo_auditd_нет_" + name), count);
        }

        /// <summary>Растр с резервом бюджета — как его отдаёт рендерер.</summary>
        private static BudgetedBitmap AuditDBitmap(int width, int height)
        {
            PdfMemoryLease lease;
            if (!PdfMemoryBudget.TryAcquire(PdfMemoryBudget.EstimateBitmapBytes(width, height), out lease))
                throw new Exception("бюджет памяти не дал " + width + "×" + height);
            return new BudgetedBitmap(new Bitmap(width, height), lease);
        }

        /// <summary>Окно свободно и остаётся свободным (отложенная работа могла начаться следом).</summary>
        private static void AuditDWaitIdle(PdfToolFormBase form)
        {
            for (int round = 0; round < 2; round++)
            {
                if (!WaitFor(delegate { return !form.IsBusy; }, 15000))
                    throw new Exception("окно не освободилось");
                Pump(form, 300);
            }
        }


        /// <summary>Вызвать непубличный метод экземпляра по имени и числу параметров (по цепочке типов).</summary>
    }
}
