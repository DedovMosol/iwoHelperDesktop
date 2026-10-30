using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// XL-3 (audit.md): сторож зависшего COM-вызова. Решения сторожа — на подставном процессе и
    /// времени; правила окон — на дочернем процессе с настоящими окнами; остановка зависшего
    /// Excel и Word — на живом Office, где вызов вешается модальным окном, долгим расчётом и
    /// отменой. Тесты падают на коде до исправления: там вызов висел до сторожа прогона.
    /// </summary>
    internal static partial class UnitTests
    {
        /// <summary>Подставной процесс: окно и разрешение остановки задаёт тест, остановки считаются.</summary>
        private sealed class Xl3FakeProcess : IWatchedProcess
        {
            public volatile string Title;
            public volatile bool AllowStop = true;
            public int Stops;

            public string ModalWindowTitle() { return Title; }

            public bool TryStop()
            {
                Interlocked.Increment(ref Stops);
                return AllowStop;
            }
        }

        private static ComWatchdog.Limits Xl3Limits(int dialogMs, int cancelMs, int stepMs)
        {
            return new ComWatchdog.Limits { PollMs = 20, DialogMs = dialogMs, CancelMs = cancelMs, StepMs = stepMs };
        }

        [Test("Аудит XL-3: сторож решает по отмене, модальному окну и пределу шага")]
        private static void AuditXl3Decide()
        {
            ComWatchdog.Limits limits = Xl3Limits(10000, 5000, 600000);
            AssertEq(WatchdogStop.None, ComWatchdog.Decide(1000, -1, -1, limits), "ничего не случилось");
            AssertEq(WatchdogStop.None, ComWatchdog.Decide(1000, 9999, 4999, limits), "сроки ещё не вышли");
            AssertEq(WatchdogStop.Dialog, ComWatchdog.Decide(10000, 10000, -1, limits), "окно держится весь срок");
            AssertEq(WatchdogStop.Cancel, ComWatchdog.Decide(10000, 10000, 5000, limits), "отмена важнее окна: человек уже решил");
            AssertEq(WatchdogStop.Timeout, ComWatchdog.Decide(600000, -1, -1, limits), "шаг дольше предела");
            AssertEq(WatchdogStop.None, ComWatchdog.Decide(long.MaxValue / 2, -1, -1, Xl3Limits(10000, 5000, 0)),
                "без предела шаг не ограничен");
        }

        [Test("Аудит XL-3: сторож останавливает процесс, только если окно держится весь срок, и один раз")]
        private static void AuditXl3DialogMustPersist()
        {
            var fake = new Xl3FakeProcess { Title = "Вопрос" };
            long now = 0;
            using (var watchdog = new ComWatchdog(fake, null, Xl3Limits(1000, 500, 0), delegate { return now; }, false))
            {
                watchdog.Poll(5000);
                AssertEq(0, fake.Stops, "вне шага сторож молчит, даже если окно есть");

                now = 10000;
                watchdog.Begin(true);
                watchdog.Poll(10000);
                watchdog.Poll(10900);
                fake.Title = null;
                watchdog.Poll(10950); // окно закрылось — отсчёт сначала
                fake.Title = "Вопрос";
                watchdog.Poll(11000);
                watchdog.Poll(11900);
                AssertEq(0, fake.Stops, "окно, закрывшееся до срока, не повод останавливать");

                watchdog.Poll(12000);
                AssertEq(1, fake.Stops, "окно продержалось срок — процесс остановлен");
                AssertEq(WatchdogStop.Dialog, watchdog.End(), "причина — окно");
                AssertEq("Вопрос", watchdog.DialogTitle, "заголовок окна сохранён для сообщения");

                watchdog.Begin(true);
                watchdog.Poll(20000);
                AssertEq(1, fake.Stops, "остановленный процесс не останавливается снова");
                AssertEq(WatchdogStop.Dialog, watchdog.End(), "причина не сбрасывается новым шагом");
            }
        }

        [Test("Аудит XL-3: отмена прерывает только прерываемый шаг и после отсрочки; окно человека — не прерывает")]
        private static void AuditXl3CancelRules()
        {
            var fake = new Xl3FakeProcess();
            bool cancel = true;
            long now = 0;
            using (var watchdog = new ComWatchdog(fake, delegate { return cancel; }, Xl3Limits(1000, 500, 0),
                delegate { return now; }, false))
            {
                watchdog.Begin(false); // запись файла
                watchdog.Poll(0);
                watchdog.Poll(10000);
                AssertEq(0, fake.Stops, "запись отменой не прерывается");
                AssertEq(WatchdogStop.None, watchdog.End(), "запись закончилась сама");

                now = 10000;
                watchdog.Begin(true);
                watchdog.Poll(10000);
                watchdog.Poll(10400);
                AssertEq(0, fake.Stops, "отсрочка: вызов успевает закончиться сам");

                fake.AllowStop = false; // экземпляром завладел человек
                watchdog.Poll(10500);
                AssertEq(1, fake.Stops, "попытка была");
                AssertEq(WatchdogStop.None, watchdog.Stop, "человеческий экземпляр не остановлен");

                fake.AllowStop = true; // человек закрыл свои окна
                watchdog.Poll(10600);
                AssertEq(WatchdogStop.Cancel, watchdog.End(), "отмена прервала шаг");
                AssertTrue(watchdog.Describe("Excel") == null, "об отмене сообщает вызывающий, не сторож");
            }
        }

        [Test("Аудит XL-3: предел длительности отсчитывается от начала шага; законченный шаг не прерывается")]
        private static void AuditXl3StepLimit()
        {
            var fake = new Xl3FakeProcess();
            long now = 0;
            using (var watchdog = new ComWatchdog(fake, null, Xl3Limits(1000, 500, 1000), delegate { return now; }, false))
            {
                watchdog.Begin(true);
                watchdog.Poll(999);
                AssertEq(WatchdogStop.None, watchdog.End(), "первый шаг уложился");
                watchdog.Poll(5000);
                AssertEq(0, fake.Stops, "между шагами сторож не останавливает");

                now = 999;
                watchdog.Begin(true);
                watchdog.Poll(1500);
                AssertEq(0, fake.Stops, "новый шаг считается от своего начала");
                watchdog.Poll(1999);
                AssertEq(WatchdogStop.Timeout, watchdog.End(), "шаг дольше предела");
                AssertEq(string.Format(Loc.T("err.office.timeout"), "Excel", 1), watchdog.Describe("Excel"),
                    "сообщение называет предел в минутах");
            }
        }

        [Test("Аудит XL-3: поток сторожа останавливает зависший процесс и завершается по Dispose; без процесса сторожа нет")]
        private static void AuditXl3WatchdogThread()
        {
            using (var inert = new ComWatchdog(null, delegate { return true; }, Xl3Limits(1, 1, 1)))
            {
                inert.Begin(true);
                Thread.Sleep(50);
                AssertEq(WatchdogStop.None, inert.End(), "процесс не определён — сторож ничего не делает");
            }

            var fake = new Xl3FakeProcess { Title = "" };
            var watchdog = new ComWatchdog(fake, null, Xl3Limits(100, 100, 0));
            try
            {
                watchdog.Begin(true);
                AssertTrue(WaitFor(delegate { return watchdog.Stop != WatchdogStop.None; }, 5000), "сторож сработал сам");
                AssertEq(WatchdogStop.Dialog, watchdog.End(), "причина — окно");
                AssertEq(string.Format(Loc.T("err.office.dialogUntitled"), "Word"), watchdog.Describe("Word"),
                    "окно без заголовка — своя формулировка");
            }
            finally
            {
                watchdog.Dispose();
            }
            int stops = fake.Stops;
            Thread.Sleep(200);
            AssertEq(stops, fake.Stops, "после Dispose опросов нет — поток завершён");
            watchdog.Dispose(); // повторный Dispose безопасен
        }

        [Test("Аудит XL-3: примечание к файлу и признаки остановленного сервера COM")]
        private static void AuditXl3NotesAndServerGone()
        {
            AssertEq("Excel ждал ответа в окне «Открытие документа»",
                MergeService.StalledNote(WatchdogStop.Dialog, "Открытие документа", MergeService.StepLimitMs), "окно с заголовком");
            AssertEq("Excel ждал ответа в диалоговом окне",
                MergeService.StalledNote(WatchdogStop.Dialog, "", MergeService.StepLimitMs), "окно без заголовка");
            AssertEq("Excel не ответил за 10 мин",
                MergeService.StalledNote(WatchdogStop.Timeout, null, MergeService.StepLimitMs), "предел шага");
            AssertEq(1, ComWatchdog.LimitMinutes(1), "меньше минуты — одна");
            AssertEq(1, ComWatchdog.LimitMinutes(60000), "ровно минута");
            AssertEq(2, ComWatchdog.LimitMinutes(60001), "вверх, а не вниз");
            AssertTrue(MergeService.IsServerGone(unchecked((int)0x800706BE)), "RPC_S_CALL_FAILED — так возвращается вызов в остановленный Excel");
            AssertTrue(MergeService.IsServerGone(unchecked((int)0x800706BA)), "RPC_S_SERVER_UNAVAILABLE");
            AssertTrue(MergeService.IsServerGone(unchecked((int)0x80010108)), "RPC_E_DISCONNECTED");
            AssertTrue(!MergeService.IsServerGone(unchecked((int)0x80004005)), "E_FAIL — обычная ошибка файла");
            AssertTrue(!MergeService.IsServerGone(unchecked((int)0x80010001)), "RPC_E_CALL_REJECTED — занят, повтор уместен");
        }

        // ---------- Правила окон: дочерний процесс с настоящими окнами ----------

        /// <summary>
        /// Дочерний процесс (<see cref="WindowChild"/>): рамка и окно поверх неё, как у Office с
        /// модальным окном. Рамка выключена (так её выключает модальное окно); показана — как у
        /// экземпляра человека, или нет — как у скрытого. Окна созданы и показаны до ответа процесса.
        /// </summary>
        private static Process Xl3StartDialogProcess(bool frameVisible, out OwnedOfficeProcess owned)
        {
            long window;
            Process child = WindowChild.Start(
                frameVisible ? WindowChild.DialogVisibleFrame : WindowChild.DialogHiddenFrame, out window);
            try
            {
                owned = OwnedOfficeProcess.FromWindow(window);
                AssertTrue(owned != null && owned.Id == child.Id, "процесс определён по окну");
                return child;
            }
            catch
            {
                WindowChild.Stop(child);
                child.Dispose();
                throw;
            }
        }

        [Test("Аудит XL-3: модальное окно скрытого экземпляра узнаётся по заголовку, процесс останавливается")]
        private static void AuditXl3ModalOfHiddenInstance()
        {
            OwnedOfficeProcess owned;
            using (Process child = Xl3StartDialogProcess(false, out owned))
            {
                try
                {
                    AssertEq(WindowChild.DialogTitle, owned.ModalWindowTitle(), "модальное окно найдено по заголовку");
                    AssertTrue(owned.TryStop(), "окна человека нет — процесс останавливается");
                    AssertTrue(child.WaitForExit(5000), "процесс действительно завершён");
                }
                finally
                {
                    WindowChild.Stop(child);
                }
            }
        }

        [Test("Аудит XL-3: окно поверх показанной рамки — экземпляр человека: не модальное «скрытое» и не останавливается")]
        private static void AuditXl3HumanInstanceSpared()
        {
            OwnedOfficeProcess owned;
            using (Process child = Xl3StartDialogProcess(true, out owned))
            {
                try
                {
                    AssertTrue(owned.ModalWindowTitle() == null, "рамка видна — это окно человека, а не скрытого экземпляра");
                    AssertTrue(!owned.TryStop(), "экземпляр человека не останавливается");
                    AssertTrue(!owned.KillIfLingering(100), "и страховкой после Quit тоже");
                    AssertTrue(!child.HasExited, "процесс жив");
                }
                finally
                {
                    WindowChild.Stop(child);
                }
            }
        }

        // ---------- Живой Office ----------

        /// <summary>
        /// Вызов, повисший в скрытом Office: возвращается ли он, почему остановлен процесс и за
        /// сколько. hang получает экземпляр и вешает его (модальное окно, расчёт).
        /// </summary>
        private static WatchdogStop Xl3HangLive(string progId, int dialogMs, Func<object, OwnedOfficeProcess> own,
            Action<object> prepare, Action<object> hang, out string title, out long elapsedMs)
        {
            WatchdogStop stop = WatchdogStop.None;
            string seen = null;
            long elapsed = 0;
            int pid = 0;
            try
            {
                AuditBOfficeSta(120000, delegate
                {
                    object app = Activator.CreateInstance(Type.GetTypeFromProgID(progId));
                    ComMessageFilter.Register();
                    try
                    {
                        prepare(app);
                        OwnedOfficeProcess owned = own(app);
                        AssertTrue(owned != null, "процесс определён");
                        pid = owned.Id;
                        using (var watchdog = new ComWatchdog(owned, null, Xl3Limits(dialogMs, 500, 0)))
                        {
                            Stopwatch clock = Stopwatch.StartNew();
                            watchdog.Begin(true);
                            Exception error = null;
                            try { hang(app); }
                            catch (Exception ex) { error = ex; }
                            elapsed = clock.ElapsedMilliseconds;
                            stop = watchdog.End();
                            seen = watchdog.DialogTitle;
                            AssertTrue(error != null, "повисший вызов вернулся ошибкой, а не результатом");
                        }
                    }
                    finally
                    {
                        try { ((dynamic)app).Quit(); } catch { } // уже остановлен
                        ComSafe.Release(app);
                        ComSafe.Collect();
                        ComMessageFilter.Revoke();
                    }
                });
            }
            finally
            {
                if (AuditBAlive(pid))
                    using (Process p = Process.GetProcessById(pid))
                        AuditBStop(p);
            }
            AssertTrue(!AuditBAlive(pid), "процесс Office завершён");
            title = seen;
            elapsedMs = elapsed;
            return stop;
        }

        [Test("Аудит XL-3: модальное окно скрытого Excel — процесс остановлен, вызов вернулся (живой Excel)")]
        private static void AuditXl3ExcelDialogLive()
        {
            AuditBRequireExcel();
            string title;
            long elapsed;
            WatchdogStop stop = Xl3HangLive("Excel.Application", 1500,
                delegate(object app) { return OwnedOfficeProcess.FromWindow((long)((dynamic)app).Hwnd); },
                delegate(object app) { ((dynamic)app).Visible = false; ((dynamic)app).DisplayAlerts = false; },
                delegate(object app) { ((dynamic)app).Dialogs[1].Show(); }, // xlDialogOpen: ждёт человека
                out title, out elapsed);
            AssertEq(WatchdogStop.Dialog, stop, "причина — окно");
            AssertTrue(!string.IsNullOrEmpty(title), "заголовок окна прочитан: «" + title + "»");
            AssertTrue(elapsed < 15000, "вызов вернулся вскоре после срока окна: " + elapsed + " мс");
        }

        [Test("Аудит XL-3: модальное окно скрытого Word — процесс остановлен, вызов вернулся (живой Word)")]
        private static void AuditXl3WordDialogLive()
        {
            AuditBRequireWord();
            string title;
            long elapsed;
            WatchdogStop stop = Xl3HangLive("Word.Application", 1500,
                delegate(object app) { return WordCom.OwnWordProcess(app); },
                delegate(object app) { ((dynamic)app).Visible = false; ((dynamic)app).DisplayAlerts = 0; ((dynamic)app).Documents.Add(); },
                delegate(object app) { ((dynamic)app).Dialogs[80].Show(); }, // wdDialogFileOpen
                out title, out elapsed);
            AssertEq(WatchdogStop.Dialog, stop, "причина — окно");
            AssertTrue(!string.IsNullOrEmpty(title), "заголовок окна прочитан: «" + title + "»");
            AssertTrue(elapsed < 15000, "вызов вернулся вскоре после срока окна: " + elapsed + " мс");
        }

        /// <summary>Три книги-источника a, b, c с подписью в A1.</summary>
        private static string[] Xl3Sources(TempDir dir)
        {
            var paths = new[] { dir.File("a.xlsx"), dir.File("b.xlsx"), dir.File("c.xlsx") };
            AuditBWithExcel(delegate(object app)
            {
                foreach (string path in paths)
                {
                    string label = Path.GetFileNameWithoutExtension(path);
                    AuditBSaveBook(app, path, delegate(dynamic ws) { ws.Range("A1").Value2 = label; });
                }
            });
            return paths;
        }

        /// <summary>Свод под наблюдением: новые процессы автоматизации Office после него не остаются.</summary>
        private static void Xl3MergeLive(Action body)
        {
            var before = new HashSet<int>(AuditBOfficeAutomationPids());
            AuditBOfficeSta(240000, body);
            AssertTrue(WaitFor(delegate
            {
                foreach (int pid in AuditBOfficeAutomationPids())
                    if (!before.Contains(pid))
                        return false;
                return true;
            }, 20000), "после работы не осталось своих процессов Office");
        }

        private static ComWatchdog.Limits Xl3MergeLimits(int dialogMs, int cancelMs, int stepMs)
        {
            return new ComWatchdog.Limits { PollMs = 100, DialogMs = dialogMs, CancelMs = cancelMs, StepMs = stepMs };
        }

        [Test("Аудит XL-3: файл, на котором Excel показал окно, пропускается с причиной, остальные сводятся (живой Excel)")]
        private static void AuditXl3MergeSkipsDialogFileLive()
        {
            AuditBRequireExcel();
            using (var dir = new TempDir("iwo_xl3_dialog_"))
            {
                MergeResult result = null;
                int restarts = 0;
                string digest = dir.File("Свод.xlsx");
                Xl3MergeLive(delegate
                {
                    string[] sources = Xl3Sources(dir);
                    var service = new MergeService();
                    service.WatchdogLimits = Xl3MergeLimits(1500, 500, MergeService.StepLimitMs);
                    service.Restarting += delegate { restarts++; };
                    service.StepHookForTests = delegate(object excel, string path)
                    {
                        if (path != null && path.EndsWith("b.xlsx", StringComparison.OrdinalIgnoreCase))
                            ((dynamic)excel).Dialogs[1].Show(); // «трудный» файл: Excel ждёт ответа в окне
                    };
                    result = service.Merge(sources, digest, new MergeOptions());
                });
                AssertEq(1, restarts, "один перезапуск Excel без «трудного» файла");
                AssertEq(3, result.Files.Count, "по записи на файл");
                AssertEq(2, result.OkCount, "a и c сведены");
                FileResult b = result.Files[1];
                AssertTrue(!b.Ok && b.Note != null && b.Note.StartsWith("Excel ждал ответа в окне «", StringComparison.Ordinal),
                    "у b — причина: " + b.Note);
                List<string> sheets = Stage1SheetNames(digest);
                AssertEq("a, c", string.Join(", ", sheets.ToArray()), "листы свода");
            }
        }

        [Test("Аудит XL-3: бесконечно занятый Excel — не «окно», а предел шага; файл пропущен с причиной (живой Excel)")]
        private static void AuditXl3MergeStepLimitLive()
        {
            AuditBRequireExcel();
            using (var dir = new TempDir("iwo_xl3_busy_"))
            {
                MergeResult result = null;
                string digest = dir.File("Свод.xlsx");
                Xl3MergeLive(delegate
                {
                    string[] sources = Xl3Sources(dir);
                    var service = new MergeService();
                    // Окно ждали бы минуту — значит, сработать может только предел шага. Предел
                    // действует на КАЖДЫЙ шаг, не только на «трудный» файл: при 2 с полный прогон
                    // на загруженной машине (x86, 30.10.2026) превысил его на обычном шаге вне
                    // файла, и тест упал гонкой таймера со скоростью Excel. 10 с — с запасом выше
                    // обычных шагов под нагрузкой и всё ещё ниже десятков секунд счёта в хуке.
                    service.WatchdogLimits = Xl3MergeLimits(60000, 500, 10000);
                    service.StepHookForTests = delegate(object excel, string path)
                    {
                        if (path == null || !path.EndsWith("b.xlsx", StringComparison.OrdinalIgnoreCase))
                            return;
                        dynamic app = excel;
                        dynamic book = app.Workbooks.Add(-4167);
                        app.Calculation = -4105; // xlCalculationAutomatic: расчёт идёт в самом вызове
                        // Десятки секунд счёта в одном вызове — как файл, на котором Excel «думает» без конца.
                        book.Sheets[1].Range("A1:A400000").Formula = "=SUMPRODUCT((ROW($A$1:$A$4000)>ROW())*1)";
                    };
                    result = service.Merge(sources, digest, new MergeOptions());
                });
                FileResult b = result.Files[1];
                AssertEq("Excel не ответил за 1 мин", b.Note, "у b — предел шага, а не модальное окно");
                AssertEq(2, result.OkCount, "a и c сведены");
            }
        }

        [Test("Аудит XL-3: отмена прерывает зависший файл, существующий свод не меняется (живой Excel)")]
        private static void AuditXl3CancelHungFileLive()
        {
            AuditBRequireExcel();
            using (var dir = new TempDir("iwo_xl3_cancel_"))
            {
                string digest = dir.File("Свод.xlsx");
                MergeResult retried = null;
                byte[] before = null;
                long elapsed = 0;
                int restarts = 0;
                Xl3MergeLive(delegate
                {
                    string[] sources = Xl3Sources(dir);
                    MergeResult first = new MergeService().Merge(new[] { sources[0] }, digest, new MergeOptions());
                    AssertEq(1, first.OkCount, "свод из a создан");
                    // Как будто b был пропущен: повтор дослиёт его в существующий свод.
                    first.Files.Add(new FileResult { FileName = "b.xlsx", FullPath = sources[1], Note = "тест" });
                    first.SkipCount = 1;
                    before = File.ReadAllBytes(digest);

                    var service = new MergeService();
                    service.WatchdogLimits = Xl3MergeLimits(60000, 1000, MergeService.StepLimitMs); // окно ждали бы минуту
                    service.Restarting += delegate { restarts++; };
                    var hung = new ManualResetEvent(false);
                    service.StepHookForTests = delegate(object excel, string path)
                    {
                        if (path == null)
                            return;
                        hung.Set();
                        ((dynamic)excel).Dialogs[1].Show();
                    };
                    var canceller = new Thread(delegate()
                    {
                        if (hung.WaitOne(60000))
                        {
                            Thread.Sleep(300);
                            service.Cancel(); // как кнопка «Отмена» из окна
                        }
                    });
                    canceller.IsBackground = true;
                    canceller.Start();
                    Stopwatch clock = Stopwatch.StartNew();
                    retried = service.RetrySkipped(digest, new MergeOptions(), first);
                    elapsed = clock.ElapsedMilliseconds;
                    canceller.Join(5000);
                    hung.Close();
                });
                AssertTrue(retried.Cancelled, "повтор отменён");
                AssertEq(0, restarts, "отмена — не зависание: Excel не перезапускается ради следующих файлов");
                AssertTrue(elapsed < 30000, "отмена сработала за отсрочку, а не через минуту окна: " + elapsed + " мс");
                AssertTrue(Xl3SameBytes(before, File.ReadAllBytes(digest)), "существующий свод не изменился");
            }
        }

        private static bool Xl3SameBytes(byte[] a, byte[] b)
        {
            if (a.Length != b.Length)
                return false;
            for (int i = 0; i < a.Length; i++)
                if (a[i] != b[i])
                    return false;
            return true;
        }

        [Test("Аудит XL-3: окно Excel при записи свода — понятная ошибка, файла нет; отмена запись не прерывает (живой Excel)")]
        private static void AuditXl3DialogWhileSavingLive()
        {
            AuditBRequireExcel();
            using (var dir = new TempDir("iwo_xl3_save_"))
            {
                string digest = dir.File("Свод.xlsx");
                Exception error = null;
                Xl3MergeLive(delegate
                {
                    string[] sources = Xl3Sources(dir);
                    var service = new MergeService();
                    service.WatchdogLimits = Xl3MergeLimits(1500, 500, MergeService.StepLimitMs);
                    var saving = new ManualResetEvent(false);
                    service.StepHookForTests = delegate(object excel, string path)
                    {
                        if (path != null)
                            return;
                        saving.Set();
                        ((dynamic)excel).Dialogs[1].Show(); // окно перед самой записью
                    };
                    var canceller = new Thread(delegate()
                    {
                        if (saving.WaitOne(60000))
                            service.Cancel(); // «Отмена» во время записи: отсрочка 0,5 с короче срока окна
                    });
                    canceller.IsBackground = true;
                    canceller.Start();
                    try { service.Merge(sources, digest, new MergeOptions()); }
                    catch (Exception ex) { error = ex; }
                    canceller.Join(5000);
                    saving.Close();
                });
                AssertTrue(error is MergeException, "понятная ошибка, а не сбой RPC: " + (error == null ? "нет" : error.GetType().Name + ": " + error.Message));
                string[] around = string.Format(Loc.T("err.office.dialog"), "Excel", "\u0001").Split('\u0001');
                AssertTrue(error.Message.StartsWith(around[0], StringComparison.Ordinal) &&
                    error.Message.EndsWith(around[1], StringComparison.Ordinal), "сообщение о модальном окне: " + error.Message);
                AssertTrue(!File.Exists(digest), "свод не создан");
            }
        }

        [Test("Аудит XL-3: окно и отмена при записи .docx — остановка Word с причиной или отменой, файла нет (живой Word)")]
        private static void AuditXl3WordWriteLive()
        {
            AuditBRequireWord();
            ComWatchdog.Limits saved = WordCom.WatchdogLimits;
            using (var dir = new TempDir("iwo_xl3_word_"))
            {
                string docx = dir.File("Итог.docx");
                Exception dialogError = null;
                Exception cancelError = null;
                long cancelMs = 0;
                try
                {
                    Xl3MergeLive(delegate
                    {
                        WordCom.WatchdogLimits = Xl3MergeLimits(1500, 500, 0);
                        try
                        {
                            WordCom.WriteDocx(docx, "итог", delegate(object word, object doc)
                            {
                                ((dynamic)word).Dialogs[80].Show(); // wdDialogFileOpen: ждёт человека
                            });
                        }
                        catch (Exception ex) { dialogError = ex; }

                        WordCom.WatchdogLimits = Xl3MergeLimits(60000, 1000, 0); // окно ждали бы минуту
                        var cancel = new ManualResetEvent(false);
                        Stopwatch clock = Stopwatch.StartNew();
                        try
                        {
                            WordCom.WriteDocx(docx, "итог", delegate(object word, object doc)
                            {
                                cancel.Set(); // «Отмена» нажата, а Word завис в вызове
                                ((dynamic)word).Dialogs[80].Show();
                            }, delegate { return cancel.WaitOne(0); });
                        }
                        catch (Exception ex) { cancelError = ex; }
                        cancelMs = clock.ElapsedMilliseconds;
                    });
                }
                finally
                {
                    WordCom.WatchdogLimits = saved;
                }
                AssertTrue(dialogError is MergeException, "окно — понятная ошибка: " + (dialogError == null ? "нет" : dialogError.GetType().Name + ": " + dialogError.Message));
                string[] around = string.Format(Loc.T("err.office.dialog"), "Word", "\u0001").Split('\u0001');
                AssertTrue(dialogError.Message.StartsWith(around[0], StringComparison.Ordinal), "сообщение о модальном окне: " + dialogError.Message);
                AssertTrue(cancelError is OperationCanceledException, "отмена — отмена: " + (cancelError == null ? "нет" : cancelError.GetType().Name + ": " + cancelError.Message));
                AssertTrue(cancelMs < 30000, "отмена за отсрочку, а не через минуту окна: " + cancelMs + " мс");
                AssertTrue(!File.Exists(docx), "файл не создан");
            }
        }
    }
}
