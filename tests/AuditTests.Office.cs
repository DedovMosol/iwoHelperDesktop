using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Xml;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Аудит B — Excel/Word COM, свод, повтор, CLI. Чистые части проверяются без Office; поведение
    /// живого COM (значения-ошибки, пароли, защита листов, оглавление, Word) — на установленном
    /// Office, а без него тест честно пропускается.
    /// </summary>
    internal static partial class UnitTests
    {
        private const int AuditBErrNA = -2146826246;   // #Н/Д: xlErrNA (2042) как HRESULT 0x800A07FA
        private const int AuditBErrDiv0 = -2146826281; // #ДЕЛ/0!: xlErrDiv0 (2007)
        private const string AuditBPartialNote = "формулы заменены не полностью";

        // ---------- 1. Значения-ошибки Excel в режиме «формулы → значения» ----------

        [Test("Аудит B: EscapeValues — код ошибки Value2 (Int32) возвращается в Excel как ErrorWrapper")]
        private static void AuditBEscapeValuesWrapsErrorScalar()
        {
            object escaped = CellText.EscapeValues(AuditBErrNA);

            var wrapper = escaped as ErrorWrapper;
            AssertTrue(wrapper != null, "Int32 в Value2 — ошибка Excel, ожидался ErrorWrapper, получено " +
                AuditBDescribe(escaped));
            AssertEq(AuditBErrNA, wrapper.ErrorCode, "код ошибки");
        }

        [Test("Аудит B: EscapeValues — ошибки в массиве области заворачиваются, остальное как было")]
        private static void AuditBEscapeValuesWrapsErrorsInArray()
        {
            var area = (object[,])Array.CreateInstance(typeof(object), new[] { 2, 3 }, new[] { 1, 1 });
            area[1, 1] = AuditBErrNA;
            area[1, 2] = AuditBErrDiv0;
            area[1, 3] = 5.0;
            area[2, 1] = "=x";
            area[2, 2] = true;
            area[2, 3] = null;

            var escaped = (object[,])CellText.EscapeValues(area);

            var na = escaped[1, 1] as ErrorWrapper;
            var div0 = escaped[1, 2] as ErrorWrapper;
            AssertTrue(na != null && na.ErrorCode == AuditBErrNA, "#Н/Д в массиве: " + AuditBDescribe(escaped[1, 1]));
            AssertTrue(div0 != null && div0.ErrorCode == AuditBErrDiv0, "#ДЕЛ/0! в массиве: " + AuditBDescribe(escaped[1, 2]));
            AssertEq(5.0, escaped[1, 3], "число (в Value2 всегда Double)");
            AssertEq("'=x", escaped[2, 1], "строка экранирована");
            AssertEq(true, escaped[2, 2], "булево");
            AssertEq(null, escaped[2, 3], "пустая ячейка");
        }

        [Test("Аудит B: «формулы → значения» оставляет ошибки ошибками, а не числами (живой Excel)")]
        private static void AuditBValuesOnlyKeepsErrorValuesLive()
        {
            AuditBRequireExcel();
            using (var dir = new TempDir("auditB_err_"))
            {
                string source = dir.File("Ошибки.xlsx");
                string output = dir.File("Свод.xlsx");
                var cells = new object[4];
                bool anyFormula = true;
                MergeResult result = null;
                AuditBOfficeSta(180000, delegate
                {
                    AuditBWithExcel(delegate(object app)
                    {
                        AuditBSaveBook(app, source, delegate(dynamic ws)
                        {
                            ws.Range("A1").Formula = "=NA()";
                            ws.Range("A2").Formula = "=1/0";
                            ws.Range("A3").Formula = "=1+1";
                            ws.Range("A4").Formula = "=\"=x\""; // строковый результат, похожий на формулу
                        });
                    });
                    result = new MergeService().Merge(new[] { source }, output, new MergeOptions { ValuesOnly = true });
                    AuditBReadBook(output, delegate(dynamic wb)
                    {
                        dynamic ws = wb.Sheets["Ошибки"];
                        for (int r = 1; r <= 4; r++)
                            cells[r - 1] = ws.Cells[r, 1].Value2;
                        anyFormula = MergeService.AnyFormula((object)ws.UsedRange.HasFormula);
                    });
                });

                AssertTrue(result.Files.Count == 1 && result.Files[0].Ok, "лист перенесён");
                AssertEq<string>(null, result.Files[0].Note, "примечаний нет: замена полная");
                AssertTrue(!anyFormula, "в своде не осталось формул");
                AssertEq(AuditBErrNA, cells[0], "A1 — значение-ошибка #Н/Д (Int32), а не число; получено " + AuditBDescribe(cells[0]));
                AssertEq(AuditBErrDiv0, cells[1], "A2 — значение-ошибка #ДЕЛ/0!; получено " + AuditBDescribe(cells[1]));
                AssertEq(2.0, cells[2], "A3 — число");
                AssertEq("=x", cells[3], "A4 — строка буквально");
            }
        }

        // ---------- 2. Повтор в свод, защищённый паролем ----------

        [Test("Аудит B: сбой открытия свода для дослияния — понятное сообщение (пароль отдельно)")]
        private static void AuditBDigestOpenErrorMessages()
        {
            AssertEq(Loc.T("err.merge.outputPassword"),
                MergeService.DigestOpenError("Неверный пароль. Убедитесь, что отключена клавиша CAPS LOCK."), "русский Office");
            AssertEq(Loc.T("err.merge.outputPassword"),
                MergeService.DigestOpenError("The password you supplied is not correct."), "английский Office");
            AssertEq(string.Format(Loc.T("err.merge.outputOpenFailed"), "Файл повреждён"),
                MergeService.DigestOpenError("  Файл\r\nповреждён "), "прочие сбои — с причиной в одну строку");
            AssertEq(string.Format(Loc.T("err.merge.outputOpenFailed"), ""),
                MergeService.DigestOpenError(null), "сообщения нет");
        }

        [Test("Аудит B: повтор в свод с паролем на открытие — ошибка, а не вечное ожидание (живой Excel)")]
        private static void AuditBRetryIntoOpenPasswordDigestLive()
        {
            AuditBRetryIntoProtectedDigest("secret", null);
        }

        [Test("Аудит B: повтор в свод с паролем на запись — ошибка, а не вечное ожидание (живой Excel)")]
        private static void AuditBRetryIntoWritePasswordDigestLive()
        {
            AuditBRetryIntoProtectedDigest(null, "secret");
        }

        private static void AuditBRetryIntoProtectedDigest(string openPassword, string writePassword)
        {
            AuditBRequireExcel();
            using (var dir = new TempDir("auditB_pwd_"))
            {
                string digest = dir.File("Свод.xlsx");
                string skipped = dir.File("Пропущенный.xlsx");
                Exception error = null;
                AuditBOfficeSta(90000, delegate
                {
                    AuditBWithExcel(delegate(object app)
                    {
                        Action<dynamic> fill = delegate(dynamic ws) { ws.Range("A1").Value2 = "данные"; };
                        AuditBSaveBook(app, skipped, fill);
                        AuditBSaveBook(app, digest, fill, openPassword, writePassword);
                    });
                    var previous = new MergeResult { OutputPath = digest, OkCount = 0, SkipCount = 1 };
                    previous.Files.Add(new FileResult { FileName = "Пропущенный.xlsx", FullPath = skipped, Note = "x" });
                    try { new MergeService().RetrySkipped(digest, new MergeOptions(), previous); }
                    catch (Exception ex) { error = ex; }
                });

                var merge = error as MergeException;
                AssertTrue(merge != null, "ожидалось MergeException, получено " +
                    (error == null ? "успешное дослияние" : error.GetType().Name + ": " + error.Message));
                AssertEq(Loc.T("err.merge.outputPassword"), merge.Message, "сообщение о пароле свода");
            }
        }

        // ---------- 3. Контекст запуска: повтор и записка читают одно и то же ----------

        [Test("Аудит B: упавший запуск не подменяет папку и параметры последнего свода")]
        private static void AuditBFailedRunKeepsDigestContext()
        {
            MergeRun digestRun = AuditBSettledRun("A", true, out MergeResult _);
            var failed = new MergeRun("B", new MergeOptions { AddToc = false }, new DateTime(2026, 2, 2), new List<string>());

            MergeRun last = MergeRun.Settle(digestRun, failed, null);

            AssertTrue(ReferenceEquals(digestRun, last), "после сбоя последним остаётся запуск A целиком");
        }

        [Test("Аудит B: отменённый запуск не подменяет папку и параметры последнего свода")]
        private static void AuditBCancelledRunKeepsDigestContext()
        {
            MergeRun digestRun = AuditBSettledRun("A", true, out MergeResult _);
            var cancelled = new MergeRun("B", new MergeOptions(), new DateTime(2026, 2, 2), new List<string>());

            MergeRun last = MergeRun.Settle(digestRun, cancelled, new MergeResult { Cancelled = true });

            AssertTrue(ReferenceEquals(digestRun, last), "после отмены последним остаётся запуск A целиком");
        }

        [Test("Аудит B: удачный запуск фиксирует свои папку, параметры, время и итог вместе")]
        private static void AuditBSuccessfulRunSettlesWholeContext()
        {
            MergeResult result;
            MergeRun run = AuditBSettledRun("A", true, out result);

            AssertEq("A", run.InputFolder, "папка");
            AssertTrue(run.Options.AddToc, "параметры этого запуска");
            AssertEq(new DateTime(2026, 1, 1, 9, 30, 0), run.StartedAt, "время этого запуска");
            AssertTrue(ReferenceEquals(result, run.Result), "итог этого запуска");
        }

        [Test("Аудит B: повтор берёт папку и параметры запуска, создавшего свод, и только его пропуски")]
        private static void AuditBRetryUsesDigestRunContext()
        {
            MergeResult result;
            MergeRun digestRun = AuditBSettledRun("A", true, out result);

            MergeRun retry = digestRun.ForRetry(new DateTime(2026, 3, 3));

            AssertEq("A", retry.InputFolder, "папка свода");
            AssertTrue(ReferenceEquals(digestRun.Options, retry.Options), "параметры свода (не последнего нажатия)");
            AssertEq(new DateTime(2026, 3, 3), retry.StartedAt, "время повтора");
            AssertEq(@"C:\in\b.xlsx", string.Join("|", retry.Paths), "набор повтора — пропущенные файлы, успешные не трогаются");
            AssertTrue(retry.Result == null, "у идущего повтора итога ещё нет");
            AssertThrowsExactly<InvalidOperationException>("повтор запуска без итога", delegate { retry.ForRetry(DateTime.Now); });
        }

        /// <summary>Запуск «A»: a.xlsx перенесён, b.xlsx пропущен; зафиксирован с итогом.</summary>
        private static MergeRun AuditBSettledRun(string folder, bool addToc, out MergeResult result)
        {
            result = new MergeResult { OutputPath = @"C:\out\Свод.xlsx", OkCount = 1, SkipCount = 1 };
            result.Files.Add(new FileResult { FileName = "a.xlsx", FullPath = @"C:\in\a.xlsx", Ok = true, SheetName = "a" });
            result.Files.Add(new FileResult { FileName = "b.xlsx", FullPath = @"C:\in\b.xlsx", Note = "x" });
            var run = new MergeRun(folder, new MergeOptions { AddToc = addToc }, new DateTime(2026, 1, 1, 9, 30, 0),
                new List<string> { @"C:\in\a.xlsx", @"C:\in\b.xlsx" });
            return MergeRun.Settle(null, run, result);
        }

        // ---------- 4. «Формулы → значения»: несостоявшаяся замена видна в отчёте ----------

        [Test("Аудит B: HasFormula — формулы есть при true и при смешанном (DBNull), нет при false")]
        private static void AuditBAnyFormulaReadsHasFormula()
        {
            AssertTrue(MergeService.AnyFormula(true), "все ячейки — формулы");
            AssertTrue(!MergeService.AnyFormula(false), "формул нет");
            AssertTrue(MergeService.AnyFormula(DBNull.Value), "часть ячеек — формулы (Variant Null)");
        }

        [Test("Аудит B: защищённый лист в режиме «формулы → значения» получает примечание (живой Excel)")]
        private static void AuditBValuesOnlyProtectedSheetNotedLive()
        {
            FileResult fr = AuditBMergeValuesOnly("Защищён.xlsx", delegate(dynamic ws)
            {
                ws.Range("A1").Formula = "=1+1";
                ws.Protect("pw");
            });

            AssertTrue(fr.Ok, "лист перенесён");
            AssertTrue((fr.Note ?? "").Contains(AuditBPartialNote),
                "формулы защищённого листа остались — нужно примечание; получено «" + fr.Note + "»");
        }

        [Test("Аудит B: формула массива, рассечённая областями SpecialCells, даёт примечание (живой Excel)")]
        private static void AuditBValuesOnlySplitArrayNotedLive()
        {
            FileResult fr = AuditBMergeValuesOnly("Массив.xlsx", delegate(dynamic ws)
            {
                ws.Range("A1:B2").FormulaArray = "=ROW(A1:B2)*10+COLUMN(A1:B2)";
                ws.Range("C1").Formula = "=1+1"; // области SpecialCells: A1:C1 и A2:B2 — массив рассечён
            });

            AssertTrue(fr.Ok, "лист перенесён");
            AssertTrue((fr.Note ?? "").Contains(AuditBPartialNote),
                "формулы массива остались — нужно примечание; получено «" + fr.Note + "»");
        }

        private static FileResult AuditBMergeValuesOnly(string fileName, Action<dynamic> fill)
        {
            AuditBRequireExcel();
            using (var dir = new TempDir("auditB_val_"))
            {
                string source = dir.File(fileName);
                MergeResult result = null;
                AuditBOfficeSta(180000, delegate
                {
                    AuditBWithExcel(delegate(object app) { AuditBSaveBook(app, source, fill); });
                    result = new MergeService().Merge(new[] { source }, dir.File("Свод.xlsx"),
                        new MergeOptions { ValuesOnly = true });
                });
                AssertEq(1, result.Files.Count, "одна запись о файле");
                return result.Files[0];
            }
        }

        // ---------- 5. Страховка от осиротевшего процесса Office ----------

        [Test("Аудит B: добивается только свой, живой, тот же и невидимый процесс")]
        private static void AuditBShouldKillOnlyOwnHiddenProcess()
        {
            var start = new DateTime(2026, 1, 1, 12, 0, 0);
            AssertTrue(OwnedOfficeProcess.ShouldKill(true, start, start, false), "свой зависший невидимый — завершить");
            AssertTrue(!OwnedOfficeProcess.ShouldKill(false, start, start, false), "уже вышел — не трогать");
            AssertTrue(!OwnedOfficeProcess.ShouldKill(true, start, start.AddSeconds(1), false),
                "PID занят другим процессом (другое время старта) — не трогать");
            AssertTrue(!OwnedOfficeProcess.ShouldKill(true, start, start, true),
                "окно показано — экземпляром пользуется человек, не трогать");
        }

        [Test("Аудит B: страховка завершает свой процесс, не вышедший за отведённое время")]
        private static void AuditBKillIfLingeringStopsHungProcess()
        {
            using (Process child = AuditBStartPing(60))
            {
                try
                {
                    OwnedOfficeProcess owned = OwnedOfficeProcess.FromProcessId(child.Id);
                    AssertTrue(owned != null, "процесс найден по PID");

                    bool killed = owned.KillIfLingering(200);

                    AssertTrue(killed, "зависший процесс должен быть завершён");
                    AssertTrue(child.WaitForExit(5000), "процесс действительно завершён");
                }
                finally
                {
                    AuditBStop(child);
                }
            }
        }

        [Test("Аудит B: страховка не трогает процесс, вышедший сам")]
        private static void AuditBKillIfLingeringLeavesExitedProcess()
        {
            using (Process child = AuditBStartPing(2)) // ~1 с и выход с кодом 0
            {
                try
                {
                    OwnedOfficeProcess owned = OwnedOfficeProcess.FromProcessId(child.Id);
                    AssertTrue(owned != null, "процесс найден по PID");

                    bool killed = owned.KillIfLingering(15000);

                    AssertTrue(!killed, "вышедший сам процесс не завершается принудительно");
                    AssertTrue(child.WaitForExit(5000), "процесс вышел");
                    AssertEq(0, child.ExitCode, "штатный код выхода, а не код принудительного завершения");
                }
                finally
                {
                    AuditBStop(child);
                }
            }
        }

        [Test("Аудит B: страховка не завершает процесс, у которого видно другое окно (запомненное уже закрыто)")]
        private static void AuditBKillIfLingeringSparesVisibleWindow()
        {
            OwnedOfficeProcess owned;
            using (Process child = AuditBStartTwoWindowProcess(true, out owned))
            {
                try
                {
                    bool killed = owned.KillIfLingering(200);

                    AssertTrue(!killed, "видимое окно — экземпляром пользуется человек, завершать нельзя");
                    AssertTrue(!child.HasExited, "процесс жив");
                }
                finally
                {
                    AuditBStop(child);
                }
            }
        }

        [Test("Аудит B: страховка завершает процесс, все окна которого скрыты (как у невидимого Excel)")]
        private static void AuditBKillIfLingeringStopsHiddenWindowProcess()
        {
            OwnedOfficeProcess owned;
            using (Process child = AuditBStartTwoWindowProcess(false, out owned))
            {
                try
                {
                    bool killed = owned.KillIfLingering(200);

                    AssertTrue(killed, "скрытые окна — экземпляр наш и завис, завершить");
                    AssertTrue(child.WaitForExit(5000), "процесс действительно завершён");
                }
                finally
                {
                    AuditBStop(child);
                }
            }
        }

        [Test("Аудит B: страховка завершает скрытый экземпляр Excel, который не вышел (живой Excel)")]
        private static void AuditBKillIfLingeringStopsHiddenExcelLive()
        {
            bool killed = AuditBKillLingeringExcel(false);

            AssertTrue(killed, "у скрытого Excel нет видимых окон — зависший экземпляр должен быть завершён");
        }

        [Test("Аудит B: страховка не завершает показанный экземпляр Excel (живой Excel)")]
        private static void AuditBKillIfLingeringSparesVisibleExcelLive()
        {
            bool killed = AuditBKillLingeringExcel(true);

            AssertTrue(!killed, "окно Excel показано — экземпляром пользуется человек, завершать нельзя");
        }

        /// <summary>
        /// Excel, который «не вышел»: ссылка на него ещё жива, когда срабатывает страховка. Свой
        /// экземпляр в любом случае не остаётся: Quit, а не вышел — добивается по PID.
        /// </summary>
        private static bool AuditBKillLingeringExcel(bool visible)
        {
            AuditBRequireExcel();
            bool killed = false;
            int pid = 0;
            try
            {
                AuditBOfficeSta(90000, delegate
                {
                    object app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                    try
                    {
                        dynamic excel = app;
                        OwnedOfficeProcess owned = OwnedOfficeProcess.FromWindow((long)excel.Hwnd);
                        AssertTrue(owned != null, "процесс Excel определён по окну");
                        pid = owned.Id;
                        excel.Visible = visible;

                        killed = owned.KillIfLingering(200);
                    }
                    finally
                    {
                        try { ((dynamic)app).Quit(); } catch { } // уже завершён страховкой
                        ComSafe.Release(app);
                        ComSafe.Collect();
                    }
                });
            }
            finally
            {
                if (AuditBAlive(pid))
                    using (Process p = Process.GetProcessById(pid))
                        if (!p.WaitForExit(15000))
                            AuditBStop(p);
            }
            return killed;
        }

        /// <summary>
        /// Дочерний процесс (WinForms в Windows PowerShell) с двумя верхними окнами — порядок как
        /// в MergeService: страховка берётся по ПЕРВОМУ окну (Application.Hwnd), затем оно
        /// уничтожается (target.Close забирает рамку книги-свода), а второе остаётся видимым или
        /// скрытым. Окна — за пределами экрана и без кнопки на панели задач.
        /// </summary>
        private static Process AuditBStartTwoWindowProcess(bool secondVisible, out OwnedOfficeProcess owned)
        {
            long window;
            Process child = WindowChild.Start(secondVisible ? WindowChild.TwoVisible : WindowChild.TwoHidden, out window);
            try
            {
                owned = OwnedOfficeProcess.FromWindow(window);
                AssertTrue(owned != null && owned.Id == child.Id, "процесс определён по окну");
                child.StandardInput.WriteLine("close-first");
                AssertEq("closed", WindowChild.ReadLine(child), "запомненное окно закрыто");
                return child;
            }
            catch
            {
                WindowChild.Stop(child);
                child.Dispose();
                throw;
            }
        }

        [Test("Аудит B: по недействительному окну процесс не определяется — страховки нет, сбоя тоже")]
        private static void AuditBOwnedProcessFromInvalidWindow()
        {
            AssertTrue(OwnedOfficeProcess.FromWindow(0) == null, "нулевое окно");
            AssertTrue(OwnedOfficeProcess.FromWindow(0x7FFFFFF0) == null, "несуществующее окно");
            AssertTrue(OwnedOfficeProcess.FromProcessId(-1) == null, "несуществующий PID");
        }

        [Test("Аудит B: процесс Excel не определён — страховки нет, и это видно в трассе")]
        private static void AuditBOwnExcelProcessFailureTraced()
        {
            var service = new MergeService();
            var traces = new List<string>();
            service.Trace += traces.Add;

            OwnedOfficeProcess fromFailedCall = service.OwnExcelProcess(new object()); // Hwnd бросает, как сбой COM
            OwnedOfficeProcess fromZeroWindow = service.OwnExcelProcess(new AuditBFakeExcel());

            AssertTrue(fromFailedCall == null && fromZeroWindow == null, "страховки нет");
            AssertEq(2, traces.Count, "по строке трассы на каждый случай: " + string.Join(" / ", traces));
        }

        [Test("Аудит B: процесс Excel определён по окну — страховка есть, трасса молчит")]
        private static void AuditBOwnExcelProcessSuccessSilent()
        {
            var traces = new List<string>();
            OwnedOfficeProcess owned = null;
            RunSta(delegate
            {
                using (var form = new System.Windows.Forms.Form()) // окно создаётся, но не показывается
                {
                    var service = new MergeService();
                    service.Trace += traces.Add;
                    owned = service.OwnExcelProcess(new AuditBFakeExcel { Hwnd = form.Handle.ToInt32() });
                }
            });

            using (Process self = Process.GetCurrentProcess())
                AssertTrue(owned != null && owned.Id == self.Id, "процесс определён по окну");
            AssertEq(0, traces.Count, "трасса без жалоб: " + string.Join(" / ", traces));
        }

        /// <summary>Подмена Excel.Application для OwnExcelProcess: только свойство Hwnd.</summary>
        internal sealed class AuditBFakeExcel
        {
            public int Hwnd { get; set; }
        }

        private static Process AuditBStartPing(int count)
        {
            var info = new ProcessStartInfo("ping.exe", "-n " + count + " 127.0.0.1");
            info.UseShellExecute = false;
            info.CreateNoWindow = true;
            info.RedirectStandardOutput = true; // вывод не нужен; буфер ping не переполнит
            return Process.Start(info);
        }

        private static void AuditBStop(Process p)
        {
            try
            {
                if (!p.HasExited)
                {
                    p.Kill();
                    p.WaitForExit(5000);
                }
            }
            catch { } // уже вышел
        }

        private static bool AuditBAlive(int pid)
        {
            if (pid == 0)
                return false;
            try
            {
                using (Process p = Process.GetProcessById(pid))
                    return !p.HasExited;
            }
            catch (ArgumentException)
            {
                return false; // процесса с таким PID уже нет
            }
        }

        // ---------- 6. PDF → Word: занятый выход ----------

        [Test("Аудит B: PDF → Word — занятый выход останавливает работу сразу, до разбора PDF")]
        private static void AuditBPdfToWordLockedOutputFailsUpFront()
        {
            using (var dir = new TempDir("auditB_busy_"))
            {
                string pdf = dir.File("in.pdf");
                WriteTextPdf(pdf, new[] { "Текст страницы" });
                string output = dir.File("out.docx");
                File.WriteAllText(output, "старый результат");
                int progressCalls = 0;
                MergeException error;
                using (new FileStream(output, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    error = AssertThrowsExactly<MergeException>("занятый выход", delegate
                    {
                        PdfToWordService.Convert(AuditBOrder(pdf), output, delegate(int done, int total) { progressCalls++; });
                    });
                }

                AssertEq(string.Format(Loc.T("err.word.fileBusy"), Loc.T("word.label.docx")), error.Message,
                    "сообщение о занятом файле");
                AssertEq(0, progressCalls, "разбор PDF не начинался");
                AssertEq("старый результат", File.ReadAllText(output), "прежний файл не тронут");
            }
        }

        [Test("Аудит B: PDF → Word — выход, занятый во время записи, даёт понятную ошибку (живой Word)")]
        private static void AuditBPdfToWordOutputLockedAtCommitLive()
        {
            AuditBRequireWord();
            using (var dir = new TempDir("auditB_commit_"))
            {
                string pdf = dir.File("in.pdf");
                WriteTextPdf(pdf, new[] { "Текст страницы" });
                string output = dir.File("out.docx");
                FileStream locker = null;
                Exception error = null;
                try
                {
                    AuditBOfficeSta(180000, delegate
                    {
                        try
                        {
                            // Точка невозврата: Word наполнен, дальше сохранение и коммит — файл занимают именно тут.
                            PdfToWordService.Convert(AuditBOrder(pdf), output, null, null, delegate
                            {
                                locker = new FileStream(output, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
                            });
                        }
                        catch (Exception ex) { error = ex; }
                    });
                }
                finally
                {
                    if (locker != null)
                        locker.Dispose();
                }

                AssertTrue(locker != null, "наполнение дошло до точки невозврата");
                var merge = error as MergeException;
                AssertTrue(merge != null, "ожидалось MergeException, получено " +
                    (error == null ? "успех" : error.GetType().Name + ": " + error.Message));
                string prefix = string.Format(Loc.T("err.word.saveFailed"), "out.docx", "");
                AssertTrue(merge.Message.StartsWith(prefix, StringComparison.Ordinal),
                    "сообщение о несохранённом файле; получено «" + merge.Message + "»");
            }
        }

        private static List<PdfPageRef> AuditBOrder(string pdf)
        {
            return new List<PdfPageRef> { new PdfPageRef { SourcePath = pdf, PageIndex = 0 } };
        }

        // ---------- 7. CLI ----------

        [Test("Аудит B: --cli с недопустимым путём итога — код 1, а не падение процесса")]
        private static void AuditBCliInvalidOutputPathReturnsOne()
        {
            using (var dir = new TempDir("auditB_cli_"))
            {
                int code = Program.RunCli(dir.Path, "out|.xlsx", new MergeOptions());

                AssertEq(1, code, "код выхода «ошибка»");
            }
        }

        [Test("Аудит B: отчёт --cli не задваивает строки файлов после перезапуска Excel")]
        private static void AuditBCliReportDropsLinesOfRestartedPass()
        {
            var report = new Program.CliReportLines();
            var a = new FileResult { FileName = "a.xlsx", FullPath = @"C:\in\a.xlsx", Ok = true, SheetName = "a" };
            var b = new FileResult { FileName = "b.xlsx", FullPath = @"C:\in\b.xlsx", Note = "x" };

            report.Trace("open: a.xlsx");
            report.FileDone(a);
            report.FileDone(b);
            report.Trace("перезапуск Excel");
            report.Restarting();
            string lineA = report.FileDone(a);
            report.FileDone(new FileResult { FileName = "b.xlsx", FullPath = @"C:\in\b.xlsx", Note = MergeService.WedgedNote });

            AssertEq(4, report.Lines.Count, "две строки трассировки и по строке на файл: " + string.Join(" / ", report.Lines));
            AssertEq("  trace: open: a.xlsx", report.Lines[0], "трассировка отменённого прохода сохраняется");
            AssertEq("  trace: перезапуск Excel", report.Lines[1], "причина перезапуска сохраняется");
            AssertEq(lineA, report.Lines[2], "файл a — один раз, из нового прохода");
            AssertEq(ReportWriter.FormatFileLine(new FileResult { FileName = "b.xlsx", Note = MergeService.WedgedNote }),
                report.Lines[3], "файл b — итог нового прохода");
        }

        // ---------- 8. Формулировки ----------

        [Test("Аудит B: пропуск из-за зависания Excel — «пропущен:» в оглавлении один раз")]
        private static void AuditBTocWedgedNoteSinglePrefix()
        {
            var fr = new FileResult { FileName = "Завис.xlsx", Note = MergeService.WedgedNote };

            AssertEq("пропущен: вызывал зависание Excel", TocBuilder.NoteCell(fr), "колонка «Примечание»");
            AssertTrue(ReportWriter.FormatFileLine(fr).IndexOf("пропущен", StringComparison.Ordinal) < 0,
                "строка отчёта уже говорит SKIPPED: " + ReportWriter.FormatFileLine(fr));
        }

        [Test("Аудит B: записка считает пропущенные ФАЙЛЫ, а потерянные листы называет отдельно")]
        private static void AuditBNoteCountsSkippedFiles()
        {
            var result = new MergeResult { OutputPath = @"C:\out\Свод.xlsx", OkCount = 1, SkipCount = 2 };
            result.Files.Add(new FileResult { FileName = "a.xlsx", FullPath = @"C:\in\a.xlsx", Ok = true, SheetName = "a · 1" });
            result.Files.Add(new FileResult { FileName = "a.xlsx", FullPath = @"C:\in\a.xlsx", Note = "лист не перенесён (x)" });
            result.Files.Add(new FileResult { FileName = "b.xlsx", FullPath = @"C:\in\b.xlsx", Note = "файл защищён паролем" });

            NoteContent note = NoteText.Build(result, @"C:\in", new MergeOptions { AllSheets = true }, new DateTime(2026, 1, 1));

            string counts = note.Body[1];
            AssertTrue(counts.Contains("Обработано файлов: 2."), counts);
            AssertTrue(counts.Contains("Пропущено файлов: 1."), "b.xlsx не дал ни одного листа, a.xlsx вошёл: " + counts);
            AssertTrue(counts.Contains("Не перенесено листов из частично включённых файлов: 1."), counts);
        }

        [Test("Аудит B: записка без частичных файлов не добавляет фразу о потерянных листах")]
        private static void AuditBNoteWithoutPartialFiles()
        {
            var result = new MergeResult { OutputPath = @"C:\out\Свод.xlsx", OkCount = 1, SkipCount = 1 };
            result.Files.Add(new FileResult { FileName = "a.xlsx", FullPath = @"C:\in\a.xlsx", Ok = true, SheetName = "a" });
            result.Files.Add(new FileResult { FileName = "b.xlsx", FullPath = @"C:\in\b.xlsx", Note = "x" });

            NoteContent note = NoteText.Build(result, @"C:\in", new MergeOptions(), new DateTime(2026, 1, 1));

            AssertEq("Обработано файлов: 2. Включено листов в сводный файл: 1. Пропущено файлов: 1.", note.Body[1], "счётчики");
        }

        [Test("Аудит B: OLE2 под .xlsx — пароль только при потоке EncryptedPackage")]
        private static void AuditBOle2MismatchNote()
        {
            AssertEq("файл защищён паролем", MergeService.Ole2MismatchNote(true), "шифрованная книга");
            string other = MergeService.Ole2MismatchNote(false);
            AssertTrue(other.IndexOf("парол", StringComparison.OrdinalIgnoreCase) < 0, "без потока о пароле ни слова: " + other);
            AssertTrue(other.Contains(".xls"), "подсказка о переименованной .xls: " + other);
        }

        [Test("Аудит B: CFB — поток EncryptedPackage в каталоге находится, Workbook (.xls) — нет")]
        private static void AuditBCfbEncryptedPackageInFirstSector()
        {
            AssertTrue(AuditBHasPackage(AuditBCfb("\u0006DataSpaces", "EncryptedPackage", "EncryptionInfo")), "шифрованная книга");
            AssertTrue(!AuditBHasPackage(AuditBCfb("Workbook", "\u0005SummaryInformation")), "книга .xls");
            AssertTrue(!AuditBHasPackage(AuditBCfb("EncryptedPackageX")), "похожее имя — не тот поток");
        }

        [Test("Аудит B: CFB — запись во втором секторе каталога находится по цепочке FAT")]
        private static void AuditBCfbEncryptedPackageInChainedSector()
        {
            // 4 записи на сектор 512 байт: шестая запись (индекс 5) — во втором секторе каталога.
            byte[] cfb = AuditBCfb("a", "b", "c", "d", "EncryptedPackage");

            AssertTrue(AuditBHasPackage(cfb), "второй сектор каталога прочитан");
        }

        [Test("Аудит B: CFB — сектор FAT за пределами заголовка находится через DIFAT")]
        private static void AuditBCfbEncryptedPackageViaDifat()
        {
            // Сектора каталога 14000 → 14001: FAT для них — 110-й (индекс 109), его номер только в DIFAT.
            const int S = 512, Dir = 14000, Fat = 13998, Difat = 13999;
            var data = new byte[S * (Dir + 3)];
            AuditBCfbHeader(data, Dir, Difat, -1);
            AuditBPut32(data, AuditBSector(Difat), Fat);                    // DIFAT[0] → сектор FAT №109
            for (int i = 1; i < S / 4; i++)
                AuditBPut32(data, AuditBSector(Difat) + i * 4, -1);
            AuditBPut32(data, AuditBSector(Fat) + (Dir % 128) * 4, Dir + 1); // цепочка каталога
            AuditBPut32(data, AuditBSector(Fat) + ((Dir + 1) % 128) * 4, -2);
            AuditBCfbEntry(data, AuditBSector(Dir), "Root Entry", 5);
            for (int i = 1; i < 4; i++)
                AuditBCfbEntry(data, AuditBSector(Dir) + i * 128, "s" + i, 2);
            AuditBCfbEntry(data, AuditBSector(Dir + 1), "EncryptedPackage", 2);

            AssertTrue(AuditBHasPackage(data), "второй сектор каталога найден через DIFAT");
        }

        [Test("Аудит B: CFB — битый, обрезанный и зацикленный каталог дают «нет», а не сбой")]
        private static void AuditBCfbBrokenDirectory()
        {
            AssertTrue(!AuditBHasPackage(new byte[0]), "пустой файл");
            AssertTrue(!AuditBHasPackage(Encoding.ASCII.GetBytes("PK\u0003\u0004 not a cfb")), "не CFB");
            byte[] cfb = AuditBCfb("a", "b", "c", "d", "EncryptedPackage");
            var truncated = new byte[3 * 512]; // заголовок, FAT и первый сектор каталога
            Array.Copy(cfb, truncated, truncated.Length);
            AssertTrue(!AuditBHasPackage(truncated), "второй сектор каталога обрезан");
            byte[] cyclic = AuditBCfb("a", "b", "c");
            AuditBPut32(cyclic, AuditBSector(0) + 1 * 4, 1); // сектор каталога ссылается сам на себя
            AssertTrue(!AuditBHasPackage(cyclic), "зацикленная цепочка завершается");
        }

        [Test("Аудит B: CFB — враждебный файл 64 МБ (самоцикл каталога и DIFAT) разбирается быстрее секунды")]
        private static void AuditBCfbHostileCycleIsFast()
        {
            const int S = 512;
            const long size = 64L * 1024 * 1024; // нули: в ZIP такой файл сжимается до сотен КБ
            int last = (int)(size / S) - 2;      // последний сектор файла (перед ним заголовок)
            using (var dir = new TempDir("auditB_cfb_"))
            {
                string path = dir.File("Враждебный.xlsx");
                using (var fs = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
                {
                    fs.SetLength(size);
                    var header = new byte[S];
                    AuditBCfbHeader(header, last, last, -1); // каталог и DIFAT — в последнем секторе
                    fs.Write(header, 0, S);
                    // Каждый int последнего сектора — его же номер: он сразу каталог без нужной
                    // записи, сектор DIFAT со ссылкой на себя и сектор FAT, где следующий сектор
                    // каталога — снова он. Сектор FAT №1023 — дальше 109 записей заголовка.
                    var sector = new byte[S];
                    for (int i = 0; i < S / 4; i++)
                        AuditBPut32(sector, i * 4, last);
                    fs.Position = AuditBSector(last);
                    fs.Write(sector, 0, S);
                }

                var watch = Stopwatch.StartNew();
                bool found = FileSignature.HasEncryptedPackage(path);
                watch.Stop();

                AssertTrue(!found, "в зацикленном каталоге потока нет");
                // Без потолка зацикленная цепочка не кончается вовсе; с потолком разбор — миллисекунды.
                // 5 с — запас на загруженный раннер CI, а не мера скорости (TS-5).
                AssertTrue(watch.ElapsedMilliseconds < 5000,
                    "разбор занял " + watch.ElapsedMilliseconds + " мс: цепочка каталога не ограничена потолком");
            }
        }

        [Test("Аудит B: CFB — поток в последнем секторе каталога под потолком находится, за потолком — «не знаем»")]
        private static void AuditBCfbDirectoryChainCap()
        {
            int cap = FileSignature.MaxDirectorySectors;

            AssertTrue(AuditBHasPackage(AuditBCfbPackageInDirSector(cap)), "сектор №" + cap + " ещё читается");
            AssertTrue(!AuditBHasPackage(AuditBCfbPackageInDirSector(cap + 1)),
                "сектор №" + (cap + 1) + " за потолком: «не знаем» → нейтральная формулировка");
        }

        /// <summary>CFB, где EncryptedPackage — первая запись n-го сектора каталога (по 4 записи, первая — Root).</summary>
        private static byte[] AuditBCfbPackageInDirSector(int n)
        {
            var streams = new List<string>();
            for (int i = 1; i < 4 * (n - 1); i++)
                streams.Add("s" + i);
            streams.Add("EncryptedPackage");
            return AuditBCfb(streams.ToArray());
        }

        [Test("Аудит B: шифрованная .xlsx — «пароль», переименованная .xls — «формат не тот» (живой Excel)")]
        private static void AuditBOle2NotesForRealFilesLive()
        {
            AuditBRequireExcel();
            using (var dir = new TempDir("auditB_ole_"))
            {
                string encrypted = dir.File("Шифр.xlsx");
                string renamed = dir.File("Старая.xlsx");
                string plain = dir.File("Обычная.xlsx");
                MergeResult result = null;
                AuditBOfficeSta(180000, delegate
                {
                    AuditBWithExcel(delegate(object app)
                    {
                        Action<dynamic> fill = delegate(dynamic ws) { ws.Range("A1").Value2 = "данные"; };
                        AuditBSaveBook(app, encrypted, fill, "secret", null);
                        AuditBSaveBook(app, dir.File("Старая.xls"), fill);
                        AuditBSaveBook(app, plain, fill);
                    });
                    File.Move(dir.File("Старая.xls"), renamed);
                    result = new MergeService().Merge(new[] { encrypted, renamed, plain }, dir.File("Свод.xlsx"), new MergeOptions());
                });

                AssertTrue(FileSignature.HasEncryptedPackage(encrypted), "у шифрованной книги Excel поток EncryptedPackage найден");
                AssertTrue(!FileSignature.HasEncryptedPackage(renamed), "у книги .xls потока нет");
                AssertEq("файл защищён паролем", AuditBNoteOf(result, "Шифр.xlsx"), "шифрованная книга");
                string renamedNote = AuditBNoteOf(result, "Старая.xlsx") ?? "";
                AssertTrue(renamedNote.IndexOf("парол", StringComparison.OrdinalIgnoreCase) < 0 && renamedNote.Contains(".xls"),
                    "переименованная .xls — не «пароль», а подсказка о формате; получено «" + renamedNote + "»");
            }
        }

        private static string AuditBNoteOf(MergeResult result, string fileName)
        {
            foreach (FileResult fr in result.Files)
                if (fr.FileName == fileName)
                    return fr.Note;
            throw new Exception("нет записи о " + fileName);
        }

        private static bool AuditBHasPackage(byte[] data)
        {
            using (var s = new MemoryStream(data, false))
                return FileSignature.HasEncryptedPackage(s);
        }

        /// <summary>
        /// Минимальный CFB версии 3 (MS-CFB): заголовок, сектор FAT №0, каталог с сектора 1.
        /// streams — потоки корня (по 4 записи каталога на сектор, первая — Root Entry).
        /// </summary>
        private static byte[] AuditBCfb(params string[] streams)
        {
            int entries = streams.Length + 1;
            int dirSectors = (entries + 3) / 4;
            var data = new byte[512 * (2 + dirSectors)];
            AuditBCfbHeader(data, 1, -2, 0);
            int fat = AuditBSector(0);
            for (int i = 0; i < 128; i++)
                AuditBPut32(data, fat + i * 4, -1);                     // FREESECT
            AuditBPut32(data, fat, -3);                                  // FATSECT: сам сектор FAT
            for (int i = 0; i < dirSectors; i++)
                AuditBPut32(data, fat + (1 + i) * 4, i + 1 < dirSectors ? 2 + i : -2);
            for (int e = 0; e < entries; e++)
            {
                int offset = AuditBSector(1 + e / 4) + (e % 4) * 128;
                if (e == 0)
                    AuditBCfbEntry(data, offset, "Root Entry", 5);
                else
                    AuditBCfbEntry(data, offset, streams[e - 1], 2);
            }
            return data;
        }

        private static void AuditBCfbHeader(byte[] data, int firstDirSector, int firstDifatSector, int fatSector0)
        {
            byte[] signature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
            Array.Copy(signature, data, signature.Length);
            AuditBPut16(data, 0x18, 0x3E);   // minor version
            AuditBPut16(data, 0x1A, 3);      // major version 3 — сектор 512 байт
            AuditBPut16(data, 0x1C, 0xFFFE); // порядок байтов
            AuditBPut16(data, 0x1E, 9);      // sector shift
            AuditBPut16(data, 0x20, 6);      // mini sector shift
            AuditBPut32(data, 0x2C, 1);      // число секторов FAT
            AuditBPut32(data, 0x30, firstDirSector);
            AuditBPut32(data, 0x38, 4096);   // mini stream cutoff
            AuditBPut32(data, 0x3C, -2);     // мини-FAT нет
            AuditBPut32(data, 0x44, firstDifatSector);
            AuditBPut32(data, 0x48, firstDifatSector < 0 ? 0 : 1);
            for (int i = 0; i < 109; i++)
                AuditBPut32(data, 0x4C + i * 4, i == 0 ? fatSector0 : -1);
        }

        private static void AuditBCfbEntry(byte[] data, int offset, string name, byte type)
        {
            byte[] bytes = Encoding.Unicode.GetBytes(name);
            Array.Copy(bytes, 0, data, offset, bytes.Length);
            AuditBPut16(data, offset + 0x40, bytes.Length + 2); // длина с завершающим нулём
            data[offset + 0x42] = type;
            AuditBPut32(data, offset + 0x44, -1); // левый, правый соседи и потомок — NOSTREAM
            AuditBPut32(data, offset + 0x48, -1);
            AuditBPut32(data, offset + 0x4C, -1);
        }

        private static int AuditBSector(int sid)
        {
            return (sid + 1) * 512;
        }

        private static void AuditBPut16(byte[] data, int offset, int value)
        {
            data[offset] = (byte)value;
            data[offset + 1] = (byte)(value >> 8);
        }

        private static void AuditBPut32(byte[] data, int offset, int value)
        {
            byte[] bytes = BitConverter.GetBytes(value);
            Array.Copy(bytes, 0, data, offset, 4);
        }

        // ---------- 9. Оглавление: ячейки с именами остаются текстом ----------

        /// <summary>
        /// Проверка КОНТРАКТА оглавления, а не регрессия правки TextToDisplay: на Excel 15.0 и прежний
        /// код (TextToDisplay = имя листа) её проходит — текст ссылки там всегда совпадал с текстом
        /// ячейки (EscapeValues добавляет только префикс-апостроф), а совпадающий текст Excel не
        /// разбирает. Регрессию ловит лишь колонка «№»: Int32 вместо double EscapeValues превратил
        /// бы в значение-ошибку.
        /// </summary>
        [Test("Аудит B: контракт оглавления — имена «=1+1» и «0012» текстом, номера числами (живой Excel)")]
        private static void AuditBTocContractLive()
        {
            AuditBRequireExcel();
            var cells = new Dictionary<string, object>();
            int links = 0;
            string firstSub = null;
            bool b2Formula = true;
            AuditBOfficeSta(120000, delegate
            {
                AuditBWithExcel(delegate(object app)
                {
                    dynamic excel = app;
                    dynamic wb = excel.Workbooks.Add(-4167);
                    try
                    {
                        dynamic first = wb.Sheets[1];
                        first.Name = "=1+1";
                        dynamic second = wb.Sheets.Add(Type.Missing, first);
                        second.Name = "0012";
                        var files = new List<FileResult>
                        {
                            new FileResult { FileName = "=1+1.xlsx", SheetName = "=1+1", Ok = true, Linkable = true },
                            new FileResult { FileName = "0012.xlsx", SheetName = "0012", Ok = true, Linkable = true },
                            new FileResult { FileName = "Завис.xlsx", Note = MergeService.WedgedNote }
                        };

                        TocBuilder.Build((object)wb, "Содержание", files);

                        dynamic toc = wb.Sheets["Содержание"];
                        foreach (string a in new[] { "A2", "B2", "B3", "C2", "D4" })
                            cells[a] = toc.Range(a).Value2;
                        b2Formula = (bool)toc.Range("B2").HasFormula;
                        links = (int)toc.Hyperlinks.Count;
                        firstSub = (string)toc.Hyperlinks[1].SubAddress;
                    }
                    finally
                    {
                        object o = wb;
                        try { wb.Close(false); } catch { }
                        ComSafe.Release(o);
                    }
                });
            });

            AssertTrue(!b2Formula, "имя листа «=1+1» не стало формулой");
            AssertEq("=1+1", cells["B2"], "B2 — имя буквально");
            AssertEq("0012", cells["B3"], "B3 — ведущие нули на месте: " + AuditBDescribe(cells["B3"]));
            AssertEq("=1+1.xlsx", cells["C2"], "C2 — имя файла буквально");
            AssertEq(1.0, cells["A2"], "№ — число, а не значение-ошибка: " + AuditBDescribe(cells["A2"]));
            AssertEq("пропущен: вызывал зависание Excel", cells["D4"], "примечание пропущенного");
            AssertEq(2, links, "ссылки только у перенесённых листов");
            AssertEq("'=1+1'!A1", firstSub, "ссылка ведёт на свой лист");
        }

        // ---------- 11. Макросы в итоговом файле ----------

        [Test("Аудит B: VBA переносится только в форматы с макросами (.xlsm/.xlsb/.xls)")]
        private static void AuditBKeepsMacrosByFormat()
        {
            AssertTrue(!MergeService.KeepsMacros(@"C:\out\Свод.xlsx"), ".xlsx отбрасывает код");
            AssertTrue(MergeService.KeepsMacros(@"C:\out\Свод.xlsm"), ".xlsm");
            AssertTrue(MergeService.KeepsMacros(@"C:\out\Свод.XLSB"), ".xlsb (регистр не важен)");
            AssertTrue(MergeService.KeepsMacros(@"C:\out\Свод.xls"), ".xls");
            AssertTrue(!MergeService.KeepsMacros(@"C:\out\Свод.csv"), "неподдерживаемый формат");
        }

        [Test("Аудит B: перенесённые макросы — предупреждение в отчёте, записке и строке статуса")]
        private static void AuditBMacrosWarningEverywhere()
        {
            var result = new MergeResult { OutputPath = @"C:\out\Свод.xlsm", OkCount = 1, MacrosCarried = true };
            result.Files.Add(new FileResult { FileName = "a.xlsm", FullPath = @"C:\in\a.xlsm", Ok = true, SheetName = "a" });

            string report = ReportWriter.BuildReport(result, @"C:\in", new MergeOptions(), new DateTime(2026, 1, 1));
            NoteContent note = NoteText.Build(result, @"C:\in", new MergeOptions(), new DateTime(2026, 1, 1));
            bool clean;
            string status = MainForm.DoneStatus(result, out clean);

            AssertTrue(report.Contains("Внимание:       " + MergeService.MacrosWarningText), "отчёт:\n" + report);
            AssertTrue(note.Tail.Contains("Внимание: " + MergeService.MacrosWarningText + "."), "записка");
            AssertTrue(status.Contains(Loc.T("err.merge.macrosCarried")), "статус: " + status);
            AssertTrue(!clean, "с предупреждением статус не «чистый» (не зелёный)");
        }

        [Test("Аудит B: предупреждение о макросах — в каталоге Loc та же русская фраза, что в отчёте и записке")]
        private static void AuditBMacrosWarningMatchesLoc()
        {
            string[] pair = Loc.Pair("err.merge.macrosCarried");

            AssertTrue(pair != null, "ключ есть в каталоге");
            AssertEq(MergeService.MacrosWarningText, pair[0], "русская строка UI совпадает с документами");
        }

        [Test("Аудит B: каталог Loc — русская строка каждой записи задана литералом (её читает dump_loc.py)")]
        private static void AuditBLocRussianTextsAreLiterals()
        {
            string src = SourceDir();
            if (src == null)
                Skip("исходники (src\\Loc.cs) рядом с тестами не найдены");
            string text = File.ReadAllText(Path.Combine(src, "Loc.cs"), Encoding.UTF8);
            var offenders = new List<string>();
            // После ключа и запятой — сразу строка; выражение (константа, вызов) выгрузка не прочтёт.
            foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(
                text, "\\bA\\(\\s*\"([^\"]+)\"\\s*,\\s*(?=[^\\s\"])"))
                offenders.Add(m.Groups[1].Value);

            AssertEq(0, offenders.Count, "русская строка не литерал: " + string.Join(", ", offenders));
        }

        [Test("Аудит B: без макросов статус свода чистый и без предупреждений")]
        private static void AuditBNoMacrosCleanStatus()
        {
            var result = new MergeResult { OutputPath = @"C:\out\Свод.xlsx", OkCount = 1 };
            result.Files.Add(new FileResult { FileName = "a.xlsx", FullPath = @"C:\in\a.xlsx", Ok = true, SheetName = "a" });

            bool clean;
            string status = MainForm.DoneStatus(result, out clean);
            string report = ReportWriter.BuildReport(result, @"C:\in", new MergeOptions(), new DateTime(2026, 1, 1));

            AssertTrue(clean, "чистый свод");
            AssertEq(string.Format(Loc.T("excel.status.doneClean"), 1), status, "статус");
            AssertTrue(!report.Contains(MergeService.MacrosWarningText), "отчёт без предупреждения");
        }

        // ---------- 12. Word: отмена во время подгонки интервалов, формат ранов ----------

        [Test("Аудит B: «Отмена» во время подгонки интервалов прерывает запись .docx (живой Word)")]
        private static void AuditBWordCancelDuringSpacingFitLive()
        {
            AuditBRequireWord();
            using (var dir = new TempDir("auditB_wcancel_"))
            {
                string path = dir.File("out.docx");
                int polls = 0;
                bool committing = false;
                Exception error = null;
                AuditBOfficeSta(180000, delegate
                {
                    try
                    {
                        // Страница одна: первый опрос — перед ней, второй — уже в подгонке интервалов.
                        WordDocxWriter.Write(AuditBPage(AuditBRun("Текст", false, false, 11, null)), path, null,
                            delegate { return ++polls >= 2; }, delegate { committing = true; });
                    }
                    catch (Exception ex) { error = ex; }
                });

                AssertTrue(error is OperationCanceledException, "ожидалась отмена, получено " +
                    (error == null ? "сохранение документа" : error.GetType().Name + ": " + error.Message));
                AssertTrue(!committing, "точка невозврата не наступила");
                AssertTrue(!File.Exists(path), "документ не сохранён");
            }
        }

        [Test("Аудит B: формат каждого рана (и после гиперссылки) переносится в .docx (живой Word)")]
        private static void AuditBWordRunFormattingLive()
        {
            AuditBRequireWord();
            using (var dir = new TempDir("auditB_wruns_"))
            {
                string path = dir.File("out.docx");
                AuditBOfficeSta(180000, delegate
                {
                    WordDocxWriter.Write(AuditBPage(
                        AuditBRun("Жирный ", true, false, 20, null),
                        AuditBRun("обычный ", false, false, 10, null),
                        AuditBRun("ссылка", false, true, 10, "https://example.com/"),
                        AuditBRun(" хвост", false, false, 10, null)), path);
                });

                XmlDocument xml = AuditBDocumentXml(path);
                var ns = new XmlNamespaceManager(xml.NameTable);
                ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                AssertEq("b,sz=40", AuditBRunFormat(xml, ns, "Жирный"), "первый ран: полужирный 20 pt");
                AssertEq("sz=20", AuditBRunFormat(xml, ns, "обычный"), "второй ран: без начертаний, 10 pt");
                AssertEq("i,sz=20,link", AuditBRunFormat(xml, ns, "ссылка"), "третий ран: курсив, гиперссылка");
                AssertEq("sz=20", AuditBRunFormat(xml, ns, "хвост"), "ран после гиперссылки: формат свой");
            }
        }

        private static OcrRun AuditBRun(string text, bool bold, bool italic, double sizePt, string uri)
        {
            return new OcrRun { Text = text, Bold = bold, Italic = italic, FontSizePt = sizePt, FontName = "Arial", Uri = uri };
        }

        private static List<PdfPageText> AuditBPage(params OcrRun[] runs)
        {
            var paragraph = new OcrParagraph { Alignment = OcrAlignment.Left, TopPt = 780, BottomPt = 760, LeftPt = 56, RightPt = 539 };
            paragraph.Runs.AddRange(runs);
            var page = new PdfPageText
            {
                PageIndex = 0, WidthPt = 595, HeightPt = 842,
                LeftMarginPt = 56, RightMarginPt = 56, TopMarginPt = 56, BottomMarginPt = 56
            };
            page.Paragraphs.Add(paragraph);
            return new List<PdfPageText> { page };
        }

        private static XmlDocument AuditBDocumentXml(string docx)
        {
            using (var zip = new ZipArchive(File.OpenRead(docx), ZipArchiveMode.Read))
            using (Stream s = zip.GetEntry("word/document.xml").Open())
            {
                var xml = new XmlDocument();
                xml.Load(s);
                return xml;
            }
        }

        /// <summary>Формат рана с текстом text: «b», «i», «sz=…» (полупункты), «link» — внутри w:hyperlink.</summary>
        private static string AuditBRunFormat(XmlDocument xml, XmlNamespaceManager ns, string text)
        {
            foreach (XmlNode t in xml.SelectNodes("//w:r/w:t", ns))
            {
                if (t.InnerText.Trim() != text)
                    continue;
                XmlNode run = t.ParentNode;
                var parts = new List<string>();
                if (AuditBOn(run.SelectSingleNode("w:rPr/w:b", ns)))
                    parts.Add("b");
                if (AuditBOn(run.SelectSingleNode("w:rPr/w:i", ns)))
                    parts.Add("i");
                XmlNode sz = run.SelectSingleNode("w:rPr/w:sz", ns);
                if (sz != null)
                    parts.Add("sz=" + sz.Attributes["val", ns.LookupNamespace("w")].Value);
                if (run.ParentNode != null && run.ParentNode.LocalName == "hyperlink")
                    parts.Add("link");
                return string.Join(",", parts.ToArray());
            }
            return "нет рана «" + text + "»";
        }

        private static bool AuditBOn(XmlNode toggle)
        {
            if (toggle == null)
                return false;
            XmlAttribute val = toggle.Attributes["val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"];
            return val == null || (val.Value != "0" && val.Value != "false");
        }

        // ---------- Обвязка живого Office ----------

        [Test("Аудит B: сторож по таймауту завершает экземпляр автоматизации, а Excel человека не трогает (живой Excel)")]
        private static void AuditBWatchdogSparesHumanExcelLive()
        {
            AuditBRequireExcel();
            int automationPid = 0;
            bool humanUpBeforeTimeout = false;
            Process human = null;
            Exception watchdog = null;
            try
            {
                try
                {
                    AuditBOfficeSta(25000, delegate
                    {
                        object app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                        try
                        {
                            automationPid = OwnedOfficeProcess.FromWindow((long)((dynamic)app).Hwnd).Id;
                            // Как человек: отдельный процесс Excel (/x), открытый, пока идёт тест.
                            human = Process.Start(new ProcessStartInfo("excel.exe", "/x") { UseShellExecute = true });
                            human.WaitForInputIdle(15000);
                            using (Process automation = Process.GetProcessById(automationPid))
                            {
                                humanUpBeforeTimeout = !automation.HasExited;
                                automation.WaitForExit(60000); // «завис»: ждём, пока сторож не завершит экземпляр
                            }
                        }
                        finally
                        {
                            ComSafe.Release(app);
                        }
                    });
                }
                catch (Exception ex)
                {
                    watchdog = ex;
                }

                AssertTrue(watchdog != null && watchdog.Message.StartsWith("Office не ответил", StringComparison.Ordinal),
                    "сторож сработал: " + (watchdog == null ? "нет" : watchdog.Message));
                AssertTrue(human != null && humanUpBeforeTimeout, "Excel человека запущен до срабатывания сторожа");
                AssertTrue(!human.HasExited, "Excel, открытый не автоматизацией, сторож не трогает");
                AssertTrue(!AuditBAlive(automationPid), "экземпляр автоматизации завершён");
            }
            finally
            {
                if (human != null)
                {
                    // Пустая книга без правок закрывается без вопросов; не вышел — добить.
                    try { human.CloseMainWindow(); } catch { }
                    if (!human.WaitForExit(15000))
                        AuditBStop(human);
                    human.Dispose();
                }
                if (AuditBAlive(automationPid))
                    using (Process p = Process.GetProcessById(automationPid))
                        AuditBStop(p);
            }
        }

        private static void AuditBRequireExcel()
        {
            if (Type.GetTypeFromProgID("Excel.Application") == null)
                SkipMissing("office", "Excel не установлен — живую проверку COM выполнить нечем");
        }

        private static void AuditBRequireWord()
        {
            if (Type.GetTypeFromProgID("Word.Application") == null)
                SkipMissing("office", "Word не установлен — живую проверку COM выполнить нечем");
        }

        /// <summary>
        /// Тело — на STA-потоке (требование Office COM) с таймаутом. Модальный запрос в скрытом
        /// экземпляре повесил бы прогон до сторожа: по таймауту завершаются ТОЛЬКО процессы
        /// автоматизации Office, которых не было до теста, и тест падает с понятной причиной.
        /// Excel или Word, открытый за это время человеком, сторож не трогает.
        /// </summary>
        private static void AuditBOfficeSta(int timeoutMs, Action body)
        {
            HashSet<int> before = AuditBOfficePids();
            Exception error = null;
            var thread = new Thread(delegate()
            {
                try { body(); }
                catch (Exception ex) { error = ex; }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
            if (!thread.Join(timeoutMs))
            {
                foreach (int pid in AuditBOfficeAutomationPids())
                {
                    if (before.Contains(pid))
                        continue;
                    try
                    {
                        using (Process p = Process.GetProcessById(pid))
                            p.Kill();
                    }
                    catch { } // уже вышел
                }
                thread.Join(10000);
                throw new Exception("Office не ответил за " + timeoutMs / 1000 + " с — модальный запрос в скрытом экземпляре?");
            }
            if (error != null)
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
        }

        /// <summary>
        /// PID процессов Excel/Word, запущенных COM: к командной строке сервера COM дописывает
        /// « -Embedding» (LocalServer32), у открытых человеком его нет. Командная строка недоступна
        /// (чужой сеанс, повышенные права) — процесс не наш, в список не попадает.
        /// </summary>
        private static List<int> AuditBOfficeAutomationPids()
        {
            var pids = new List<int>();
            using (var query = new System.Management.ManagementObjectSearcher(
                "SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name = 'EXCEL.EXE' OR Name = 'WINWORD.EXE'"))
            using (System.Management.ManagementObjectCollection found = query.Get())
            {
                foreach (System.Management.ManagementBaseObject process in found)
                    using (process)
                    {
                        string commandLine = process["CommandLine"] as string;
                        if (commandLine != null && commandLine.IndexOf("-Embedding", StringComparison.OrdinalIgnoreCase) >= 0)
                            pids.Add((int)(uint)process["ProcessId"]);
                    }
            }
            return pids;
        }

        private static HashSet<int> AuditBOfficePids()
        {
            var pids = new HashSet<int>();
            foreach (string name in new[] { "EXCEL", "WINWORD" })
                foreach (Process p in Process.GetProcessesByName(name))
                    using (p)
                        pids.Add(p.Id);
            return pids;
        }

        /// <summary>Скрытый Excel на время body; Quit + Release + Collect в finally (правила COM проекта).</summary>
        private static void AuditBWithExcel(Action<object> body)
        {
            object app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            ComMessageFilter.Register();
            try
            {
                dynamic excel = app;
                excel.Visible = false;
                excel.DisplayAlerts = false;
                body(app);
            }
            finally
            {
                try { ((dynamic)app).Quit(); } catch { }
                ComSafe.Release(app);
                ComSafe.Collect();
                ComMessageFilter.Revoke();
            }
        }

        /// <summary>Новая книга из одного листа: fill наполняет лист, затем SaveAs по расширению path.</summary>
        private static void AuditBSaveBook(object app, string path, Action<dynamic> fill,
            string openPassword = null, string writePassword = null)
        {
            dynamic excel = app;
            dynamic wb = excel.Workbooks.Add(-4167);
            try
            {
                fill(wb.Sheets[1]);
                wb.SaveAs(path, OutputFormats.FileFormatFor(path),
                    (object)openPassword ?? Type.Missing, (object)writePassword ?? Type.Missing);
            }
            finally
            {
                object o = wb;
                try { wb.Close(false); } catch { }
                ComSafe.Release(o);
            }
        }

        /// <summary>Открыть книгу только для чтения в отдельном скрытом Excel и прочитать её.</summary>
        private static void AuditBReadBook(string path, Action<dynamic> read)
        {
            AuditBWithExcel(delegate(object app)
            {
                dynamic excel = app;
                dynamic wb = excel.Workbooks.Open(path, 0, true);
                try { read(wb); }
                finally
                {
                    object o = wb;
                    try { wb.Close(false); } catch { }
                    ComSafe.Release(o);
                }
            });
        }

        private static string AuditBDescribe(object value)
        {
            return value == null ? "null" : value.GetType().Name + ":" + value;
        }
    }
}
