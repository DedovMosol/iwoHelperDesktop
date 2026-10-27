using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Windows.Forms;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Регрессии аудита оболочки (A): запуск, обновления, окна, история, атомарная запись,
    /// уборка при старте. Каждый тест воспроизводит найденный дефект, а не пересказывает код.
    /// </summary>
    internal static partial class UnitTests
    {
        // ---------- 4. WindowPlacement: переполнение на испорченных координатах ----------

        /// <summary>
        /// Испорченная строка настроек (x = 2147483000) раньше переполняла «x + w» в
        /// отрицательное число, проверка правого края молчала, и окно навсегда открывалось
        /// за пределами экрана. Проверяем все крайние значения на двух мониторах.
        /// </summary>
        [Test("Аудит A: окно с крайними координатами возвращается на экран")]
        private static void AuditWindowClampSurvivesExtremeCoordinates()
        {
            var areas = new[] { new Rectangle(0, 0, 1920, 1040), new Rectangle(-1280, 0, 1280, 1024) };
            int[] extremes = { int.MaxValue, 2147483000, int.MaxValue - 799, int.MinValue, -2147483000, 0 };
            foreach (int x in extremes)
                foreach (int y in extremes)
                {
                    Rectangle r = WindowPlacement.ClampToWorkingArea(new Rectangle(x, y, 800, 600),
                        areas, new Size(400, 300));
                    AssertTrue(InsideArea(r, areas[0]) || InsideArea(r, areas[1]),
                        "x=" + x + ", y=" + y + " → " + r + " вне рабочих областей");
                    AssertEq(800, r.Width, "ширина сохранена при x=" + x);
                    AssertEq(600, r.Height, "высота сохранена при y=" + y);
                }
            Rectangle huge = WindowPlacement.ClampToWorkingArea(
                new Rectangle(100, 100, int.MaxValue, int.MaxValue), areas, Size.Empty);
            AssertTrue(InsideArea(huge, areas[0]) || InsideArea(huge, areas[1]),
                "огромный размер ужат до экрана: " + huge);
        }

        /// <summary>
        /// Вложенность в long: Rectangle.Contains сам считает «X + Width» в int и на тех же
        /// крайних координатах переполняется, подтверждая заэкранное окно как «внутри».
        /// </summary>
        private static bool InsideArea(Rectangle r, Rectangle area)
        {
            return r.X >= area.X && r.Y >= area.Y &&
                (long)r.X + r.Width <= (long)area.X + area.Width &&
                (long)r.Y + r.Height <= (long)area.Y + area.Height;
        }

        // ---------- 8. AtomicOutput: длинный путь ----------

        /// <summary>
        /// Временный файл, журнал и резервная копия длиннее цели на 41 символ. Без
        /// longPathAware путь цели в 219–259 символов допустим сам, но рядом с ним ничего
        /// не создаётся — раньше это оборачивалось «не найдена часть пути». Теперь — понятный
        /// отказ ДО записи, а путь на границе по-прежнему работает.
        /// </summary>
        [Test("Аудит A: слишком длинный путь результата — понятный отказ до записи")]
        private static void AuditAtomicOutputRejectsTooLongPath()
        {
            using (var tmp = new TempDir("iwo_lp_"))
            {
                string sub = Path.Combine(tmp.Path, new string('d', Math.Max(1, 150 - tmp.Path.Length - 1)));
                Directory.CreateDirectory(sub);
                string tooLong = Path.Combine(sub, new string('n', 219 - sub.Length - 1 - 4) + ".pdf");
                AssertEq(219, tooLong.Length, "длина цели для проверки");
                string message = null;
                try
                {
                    using (new AtomicOutput(tooLong)) { }
                }
                catch (MergeException ex) { message = ex.Message; }
                AssertTrue(message != null, "длинный путь отвергнут MergeException");
                AssertTrue(message.IndexOf("err.output", StringComparison.Ordinal) < 0,
                    "сообщение локализовано, а не ключ: " + message);
                AssertTrue(message.Contains("219"), "сообщение называет длину пути: " + message);
                AssertEq(0, Directory.GetFiles(sub).Length, "ни журнал, ни временный файл не созданы");

                string fits = Path.Combine(sub, new string('m', 218 - sub.Length - 1 - 4) + ".pdf");
                using (var output = new AtomicOutput(fits))
                {
                    File.WriteAllText(output.TempPath, "ok");
                    output.Commit();
                }
                AssertEq("ok", File.ReadAllText(fits), "путь на границе публикуется как раньше");
                AssertEq(1, Directory.GetFiles(sub).Length, "после публикации служебных файлов нет");
            }
        }

        // ---------- 7. OperationHistory: неразобранные записи ----------

        private static void SeedHistory(params string[] lines)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(AppPaths.HistoryFile));
            File.WriteAllLines(AppPaths.HistoryFile, lines);
        }

        /// <summary>
        /// Строка «e=…», которую эта версия не разобрала (испорчена или из более новой
        /// версии), раньше переживала очистку, выключение и автоочистку: путь к файлу
        /// оставался на диске вопреки PRIVACY.md. Чужие НЕзаписи при этом беречь надо.
        /// </summary>
        [Test("Аудит A: очистка, выключение и автоочистка истории стирают неразобранные записи")]
        private static void AuditHistoryDropsUnparsedEntries()
        {
            InIsolatedSettings("iwo_hist_unknown_", delegate
            {
                const string secret = @"C:\Тайное\отчёт.pdf";
                string now = DateTime.UtcNow.Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture);
                string old = DateTime.UtcNow.AddDays(-60).Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture);
                string[] seed =
                {
                    "enabled=True",
                    "autoClearDays=0",
                    "e=не-число\top\t" + secret,
                    "e=" + now + "\top\t" + secret + "\tполе-из-будущего",
                    "e=" + now + "\top\tC:\\обычный.pdf",
                    "future.key=keep"
                };

                SeedHistory(seed);
                AssertTrue(OperationHistory.Clear(), "очистка выполнена");
                string text = File.ReadAllText(AppPaths.HistoryFile);
                AssertTrue(text.IndexOf(secret, StringComparison.Ordinal) < 0, "очистка стёрла неразобранные пути");
                AssertTrue(text.IndexOf("обычный", StringComparison.Ordinal) < 0, "очистка стёрла обычные записи");
                AssertTrue(text.Contains("future.key=keep"), "чужая настройка после очистки уцелела");

                SeedHistory(seed);
                AssertTrue(OperationHistory.SetEnabled(false), "выключение выполнено");
                text = File.ReadAllText(AppPaths.HistoryFile);
                AssertTrue(text.IndexOf(secret, StringComparison.Ordinal) < 0, "выключение стёрло неразобранные пути");
                AssertTrue(text.Contains("future.key=keep"), "чужая настройка после выключения уцелела");

                SeedHistory("enabled=True", "autoClearDays=0",
                    "e=не-число\top\t" + secret,
                    "e=" + old + "\top\t" + secret + "\tполе-из-будущего",
                    "e=" + now + "\top\tC:\\свежее.pdf\tполе-из-будущего",
                    "future.key=keep");
                AssertTrue(OperationHistory.SetAutoClear(30), "автоочистка включена");
                text = File.ReadAllText(AppPaths.HistoryFile);
                AssertTrue(text.IndexOf(secret, StringComparison.Ordinal) < 0,
                    "автоочистка стёрла старые и недатируемые неразобранные записи");
                AssertTrue(text.Contains("свежее.pdf"), "свежая запись нового формата сохранена");
                AssertTrue(text.Contains("future.key=keep"), "чужая настройка после автоочистки уцелела");
            });
        }

        // ---------- 9. StartupSweep: изоляция сбоев по журналам ----------

        /// <summary>
        /// Удаление разобранного журнала стояло вне try: один журнал, который нельзя
        /// удалить, обрывал уборку всей папки — следующие журналы не восстанавливались,
        /// старые временные файлы оставались. Запрет удаления ставим через ACL.
        /// </summary>
        [Test("Аудит A: неудаляемый журнал не срывает уборку остальной папки")]
        private static void AuditSweepSurvivesUndeletableJournal()
        {
            using (var tmp = new TempDir("iwo_sweep_acl_"))
            {
                string a = new string('a', 32), b = new string('b', 32);
                string marker1 = tmp.File("one.pdf.iwo-" + a + ".txn");
                File.WriteAllText(tmp.File("one.pdf"), "current");
                File.WriteAllText(tmp.File("one.iwo-" + a + ".pdf"), "new");
                File.WriteAllText(marker1, "");
                string target2 = tmp.File("two.pdf");
                File.WriteAllText(tmp.File("two.iwo-" + b + ".pdf"), "new2");
                File.WriteAllText(tmp.File("two.pdf.iwo-" + b + ".bak"), "old2");
                File.WriteAllText(tmp.File("two.pdf.iwo-" + b + ".txn"), "");
                string orphan = tmp.File("x.iwo-" + new string('c', 32) + ".pdf");
                File.WriteAllText(orphan, "x");
                File.SetLastWriteTimeUtc(orphan, DateTime.UtcNow.AddHours(-2));

                SecurityIdentifier user = WindowsIdentity.GetCurrent().User;
                var fileDeny = new FileSystemAccessRule(user, FileSystemRights.Delete, AccessControlType.Deny);
                var dirDeny = new FileSystemAccessRule(user, FileSystemRights.DeleteSubdirectoriesAndFiles,
                    InheritanceFlags.None, PropagationFlags.None, AccessControlType.Deny);
                FileSecurity fileAcl = File.GetAccessControl(marker1);
                fileAcl.AddAccessRule(fileDeny);
                File.SetAccessControl(marker1, fileAcl);
                DirectorySecurity dirAcl = Directory.GetAccessControl(tmp.Path);
                dirAcl.AddAccessRule(dirDeny);
                Directory.SetAccessControl(tmp.Path, dirAcl);
                try
                {
                    StartupSweep.Sweep(new[] { tmp.Path }, DateTime.UtcNow);
                    if (!File.Exists(marker1))
                        Skip("запрет удаления через ACL не сработал на этой машине");
                    AssertTrue(File.Exists(target2) && File.ReadAllText(target2) == "old2",
                        "второй журнал восстановлен несмотря на сбой первого");
                    AssertTrue(!File.Exists(orphan), "старый временный файл убран");
                }
                finally
                {
                    dirAcl = Directory.GetAccessControl(tmp.Path);
                    dirAcl.RemoveAccessRule(dirDeny);
                    Directory.SetAccessControl(tmp.Path, dirAcl);
                    fileAcl = File.GetAccessControl(marker1);
                    fileAcl.RemoveAccessRule(fileDeny);
                    File.SetAccessControl(marker1, fileAcl);
                }
            }
        }

        // ---------- 6. ShellContext: подписка на язык и модальные окна ----------

        private static int LocChangedSubscribers()
        {
            var handler = (Delegate)typeof(Loc).GetField("Changed",
                BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
            return handler == null ? 0 : handler.GetInvocationList().Length;
        }

        /// <summary>Обработчик сигнала повторного запуска, который оболочка отдала слушателю.</summary>
        private static Action ActivationCallback(ShellContext ctx)
        {
            object listener = typeof(ShellContext).GetField("_activation",
                BindingFlags.Instance | BindingFlags.NonPublic).GetValue(ctx);
            if (listener == null)
                return null;
            return (Action)listener.GetType().GetField("_onActivate",
                BindingFlags.Instance | BindingFlags.NonPublic).GetValue(listener);
        }

        /// <summary>
        /// Статическое событие держит оболочку живой: без отписки в Dispose следующая смена
        /// языка пересобирала бы окна уже уничтоженного контекста.
        /// </summary>
        [Test("Аудит A: оболочка отписывается от смены языка в Dispose")]
        private static void AuditShellUnsubscribesFromLanguage()
        {
            InIsolatedSettings("iwo_shell_loc_", delegate
            {
                int before = LocChangedSubscribers();
                var ctx = new ShellContext();
                try
                {
                    AssertEq(before + 1, LocChangedSubscribers(), "оболочка подписана на смену языка");
                }
                finally
                {
                    Form hub = HubOf(ctx);
                    ctx.Dispose();
                    if (hub != null && !hub.IsDisposed)
                        hub.Close();
                }
                AssertEq(before, LocChangedSubscribers(), "после Dispose подписки не осталось");
            });
        }

        /// <summary>
        /// Повторный запуск ярлыка, пока открыто модальное окно, создавал НОВЫЙ хаб поверх
        /// модального цикла; смена языка в нём закрывала владельца модального окна. Теперь
        /// сигнал поднимает само модальное окно, а без него по-прежнему открывает хаб.
        /// </summary>
        [Test("Аудит A: сигнал повторного запуска при модальном окне не создаёт хаб")]
        private static void AuditActivationPrefersModalWindow()
        {
            InIsolatedSettings("iwo_shell_modal_", delegate
            {
                var ctx = new ShellContext();
                Form spawned = null;
                string failure = null;
                try
                {
                    Action activate = ActivationCallback(ctx);
                    if (activate == null)
                        Skip("слушатель повторного запуска не создан в этом сеансе");
                    HubOf(ctx).Close(); // хаб закрыт, открыт только диалог — исходный сценарий
                    using (var dialog = new Form())
                    {
                        dialog.ShowInTaskbar = false;
                        dialog.StartPosition = FormStartPosition.Manual;
                        dialog.Bounds = new Rectangle(-3000, -3000, 200, 120);
                        dialog.Shown += delegate
                        {
                            try
                            {
                                activate();
                                spawned = HubOf(ctx);
                            }
                            catch (Exception ex) { failure = ex.GetType().Name + ": " + ex.Message; }
                            finally { dialog.Close(); }
                        };
                        dialog.ShowDialog();
                    }
                    AssertTrue(failure == null, "сигнал при модальном окне: " + failure);
                    AssertTrue(spawned == null, "поверх модального окна хаб не создан");

                    activate();
                    spawned = HubOf(ctx);
                    AssertTrue(spawned != null && !spawned.IsDisposed, "без модального окна сигнал открывает хаб");
                }
                finally
                {
                    ctx.Dispose();
                    if (spawned != null && !spawned.IsDisposed)
                        spawned.Close();
                }
            });
        }

        // ---------- 3. Сводка «что нового»: обрезка только в диалоге обновления ----------

        private static string EightBulletNotes()
        {
            return "{\"9.9.9\":{\"ru\":\"• 1\\n• 2\\n• 3\\n• 4\\n• 5\\n• 6\\n• 7\\n• 8\",\"en\":\"• one\"}}";
        }

        /// <summary>
        /// Шесть строк — предел ОКНА обновления. Раньше обрезка сидела в общем разборе, и
        /// локальное окно «Что нового» теряло пункты с седьмого.
        /// </summary>
        [Test("Аудит A: разбор сводки не обрезает пункты локального окна")]
        private static void AuditWhatsNewParserKeepsAllBullets()
        {
            string notes = UpdateChecker.ExtractNotes(EightBulletNotes(), "9.9.9", "ru");
            AssertEq(8, notes.Split('\n').Length, "разбор отдаёт все восемь строк");
            AssertTrue(notes.IndexOf('…') < 0, "многоточие обрезки в разбор не попадает");

            List<string> items = WhatsNewCatalog.ParseItems(EightBulletNotes(), "9.9.9", "ru");
            AssertEq(8, items.Count, "локальное окно «Что нового» получает все восемь пунктов");
            AssertTrue(items[7].Contains("8"), "восьмой пункт на месте: " + items[7]);
        }

        [Test("Аудит A: сводка диалога обновления по-прежнему ограничена шестью строками")]
        private static void AuditUpdateDialogNotesStayLimited()
        {
            string notes = UpdateChecker.DialogNotes(EightBulletNotes(), "9.9.9", "ru");
            string[] lines = notes.Replace("\r\n", "\n").Split('\n');
            AssertEq(7, lines.Length, "шесть строк и многоточие");
            AssertEq("…", lines[6], "отброшенное заменено многоточием");
            AssertEq("", UpdateChecker.DialogNotes("мусор", "9.9.9", "ru"), "битый файл — пустая сводка");
        }

        // ---------- 1. Проверка обновлений при запуске: fail closed ----------

        /// <summary>
        /// Load() на нечитаемом файле отдаёт умолчания (проверять — да), и отказ пользователя
        /// от сетевой проверки молча игнорировался. Решение — только по достоверному снимку.
        /// </summary>
        [Test("Аудит A: непрочитанные настройки не разрешают проверку обновлений при запуске")]
        private static void AuditStartupUpdateCheckFailsClosed()
        {
            var on = new UserSettings { UpdateCheckOnStart = true };
            var off = new UserSettings { UpdateCheckOnStart = false };
            AssertTrue(UpdateChecker.AllowsStartupCheck(true, on), "достоверно включено — проверяем");
            AssertTrue(!UpdateChecker.AllowsStartupCheck(true, off), "достоверно выключено — нет");
            AssertTrue(!UpdateChecker.AllowsStartupCheck(false, on), "непрочитанный файл (умолчания) — нет");
            AssertTrue(!UpdateChecker.AllowsStartupCheck(true, null), "нет снимка — нет");

            InIsolatedSettings("iwo_upd_closed_", delegate
            {
                UserSettings settings;
                AssertTrue(UpdateUi.StartupCheckEnabled(out settings),
                    "файла нет — это достоверные умолчания, проверка разрешена");
                string path = AppPaths.SettingsFile;
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, "");
                using (new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    AssertTrue(!UpdateUi.StartupCheckEnabled(out settings),
                        "файл занят и не прочитан — в сеть не идём");
                }
            });
        }

        // ---------- 2. TLS: не пинить 1.2 поверх SystemDefault ----------

        [Test("Аудит A: SystemDefault протоколов не превращается в один TLS 1.2")]
        private static void AuditTlsKeepsSystemDefault()
        {
            AssertEq(System.Net.SecurityProtocolType.SystemDefault,
                UpdateChecker.WithTls12(System.Net.SecurityProtocolType.SystemDefault),
                "SystemDefault оставлен ОС");
            AssertEq(System.Net.SecurityProtocolType.Tls | System.Net.SecurityProtocolType.Tls12,
                UpdateChecker.WithTls12(System.Net.SecurityProtocolType.Tls),
                "к явному списку Tls12 добавляется");
            AssertEq(System.Net.SecurityProtocolType.Tls12,
                UpdateChecker.WithTls12(System.Net.SecurityProtocolType.Tls12), "уже есть — без изменений");
        }

        // ---------- 5. Распаковка вшитых документов ----------

        /// <summary>
        /// Раньше копия той же ДЛИНЫ считалась актуальной (правка без смены размера не
        /// доезжала), а перезапись шла усечением на месте. Теперь — побайтное сравнение и
        /// публикация через временный файл: служебных файлов после записи не остаётся.
        /// </summary>
        [Test("Аудит A: распаковка документа сравнивает содержимое и пишет атомарно")]
        private static void AuditEmbeddedFileWritesOnContentChange()
        {
            using (var tmp = new TempDir("iwo_embed_"))
            {
                string path = tmp.File("manual.docx");
                AssertTrue(EmbeddedFile.WriteIfChanged(path, new byte[] { 1, 2, 3 }), "файла нет — записан");
                AssertTrue(!EmbeddedFile.WriteIfChanged(path, new byte[] { 1, 2, 3 }), "то же содержимое — не трогаем");
                AssertTrue(EmbeddedFile.WriteIfChanged(path, new byte[] { 1, 9, 3 }),
                    "та же длина, другое содержимое — переписан");
                byte[] now = File.ReadAllBytes(path);
                AssertEq(9, (int)now[1], "на диске новое содержимое");
                AssertEq(1, Directory.GetFiles(tmp.Path).Length, "временных и служебных файлов не осталось");

                // Сбой записи (копия открыта без общего доступа) оставляет прежнюю копию целой.
                using (new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    bool threw = false;
                    try { EmbeddedFile.WriteIfChanged(path, new byte[] { 7, 7, 7, 7 }); }
                    catch (Exception) { threw = true; }
                    AssertTrue(threw, "занятую копию заменить нельзя — ошибка наружу");
                }
                AssertEq(3, File.ReadAllBytes(path).Length, "прежняя копия не усечена");
                AssertEq(9, (int)File.ReadAllBytes(path)[1], "прежняя копия цела");
            }
        }

        // ---------- 9. Уборка расшифрованных копий во %TEMP% ----------

        [Test("Аудит A: уборка %TEMP% удаляет только свои старые копии PDF")]
        private static void AuditSweepTempCopies()
        {
            using (var tmp = new TempDir("iwo_tmpsweep_"))
            {
                DateTime now = DateTime.UtcNow;
                string g = Guid.NewGuid().ToString("N");
                Func<string, bool, string> file = delegate(string name, bool old)
                {
                    string p = tmp.File(name);
                    File.WriteAllText(p, "x");
                    File.SetLastWriteTimeUtc(p, old ? now.AddHours(-2) : now.AddMinutes(-10));
                    return p;
                };
                string oldDec = file("iwo_pg_dec_" + g + ".pdf", true);
                string oldOps = file("iwo_ops_" + g + ".pdf", true);
                string freshOps = file("iwo_ops_" + new string('a', 32) + ".pdf", false);
                string foreign = file("iwo_ops_report.pdf", true);
                string foreign2 = file("iwo_pg_dec_" + g + ".pdf.keep", true);
                string other = file("someone_" + g + ".pdf", true);

                string oldDir = Path.Combine(tmp.Path, "iwo_bg_" + g);
                Directory.CreateDirectory(Path.Combine(oldDir, "src0"));
                string inner = Path.Combine(oldDir, "src0", "decrypted.pdf");
                File.WriteAllText(inner, "x");
                File.SetLastWriteTimeUtc(inner, now.AddHours(-3));
                Directory.SetLastWriteTimeUtc(Path.Combine(oldDir, "src0"), now.AddHours(-3));
                Directory.SetLastWriteTimeUtc(oldDir, now.AddHours(-3));

                // Папка старая по своему времени, но внутри идёт свежая запись — живой рендер.
                string liveDir = Path.Combine(tmp.Path, "iwo_bg_" + new string('b', 32));
                Directory.CreateDirectory(liveDir);
                string liveFile = Path.Combine(liveDir, "bg-0001.png");
                File.WriteAllText(liveFile, "x");
                File.SetLastWriteTimeUtc(liveFile, now.AddMinutes(-1));
                Directory.SetLastWriteTimeUtc(liveDir, now.AddHours(-3));

                int removed = StartupSweep.SweepTempCopies(tmp.Path, now);
                AssertEq(3, removed, "удалены две старые копии и одна старая папка");
                AssertTrue(!File.Exists(oldDec) && !File.Exists(oldOps), "старые копии удалены");
                AssertTrue(!Directory.Exists(oldDir), "старая папка подложек удалена");
                AssertTrue(File.Exists(freshOps), "свежая копия не тронута");
                AssertTrue(Directory.Exists(liveDir), "папка с живой записью не тронута");
                AssertTrue(File.Exists(foreign) && File.Exists(foreign2) && File.Exists(other),
                    "чужие имена не тронуты");
                AssertEq(0, StartupSweep.SweepTempCopies(Path.Combine(tmp.Path, "нет-такой"), now),
                    "нет корня — тихо ноль");
            }
        }

        // ---------- 11. Право на передний план — только слушателю ----------

        [Test("Аудит A: процесс слушателя повторного запуска находится по окну")]
        private static void AuditListenerProcessIsFound()
        {
            RunSta(delegate
            {
                IDisposable listener = SingleInstance.Listen(delegate { });
                if (listener == null)
                    Skip("окно-слушатель не создаётся в этом сеансе");
                try
                {
                    int self = System.Diagnostics.Process.GetCurrentProcess().Id;
                    AssertTrue(SingleInstance.ListenerProcessIds().Contains(self),
                        "свой слушатель найден — право получит именно этот процесс");
                }
                finally { listener.Dispose(); }
            });
        }
    }
}
