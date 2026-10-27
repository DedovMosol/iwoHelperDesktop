using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Тесты этапа «2. Тесты и CI» плана аудита (audit.md): сам раннер и обвязка тестов —
    /// строгий режим, дочерний процесс с окнами.
    /// </summary>
    internal static partial class UnitTests
    {
        [Test("Аудит TS-4: строгий режим допускает только пропуски из-за названной отсутствующей зависимости")]
        private static void AuditTs4StrictSkipPolicy()
        {
            HashSet<string> optional = OptionalDependencies(" office, Printer ,, ");
            AssertEq(2, optional.Count, "пустые элементы и пробелы отброшены");
            AssertTrue(SkipAllowed("office", optional), "Office не установлен — допустимо");
            AssertTrue(SkipAllowed("OFFICE", optional), "имя без учёта регистра");
            AssertTrue(SkipAllowed("printer", optional), "второй элемент списка");
            AssertTrue(!SkipAllowed("ghostscript", optional), "Ghostscript не назван — пропуск недопустим");
            AssertTrue(!SkipAllowed(null, optional), "пропуск без зависимости недопустим всегда");
            AssertEq(0, OptionalDependencies(null).Count, "переменная не задана — допустимых нет");
        }

        [Test("Аудит TS-14: [Test], который раннер не выполнит, отмечается провалом, а не теряется")]
        private static void AuditTs14MisplacedTestsReported()
        {
            List<string> ours = MisplacedTestMethods(typeof(UnitTests).Assembly.GetTypes());
            AssertEq(0, ours.Count, "все [Test] этой сборки выполняются: " + string.Join("; ", ours.ToArray()));

            // Зонд — в отдельной динамической сборке, чтобы не ломать сам прогон: метод экземпляра
            // и статический метод с параметром, оба с [Test].
            AssemblyBuilder assembly = AppDomain.CurrentDomain.DefineDynamicAssembly(
                new AssemblyName("Ts14Probe"), AssemblyBuilderAccess.Run);
            TypeBuilder type = assembly.DefineDynamicModule("Ts14Probe").DefineType("Probe", TypeAttributes.Public);
            var test = new CustomAttributeBuilder(typeof(TestAttribute).GetConstructor(new[] { typeof(string) }),
                new object[] { "зонд" });
            MethodBuilder instance = type.DefineMethod("Instance", MethodAttributes.Public, typeof(void), Type.EmptyTypes);
            instance.GetILGenerator().Emit(OpCodes.Ret);
            instance.SetCustomAttribute(test);
            MethodBuilder withArg = type.DefineMethod("WithArg", MethodAttributes.Public | MethodAttributes.Static,
                typeof(void), new[] { typeof(int) });
            withArg.GetILGenerator().Emit(OpCodes.Ret);
            withArg.SetCustomAttribute(test);
            List<string> probe = MisplacedTestMethods(new[] { type.CreateType() });
            AssertEq(2, probe.Count, "оба невыполнимых [Test] найдены: " + string.Join("; ", probe.ToArray()));
        }

        [Test("Аудит TS-8: оболочка, уничтоженная раньше своих окон, закрывает хаб и инструменты")]
        private static void AuditTs8ShellDisposeClosesWindows()
        {
            System.Windows.Forms.Form hub = null, tool = null;
            InIsolatedSettings("iwo_ts8_shell_", delegate
            {
                var ctx = new ShellContext();
                ctx.OpenTool("split", "Split", delegate(Action back) { return new PdfSplitForm(back); }, HubLevel.Pdf);
                hub = HubOf(ctx);
                tool = FindOpenTool(ctx, typeof(PdfSplitForm));
                AssertTrue(hub != null && tool != null, "хаб и инструмент открыты");
                ctx.Dispose();
            });
            AssertTrue(hub.IsDisposed, "хаб закрыт вместе с оболочкой");
            AssertTrue(tool.IsDisposed, "инструмент закрыт вместе с оболочкой");
        }

        [Test("Аудит TS-8: предпросмотр, уничтоженный без закрытия, отпускает растр и его резерв памяти")]
        private static void AuditTs8PreviewDisposeReleases()
        {
            System.Drawing.Bitmap owned = null;
            long budget = PdfMemoryBudget.Used;
            InIsolatedSettings("iwo_ts8_preview_", delegate
            {
                using (var dir = new TempDir("iwo_ts8_preview_"))
                {
                    PagePreviewForm preview = AuditDNewPreview(dir.File("нет.pdf"));
                    preview.Show();
                    AuditDJoinPreviewWorker(preview);
                    System.Windows.Forms.Application.DoEvents();
                    BudgetedBitmap page = AuditDBitmap(8, 8);
                    owned = page.Bitmap;
                    SetInternal(preview, "_ownedImage", page);
                    preview.Dispose(); // без Close: FormClosed не поднимается (проверено)
                }
            });
            AssertThrowsExactly<ArgumentException>("растр окна освобождён Dispose", delegate { owned.GetPixel(0, 0); });
            AssertEq(budget, PdfMemoryBudget.Used, "резерв памяти растра вернулся");
        }

        [Test("Аудит TS-17: дочерний процесс с окнами — неизвестная раскладка даёт код 2, закрытый stdin завершает процесс")]
        private static void AuditTs17WindowChildLifecycle()
        {
            var info = new ProcessStartInfo(Assembly.GetEntryAssembly().Location, WindowChild.Switch + " no-such-layout");
            info.UseShellExecute = false;
            info.CreateNoWindow = true;
            info.RedirectStandardError = true;
            using (Process bad = Process.Start(info))
            {
                try
                {
                    AssertTrue(bad.WaitForExit(30000), "процесс с неизвестной раскладкой завершился сам");
                    AssertEq(2, bad.ExitCode, "код неизвестной раскладки");
                }
                finally
                {
                    WindowChild.Stop(bad);
                }
            }

            long window;
            using (Process child = WindowChild.Start(WindowChild.TwoHidden, out window))
            {
                try
                {
                    AssertTrue(window != 0, "окно сообщено");
                    // Команда с BOM — как её пишет Process.StandardInput при консоли UTF-8 (раннеры CI):
                    // байтами, чтобы проверка не зависела от кодовой страницы этой машины.
                    byte[] command = Encoding.UTF8.GetBytes("\uFEFFclose-first\r\n");
                    child.StandardInput.BaseStream.Write(command, 0, command.Length);
                    child.StandardInput.BaseStream.Flush();
                    AssertEq("closed", WindowChild.ReadLine(child), "команда с BOM узнана");
                    child.StandardInput.Close(); // как родитель, который завершился, не убрав за собой
                    AssertTrue(child.WaitForExit(10000), "без родителя дочерний процесс не остаётся");
                }
                finally
                {
                    WindowChild.Stop(child);
                }
            }
        }
    }
}
