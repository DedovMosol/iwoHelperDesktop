using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Регрессии этапа «Интерфейс» (UI-*): масштаб DPI, доступность, высокий контраст.
    /// </summary>
    internal static partial class UnitTests
    {
        // ---------- UI-1: режим DPI задаёт app.config, а не манифест ----------

        /// <summary>
        /// WinForms включает свою поддержку высокого DPI (перестройку окна при переходе на
        /// монитор с другим масштабом) только по ключу DpiAwareness в app.config. В 1.18.4
        /// процесс был PerMonitorV2 по манифесту, а WinForms об этом не знал: проба на сборке
        /// с тем манифестом показала EnableDpiChangedMessageHandling = False, с app.config — True.
        /// Объявлять DPI в манифесте Microsoft больше не рекомендует (оно перекрывает app.config),
        /// а при нём --selftest не отличил бы пропавший .config от лежащего рядом: окно было бы
        /// per-monitor и так. Поэтому стережём здесь все звенья: значение в build\app.config, его
        /// подключение к сборке, отсутствие DPI и наличие Windows 10 в манифесте, установщик.
        /// Что файл действительно лежит рядом с exe и сработал, проверяет --selftest (код 5).
        /// </summary>
        [Test("Аудит UI-1: режим DPI объявлен только в app.config и подключён к сборке")]
        private static void AuditDpiModeComesFromAppConfig()
        {
            string root = RepoRoot();
            AssertTrue(root != null, "корень репозитория найден");

            XDocument config = XDocument.Load(Path.Combine(root, "build", "app.config"));
            XElement section = config.Root.Element("System.Windows.Forms.ApplicationConfigurationSection");
            AssertTrue(section != null, "в app.config есть секция WinForms");
            string[] dpi = section.Elements("add")
                .Where(e => (string)e.Attribute("key") == "DpiAwareness")
                .Select(e => (string)e.Attribute("value")).ToArray();
            AssertEq(1, dpi.Length, "ключ DpiAwareness ровно один");
            AssertEq("PerMonitorV2", dpi[0], "режим — PerMonitorV2");

            XDocument project = XDocument.Load(Path.Combine(root, "iwoHelperDesktop.csproj"));
            string[] appConfig = project.Descendants("AppConfig").Select(e => e.Value.Trim()).ToArray();
            AssertEq(1, appConfig.Length, "AppConfig задан один раз");
            AssertEq(@"build\app.config", appConfig[0], "сборка кладёт build\\app.config рядом с exe");

            XDocument manifest = XDocument.Load(Path.Combine(root, "build", "app.manifest"));
            string[] declared = manifest.Descendants()
                .Where(e => e.Name.LocalName == "dpiAware" || e.Name.LocalName == "dpiAwareness")
                .Select(e => e.Name.LocalName).ToArray();
            AssertEq(0, declared.Length, "манифест не объявляет DPI: " + string.Join(", ", declared));
            AssertTrue(manifest.Descendants().Any(e => e.Name.LocalName == "supportedOS" &&
                    (string)e.Attribute("Id") == "{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"),
                "манифест объявляет Windows 10: без этого WinForms не включает PerMonitorV2");

            string installer = File.ReadAllText(Path.Combine(root, "installer", "iwoHelperDesktop.iss"));
            AssertTrue(installer.Contains("Source: \"{#DistDir}\\{#AppExe}.config\"; DestDir: \"{app}\""),
                "установщик ставит .config рядом с exe");
        }

        /// <summary>
        /// Обратная сторона проверки --selftest: она обязана уметь провалиться. У тестового
        /// процесса нет ни .config, ни DPI в манифесте — его окно DPI-unaware, и ответ «да»
        /// здесь означал бы, что самопроверка не заметит пропавший из пакета файл.
        /// </summary>
        [Test("Аудит UI-1: окно процесса без app.config не считается per-monitor")]
        private static void AuditDpiAwarenessRejectsUnawareWindow()
        {
            AssertTrue(!File.Exists(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile),
                "у тестового exe нет .config — иначе проверка ничего не доказывает");
            using (var form = new Form())
                AssertTrue(!DpiAwareness.IsPerMonitor(form.Handle), "окно тестового процесса DPI-unaware");
        }
    }
}
