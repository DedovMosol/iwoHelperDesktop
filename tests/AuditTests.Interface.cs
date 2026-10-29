using System;
using System.Collections.Generic;
using System.Drawing;
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

namespace ExcelMerger.Tests
{
    internal static partial class UnitTests
    {
        // ---------- UI-1: масштаб под DPI после постройки окна ----------

        /// <summary>
        /// Все окна, которые проверяются под масштабом экрана: инструменты, диалоги, хаб и
        /// «Что нового». Отказ сборки — в offenders, а не исключением: одно сломанное окно
        /// не должно прятать остальные.
        /// </summary>
        private static List<Form> ScaledWindows(List<string> offenders)
        {
            var windows = new List<Form>(NewAllToolWindows());
            windows.AddRange(NewAllDialogs(offenders));
            windows.Add(new StartForm());
            windows.Add(new WhatsNewForm(WhatsNewCatalog.CurrentVersion));
            try
            {
                windows.Add((Form)Construct(typeof(PasswordPromptDialog)));
                // Длинная подпись: ширину диалога задаёт уже она, а не минимум в 300 логических.
                windows.Add((Form)typeof(NumberPromptDialog).GetConstructors(
                        System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)[0]
                    .Invoke(new object[] { "Проверка раскладки",
                        "Переместить выделенные страницы после страницы (от 1 до 12345):",
                        "Переместить", 1, 12345, 1 }));
            }
            catch (Exception ex) { offenders.Add("диалог ввода — не собрался: " + Root(ex).Message); }
            return windows;
        }

        /// <summary>
        /// Главная проверка UI-1. На машине с 96 DPI масштаб экрана изображает шов
        /// <see cref="DpiScale.EmulatedScale"/>: шрифты растут, как выросли бы на мониторе
        /// 125/150 %, а окна масштабируются тем же автомасштабом WinForms. До исправления на
        /// 150 % шрифты росли, а раскладка оставалась в пикселях 96 DPI — подписи обрезались.
        /// Проверяем: ничего не свисает и не наложено, подписи кнопок целы — на собственном
        /// размере окна и на минимальном; окно выросло пропорционально масштабу (двойной рост
        /// замеров текста дал бы +20 %, недоросшая раскладка — меньше); раскрытие секций в
        /// «О программе» после масштаба не сжимает окно обратно до логических пикселей.
        /// Окно, которому при этом масштабе не хватает экрана этой машины, не проверяется —
        /// Windows обрезает его по экрану, и это уже другой вопрос (см. audit.md, UI-1).
        /// </summary>
        [Test("Аудит UI-1: при 125 % и 150 % окна целы, пропорциональны, подписи не обрезаны")]
        private static void AuditWindowsSurviveDisplayScale()
        {
            var offenders = new List<string>();
            var notChecked = new List<string>();
            var baseline = new Dictionary<string, Size>();
            var edges = new Dictionary<string, Dictionary<string, int[]>>();
            var minimum = new Dictionary<string, Size>();
            int checkedCount = 0;
            // Дважды: на настоящем экране и на экране CI (1024×768, панель задач 48 px). Окно,
            // которому маленького экрана не хватает, ведёт себя иначе (прокрутка, ужатие) —
            // эту ветку CI однажды и уронил, хотя здесь, на 1080p, всё проходило.
            foreach (Rectangle? fake in new Rectangle?[] { null, new Rectangle(0, 0, 1024, 720) })
            {
            ScaledForm.WorkingAreaForTests = fake == null ? (Func<Rectangle, Rectangle>)null
                : delegate(Rectangle near) { return fake.Value; };
            string screenName = fake == null ? "" : " [экран 1024×768]";
            baseline.Clear(); edges.Clear(); minimum.Clear();
            try
            {
            foreach (float scale in new[] { 1f, 1.25f, 1.5f })
            {
                DpiScale.EmulatedScale = scale;
                try
                {
                    // Шов действительно растит шрифты — иначе проверка была бы на 100 % трижды.
                    AssertEq(9.75f * scale, Ui.Font(9.75f).Size, "шрифт окна при масштабе " + scale);
                    InIsolatedSettings("iwo_dpi_", delegate
                    {
                        List<Form> windows = ScaledWindows(offenders);
                        for (int index = 0; index < windows.Count; index++)
                            using (Form f = windows[index])
                            {
                                // Один тип окна бывает в списке дважды (короткая и длинная подпись).
                                string name = f.GetType().Name + "#" + index;
                                string where = name + " @" + (int)Math.Round(scale * 100) + "%" + screenName;
                                try
                                {
                                    // Минимум — заданный конструктором (замысел окна), до показа: «Свод
                                    // Excel» после показа сам ставит минимумом свою высоту по факту, и
                                    // такой «минимум» на маленьком экране не помещался бы никогда —
                                    // окно, вылезшее за экран, тогда молча уходило бы в «экран мал».
                                    Size designMinimum = f.MinimumSize;
                                    f.Show();
                                    f.PerformLayout();
                                    if (scale == 1f)
                                    {
                                        baseline[name] = f.ClientSize;
                                        minimum[name] = designMinimum;
                                        edges[name] = CollectEdges(f);
                                        f.Close();
                                        continue;
                                    }
                                    Rectangle area = ScaledForm.WorkingArea(f.Bounds);
                                    // Помещается ли минимум — по минимуму 100 %, умноженному на масштаб:
                                    // MinimumSize живого окна Windows сам ужимает до размеров экрана
                                    // (на CI 1024×768 он «помещался» всегда).
                                    Size min100 = minimum[name];
                                    bool minimumFits = min100.Width * scale <= area.Width &&
                                        min100.Height * scale <= area.Height;
                                    if (f.Width > area.Width || f.Height > area.Height)
                                    {
                                        // Больше экрана законно только тогда, когда экрану не хватает
                                        // минимума окна. «Свод Excel» выше экрана прокручивается (UI-24).
                                        // Диалоги с неизменяемой рамкой не ужимаются (ScaledForm.FitToScreen).
                                        bool sizable = f.FormBorderStyle == FormBorderStyle.Sizable;
                                        if (sizable && minimumFits)
                                            offenders.Add(where + " → окно " + f.Size + " больше рабочей области " +
                                                area.Size + ", хотя минимум " + f.MinimumSize + " в неё помещается");
                                        else
                                            notChecked.Add(where);
                                        f.Close();
                                        continue;
                                    }
                                    checkedCount++;
                                    CheckScaledLayout(f, where, offenders);
                                    CheckScrollReachesEnd(f, where, offenders);
                                    CheckButtonCaptionsFit(f, where, offenders);
                                    if (CheckGrewWithScale(f, baseline[name], scale, area, where, offenders))
                                        CheckEdgesScaled(f, edges[name], scale, where, offenders);
                                    var about = f as AboutForm;
                                    if (about != null)
                                    {
                                        int width = about.ClientSize.Width;
                                        about.ProjectToggle.PerformClick();
                                        about.SupportToggle.PerformClick();
                                        about.PerformLayout();
                                        if (about.ClientSize.Width != width)
                                            offenders.Add(where + " → раскрытие секций изменило ширину: " +
                                                width + " → " + about.ClientSize.Width);
                                        CheckScaledLayout(about, where + " (секции раскрыты)", offenders);
                                        // Раскрытое окно растёт до рабочей области СВОЕГО экрана, дальше
                                        // прокручивается; минимум высоты (440) — единственное исключение.
                                        int frameH = about.Height - about.ClientSize.Height;
                                        if (about.Height > area.Height && DpiScale.Device(about, 440) + frameH <= area.Height)
                                            offenders.Add(where + " → раскрытое окно выше экрана: " +
                                                about.Height + " > " + area.Height);
                                    }
                                    if (f.MinimumSize.Width > 0 && minimumFits)
                                    {
                                        f.Size = f.MinimumSize;
                                        f.PerformLayout();
                                        CheckScaledLayout(f, where + " MIN", offenders);
                                        CheckScrollReachesEnd(f, where + " MIN", offenders);
                                    }
                                    // Развёрнутое окно размер не выбирает: раскладке, которой не
                                    // хватает экрана, остаётся только прокрутка (UI-24).
                                    if (f.FormBorderStyle == FormBorderStyle.Sizable && f.MaximizeBox)
                                    {
                                        Size normal = f.Size;
                                        f.WindowState = FormWindowState.Maximized;
                                        f.PerformLayout();
                                        CheckScaledLayout(f, where + " MAX", offenders);
                                        CheckScrollReachesEnd(f, where + " MAX", offenders);
                                        f.WindowState = FormWindowState.Normal;
                                        // Раскладка развёрнутого окна не трогает размер, к которому
                                        // оно вернётся: иначе «Восстановить» открывало бы другое окно.
                                        if (f.Size != normal)
                                            offenders.Add(where + " → после развёртывания размер " +
                                                f.Size + " вместо " + normal);
                                    }
                                    f.Close();
                                }
                                catch (Exception ex) { offenders.Add(where + " — " + Root(ex).Message); }
                            }
                    });
                }
                finally { DpiScale.EmulatedScale = 1f; }
            }
            }
            finally { ScaledForm.WorkingAreaForTests = null; }
            }
            if (notChecked.Count > 0)
                Console.WriteLine("      UI-1: экран мал для " + string.Join(", ", notChecked.ToArray()));
            AssertTrue(checkedCount >= 8, "под масштабом проверено окон: " + checkedCount);
            AssertTrue(offenders.Count == 0, "под масштабом экрана: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// Прокручиваемое окно обязано докручиваться до своего нижнего ряда: после прокрутки
        /// до конца ни один видимый контрол не ниже клиентской области. Так ловится холст,
        /// заданный короче раскладки, — привязанный к низу ряд тогда остаётся недостижимым.
        /// </summary>
        private static void CheckScrollReachesEnd(Form f, string where, List<string> offenders)
        {
            if (!f.AutoScroll || f.AutoScrollMinSize.IsEmpty)
                return;
            f.AutoScrollPosition = new Point(0, f.AutoScrollMinSize.Height);
            f.PerformLayout();
            // Прокрученное окно — та же раскладка: холст не должен ужаться на сдвиг прокрутки.
            CheckScaledLayout(f, where + " (прокручено)", offenders);
            foreach (Control c in f.Controls)
                if (c.Visible && c.Bottom > f.ClientSize.Height)
                    offenders.Add(where + " → прокрутка до конца не достаёт " + Describe(c) + " (низ " +
                        c.Bottom + " при клиенте " + f.ClientSize.Height + ")");
            f.AutoScrollPosition = Point.Empty;
            f.PerformLayout();
            // С клавиатуры: фокус на самом нижнем доступном контроле прокручивает окно к нему.
            Control lowest = f.Controls.Cast<Control>()
                .Where(c => c.Visible && c.CanSelect && c.TabStop)
                .OrderByDescending(c => c.Bottom).FirstOrDefault();
            if (lowest != null)
            {
                // Как Tab: фокус приходит с другого контрола (верхнего), а не «остаётся» на нём.
                Control upper = f.Controls.Cast<Control>()
                    .Where(c => c.Visible && c.CanSelect && c.TabStop && !ReferenceEquals(c, lowest))
                    .OrderBy(c => c.Top).FirstOrDefault();
                if (upper != null)
                    f.ActiveControl = upper;
                f.ActiveControl = lowest;
                if (lowest.Top < 0 || lowest.Bottom > f.ClientSize.Height)
                    offenders.Add(where + " → фокус на " + Describe(lowest) + " не прокрутил к нему окно (" +
                        lowest.Bounds + " при клиенте " + f.ClientSize + ")");
                // Раскладка не отбирает прокрутку у человека: прокрутили вниз колесом, фокус
                // остался наверху, окно перераскладывается (появились ссылки результата) — вид
                // остаётся там, куда его прокрутили.
                if (upper != null)
                {
                    f.ActiveControl = upper;
                    f.AutoScrollPosition = new Point(0, f.AutoScrollMinSize.Height);
                    Point scrolled = f.AutoScrollPosition;
                    f.PerformLayout();
                    if (f.AutoScrollPosition != scrolled)
                        offenders.Add(where + " → раскладка вернула прокрутку к фокусу: " + scrolled +
                            " → " + f.AutoScrollPosition);
                }
                f.AutoScrollPosition = Point.Empty;
            }
        }

        /// <summary>
        /// Константы, которые работают уже после показа окна и потому мимо автомасштаба:
        /// поля вокруг страницы в «вписать» сравнения и линия вставки при перетаскивании файлов
        /// в «Своде Excel». При 150 % обе обязаны вырасти в полтора раза — найдены перепроверкой
        /// шага 2 UI-1, раскладочные тесты их не видят (они меряют контролы, а не масштаб
        /// страницы и не временную линию).
        /// </summary>
        [Test("Аудит UI-1: поля «вписать» и линия вставки растут с масштабом экрана")]
        private static void AuditRuntimeConstantsFollowScale()
        {
            foreach (float scale in new[] { 1f, 1.5f })
            {
                DpiScale.EmulatedScale = scale;
                try
                {
                    using (var dir = new TempDir("iwo_dpi_fit_"))
                    {
                        PdfReviewResult result = AuditLiveResult(dir);
                        InIsolatedSettings("iwo_dpi_fit_s_", delegate
                        {
                            using (var form = new PdfReviewForm(null))
                            {
                                form.Show();
                                Application.DoEvents();
                                InvokeReviewApplyResult(form, result);
                                PdfReviewPageView view = form.UnifiedSourceView;
                                AssertTrue(WaitFor(delegate { return AuditViewIdle(view); }), "страница отрисована");
                                Button fit = view.Controls.OfType<Button>().First(b => b.Text == Loc.T("preview.fit"));
                                fit.PerformClick();
                                Image image = view.ViewportPanel.Controls.OfType<PictureBox>().Single().Image;
                                int margin = (int)Math.Round(24 * scale);
                                Size area = view.ViewportPanel.ClientSize;
                                double expected = PreviewZoom.Fit(image.Size,
                                    new Size(area.Width - margin, area.Height - margin));
                                AssertTrue(Math.Abs(view.ZoomScale - expected) < 1e-9,
                                    "«вписать» при " + scale + ": масштаб " + view.ZoomScale + ", ждали " + expected +
                                    " (поля " + margin + " px)");
                                form.Close();
                            }
                        });
                    }
                    InIsolatedSettings("iwo_dpi_drop_", delegate
                    {
                        using (var form = new MainForm(delegate { }))
                        {
                            form.Show();
                            const System.Reflection.BindingFlags Any =
                                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
                            typeof(MainForm).GetMethod("ShowDropLine", Any).Invoke(form, new object[] { 0 });
                            var line = (Control)typeof(MainForm).GetField("_dropLine", Any).GetValue(form);
                            var list = (Control)typeof(MainForm).GetField("_list", Any).GetValue(form);
                            int inset = (int)Math.Round(2 * scale);
                            AssertEq(inset, line.Height, "толщина линии вставки при " + scale);
                            AssertEq(list.Left + inset, line.Left, "отступ линии слева при " + scale);
                            AssertEq(list.Width - 2 * inset, line.Width, "ширина линии при " + scale);
                            form.Close();
                        }
                    });
                }
                finally { DpiScale.EmulatedScale = 1f; }
            }
        }

        /// <summary>
        /// Нижний ряд «Свода Excel»: ссылки на результат и «Повторить пропущенные» появляются
        /// только после свода, поэтому раскладочные тесты их не видят (скрытые ссылки не
        /// проверяет никто, а ссылка — не кнопка). Ширину минимума окна держат именно они:
        /// на минимальной ширине, при каждом масштабе, на обоих языках и с полосой прокрутки,
        /// когда экрану не хватает высоты, ни одна ссылка не заходит под кнопку.
        /// </summary>
        [Test("Аудит UI-24: нижние ссылки «Свода Excel» не заходят под «Повторить» на минимальной ширине")]
        private static void AuditExcelBottomRowFitsAtMinimumWidth()
        {
            var offenders = new List<string>();
            Lang was = Loc.Current;
            const System.Reflection.BindingFlags Any =
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
            try
            {
                foreach (Lang lang in new[] { Lang.Ru, Lang.En })
                    foreach (float scale in new[] { 1f, 1.25f, 1.5f })
                    {
                        Loc.Init(lang);
                        DpiScale.EmulatedScale = scale;
                        try
                        {
                            InIsolatedSettings("iwo_ui24_row_", delegate
                            {
                                using (var form = new MainForm(delegate { }))
                                {
                                    form.Show();
                                    form.Size = form.MinimumSize;
                                    form.PerformLayout();
                                    string where = lang + " @" + (int)Math.Round(scale * 100) + "% " +
                                        form.ClientSize + (form.AutoScroll ? " (прокрутка)" : "");
                                    var retry = (Control)typeof(MainForm).GetField("_btnRetry", Any).GetValue(form);
                                    retry.Visible = true;
                                    foreach (string name in new[] { "_lnkOpenFile", "_lnkOpenFolder", "_lnkOpenReport", "_lnkNote" })
                                    {
                                        var link = (Control)typeof(MainForm).GetField(name, Any).GetValue(form);
                                        link.Visible = true;
                                        if (link.Bounds.IntersectsWith(retry.Bounds))
                                            offenders.Add(where + " → «" + link.Text + "» " + link.Bounds +
                                                " под «" + retry.Text + "» " + retry.Bounds);
                                    }
                                    form.Close();
                                }
                            });
                        }
                        finally { DpiScale.EmulatedScale = 1f; }
                    }
            }
            finally { Loc.Init(was); }
            AssertTrue(offenders.Count == 0, "нижний ряд свода: " + string.Join(" | ", offenders.ToArray()));
        }

        // ---------- UI-25: детали рисованных контролов растут с масштабом ----------

        /// <summary>
        /// Прямоугольник пикселей, нарисованных контролом и прошедших фильтр цвета, внутри
        /// области region (пусто — весь контрол). Область нужна там, где рядом текст: каймы
        /// ClearType у тёмных букв цветные и прошли бы фильтр «красного» значка.
        /// </summary>
        private static Rectangle InkBounds(Control c, Func<Color, bool> ink, Rectangle region = default(Rectangle))
        {
            using (var bmp = new Bitmap(c.Width, c.Height))
            {
                c.DrawToBitmap(bmp, new Rectangle(Point.Empty, c.Size));
                Rectangle area = region.IsEmpty ? new Rectangle(Point.Empty, c.Size)
                    : Rectangle.Intersect(region, new Rectangle(Point.Empty, c.Size));
                int left = int.MaxValue, top = int.MaxValue, right = -1, bottom = -1;
                for (int y = area.Top; y < area.Bottom; y++)
                    for (int x = area.Left; x < area.Right; x++)
                        if (ink(bmp.GetPixel(x, y)))
                        {
                            left = Math.Min(left, x); top = Math.Min(top, y);
                            right = Math.Max(right, x); bottom = Math.Max(bottom, y);
                        }
                return right < 0 ? Rectangle.Empty : Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
            }
        }

        /// <summary>
        /// Корень из числа «пустых» пикселей в левом верхнем квадрате side×side: срезанный
        /// скруглением угол — это r²·(1 − π/4), так что корень пропорционален радиусу. Площадь
        /// меряется устойчивее, чем край дуги, — тот на радиусе 10 px ловит ±1 px сглаживания.
        /// </summary>
        private static double CornerCut(Control c, int side, Func<Color, bool> ink)
        {
            using (var bmp = new Bitmap(c.Width, c.Height))
            {
                c.DrawToBitmap(bmp, new Rectangle(Point.Empty, c.Size));
                int empty = 0;
                for (int y = 0; y < side; y++)
                    for (int x = 0; x < side; x++)
                        if (!ink(bmp.GetPixel(x, y)))
                            empty++;
                return Math.Sqrt(empty);
            }
        }

        private static bool IsAccentInk(Color c) { return c.G - c.R > 40 && c.G - c.B > 20; }
        private static bool IsRedInk(Color c) { return c.R > 150 && c.G < 120 && c.B < 120; }
        private static bool IsWhiteInk(Color c) { return c.R > 225 && c.G > 225 && c.B > 225; }

        /// <summary>
        /// Детали, которые контролы рисуют сами, растут вместе с масштабом экрана: квадрат
        /// флажка, скругление кнопки, значок карточки инструмента и карточки недавнего файла,
        /// поля заголовка шапки, значок сообщения, запасной флаг глобуса, единицы сетки
        /// страниц. До UI-25 при 150 % текст вырастал, а всё это оставалось размером 100 %.
        /// Меряем по пикселям: контрол рисуется в картинку размером, который ему дал бы
        /// автомасштаб окна, и сравниваются размеры нарисованного при 100 % и при 150 %.
        /// </summary>
        [Test("Аудит UI-25: рисованные детали контролов растут с масштабом экрана")]
        private static void AuditOwnerDrawnDetailsFollowScale()
        {
            var measured = new Dictionary<string, double[]>();
            Action<string, int, double> put = delegate(string what, int index, double value)
            {
                double[] pair;
                if (!measured.TryGetValue(what, out pair))
                    measured[what] = pair = new double[2];
                pair[index] = value;
            };
            Type glyphType = typeof(StartForm).GetNestedType("GlyphButton", System.Reflection.BindingFlags.NonPublic);
            using (var flag = new Bitmap(20, 14))
            {
                using (Graphics g = Graphics.FromImage(flag))
                    g.Clear(Color.FromArgb(220, 20, 20));
                float[] scales = { 1f, 1.5f };
                for (int i = 0; i < scales.Length; i++)
                {
                    float s = scales[i];
                    Func<int, int> d = delegate(int v) { return (int)Math.Round(v * s); };
                    DpiScale.EmulatedScale = s;
                    try
                    {
                        using (var box = new AccentCheckBox { Text = "Флажок", Checked = true, Size = new Size(d(200), d(30)) })
                            put("квадрат флажка", i, InkBounds(box, IsAccentInk).Width);
                        using (var button = new RoundedButton(true) { Text = "", Size = new Size(d(120), d(38)) })
                            put("скругление кнопки", i, CornerCut(button, d(16), IsAccentInk));
                        // Значок — в верхней половине карточки, текст под ним в замер не попадает.
                        using (var card = new ChoiceCard(CardGlyph.Pdf, "Карточка", "Описание") { Size = new Size(d(240), d(250)) })
                            put("значок карточки", i, InkBounds(card, IsRedInk, new Rectangle(0, 0, card.Width, card.Height * 55 / 100)).Width);
                        using (var recent = new RecentCard(new HistoryEntry { Operation = "pdf.merged", Path = @"C:\нет\документ.pdf", WhenUtc = DateTime.UtcNow }))
                        {
                            recent.Size = new Size(d(200), d(36));
                            // Значок — у левого края (до 7 + 24 + 4 логических px), правее — имя файла.
                            put("значок недавнего", i, InkBounds(recent, IsRedInk, new Rectangle(0, 0, d(35), recent.Height)).Width);
                        }
                        using (var header = new HeaderBand("Заголовок", "", Theme.HubBlue, Theme.HubBlueDark) { Size = new Size(d(780), d(76)) })
                        {
                            // Поле — от края шапки до начала текста; собственный отступ первой буквы
                            // (он нелинеен из-за хинтинга) меряем тем же шрифтом с нуля и вычитаем.
                            Font title = Ui.Font(header, 15f, FontStyle.Bold);
                            int bearing;
                            using (var probe = new Bitmap(d(400), d(60)))
                            {
                                using (Graphics g = Graphics.FromImage(probe))
                                {
                                    g.Clear(Theme.HubBlue);
                                    TextRenderer.DrawText(g, "Заголовок", title, Point.Empty, Color.White,
                                        TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding);
                                }
                                bearing = int.MaxValue;
                                for (int y = 0; y < probe.Height; y++)
                                    for (int x = 0; x < probe.Width; x++)
                                        if (IsWhiteInk(probe.GetPixel(x, y)))
                                            bearing = Math.Min(bearing, x);
                            }
                            Rectangle ink = InkBounds(header, IsWhiteInk);
                            put("поле заголовка шапки", i, ink.Left - bearing);
                            put("высота заголовка шапки", i, ink.Height);
                        }
                        // Поле подписи: длинная подпись обрезается многоточием и занимает всю
                        // ширину между полями. Знак узкий («|»), чтобы недобор до края при
                        // обрезке был меньше пикселя-двух, — тогда левый край чернил и есть поле.
                        using (var button = new RoundedButton(false) { Text = new string('|', 200), Size = new Size(d(120), d(30)) })
                            put("поле подписи кнопки", i, InkBounds(button,
                                // ClearType красит узкие штрихи в цветные субпиксели, поэтому «чернила» —
                                // любой заметно не белый канал; рамку кнопки отрезает область замера.
                                c => Math.Min(c.R, Math.Min(c.G, c.B)) < 200,
                                new Rectangle(d(4), button.Height / 3, button.Width - d(8), button.Height / 3)).Left);
                        // Предпочтительный размер флажка: место под квадрат и просвет до подписи.
                        using (var box = new AccentCheckBox { Text = "Флажок" })
                            put("место квадрата в размере флажка", i,
                                box.GetPreferredSize(Size.Empty).Width - TextRenderer.MeasureText(box.Text, box.Font).Width);
                        using (var glyph = (Control)Activator.CreateInstance(glyphType,
                            new object[] { "", 15f, "IwoNoSuchFontFamily", (Func<Image>)delegate { return flag; } }))
                        {
                            glyph.Size = new Size(d(30), d(30));
                            put("запасной флаг глобуса", i, InkBounds(glyph, IsRedInk).Width);
                        }
                        using (var message = (Form)Construct(typeof(MessageForm)))
                        {
                            var icon = (Image)typeof(MessageForm).GetField("_icon",
                                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(message);
                            put("значок сообщения", i, icon.Width);
                        }
                        using (var grid = new PdfPageGrid())
                            put("единица сетки страниц", i, (int)typeof(PdfPageGrid).GetMethod("Px",
                                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                                .Invoke(grid, new object[] { 20 }));
                    }
                    finally { DpiScale.EmulatedScale = 1f; }
                }
            }
            var offenders = new List<string>();
            foreach (KeyValuePair<string, double[]> m in measured)
            {
                double ratio = m.Value[0] > 0 ? m.Value[1] / m.Value[0] : 0;
                if (ratio < 1.35 || ratio > 1.65)
                    offenders.Add(m.Key + ": " + m.Value[0] + " → " + m.Value[1] + " (×" + ratio.ToString("0.00") + ")");
            }
            AssertEq(11, measured.Count, "измерено деталей");
            AssertTrue(offenders.Count == 0, "при 150 % не выросло в 1,5 раза: " + string.Join(" | ", offenders.ToArray()));

            // Пределы скругления и поле подписи растут с масштабом, а без него — прежние.
            AssertEq(9.5f, RoundedButton.RadiusFor(38), "радиус кнопки 38 px при 100 %");
            AssertEq(14.25f, RoundedButton.RadiusFor(57, 1.5f), "та же кнопка при 150 % — пропорционально");
            AssertEq(15f, RoundedButton.RadiusFor(200, 1.5f), "верхний предел при 150 % — 15");
            AssertEq(10, RoundedButton.TextPadFor(160), "поле подписи при 100 %");
            AssertEq(15, RoundedButton.TextPadFor(240, 1.5f), "поле подписи при 150 %");

            // Шрифт рисования на мониторе с другим DPI (PerMonitorV2): одного монитора у этой
            // машины не хватает, поэтому проверяем сам пересчёт от системного DPI.
            int system;
            using (var probe = new Label())
                system = probe.DeviceDpi; // без хэндла — системный
            AssertEq(1f, DpiScale.MonitorRatio(system), "на системном мониторе — без пересчёта");
            AssertEq(1.5f, DpiScale.MonitorRatio(system * 3 / 2), "монитор на 50 % крупнее — кегль ×1,5");
            AssertEq(0.75f, DpiScale.MonitorRatio(system * 3 / 4), "монитор мельче — кегль меньше");
        }

        // ---------- UI-4, UI-10, UI-21: рисуемые кнопки — одно поведение ----------

        private static readonly System.Reflection.BindingFlags AnyInstance =
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic |
            System.Reflection.BindingFlags.Public;

        /// <summary>Вызвать защищённый обработчик контрола (OnKeyDown, OnMouseEnter…) как его зовёт WinForms.</summary>
        private static void Raise(Control c, string handler, EventArgs e)
        {
            typeof(Control).GetMethod(handler, AnyInstance).Invoke(c, new object[] { e });
        }

        /// <summary>Три рисуемые кнопки приложения — каждая в своём хосте того размера, что на экране.</summary>
        private static List<PushControl> NewPushControls()
        {
            Type glyph = typeof(StartForm).GetNestedType("GlyphButton", System.Reflection.BindingFlags.NonPublic);
            var globe = (PushControl)Activator.CreateInstance(glyph,
                new object[] { "", 15f, "Segoe MDL2 Assets", null });
            globe.ForeColor = Theme.HubBlue;
            globe.AccessibleName = "Язык";
            globe.Size = new Size(30, 30);
            return new List<PushControl>
            {
                new ChoiceCard(CardGlyph.Pdf, "Карточка", "Описание") { Size = new Size(240, 250) },
                new RecentCard(new HistoryEntry { Operation = "pdf.merged", Path = @"C:\нет\документ.pdf" }) { Size = new Size(200, 36) },
                globe
            };
        }

        private static int CountDiff(Bitmap a, Bitmap b)
        {
            int n = 0;
            for (int y = 0; y < a.Height; y++)
                for (int x = 0; x < a.Width; x++)
                    if (a.GetPixel(x, y) != b.GetPixel(x, y))
                        n++;
            return n;
        }

        private static Bitmap Snapshot(Control c)
        {
            var bmp = new Bitmap(c.Width, c.Height);
            c.DrawToBitmap(bmp, new Rectangle(Point.Empty, c.Size));
            return bmp;
        }

        /// <summary>
        /// Карточка инструмента, карточка недавнего файла и кнопка-значок — одно поведение
        /// (PushControl). До этого их было три копии, и они разошлись: кнопка-значок без роли
        /// кнопки (диктор говорил «клиент», UI-4), нажатие по нажатию клавиши у двух и по
        /// отпусканию у третьей (UI-21), фокус, неотличимый от наведения и гаснущий при уходе
        /// мыши (UI-10). Проверяем по каждой: роль, имя и действие «Нажать» для диктора, причём
        /// действие нажимает; Enter и Пробел нажимают ровно один раз и только по отпусканию;
        /// отпускание, пришедшее после ухода фокуса, не нажимает; кольцо фокуса рисуется,
        /// отличается от наведения и не гаснет, когда мышь уходит.
        /// </summary>
        [Test("Аудит UI-4/UI-10: рисуемые кнопки — роль для диктора, Enter/Пробел, кольцо фокуса")]
        private static void AuditPushControlsBehaveAsButtons()
        {
            var offenders = new List<string>();
            foreach (PushControl c in NewPushControls())
                using (var host = new Form { ClientSize = new Size(c.Width + 40, c.Height + 40), ShowInTaskbar = false })
                {
                    string name = c.GetType().Name;
                    var other = new Button { Bounds = new Rectangle(0, 0, 10, 10) };
                    host.Controls.Add(other);
                    c.Location = new Point(20, 20);
                    host.Controls.Add(c);
                    host.Show();
                    int clicks = 0;
                    c.Click += delegate { clicks++; };

                    AccessibleObject acc = c.AccessibilityObject;
                    if (acc.Role != AccessibleRole.PushButton) offenders.Add(name + ": роль " + acc.Role);
                    if (string.IsNullOrEmpty(acc.Name)) offenders.Add(name + ": у диктора нет имени");
                    if (acc.DefaultAction != Loc.T("common.a11y.press")) offenders.Add(name + ": действие «" + acc.DefaultAction + "»");
                    acc.DoDefaultAction();
                    if (clicks != 1) offenders.Add(name + ": действие диктора нажало " + clicks + " раз");

                    foreach (Keys key in new[] { Keys.Enter, Keys.Space })
                    {
                        clicks = 0;
                        Raise(c, "OnKeyDown", new KeyEventArgs(key));
                        if (clicks != 0) offenders.Add(name + ": " + key + " нажал по нажатию клавиши");
                        Raise(c, "OnKeyUp", new KeyEventArgs(key));
                        if (clicks != 1) offenders.Add(name + ": " + key + " — нажатий " + clicks + " вместо одного");
                        Raise(c, "OnKeyUp", new KeyEventArgs(key));
                        if (clicks != 1) offenders.Add(name + ": отпускание без нажатия тоже нажало");
                    }
                    clicks = 0;
                    Raise(c, "OnKeyDown", new KeyEventArgs(Keys.Space));
                    Raise(c, "OnLostFocus", EventArgs.Empty);
                    Raise(c, "OnKeyUp", new KeyEventArgs(Keys.Space));
                    if (clicks != 0) offenders.Add(name + ": отпускание после ухода фокуса нажало");
                    clicks = 0;
                    var left = new MouseEventArgs(MouseButtons.Left, 1, 5, 5, 0);
                    Raise(c, "OnMouseDown", left);
                    Raise(c, "OnKeyDown", new KeyEventArgs(Keys.Space));
                    Raise(c, "OnKeyUp", new KeyEventArgs(Keys.Space));
                    Raise(c, "OnMouseUp", left);
                    if (clicks != 0) offenders.Add(name + ": Пробел при зажатой мыши нажал " + clicks + " раз");
                    Raise(c, "OnKeyDown", new KeyEventArgs(Keys.Enter));
                    Raise(c, "OnMouseDown", left);
                    Raise(c, "OnMouseUp", left);
                    Raise(c, "OnKeyUp", new KeyEventArgs(Keys.Enter));
                    if (clicks != 0) offenders.Add(name + ": Enter, перехваченный мышью, нажал " + clicks + " раз");
                    Raise(c, "OnKeyDown", new KeyEventArgs(Keys.Space));
                    Raise(c, "OnMouseLeave", EventArgs.Empty);
                    Raise(c, "OnKeyUp", new KeyEventArgs(Keys.Space));
                    if (clicks != 1) offenders.Add(name + ": уход мыши отменил нажатие Пробелом");
                    c.AccessibleDefaultActionDescription = "Открыть";
                    if (acc.DefaultAction != "Открыть") offenders.Add(name + ": своё описание действия заменено общим");
                    c.AccessibleDefaultActionDescription = null;

                    other.Focus();
                    Application.DoEvents();
                    using (Bitmap idle = Snapshot(c))
                    {
                        Raise(c, "OnMouseEnter", EventArgs.Empty);
                        using (Bitmap hover = Snapshot(c))
                        {
                            Raise(c, "OnMouseLeave", EventArgs.Empty);
                            c.Focus();
                            Application.DoEvents();
                            if (!c.Focused) { offenders.Add(name + ": не получил фокус"); continue; }
                            using (Bitmap focused = Snapshot(c))
                            {
                                if (CountDiff(idle, focused) < 20) offenders.Add(name + ": фокус не виден");
                                if (CountDiff(hover, focused) < 20) offenders.Add(name + ": фокус неотличим от наведения");
                                Raise(c, "OnMouseEnter", EventArgs.Empty);
                                Raise(c, "OnMouseLeave", EventArgs.Empty);
                                using (Bitmap after = Snapshot(c))
                                    if (CountDiff(focused, after) != 0) offenders.Add(name + ": уход мыши погасил фокус");
                            }
                        }
                    }
                    host.Close();
                }
            AssertTrue(offenders.Count == 0, "рисуемые кнопки: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// Текст пунктов «Что нового» нарисован вручную (NonSelectableText): без имени диктор
        /// в режиме просмотра читал пустой «клиент» (UI-5). Имя обязано идти за текстом.
        /// </summary>
        [Test("Аудит UI-5: текст «Что нового» доступен диктору")]
        private static void AuditNonSelectableTextIsReadable()
        {
            using (var text = new NonSelectableText { Text = "Первый пункт" })
            {
                AssertEq(AccessibleRole.StaticText, text.AccessibilityObject.Role, "роль — статический текст");
                AssertEq("Первый пункт", text.AccessibilityObject.Name, "имя — сам текст");
                text.Text = "Второй";
                AssertEq("Второй", text.AccessibilityObject.Name, "имя идёт за текстом");
            }
        }

        /// <summary>
        /// Порядок Tab нижнего ряда хаба — как на экране: «Настройки», карточки недавних,
        /// «О программе», шапка последней. Карточки приходят после показа окна, и до UI-6
        /// получали индекс после шапки — к ним шли через выбор языка.
        /// </summary>
        [Test("Аудит UI-6: карточки недавних в порядке Tab между настройками и «О программе»")]
        private static void AuditRecentCardsTabOrder()
        {
            var offenders = new List<string>();
            string dir = Path.Combine(Path.GetTempPath(), "iwo_ui6_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                InIsolatedSettings("iwo_ui6_s_", delegate
                {
                    for (int i = 0; i < 3; i++)
                    {
                        string file = Path.Combine(dir, "итог" + i + ".pdf");
                        MakeEmptyPagesPdf(file, 1);
                        OperationHistory.Record("pdf.merged", file);
                    }
                    using (var hub = new StartForm())
                    {
                        hub.Show();
                        PumpUntil(delegate { return CountRecentCards(hub) >= 3; }, 15000);
                        // Настоящий обход Tab (SelectNextControl), а не сортировка по TabIndex:
                        // при равных индексах порядок решает z-порядок, и сортировка его не видит.
                        var settings = (Control)typeof(StartForm).GetField("_settingsButton", AnyInstance).GetValue(hub);
                        var about = (Control)typeof(StartForm).GetField("_aboutButton", AnyInstance).GetValue(hub);
                        Control header = hub.Controls.Cast<Control>().First(c => c is HeaderBand);
                        List<Control> cards = hub.Controls.OfType<RecentCard>().Where(c => c.Visible)
                            .OrderBy(c => c.Left).Cast<Control>().ToList();
                        if (cards.Count != 3) offenders.Add("видимых карточек " + cards.Count);
                        Control at = settings;
                        foreach (Control want in new List<Control>(cards) { about })
                        {
                            hub.SelectNextControl(at, true, true, true, true);
                            Control got = hub.ActiveControl;
                            if (got != want)
                                offenders.Add("после «" + at.AccessibleName + "» Tab привёл к " +
                                    (got == null ? "null" : got.GetType().Name + " «" + got.AccessibleName + "»") +
                                    " вместо «" + want.AccessibleName + "»");
                            at = want;
                        }
                        if (hub.Controls.Cast<Control>().Any(c => c != header && c.TabIndex >= header.TabIndex))
                            offenders.Add("шапка не последняя");
                        hub.Close();
                    }
                });
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
            AssertTrue(offenders.Count == 0, "порядок Tab хаба: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>
        /// Цвет палитры задаётся один раз — в Theme. Копия числами в другом файле расходится с
        /// палитрой при первой её правке (UI-21: красный значка PDF, серые рамки, цвет
        /// заголовков окон в пяти местах). Совпадение по смыслу разное — заведите в Theme имя.
        /// </summary>
        [Test("Аудит UI-21: цвета палитры не повторяются числами вне Theme")]
        private static void AuditNoThemeColorCopies()
        {
            string src = SourceDir();
            AssertTrue(src != null, "каталог src найден");
            var palette = new Dictionary<int, string>();
            foreach (System.Reflection.FieldInfo field in typeof(Theme).GetFields(
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static))
            {
                if (field.FieldType != typeof(Color))
                    continue;
                var color = (Color)field.GetValue(null);
                if (color.A == 255 && !palette.ContainsKey(color.ToArgb()))
                    palette[color.ToArgb()] = field.Name;
            }
            AssertTrue(palette.Count > 30, "палитра прочитана (" + palette.Count + " цветов)");
            var literal = new System.Text.RegularExpressions.Regex(
                @"Color\.FromArgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)");
            var offenders = new List<string>();
            foreach (string file in Directory.GetFiles(src, "*.cs", SearchOption.AllDirectories))
            {
                if (Path.GetFileName(file) == "Theme.cs")
                    continue;
                string[] lines = File.ReadAllLines(file);
                for (int i = 0; i < lines.Length; i++)
                    foreach (System.Text.RegularExpressions.Match m in literal.Matches(lines[i]))
                    {
                        int argb = Color.FromArgb(int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value),
                            int.Parse(m.Groups[3].Value)).ToArgb();
                        string name;
                        if (palette.TryGetValue(argb, out name))
                            offenders.Add(Path.GetFileName(file) + ":" + (i + 1) + " = Theme." + name);
                    }
            }
            AssertTrue(offenders.Count == 0, "копии цветов Theme: " + string.Join(", ", offenders.ToArray()));
            using (var host = new Panel())
            {
                Label title = Ui.Title(host, "Заголовок", 0, 0);
                AssertEq(Theme.TitleText, title.ForeColor, "цвет заголовка окна");
                AssertTrue(title.Font.Bold && Math.Abs(title.Font.SizeInPoints - 14f) < 0.01f, "кегль заголовка — 14, жирный");
            }
        }
        /// <summary>
        /// Давность автоочистки выбирают в двух окнах — истории (Настройки) и статистики. Список
        /// «пункт — дни» один (AutoClearCombo); проверяем его и то, что оба окна им пользуются
        /// по-настоящему: показывают сохранённое и сохраняют выбранное.
        /// </summary>
        [Test("Аудит UI-21: давность автоочистки — один список на Настройки и Статистику")]
        private static void AuditAutoClearChoicesShared()
        {
            using (var combo = new AutoClearCombo())
            {
                AssertEq(4, combo.Items.Count, "пунктов");
                AssertEq(Loc.T("stats.auto.30days"), (string)combo.Items[3], "последний пункт — 30 дней");
                foreach (int days in new[] { 0, 1, 7, 30 })
                {
                    combo.Days = days;
                    AssertEq(days, combo.Days, "давность " + days + " туда и обратно");
                }
                combo.Days = 5;
                AssertEq(0, combo.SelectedIndex, "неизвестная давность — «выключена»");
                AssertEq(0, combo.Days, "и читается как 0");
            }
            InIsolatedSettings("iwo_ui21_auto_", delegate
            {
                AssertTrue(OperationHistory.SetAutoClear(7), "давность истории сохранена");
                AssertTrue(UsageStats.SetAutoClear(30), "давность статистики сохранена");
                using (var settings = new SettingsForm())
                using (var stats = new StatsForm())
                {
                    var history = (AutoClearCombo)typeof(SettingsForm).GetField("_historyAuto", AnyInstance).GetValue(settings);
                    var counters = (AutoClearCombo)typeof(StatsForm).GetField("_cmbAuto", AnyInstance).GetValue(stats);
                    AssertEq(7, history.Days, "Настройки показывают сохранённую давность истории");
                    AssertEq(30, counters.Days, "Статистика показывает сохранённую давность счётчиков");
                    // Дни, не равные номеру пункта: сохранить номер вместо дней — тоже ошибка.
                    history.Days = 30;
                    counters.Days = 7;
                    AssertEq(30, OperationHistory.Load().AutoClearDays, "выбор в Настройках сохранён");
                    AssertEq(7, UsageStats.Load().AutoClearDays, "выбор в Статистике сохранён");
                }
            });
        }
        /// <summary>
        /// Реквизиты в «Что нового» и «О программе» — одно поле (Ui.SelectableValue): только
        /// для чтения, по Tab, на фоне своего блока. Раньше у каждого окна была своя копия.
        /// </summary>
        [Test("Аудит UI-21: поля реквизитов одинаковы в «О программе» и «Что нового»")]
        private static void AuditDonationFieldsShared()
        {
            InIsolatedSettings("iwo_ui21_don_", delegate
            {
                using (var about = new AboutForm())
                using (var news = new WhatsNewForm(WhatsNewCatalog.CurrentVersion))
                {
                    foreach (Panel panel in new[] { about.SupportPanel, news.SupportPanel })
                        foreach (string value in new[] { AboutForm.DonationAccount, AboutForm.DonationBank })
                        {
                            TextBox box = panel.Controls.OfType<TextBox>().FirstOrDefault(t => t.Text == value);
                            string where = panel.FindForm().GetType().Name + " «" + value + "»";
                            AssertTrue(box != null, where + ": поле есть");
                            AssertTrue(box.ReadOnly && box.TabStop && box.BorderStyle == BorderStyle.None, where + ": только чтение, по Tab, без рамки");
                            AssertEq(panel.BackColor, box.BackColor, where + ": фон блока");
                        }
                }
            });
        }
        /// <summary>Цвета фона, текста и ссылки каждого контрола окна — по самому контролу.</summary>
        private static Dictionary<Control, Color[]> ColorsOf(Control root)
        {
            var into = new Dictionary<Control, Color[]>();
            var stack = new Stack<Control>();
            stack.Push(root);
            while (stack.Count > 0)
            {
                Control c = stack.Pop();
                var link = c as LinkLabel;
                into[c] = new[] { c.BackColor, c.ForeColor, link == null ? Color.Empty : link.LinkColor };
                foreach (Control child in c.Controls)
                    stack.Push(child);
            }
            return into;
        }
        /// <summary>В высоком контрасте у каждого контрола окна фон, текст и ссылка — системные.</summary>
        private static void CheckSystemColors(Control root, string where, List<string> offenders)
        {
            int reported = 0;
            foreach (KeyValuePair<Control, Color[]> c in ColorsOf(root))
                for (int k = 0; k < 3; k++)
                    if (!c.Value[k].IsEmpty && !c.Value[k].IsSystemColor && reported++ < 3)
                        offenders.Add(where + " → " + c.Key.GetType().Name + " «" + c.Key.Text + "»: " +
                            new[] { "фон", "текст", "ссылка" }[k] + " " + c.Value[k]);
        }
        /// <summary>
        /// Высокий контраст (UI-7). Окна задают цвета явно, а явный цвет в WinForms главнее
        /// системного: в чёрной теме окно оставалось белым. Проверяем каждое окно приложения:
        /// режим, включённый при открытом окне, делает все цвета системными — и у контрола,
        /// добавленного позже, и у подписи, перекрашенной позже; выключенный режим возвращает
        /// окну его цвета, причём перекрашенной подписи — последний заданный. И окно, открытое
        /// уже в режиме, сразу системное.
        /// </summary>
        [Test("Аудит UI-7: в высоком контрасте окна берут системные цвета и возвращают свои после")]
        private static void AuditWindowsFollowHighContrast()
        {
            var offenders = new List<string>();
            InIsolatedSettings("iwo_ui7_", delegate
            {
                foreach (Form f in ScaledWindows(offenders))
                    using (f)
                    {
                        string name = f.GetType().Name;
                        try
                        {
                            f.Show();
                            Dictionary<Control, Color[]> was = ColorsOf(f);
                            Label repainted = was.Keys.OfType<Label>().FirstOrDefault(l => !(l is LinkLabel));
                            Theme.HighContrastForTests = true;
                            Raise(f, "OnSystemColorsChanged", EventArgs.Empty);
                            CheckSystemColors(f, name + " (режим включён)", offenders);
                            var late = new Label { Text = "добавлена в режиме", ForeColor = Theme.ErrRed, BackColor = Color.White };
                            f.Controls.Add(late);
                            if (repainted != null)
                                repainted.ForeColor = Theme.OkGreen;
                            CheckSystemColors(f, name + " (после добавления и перекраски)", offenders);
                            Theme.HighContrastForTests = false;
                            Raise(f, "OnSystemColorsChanged", EventArgs.Empty);
                            foreach (KeyValuePair<Control, Color[]> c in ColorsOf(f))
                            {
                                Color[] before;
                                if (c.Key == late || c.Key == repainted || !was.TryGetValue(c.Key, out before))
                                    continue;
                                if (!before.SequenceEqual(c.Value))
                                {
                                    offenders.Add(name + " → " + c.Key.GetType().Name + " «" + c.Key.Text +
                                        "»: после режима " + string.Join("/", c.Value.Select(v => v.Name).ToArray()) +
                                        " вместо " + string.Join("/", before.Select(v => v.Name).ToArray()));
                                    break;
                                }
                            }
                            if (late.ForeColor != Theme.ErrRed || late.BackColor != Color.White)
                                offenders.Add(name + ": добавленная в режиме подпись не вернула свои цвета");
                            if (repainted != null && repainted.ForeColor != Theme.OkGreen)
                                offenders.Add(name + ": перекрашенная в режиме подпись вернула не последний цвет, а " + repainted.ForeColor);
                        }
                        catch (Exception ex) { offenders.Add(name + " — " + Root(ex).Message); }
                        finally { Theme.HighContrastForTests = null; f.Close(); }
                    }
                Theme.HighContrastForTests = true;
                try
                {
                    foreach (Form f in ScaledWindows(offenders))
                        using (f)
                        {
                            f.Show();
                            CheckSystemColors(f, f.GetType().Name + " (открыто в режиме)", offenders);
                            f.Close();
                        }
                }
                finally { Theme.HighContrastForTests = null; }
            });
            AssertTrue(offenders.Count == 0, "высокий контраст окон: " + string.Join(" | ", offenders.ToArray()));
        }
        /// <summary>
        /// Рисованные контролы в высоком контрасте рисуют только системными цветами (UI-7):
        /// палитра приложения — акцентный зелёный, серые рамки, тёмно-серый текст — уступает
        /// цветам, которые выбрал человек. Каждый контрол рисуется дважды, без режима и в нём.
        /// Цвета, занявшие заметную долю картинки без режима и не системные (заливки, рамки и
        /// их сглаживание, текст палитры), в режиме встречаться не должны. Сравнивать с
        /// перечнем системных цветов напрямую нельзя: серая рамка палитры лежит на той же
        /// прямой, что и сглаживание чёрного по светлому. Значок документа на карточках —
        /// рисунок содержимого и остаётся в своих цветах.
        /// </summary>
        [Test("Аудит UI-7: рисованные контролы в высоком контрасте — только системные цвета")]
        private static void AuditOwnerDrawnFollowHighContrast()
        {
            var offenders = new List<string>();
            var allowed = new HashSet<int>(new[]
            {
                SystemColors.Control, SystemColors.ControlText, SystemColors.Window, SystemColors.WindowText,
                SystemColors.Highlight, SystemColors.HighlightText, SystemColors.GrayText
            }.Select(c => c.ToArgb()));
            var glyphInk = new HashSet<int>(new[] { Theme.PdfRed, Theme.PdfRedDark, Color.White }.Select(c => c.ToArgb()));
            Type glyphType = typeof(StartForm).GetNestedType("GlyphButton", System.Reflection.BindingFlags.NonPublic);
            Func<Control> globe = delegate
            {
                var g = (Control)Activator.CreateInstance(glyphType, new object[] { "A", 15f, "Segoe UI", null });
                g.ForeColor = Theme.SettingsBlue; // как значок настроек: цвет и подложка наведения
                glyphType.GetField("HoverFill").SetValue(g, Color.FromArgb(30, Theme.SettingsBlue));
                g.Size = new Size(30, 30);
                return g;
            };
            Action<Control> idle = delegate { };
            Action<Control> hover = delegate(Control c) { Raise(c, "OnMouseEnter", EventArgs.Empty); };
            Action<Control> focus = delegate(Control c) { c.Focus(); };
            Action<Control> disabled = delegate(Control c) { c.Enabled = false; };
            var subjects = new List<Tuple<string, Func<Control>, Action<Control>, bool>>
            {
                Tuple.Create("главная кнопка", (Func<Control>)(() => new RoundedButton(ButtonLook.Primary) { Text = "Сохранить", Size = new Size(140, 36) }), idle, false),
                Tuple.Create("главная кнопка под мышью", (Func<Control>)(() => new RoundedButton(ButtonLook.Primary) { Text = "Сохранить", Size = new Size(140, 36) }), hover, false),
                Tuple.Create("обычная кнопка в фокусе", (Func<Control>)(() => new RoundedButton(ButtonLook.Secondary) { Text = "Открыть", Size = new Size(140, 36) }), focus, false),
                Tuple.Create("недоступная кнопка", (Func<Control>)(() => new RoundedButton(ButtonLook.Secondary) { Text = "Открыть", Size = new Size(140, 36) }), disabled, false),
                Tuple.Create("кнопка на тёмном", (Func<Control>)(() => new RoundedButton(ButtonLook.OnDark) { Text = "+", Size = new Size(36, 36) }), idle, false),
                Tuple.Create("флажок", (Func<Control>)(() => new AccentCheckBox { Text = "Флажок", Size = new Size(160, 30) }), idle, false),
                Tuple.Create("отмеченный флажок в фокусе", (Func<Control>)(() => new AccentCheckBox { Text = "Флажок", Checked = true, Size = new Size(160, 30) }), focus, false),
                Tuple.Create("карточка инструмента", (Func<Control>)(() => new ChoiceCard(CardGlyph.Pdf, "Карточка", "Описание карточки") { Size = new Size(240, 250) }), idle, true),
                Tuple.Create("карточка под мышью в фокусе", (Func<Control>)(() => new ChoiceCard(CardGlyph.Pdf, "Карточка", "Описание карточки") { Size = new Size(240, 250) }),
                    (Action<Control>)(c => { focus(c); hover(c); }), true),
                Tuple.Create("карточка недавнего", (Func<Control>)(() => new RecentCard(new HistoryEntry { Operation = "pdf.merged", Path = @"C:\нет\документ.pdf" }) { Size = new Size(200, 36) }), hover, true),
                Tuple.Create("кнопка-значок под мышью", globe, hover, false),
                Tuple.Create("кнопка-значок в фокусе", globe, focus, false),
                Tuple.Create("шапка", (Func<Control>)(() => new HeaderBand("Заголовок окна", "Подпись", Theme.HubBlue, Theme.HubBlueDark) { Size = new Size(400, 70) }), idle, false)
            };
            try
            {
                foreach (var subject in subjects)
                {
                    Theme.HighContrastForTests = false;
                    Dictionary<int, int> normal = DominantInks(subject.Item2, subject.Item3);
                    Theme.HighContrastForTests = true;
                    Dictionary<int, int> contrast = DominantInks(subject.Item2, subject.Item3);
                    bool any = false;
                    foreach (KeyValuePair<int, int> ink in contrast)
                    {
                        if (allowed.Contains(ink.Key) && ink.Key != SystemColors.Control.ToArgb())
                            any = true; // нарисовано хоть что-то системным цветом, кроме фона
                        if (normal.ContainsKey(ink.Key) && !allowed.Contains(ink.Key) &&
                            !(subject.Item4 && glyphInk.Contains(ink.Key)))
                            offenders.Add(subject.Item1 + ": цвет палитры " + Color.FromArgb(ink.Key) + " × " + ink.Value);
                    }
                    if (!any)
                        offenders.Add(subject.Item1 + ": в режиме нет ни одного системного цвета, кроме фона");
                }
            }
            finally { Theme.HighContrastForTests = null; }
            AssertTrue(offenders.Count == 0, "рисованные контролы в высоком контрасте: " + string.Join(" | ", offenders.ToArray()));
        }
        /// <summary>
        /// Цвета, занявшие заметную долю картинки контрола (не меньше 0,5 % и не меньше 12 px):
        /// заливки, рамки, кольца, текст. Контрол стоит в окне приложения (ScaledForm) с фоном
        /// системного цвета, и рисуется в нужном состоянии. Фон один и тот же в обоих режимах:
        /// полупрозрачная подложка палитры над разным фоном дала бы разные смеси, и сравнение
        /// не узнало бы в смеси режима ту же подложку (сброс белого фона окна проверяет тест окон).
        /// </summary>
        private static Dictionary<int, int> DominantInks(Func<Control> make, Action<Control> state)
        {
            Control c = make();
            using (var host = new ScaledForm { ClientSize = new Size(c.Width + 40, c.Height + 40), ShowInTaskbar = false })
            {
                c.Location = new Point(20, 20);
                host.Controls.Add(c);
                host.Show();
                state(c);
                Application.DoEvents();
                var counts = new Dictionary<int, int>();
                using (Bitmap bmp = Snapshot(c))
                {
                    for (int y = 0; y < bmp.Height; y++)
                        for (int x = 0; x < bmp.Width; x++)
                        {
                            int argb = bmp.GetPixel(x, y).ToArgb();
                            int n;
                            counts.TryGetValue(argb, out n);
                            counts[argb] = n + 1;
                        }
                    int least = Math.Max(12, bmp.Width * bmp.Height / 200);
                    host.Close();
                    return counts.Where(p => p.Value >= least).ToDictionary(p => p.Key, p => p.Value);
                }
            }
        }
        /// <summary>Те же проверки целости, что у раскладочных тестов на 100 %.</summary>
        private static void CheckScaledLayout(Form f, string where, List<string> offenders)
        {
            CheckFits(f, where, offenders);
            CheckControlsDoNotOverlap(f, where, offenders);
            CheckButtonsDoNotOverlap(f, where, offenders);
        }

        /// <summary>
        /// Клиентская область выросла в scale раз (±8 %) против той же раскладки на 100 % —
        /// если такое окно помещается на экран; иначе его законно ужал FitToScreen. Возвращает,
        /// поместилось ли окно целиком: только тогда с ним сравнимы и места контролов.
        /// </summary>
        private static bool CheckGrewWithScale(Form f, Size b, float scale,
            Rectangle area, string where, List<string> offenders)
        {
            if (b.Width <= 0 || b.Height <= 0)
                return false;
            Size frame = f.Size - f.ClientSize;
            bool fitsW = b.Width * scale + frame.Width <= area.Width;
            bool fitsH = b.Height * scale + frame.Height <= area.Height;
            double rw = f.ClientSize.Width / (b.Width * scale);
            double rh = f.ClientSize.Height / (b.Height * scale);
            // Допуск несимметричный. Вниз — 8 %: на 96 DPI шов растит шрифт, но не системные
            // метрики (рамка поля ввода, квадрат флажка), и поля выходят на 5 % ниже, чем на
            // настоящем мониторе. Вверх — 4,5 %: лишний рост — это и есть след двойного масштаба.
            if (fitsW && (rw < 0.92 || rw > 1.045))
                offenders.Add(where + " → ширина " + f.ClientSize.Width + " вместо ~" + (int)(b.Width * scale));
            if (fitsH && (rh < 0.92 || rh > 1.045))
                offenders.Add(where + " → высота " + f.ClientSize.Height + " вместо ~" + (int)(b.Height * scale));
            return fitsW && fitsH;
        }

        /// <summary>
        /// Место каждого контрола — от того края родителя, к которому он привязан (у правого
        /// и нижнего якоря — расстояние до правого/нижнего края). Путь — индексы в Controls:
        /// окна строятся одинаково при любом масштабе, поэтому пути совпадают.
        /// </summary>
        private static Dictionary<string, int[]> CollectEdges(Control root)
        {
            var into = new Dictionary<string, int[]>();
            CollectEdges(root, "", into);
            return into;
        }

        /// <summary>Место (от своего якоря), размер и кегль ×100 каждого контрола.</summary>
        private static void CollectEdges(Control parent, string path, Dictionary<string, int[]> into)
        {
            for (int i = 0; i < parent.Controls.Count; i++)
            {
                Control c = parent.Controls[i];
                // Высоту регулятора задают системные метрики, а не шрифт: шов масштаба (96 DPI)
                // её не растит, хотя на настоящем мониторе 150 % она выросла бы. Детей таблицы
                // и стрелки поля-счётчика раскладывают сами таблица и поле — по своим стилям и
                // системным метрикам, сравнивать их места незачем.
                if (c is TrackBar || parent is TableLayoutPanel || parent is UpDownBase)
                    continue;
                string p = path + "/" + i + c.GetType().Name;
                bool fromRight = (c.Anchor & AnchorStyles.Right) != 0 && (c.Anchor & AnchorStyles.Left) == 0;
                bool fromBottom = (c.Anchor & AnchorStyles.Bottom) != 0 && (c.Anchor & AnchorStyles.Top) == 0;
                into[p] = new[]
                {
                    fromRight ? parent.ClientSize.Width - c.Right : c.Left,
                    fromBottom ? parent.ClientSize.Height - c.Bottom : c.Top,
                    c.Width, c.Height, (int)Math.Round(c.Font.Size * 100)
                };
                CollectEdges(c, p, into);
            }
        }

        /// <summary>
        /// Места и размеры контролов выросли вместе с масштабом, а кегль — ровно в scale раз.
        /// Размер окна ловит не всё: лишний отступ под каждым полем ввода, умноженный второй
        /// раз, менял высоту диалога на 10 %, а место последнего поля — на 15 %; подпись,
        /// ужатая шрифтом по замеру в чужих единицах, на размерах не видна вовсе.
        /// </summary>
        private static void CheckEdgesScaled(Form f, Dictionary<string, int[]> was,
            float scale, string where, List<string> offenders)
        {
            int reported = 0;
            foreach (KeyValuePair<string, int[]> now in CollectEdges(f))
            {
                int[] b;
                if (!was.TryGetValue(now.Key, out b))
                    continue;
                // Кегль — ровно в scale раз; разрешена одна ступень вниз (0,5 пункта): подпись,
                // ужимаемая шрифтом до ширины кнопки (SettingsForm.FitFont), на нелинейных
                // метриках шрифта может уступить одну ступень на одном масштабе и не на другом.
                double font = b[4] * scale;
                if ((now.Value[4] > font + 1 || now.Value[4] < font - 50 * scale - 1) && reported++ < 3)
                    offenders.Add(where + " → " + now.Key + ": кегль " + now.Value[4] / 100.0 +
                        " вместо " + b[4] * scale / 100.0);
                for (int k = 0; k < 4; k++)
                {
                    int[] axis = { b[k], now.Value[k] };
                    double expected = axis[0] * scale;
                    double drift = axis[1] - expected;
                    // Вниз шире, вверх уже — по той же причине, что у размера окна; вверх
                    // чуть шире, чем у окна: место кнопки складывается из нескольких замеров
                    // текста, и нелинейность метрик шрифта копится (длинная подпись — +6 %).
                    if (drift <= Math.Abs(expected) * 0.065 + 6 && -drift <= Math.Abs(expected) * 0.08 + 6)
                        continue;
                    if (reported++ < 3)
                        offenders.Add(where + " → " + now.Key + ": " + axis[1] + " вместо ~" + (int)expected);
                }
            }
        }

        /// <summary>
        /// Масштаб ставится перед созданием хэндла (<see cref="ScaledForm"/>), значит, ни одно
        /// окно не имеет права создать хэндл в конструкторе: тогда WinForms масштабировал бы
        /// недостроенное окно, и всё, что добавлено после, осталось бы в пикселях 96 DPI.
        /// Так и случилось при разработке: Screen.FromControl в раскладке «Свода Excel»
        /// создавал хэндл, и окно выходило в 1,44 раза шире. Заодно: каждое окно приложения
        /// наследует ScaledForm — новое окно, забывшее об этом, не масштабировалось бы вовсе.
        /// </summary>
        [Test("Аудит UI-1: окна масштабируются после постройки и все наследуют ScaledForm")]
        private static void AuditFormsScaleAfterConstruction()
        {
            var offenders = new List<string>();
            foreach (Type t in typeof(ScaledForm).Assembly.GetTypes())
            {
                // Свои окна тестов (зонды раскладки) — не окна приложения.
                if (!typeof(Form).IsAssignableFrom(t) || t == typeof(ScaledForm) || t.Namespace == "ExcelMerger.Tests")
                    continue;
                // Предпросмотр страницы считает размер от рабочей области сам (AutoScaleMode.None).
                if (t == typeof(PagePreviewForm))
                    continue;
                if (!typeof(ScaledForm).IsAssignableFrom(t))
                    offenders.Add(t.Name + " не наследует ScaledForm");
            }
            InIsolatedSettings("iwo_dpi_ctor_", delegate
            {
                foreach (Form f in ScaledWindows(offenders))
                    using (f)
                    {
                        var scaled = f as ScaledForm;
                        if (f.IsHandleCreated || (scaled != null && scaled.IsScaled))
                            offenders.Add(f.GetType().Name + ": хэндл создан в конструкторе");
                        if (f.AutoScaleMode != AutoScaleMode.Inherit)
                            offenders.Add(f.GetType().Name + ": масштаб задан до постройки (" + f.AutoScaleMode + ")");
                        IntPtr handle = f.Handle;
                        if (scaled == null || !scaled.IsScaled || f.AutoScaleMode != AutoScaleMode.Dpi)
                            offenders.Add(f.GetType().Name + ": после создания хэндла масштаб не включён");
                    }
            });
            AssertTrue(offenders.Count == 0, "масштаб окон: " + string.Join(" | ", offenders.ToArray()));
        }

        /// <summary>Зонд окна заданной клиентской области с изменяемой рамкой.</summary>
        private sealed class SizeProbe : ScaledForm
        {
            internal SizeProbe(Size client)
            {
                FormBorderStyle = FormBorderStyle.Sizable;
                StartPosition = FormStartPosition.Manual;
                ShowInTaskbar = false;
                ClientSize = client;
            }
        }
        /// <summary>
        /// Окно шире предельного размера Windows (MaxWindowTrackSize) WinForms ужимает в Size ещё
        /// до хэндла, а ClientSize оставляет заданным. FitToScreen считал рамку как Size − ClientSize,
        /// получал отрицательную и недоужимал окно: на CI (экран 1024) «Что нового» шириной 1193
        /// открывалось на 1044 px. Здесь предел — настоящего экрана, а рабочая область — 1024×720.
        /// </summary>
        [Test("Аудит UI-1: окно шире предельного размера Windows всё равно вписывается в экран")]
        private static void AuditFitsScreenBeyondMaxTrackSize()
        {
            Size max = SystemInformation.MaxWindowTrackSize;
            var area = new Rectangle(0, 0, 1024, 720);
            ScaledForm.WorkingAreaForTests = delegate(Rectangle near) { return area; };
            try
            {
                foreach (float scale in new[] { 1f, 1.5f })
                {
                    DpiScale.EmulatedScale = scale;
                    try
                    {
                        using (var f = new SizeProbe(new Size(max.Width + 300, 300)))
                        {
                            AssertTrue(f.Width <= max.Width && f.ClientSize.Width > max.Width,
                                "предпосылка: Size ужат до предела Windows, ClientSize — нет");
                            f.Show();
                            AssertTrue(f.Width <= area.Width && f.Height <= area.Height,
                                "окно " + f.Size + " вписано в " + area.Size + " при " + scale);
                            f.Close();
                        }
                    }
                    finally { DpiScale.EmulatedScale = 1f; }
                }
            }
            finally { ScaledForm.WorkingAreaForTests = null; }
        }
        /// <summary>Зонд порядка: раскладка, увидевшая IsScaled раньше самого масштаба.</summary>
        private sealed class ScaleOrderProbe : ScaledForm
        {
            internal int EarlyLayouts;

            internal ScaleOrderProbe()
            {
                ClientSize = new Size(300, 200);
                Controls.Add(new Button { Bounds = new Rectangle(10, 10, 100, 30) });
            }

            protected override void OnLayout(LayoutEventArgs levent)
            {
                base.OnLayout(levent);
                // Закончив масштаб, WinForms приравнивает AutoScaleDimensions к текущим;
                // до второй настройки режим ещё не Dpi.
                if (IsScaled && (AutoScaleMode != AutoScaleMode.Dpi ||
                                 AutoScaleDimensions != CurrentAutoScaleDimensions))
                    EarlyLayouts++;
            }
        }

        /// <summary>
        /// Наследник ScaledForm переводит константы по IsScaled, поэтому IsScaled обязан
        /// становиться истинным не раньше самого масштаба. Две настройки автомасштаба
        /// запускают раскладку каждая; без SuspendLayout вокруг них раскладка шла между ними —
        /// с IsScaled = true на ещё логическом окне, и всё, что она ставила в пикселях экрана,
        /// WinForms затем умножал ещё раз (то, что не пересчитывается заново, так и оставалось).
        /// </summary>
        [Test("Аудит UI-1: раскладка не видит IsScaled раньше самого масштаба")]
        private static void AuditScaledFlagFollowsScale()
        {
            DpiScale.EmulatedScale = 1.5f;
            try
            {
                using (var probe = new ScaleOrderProbe())
                {
                    IntPtr handle = probe.Handle;
                    AssertEq(0, probe.EarlyLayouts, "раскладок с IsScaled до масштаба");
                    AssertEq(new Size(450, 300), probe.ClientSize, "окно выросло ровно в 1,5 раза");
                    AssertEq(new Rectangle(15, 15, 150, 45), probe.Controls[0].Bounds, "и его содержимое тоже");
                }
            }
            finally { DpiScale.EmulatedScale = 1f; }
        }

        /// <summary>
        /// Переводы единиц и вписывание окна в экран — чистые функции, проверяем числами:
        /// именно в них сидят ошибки на полпикселя, которые потом обрезают последнюю букву.
        /// </summary>
        [Test("Аудит UI-1: переводы логических единиц и вписывание окна в экран")]
        private static void AuditDpiScaleArithmetic()
        {
            // 150 % на 96 DPI: шов умножает системный DPI.
            DpiScale.EmulatedScale = 1.5f;
            try
            {
                AssertEq(100, DpiScale.ToLogical(150), "150 px экрана — 100 логических");
                AssertEq(101, DpiScale.ToLogical(151), "с запасом вверх: место не меньше замеренного");
                AssertEq(new Size(20, 14), DpiScale.ToLogical(new Size(30, 21)), "размер переводится по обеим осям");
                using (var free = new Label())
                {
                    AssertEq(30, DpiScale.Px(free, 20), "контрол вне окна приложения — уже в пикселях экрана");
                    AssertEq(45, DpiScale.FromDevice(free, 45), "замер для такого контрола не переводится");
                }
                using (var form = new AboutForm())
                {
                    AssertEq(20, DpiScale.Px(form, 20), "окно до масштаба — логическое");
                    AssertEq(30, DpiScale.FromDevice(form, 45), "замер в конструкторе переводится в логические");
                    AssertEq(450, DpiScale.MeasureWidth(form, 300), "перенос текста меряется на ширине экрана");
                    IntPtr handle = form.Handle;
                    AssertEq(30, DpiScale.Px(form, 20), "после масштаба — пиксели экрана");
                    AssertEq(45, DpiScale.FromDevice(form, 45), "после масштаба замер как есть");
                    AssertEq(300, DpiScale.MeasureWidth(form, 300), "после масштаба ширина как есть");
                }
            }
            finally { DpiScale.EmulatedScale = 1f; }

            var frame = new Size(16, 39);
            var area = new Size(1920, 1032);
            AssertEq(new Size(780, 660), ScaledForm.FitClient(new Size(780, 660), new Size(660, 540), frame, area, 1.5f),
                "помещается — размер не трогаем");
            AssertEq(new Size(780, 662), ScaledForm.FitClient(new Size(780, 785), new Size(760, 0), frame, area, 1.5f),
                "выше экрана: (1032 − 39) / 1,5 = 662 логических");
            AssertEq(new Size(780, 691), ScaledForm.FitClient(new Size(780, 785), new Size(820, 730), frame, area, 1.5f),
                "не ниже минимума окна, даже если экран меньше");
            AssertEq(new Size(1269, 500), ScaledForm.FitClient(new Size(1500, 500), Size.Empty, frame, area, 1.5f),
                "по ширине так же: (1920 − 16) / 1,5 = 1269");
        }
    }
}
