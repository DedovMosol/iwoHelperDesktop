using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;

// Снимки РЕАЛЬНЫХ окон приложения для инструкции. Грузим собранный exe как сборку,
// строим окна, наполняем их документами-образцами и печатаем окно в PNG.
// Окна с раскрытым меню снимаем с экрана: выпадающее меню — отдельное окно Windows,
// в PrintWindow формы оно не попадает.
static class Shots
{
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern void mouse_event(uint flags, uint dx, uint dy,
        uint data, UIntPtr extraInfo);
    [DllImport("user32.dll")] static extern uint GetGuiResources(IntPtr process,
        uint flags);

    static Assembly _app;
    static Type _loc, _lang;
    static string _out, _samples;
    static string _reviewLeft, _reviewRight;
    const string ReviewFixture1Sha256 =
        "21a45a5d7b943e919ae3d764d31526fb026a870410e22b77d185dd3ea79ed9b5";
    const string ReviewFixture2Sha256 =
        "ac3a87fed669596e417de82fc7cbd0707083c51bbe09638e199808658a7e96e7";
    const long ReviewFixture1Bytes = 219873;
    const long ReviewFixture2Bytes = 198045;
    static int _failures;
    static bool _en; // язык набора снимков: образцы-заполнители тоже должны быть на нём
    static readonly List<string> _log = new List<string>();
    static Rectangle _popup = Rectangle.Empty; // экранные границы раскрытого меню
    static Form _backdrop;

    // Имена образцов ищем ПО МАСКЕ, а не по литералу: у английского набора они английские,
    // и держать здесь вторую таблицу имён значило бы завести ещё одно место, где они врозь.
    static string Sample(string ruStart, string enStart)
    {
        foreach (string path in Directory.GetFiles(_samples, "*.pdf"))
        {
            string name = Path.GetFileName(path);
            if (name.StartsWith(ruStart, StringComparison.OrdinalIgnoreCase)
                || name.StartsWith(enStart, StringComparison.OrdinalIgnoreCase))
                return path;
        }
        throw new FileNotFoundException("нет образца " + ruStart + " / " + enStart + " в " + _samples);
    }

    static string Doc { get { return Sample("Образец", "Sample"); } }
    static string Photo { get { return SampleImage("Снимок", "Page photo"); } }
    /// <summary>Образец-снимок (его добавляют страницей) — по той же маске имени, что и PDF.</summary>
    static string SampleImage(string ruStart, string enStart)
    {
        foreach (string path in Directory.GetFiles(_samples, "*.jpg"))
        {
            string name = Path.GetFileName(path);
            if (name.StartsWith(ruStart, StringComparison.OrdinalIgnoreCase)
                || name.StartsWith(enStart, StringComparison.OrdinalIgnoreCase))
                return path;
        }
        throw new FileNotFoundException("нет образца-снимка " + ruStart + " / " + enStart + " в " + _samples);
    }

    static string AppA { get { return Sample("Приложение А", "Appendix A"); } }
    static string AppB { get { return Sample("Приложение Б", "Appendix B"); } }
    static string XlDir
    {
        get
        {
            foreach (string dir in Directory.GetDirectories(_samples))
                return dir; // папка книг в наборе одна, как бы она ни называлась
            throw new DirectoryNotFoundException("нет папки книг Excel в " + _samples);
        }
    }

    [STAThread]
    static int Main(string[] args)
    {
        if (args == null || args.Length < 3)
        {
            Console.Error.WriteLine("Usage: Shots.exe <app.exe> <out-dir> <samples-dir> [ru|en] [scenario ...]");
            return 2;
        }

        _app = Assembly.LoadFrom(args[0]);
        _out = args[1];
        _samples = args[2];
        _reviewLeft = Environment.GetEnvironmentVariable("IWO_REVIEW_LEFT");
        _reviewRight = Environment.GetEnvironmentVariable("IWO_REVIEW_RIGHT");
        // Язык интерфейса на снимках: в английском руководстве русские окна выглядели бы
        // так же неуместно, как английские в русском. По умолчанию — русский, как было.
        string lang = args.Length > 3 ? args[3] : "ru";
        _en = lang.StartsWith("en", StringComparison.OrdinalIgnoreCase);
        Directory.CreateDirectory(_out);

        _app.GetType("ExcelMerger.AppPaths")
            .GetMethod("SetRootForTests", BindingFlags.NonPublic | BindingFlags.Static)
            .Invoke(null, new object[] { Path.Combine(_out, "appdata") });

        _loc = _app.GetType("ExcelMerger.Loc");
        _lang = _app.GetType("ExcelMerger.Lang");
        object language = _loc.GetMethod("Parse", new[] { typeof(string) }).Invoke(null, new object[] { lang });
        _loc.GetMethod("Init", new[] { _lang }).Invoke(null, new[] { language });

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        // Нейтральная подложка во весь экран: на снимках не должно быть видно ни рабочего
        // стола, ни чужих окон — инструкцию будут читать посторонние люди.
        _backdrop = new Form();
        _backdrop.FormBorderStyle = FormBorderStyle.None;
        _backdrop.BackColor = Color.FromArgb(232, 234, 238);
        _backdrop.ShowInTaskbar = false;
        _backdrop.StartPosition = FormStartPosition.Manual;
        _backdrop.Bounds = Screen.PrimaryScreen.Bounds;
        _backdrop.Show();

        var only = args.Length > 4 ? new List<string>(args).GetRange(4, args.Length - 4) : new List<string>();
        int selected = 0;
        foreach (KeyValuePair<string, Action> s in Scenarios())
        {
            if (only.Count > 0 && !only.Contains(s.Key)) continue;
            // Локальные 1.pdf/2.pdf не входят в набор руководства и не должны становиться
            // его неявной зависимостью. Review запускается явно либо когда переданы оба пути.
            if (s.Key == "review" && only.Count == 0 && !ReviewInputsAvailable()) continue;
            selected++;
            try { s.Value(); }
            catch (Exception ex)
            {
                _failures++;
                Exception root = RootCause(ex);
                Say(s.Key + " -> ERR " + root.GetType().Name + ": " + root.Message);
            }
        }
        if (only.Count > 0 && selected == 0)
        {
            _failures++;
            Say("ERR: ни один запрошенный сценарий не найден: " + string.Join(", ", only.ToArray()));
        }

        try { _backdrop.Close(); _backdrop.Dispose(); } catch { }
        string report = string.Join(Environment.NewLine, _log.ToArray());
        Console.WriteLine(report);
        File.WriteAllText(Path.Combine(_out, "shots.log"), report + Environment.NewLine,
            new UTF8Encoding(false));
        int exitCode = _failures == 0 ? 0 : 1;
        ExitLikeApplication(exitCode);
        return exitCode;
    }

    static void Say(string line) { _log.Add(line); }

    /// <summary>
    /// Сценарий грузит тот же Windows.Data.Pdf, что production GUI. Обычный возврат из Main
    /// снова включает известный аварийный native DLL_PROCESS_DETACH, поэтому ПОСЛЕ записи
    /// всех результатов завершаемся тем же проверенным путём, что и само приложение.
    /// </summary>
    static void ExitLikeApplication(int code)
    {
        Type fastExit = _app.GetType("ExcelMerger.FastExit");
        MethodInfo now = fastExit == null ? null : fastExit.GetMethod("Now",
            BindingFlags.Public | BindingFlags.Static);
        if (now != null)
            now.Invoke(null, new object[] { code });
    }

    // ---------- сценарии ----------

    static IEnumerable<KeyValuePair<string, Action>> Scenarios()
    {
        yield return S("hub", Hub);
        yield return S("hub-pdf", HubPdf);
        yield return S("hub-lang", HubLanguage);
        yield return S("excel", Excel);
        yield return S("excel-menu", ExcelMenu);
        yield return S("merge", Merge);
        yield return S("merge-ctx", MergeContextMenu);
        yield return S("merge-menu", MergeMenu);
        yield return S("merge-compress", MergeCompression);
        yield return S("split", Split);
        yield return S("split-modes", SplitModes);
        yield return S("split-ranges", SplitRanges);
        yield return S("split-everyn", SplitEveryN);
        yield return S("split-bookmarks", SplitBookmarks);
        yield return S("split-template", SplitTemplate);
        yield return S("ops", Ops);
        yield return S("ops-images", OpsImages);
        yield return S("ops-dpi", OpsDpi);
        yield return S("ocr", Ocr);
        yield return S("pptx", Pptx);
        yield return S("review", Review);
        yield return S("review-layout", ReviewLayout);
        yield return S("whatsnew", WhatsNew);
        yield return S("settings", Settings);
        yield return S("preview", Preview);
        yield return S("metadata", Metadata);
        yield return S("about", About);
        yield return S("stats", Stats);
        yield return S("help-merge", HelpMerge);
        yield return S("help-excel", HelpExcel);
        yield return S("help-split", HelpSplit);
        yield return S("help-ocr", HelpOcr);
        yield return S("shortcuts", Shortcuts);
        yield return S("goto", GoToPage);
    }

    static KeyValuePair<string, Action> S(string name, Action body)
    {
        return new KeyValuePair<string, Action>(name, body);
    }

    // Хаб с 1.17.9 держит один размер на всех уровнях и задаёт его сам — свой размер ему
    // не навязываем, иначе на снимке будет окно, которого пользователь никогда не увидит.
    static void Hub()
    {
        Form f = New("ExcelMerger.StartForm");
        Place(f, 0, 0);
        Shot(f, "hub");
        Kill(f);
    }

    /// <summary>Раздел PDF: шесть инструментов сеткой 3×2.</summary>
    static void HubPdf()
    {
        Form f = New("ExcelMerger.StartForm");
        Place(f, 0, 0);
        Type level = _app.GetType("ExcelMerger.HubLevel");
        f.GetType().GetMethod("ShowLevel", BindingFlags.NonPublic | BindingFlags.Instance)
            .Invoke(f, new[] { Enum.Parse(level, "Pdf") });
        Pump(600);
        Shot(f, "hub-pdf");
        Kill(f);
    }

    static void HubLanguage()
    {
        Form f = New("ExcelMerger.StartForm");
        Place(f, 0, 0);
        Control globe = FindByType(f, "GlyphButton");
        globe.GetType().GetMethod("InvokeOnClick", BindingFlags.NonPublic | BindingFlags.Instance)
            .Invoke(globe, new object[] { globe, EventArgs.Empty });
        Pump(600);
        // Границы раскрытого меню запоминаем ЯВНО: снимок берёт с экрана только их, а окно
        // печатает PrintWindow'ом. Без этого меню в кадр не попадёт вовсе — оно отдельное окно.
        foreach (Form open in Application.OpenForms)
            if (open != f && open != _backdrop && open.Visible)
                _popup = open.Bounds;
        if (_popup.IsEmpty)
        {
            var menu = (ContextMenuStrip)Field(f, "_langMenu");
            if (menu != null && menu.Visible)
                _popup = menu.Bounds;
        }
        ScreenShot(f, "hub-lang");
        Kill(f);
    }

    static void Excel()
    {
        Form f = Tool("ExcelMerger.MainForm");
        Place(f, 1080, 760);
        SetText(f, "_txtInput", XlDir);
        Pump(2500); // папка сканируется в фоне
        Shot(f, "excel");
        Kill(f);
    }

    static void ExcelMenu()
    {
        Form f = Tool("ExcelMerger.MainForm");
        Place(f, 1080, 760);
        SetText(f, "_txtInput", XlDir);
        Pump(2000);
        OpenMainMenu(f);
        ScreenShot(f, "excel-menu");
        Kill(f);
    }

    static void Merge()
    {
        Form f = Loaded("ExcelMerger.PdfMergeForm", new[] { Doc, AppA });
        Shot(f, "merge");
        Kill(f);
    }

    static void MergeContextMenu()
    {
        Form f = Loaded("ExcelMerger.PdfMergeForm", new[] { Doc });
        Control grid = FindByType(f, "PdfPageGrid");
        Select(grid, 0);
        Control list = FindByType(grid, "ScrollList");
        ShowMenu(list.ContextMenuStrip, list, new Point(90, 130));
        ScreenShot(f, "merge-ctx");
        Kill(f);
    }

    static void MergeMenu()
    {
        Form f = Loaded("ExcelMerger.PdfMergeForm", new[] { Doc, AppA });
        OpenMainMenu(f);
        ScreenShot(f, "merge-menu");
        Kill(f);
    }

    static void MergeCompression()
    {
        Form f = Loaded("ExcelMerger.PdfMergeForm", new[] { Doc });
        Control picker = FindByType(f, "CompressionPicker");
        var combo = (ComboBox)FindByType(picker, "ComboBox");
        DropDown(combo);
        ScreenShot(f, "merge-compress", false);
        Kill(f);
    }

    static void Split() { SplitShot("split", 0, null); }
    static void SplitRanges() { SplitShot("split-ranges", 1, "1-3, 5, 8-"); }
    static void SplitEveryN() { SplitShot("split-everyn", 2, null); }
    static void SplitBookmarks() { SplitShot("split-bookmarks", 3, null); }

    static void SplitShot(string name, int mode, string ranges)
    {
        Form f = Loaded("ExcelMerger.PdfSplitForm", new[] { Doc });
        if (mode > 0)
        {
            var combo = (ComboBox)Field(f, "_cmbMode");
            combo.SelectedIndex = mode;
            Pump(400);
            // v1.18.4: _txtRanges заменён на _cmbRanges (ComboBox с presets)
            if (ranges != null)
            {
                var cmbRanges = (ComboBox)Field(f, "_cmbRanges");
                cmbRanges.Text = ranges;
            }
            Pump(300);
        }
        else
        {
            Control grid = FindByType(f, "PdfPageGrid");
            Select(grid, 0, 1, 2);
            Pump(300);
        }
        Shot(f, name);
        Kill(f);
    }

    static void SplitModes()
    {
        Form f = Loaded("ExcelMerger.PdfSplitForm", new[] { Doc });
        var combo = (ComboBox)Field(f, "_cmbMode");
        DropDown(combo);
        ScreenShot(f, "split-modes", false);
        Kill(f);
    }

    /// <summary>«Прочие операции» — с 1.17.9 отдельное окно, а не меню внутри разделения.</summary>
    static void Ops()
    {
        Form f = Loaded("ExcelMerger.PdfOpsForm", new[] { Doc }, 0, 0);
        Shot(f, "ops");
        Kill(f);
    }

    /// <summary>
    /// Картинка, добавленная страницей к страницам документа (с 1.18.1). Зовём тот же метод,
    /// что и кнопка «Добавить картинки…», — снимок должен показывать настоящий путь, а не
    /// подстроенную сетку.
    /// </summary>
    static void OpsImages()
    {
        Form f = Loaded("ExcelMerger.PdfOpsForm", new[] { Doc }, 0, 0);
        f.GetType().GetMethod("AddImages", BindingFlags.Instance | BindingFlags.NonPublic)
            .Invoke(f, new object[] { new[] { Photo } });
        Pump(4000); // обёртывание картинки и её миниатюра идут в фоне
        Shot(f, "ops-images");
        Kill(f);
    }

    /// <summary>Выбор разрешения при сохранении страниц картинками.</summary>
    static void OpsDpi()
    {
        Form f = Loaded("ExcelMerger.PdfOpsForm", new[] { Doc }, 0, 0);
        var button = (Button)Field(f, "_btnExtract");
        button.PerformClick();
        Pump(700);
        var menu = (ContextMenuStrip)Field(f, "_extractMenu");
        var imagesItem = (ToolStripMenuItem)menu.Items[0];
        imagesItem.ShowDropDown();
        Pump(400);
        _popup = imagesItem.DropDown.Bounds;
        ScreenShot(f, "ops-dpi");
        Kill(f);
    }

    static void SplitTemplate()
    {
        Form f = Loaded("ExcelMerger.PdfSplitForm", new[] { Doc });
        var combo = (ComboBox)Field(f, "_cmbMode");
        combo.SelectedIndex = 2;
        Pump(400);
        var box = (TextBox)Field(f, "_txtNameTemplate");
        box.Text = "[BASENAME]_[FILENUMBER###]";
        ShowMenu(box.ContextMenuStrip, box, new Point(30, 20));
        ScreenShot(f, "split-template");
        Kill(f);
    }

    static void Ocr()
    {
        Form f = Loaded("ExcelMerger.OcrForm", new[] { Doc, AppB });
        Shot(f, "ocr");
        Kill(f);
    }

    /// <summary>«PDF → PowerPoint» — шестой инструмент, появился в 1.18.0.</summary>
    static void Pptx()
    {
        Form f = Loaded("ExcelMerger.PptxForm", new[] { Doc, AppB });
        Shot(f, "pptx");
        Kill(f);
    }

    static void ReviewLayout()
    {
        Form form = Tool("ExcelMerger.PdfReviewForm");
        Place(form, 1120, 760);
        Shot(form, "review-layout-unified");
        ((Button)Field(form, "_sideModeButton")).PerformClick();
        Pump(400);
        AssertReviewLegendsAndLayout(form, "default-empty");
        Shot(form, "review-layout-side");
        form.Size = form.MinimumSize;
        Pump(500);
        AssertReviewLegendsAndLayout(form, "minimum-empty");
        Shot(form, "review-layout-side-minimum");
        Kill(form);
    }

    /// <summary>
    /// Приёмка настоящего собранного PdfReviewForm на локальных 1.pdf/2.pdf. В отличие от
    /// снимков руководства, это диагностический сценарий: он ждёт каждую асинхронную стадию
    /// по состоянию, снимает все три изменённые пары и не подменяет приложение тестовым UI.
    /// </summary>
    static void Review()
    {
        string leftPath, rightPath;
        ResolveReviewInputs(out leftPath, out rightPath);
        ReviewFileStamp leftBefore = ReviewStamp(leftPath);
        ReviewFileStamp rightBefore = ReviewStamp(rightPath);
        bool reverse = IsReviewFixture(leftBefore, ReviewFixture2Sha256,
            ReviewFixture2Bytes) && IsReviewFixture(rightBefore,
            ReviewFixture1Sha256, ReviewFixture1Bytes);
        if (reverse)
        {
            EnsureReviewFixture(leftBefore, ReviewFixture2Sha256,
                ReviewFixture2Bytes, "2.pdf (left)");
            EnsureReviewFixture(rightBefore, ReviewFixture1Sha256,
                ReviewFixture1Bytes, "1.pdf (right)");
        }
        else
        {
            EnsureReviewFixture(leftBefore, ReviewFixture1Sha256,
                ReviewFixture1Bytes, "1.pdf (left)");
            EnsureReviewFixture(rightBefore, ReviewFixture2Sha256,
                ReviewFixture2Bytes, "2.pdf (right)");
        }
        Ensure(!string.Equals(leftBefore.Path, rightBefore.Path,
            StringComparison.OrdinalIgnoreCase), "Review: стороны должны быть разными файлами");

        var manifest = new StringBuilder();
        manifest.AppendLine("PDF REVIEW COMPILED-FORM ACCEPTANCE");
        manifest.AppendLine("app=" + Path.GetFullPath(_app.Location));
        manifest.AppendLine("language=" + (_en ? "en" : "ru"));
        manifest.AppendLine("os=" + Environment.OSVersion);
        manifest.AppendLine("process-bits=" + (IntPtr.Size * 8));
        manifest.AppendLine("high-contrast=" + SystemInformation.HighContrast);
        manifest.AppendLine("direction=" + (reverse ? "2.pdf→1.pdf" : "1.pdf→2.pdf"));
        manifest.AppendLine("before-left=" + leftBefore.ToLine());
        manifest.AppendLine("before-right=" + rightBefore.ToLine());

        Form f = null;
        Exception failure = null;
        try
        {
            f = Tool("ExcelMerger.PdfReviewForm");
            Place(f, 1120, 760);
            Type acceptor = _app.GetType("ExcelMerger.IFileAcceptor");
            acceptor.GetMethod("AcceptFiles").Invoke(f,
                new object[] { new[] { leftPath, rightPath } });

            WaitUntil("Review: проверка двух источников", delegate
            {
                return !ReviewBool(f, "_leftSourceChecking") &&
                    !ReviewBool(f, "_rightSourceChecking") &&
                    string.Equals((string)Field(f, "_leftFile"), leftPath,
                        StringComparison.OrdinalIgnoreCase) &&
                    string.Equals((string)Field(f, "_rightFile"), rightPath,
                        StringComparison.OrdinalIgnoreCase) &&
                    ((Button)Field(f, "_compare")).Enabled;
            }, 45000, delegate { return ReviewState(f); });

            Button compare = (Button)Field(f, "_compare");
            compare.PerformClick();
            WaitUntil("Review: document-wide сравнение", delegate
            {
                return Field(f, "_result") != null && !ReviewWorking(f);
            }, 180000, delegate { return ReviewState(f); });
            object unifiedView = Field(f, "_unifiedSource");
            WaitUntil("Review: единый документ", delegate
            {
                return Member(unifiedView, "ViewState").ToString() == "Ready" &&
                    !ReviewBool(unifiedView, "_renderWorker");
            }, 45000, delegate { return ReviewState(f); });
            ReviewShot(f, "review", false);
            ((Button)Field(f, "_sideModeButton")).PerformClick();
            Pump(300);
            WaitReviewPanes(f, "первая пара");
            object result = Field(f, "_result");
            AssertReviewFixtureResult(result, reverse);
            string semanticBefore = ReviewResultDetails(result);
            manifest.AppendLine();
            manifest.Append(semanticBefore);
            AssertReviewLegendsAndLayout(f, "normal");
            AssertReviewVisualPair(f, result, 0, manifest, true);
            AssertReviewSelectionClipboard(f, manifest);

            ListBox pairs = (ListBox)Field(f, "_pairs");
            Ensure(pairs.SelectedIndex == 0,
                "Review: после сравнения должна быть выбрана первая пара");
            ReviewShot(f, "review-page-1", true);

            Ensure(InvokeReviewKey(f, Keys.F3), "Review: F3 должен быть обработан формой");
            WaitUntil("Review: F3 → вторая изменённая пара", delegate
            {
                return pairs.SelectedIndex == 1;
            }, 5000, delegate { return ReviewState(f); });
            WaitReviewPanes(f, "вторая пара");
            AssertReviewVisualPair(f, result, 1, manifest, false);
            ReviewShot(f, "review-page-2", false);

            Ensure(InvokeReviewKey(f, Keys.F3), "Review: повторный F3 должен быть обработан");
            WaitUntil("Review: F3 → третья изменённая пара", delegate
            {
                return pairs.SelectedIndex == 2;
            }, 5000, delegate { return ReviewState(f); });
            WaitReviewPanes(f, "третья пара");
            AssertReviewVisualPair(f, result, 2, manifest, false);
            ReviewShot(f, "review-page-3", false);

            Ensure(InvokeReviewKey(f, Keys.Shift | Keys.F3),
                "Review: Shift+F3 должен быть обработан формой");
            Ensure(pairs.SelectedIndex == 1,
                "Review: Shift+F3 должен вернуть предыдущую изменённую пару");

            pairs.SelectedIndex = 0;
            WaitReviewPanes(f, "возврат к первой паре");
            f.Size = f.MinimumSize;
            Pump(800);
            AssertReviewLegendsAndLayout(f, "minimum");
            ReviewShot(f, "review-minimum", false);

            f.Size = new Size(1120, 760);
            Pump(600);
            TextBox leftPage = (TextBox)Field(f, "_leftPageInput");
            TextBox rightPage = (TextBox)Field(f, "_rightPageInput");
            leftPage.Text = "3";
            Ensure(InvokeReviewNavigate(f, true),
                "Review: левая физическая страница 3 должна открыться");
            WaitReviewPanes(f, "независимая левая страница 3");
            Ensure(ReviewInt(f, "_leftRowIndex") == 2 &&
                   ReviewInt(f, "_rightRowIndex") == 0,
                "Review: левая навигация не должна сдвигать правую сторону");

            rightPage.Text = "2";
            Ensure(InvokeReviewNavigate(f, false),
                "Review: правая физическая страница 2 должна открыться");
            WaitReviewPanes(f, "независимая правая страница 2");
            Ensure(ReviewInt(f, "_leftRowIndex") == 2 &&
                   ReviewInt(f, "_rightRowIndex") == 1,
                "Review: правая навигация не должна сдвигать левую сторону");
            Ensure(leftPage.Text == "3" && rightPage.Text == "2",
                "Review: независимые поля должны показывать 3 слева и 2 справа");
            ReviewShot(f, "review-independent-pages", false);

            pairs.SelectedIndex = 0;
            WaitReviewPanes(f, "синхронизация перед wheel");
            object leftView = Field(f, "_leftSource");
            object rightView = Field(f, "_rightSource");
            Panel leftViewport = (Panel)Member(leftView, "_viewport");
            Point leftPoint = ScreenCenter(leftViewport);
            double leftScaleBefore = ReviewDouble(leftView, "ZoomScale");
            double rightScaleBefore = ReviewDouble(rightView, "ZoomScale");
            Point rightOffsetBefore = (Point)Member(rightView, "ScrollOffset");
            object rightPageBefore = Member(rightView, "_page");

            Ensure(InvokeReviewWheel(f, leftPoint, 120, true),
                "Review: Ctrl+wheel слева должен быть маршрутизирован");
            Ensure(ReviewDouble(leftView, "ZoomScale") > leftScaleBefore,
                "Review: Ctrl+wheel должен увеличить только левую страницу");
            Ensure(Math.Abs(ReviewDouble(rightView, "ZoomScale") - rightScaleBefore) < 0.000001,
                "Review: левый Ctrl+wheel не должен менять правый масштаб");
            for (int step = 0; ReviewDouble(leftView, "ZoomScale") < 1.0 && step < 20; step++)
                Ensure(InvokeReviewWheel(f, leftPoint, 120, true),
                    "Review: левая страница должна увеличиваться до прокручиваемого масштаба");

            Point scrollBefore = (Point)Member(leftView, "ScrollOffset");
            Ensure(InvokeReviewWheel(f, leftPoint, -120, false),
                "Review: обычное колесо должно прокрутить левую страницу");
            Point scrollAfter = (Point)Member(leftView, "ScrollOffset");
            Ensure(scrollAfter != scrollBefore && ReviewInt(f, "_leftRowIndex") == 0,
                "Review: колесо сначала меняет локальный offset, а не страницу");

            for (int step = 0; ReviewInt(f, "_leftRowIndex") == 0 && step < 240; step++)
            {
                InvokeReviewWheel(f, leftPoint, -120, false);
                Application.DoEvents();
            }
            Ensure(ReviewInt(f, "_leftRowIndex") == 1,
                "Review: колесо на нижней границе должно открыть следующую левую страницу");
            Ensure(ReviewInt(f, "_rightRowIndex") == 0,
                "Review: boundary-wheel слева не должен листать правую страницу");
            WaitReviewPanes(f, "wheel → следующая левая страница");
            Ensure(Math.Abs(ReviewDouble(rightView, "ZoomScale") - rightScaleBefore) < 0.000001 &&
                   (Point)Member(rightView, "ScrollOffset") == rightOffsetBefore &&
                   ReferenceEquals(Member(rightView, "_page"), rightPageBefore),
                "Review: wheel слева должен сохранить правую страницу, zoom и offset");
            ReviewShot(f, "review-wheel-next-page", false);

            pairs.SelectedIndex = 0;
            WaitReviewPanes(f, "синхронизация перед lifecycle-приёмкой");
            AssertReviewDoubleClickPreview(f, result, semanticBefore, manifest);
            AssertReviewCompiledViewStates(f, result, manifest);
            AssertReviewCaptureNavigationLifecycle(f, pairs, manifest);
            AssertReviewRapidLifecycle(f, result, pairs, semanticBefore, manifest);

            Ensure(ReferenceEquals(result, Field(f, "_result")),
                "Review: навигация не должна заменять semantic result");
            Ensure(semanticBefore == ReviewResultDetails(result),
                "Review: F3, поля страниц и wheel не должны менять операции и статистику");
            manifest.AppendLine("ui=normal+minimum; pages=1,2,3; F3=ok; Shift+F3=ok");
            manifest.AppendLine("independent-navigation=left:3/right:2; wheel=local+boundary; ctrl-wheel=local");
            manifest.AppendLine("manual-pair=not-driven (two sequential modal dialogs; semantic invariance is covered by the automated suite)");
            manifest.AppendLine("grayscale=review-page-1-grayscale.png");
            manifest.AppendLine("high-contrast-renderer=review-page-1-high-contrast-left.png+review-page-1-high-contrast-right.png; os-level=" +
                (SystemInformation.HighContrast ? "performed" : "not-performed (system setting left unchanged)"));
        }
        catch (Exception ex)
        {
            failure = RootCause(ex);
            manifest.AppendLine("failure=" + failure.GetType().Name + ": " + failure.Message);
            if (f != null && !f.IsDisposed)
            {
                try { Shot(f, "review-failure"); }
                catch (Exception shotError)
                {
                    manifest.AppendLine("failure-shot=" + RootCause(shotError).Message);
                }
            }
        }
        finally
        {
            if (f != null)
                Kill(f);
            try
            {
                ReviewFileStamp leftAfter = ReviewStamp(leftPath);
                ReviewFileStamp rightAfter = ReviewStamp(rightPath);
                manifest.AppendLine("after-left=" + leftAfter.ToLine());
                manifest.AppendLine("after-right=" + rightAfter.ToLine());
                if (!leftBefore.SameAs(leftAfter) || !rightBefore.SameAs(rightAfter))
                {
                    var changed = new Exception(
                        "Review: локальные 1.pdf/2.pdf изменились во время read-only приёмки");
                    if (failure == null) failure = changed;
                    manifest.AppendLine("integrity=FAILED");
                }
                else
                {
                    manifest.AppendLine("integrity=unchanged");
                }
            }
            catch (Exception stampError)
            {
                if (failure == null) failure = RootCause(stampError);
                manifest.AppendLine("integrity=ERROR " + RootCause(stampError).Message);
            }
            File.WriteAllText(Path.Combine(_out, "review-manifest.txt"),
                manifest.ToString(), new UTF8Encoding(false));
        }
        if (failure != null)
            throw failure;
        Say("review -> ACCEPTED; manifest=review-manifest.txt");
    }

    sealed class ReviewFileStamp
    {
        public string Path;
        public long Length;
        public DateTime LastWriteUtc;
        public string Sha256;

        public string ToLine()
        {
            return "path=" + Path + "; bytes=" + Length + "; sha256=" + Sha256 +
                "; mtimeUtc=" + LastWriteUtc.ToString("o",
                    System.Globalization.CultureInfo.InvariantCulture);
        }

        public bool SameAs(ReviewFileStamp other)
        {
            return other != null && Length == other.Length &&
                LastWriteUtc == other.LastWriteUtc &&
                string.Equals(Sha256, other.Sha256, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(Path, other.Path, StringComparison.OrdinalIgnoreCase);
        }
    }

    static ReviewFileStamp ReviewStamp(string path)
    {
        string full = Path.GetFullPath(path);
        var info = new FileInfo(full);
        using (var sha = SHA256.Create())
        using (var stream = new FileStream(full, FileMode.Open, FileAccess.Read,
            FileShare.Read))
        {
            byte[] hash = sha.ComputeHash(stream);
            var hex = new StringBuilder(hash.Length * 2);
            foreach (byte value in hash)
                hex.Append(value.ToString("x2",
                    System.Globalization.CultureInfo.InvariantCulture));
            return new ReviewFileStamp
            {
                Path = full,
                Length = info.Length,
                LastWriteUtc = info.LastWriteTimeUtc,
                Sha256 = hex.ToString()
            };
        }
    }

    static bool IsReviewFixture(ReviewFileStamp stamp, string sha256, long bytes)
    {
        return stamp != null && stamp.Length == bytes &&
            string.Equals(stamp.Sha256, sha256, StringComparison.OrdinalIgnoreCase);
    }

    static void EnsureReviewFixture(ReviewFileStamp stamp, string sha256, long bytes,
        string name)
    {
        Ensure(IsReviewFixture(stamp, sha256, bytes),
            "Review: " + name + " не совпадает с принятой локальной фикстурой; " +
            (stamp == null ? "stamp=null" : stamp.ToLine()));
    }

    static bool ReviewInputsAvailable()
    {
        string left, right;
        return TryResolveReviewInputs(out left, out right);
    }

    static void ResolveReviewInputs(out string left, out string right)
    {
        if (TryResolveReviewInputs(out left, out right))
            return;
        throw new FileNotFoundException(
            "для сценария review нужны неизменённые 1.pdf и 2.pdf в корне репозитория; " +
            "либо задайте IWO_REVIEW_LEFT и IWO_REVIEW_RIGHT (оба пути), либо IWO_REVIEW_FIXTURE_DIR");
    }

    static bool TryResolveReviewInputs(out string left, out string right)
    {
        left = right = null;
        bool explicitLeft = !string.IsNullOrWhiteSpace(_reviewLeft);
        bool explicitRight = !string.IsNullOrWhiteSpace(_reviewRight);
        if (explicitLeft || explicitRight)
        {
            if (!explicitLeft || !explicitRight)
                return false;
            return TryReviewPair(Path.GetFullPath(_reviewLeft),
                Path.GetFullPath(_reviewRight), out left, out right);
        }

        var directories = new List<string>();
        string fixtureDir = Environment.GetEnvironmentVariable("IWO_REVIEW_FIXTURE_DIR");
        AddReviewDirectory(directories, fixtureDir);
        AddReviewDirectory(directories, Environment.CurrentDirectory);
        string appDir = Path.GetDirectoryName(_app.Location);
        AddReviewDirectory(directories, appDir);
        AddReviewDirectory(directories, ParentDirectory(appDir, 1));
        AddReviewDirectory(directories, ParentDirectory(_samples, 3));
        foreach (string directory in directories)
        {
            if (TryReviewPair(Path.Combine(directory, "1.pdf"),
                Path.Combine(directory, "2.pdf"), out left, out right))
                return true;
        }
        return false;
    }

    static bool TryReviewPair(string candidateLeft, string candidateRight,
        out string left, out string right)
    {
        left = right = null;
        if (string.IsNullOrWhiteSpace(candidateLeft) ||
            string.IsNullOrWhiteSpace(candidateRight) ||
            !File.Exists(candidateLeft) || !File.Exists(candidateRight))
            return false;
        left = Path.GetFullPath(candidateLeft);
        right = Path.GetFullPath(candidateRight);
        return true;
    }

    static void AddReviewDirectory(List<string> directories, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;
        string full;
        try { full = Path.GetFullPath(path); }
        catch { return; }
        foreach (string existing in directories)
            if (string.Equals(existing, full, StringComparison.OrdinalIgnoreCase))
                return;
        if (Directory.Exists(full))
            directories.Add(full);
    }

    static string ParentDirectory(string path, int levels)
    {
        if (string.IsNullOrWhiteSpace(path))
            return null;
        DirectoryInfo current;
        try { current = new DirectoryInfo(Path.GetFullPath(path)); }
        catch { return null; }
        for (int i = 0; i < levels && current != null; i++)
            current = current.Parent;
        return current == null ? null : current.FullName;
    }

    static void WaitReviewPanes(Form f, string context)
    {
        object left = Field(f, "_leftSource");
        object right = Field(f, "_rightSource");
        WaitUntil("Review: рендер — " + context, delegate
        {
            return ReviewPaneTerminal(left) && ReviewPaneTerminal(right);
        }, 45000, delegate { return ReviewState(f); });
        Ensure(ReviewPaneReady(left) && ReviewPaneReady(right),
            "Review: обе стороны обязаны иметь видимый растр (" + context + "); " +
            ReviewState(f));
    }

    static bool ReviewPaneTerminal(object pane)
    {
        string state = Convert.ToString(Member(pane, "ViewState"),
            System.Globalization.CultureInfo.InvariantCulture);
        return !ReviewBool(pane, "_renderWorker") &&
            (state == "Ready" || state == "Unavailable" ||
             state == "MissingCounterpart" || state == "Empty");
    }

    static bool ReviewPaneReady(object pane)
    {
        return Convert.ToString(Member(pane, "ViewState"),
                   System.Globalization.CultureInfo.InvariantCulture) == "Ready" &&
               (bool)Member(pane, "HasVisiblePage") &&
               !ReviewBool(pane, "_renderWorker");
    }

    static void AssertReviewDoubleClickPreview(Form f, object result,
        string semanticBefore, StringBuilder manifest)
    {
        object view = Field(f, "_leftSource");
        Control surface = ReviewSurface(f, true);
        object page = Member(view, "_page");
        Ensure(page != null && ReviewPaneReady(view),
            "Review preview: левая ready page отсутствует");
        int pageIndex = ReviewInt(page, "PageIndex");
        string sourcePath = Convert.ToString(Member(page, "SourcePath"),
            System.Globalization.CultureInfo.InvariantCulture);
        string expectedCaption = string.Format(Text("preview.title"), pageIndex + 1);
        double zoomBefore = ReviewDouble(view, "ZoomScale");
        Point offsetBefore = (Point)Member(view, "ScrollOffset");
        object modelBefore = Member(surface, "SelectionModel");
        Type previewType = _app.GetType("ExcelMerger.PagePreviewForm");
        Ensure(previewType != null, "Review preview: PagePreviewForm не найден");

        bool opened = false;
        bool rendered = false;
        Exception callbackFailure = null;
        var elapsed = Stopwatch.StartNew();
        var closer = new System.Windows.Forms.Timer();
        closer.Interval = 50;
        closer.Tick += delegate
        {
            Form preview = null;
            foreach (Form open in Application.OpenForms)
                if (previewType.IsInstanceOfType(open))
                {
                    preview = open;
                    break;
                }
            try
            {
                if (preview == null)
                {
                    if (elapsed.ElapsedMilliseconds > 45000)
                        callbackFailure = new TimeoutException(
                            "Review preview: modal PagePreviewForm не открылся за 45000 мс");
                    return;
                }
                opened = true;
                Ensure(preview.Owner == f,
                    "Review preview: modal owner должен быть PdfReviewForm");
                Ensure(ReferenceEquals(Member(preview, "_page"), page),
                    "Review preview: double-click передал не текущую PdfPageRef");
                Ensure(preview.Text == expectedCaption,
                    "Review preview: caption не соответствует физической странице");
                Bitmap image = Member(preview, "_image") as Bitmap;
                if (image == null)
                {
                    if (elapsed.ElapsedMilliseconds <= 45000)
                        return;
                    callbackFailure = new TimeoutException(
                        "Review preview: растр fixture не появился за 45000 мс");
                }
                else
                {
                    Ensure(image.Width > 0 && image.Height > 0,
                        "Review preview: показан пустой растр");
                    rendered = true;
                }
            }
            catch (Exception ex)
            {
                callbackFailure = RootCause(ex);
            }
            finally
            {
                if (preview != null && (rendered || callbackFailure != null))
                {
                    closer.Stop();
                    preview.Close();
                }
            }
        };

        surface.Focus();
        Ensure(surface.Focused, "Review preview: page surface не получил focus");
        closer.Start();
        try
        {
            MethodInfo onDoubleClick = typeof(Control).GetMethod("OnDoubleClick",
                BindingFlags.Instance | BindingFlags.NonPublic);
            Ensure(onDoubleClick != null,
                "Review preview: Control.OnDoubleClick не найден");
            onDoubleClick.Invoke(surface, new object[] { EventArgs.Empty });
        }
        finally
        {
            closer.Stop();
            closer.Dispose();
        }
        if (callbackFailure != null)
            throw callbackFailure;
        Ensure(opened && rendered,
            "Review preview: compiled surface double-click не открыл готовый preview");
        Pump(400);
        foreach (Form open in Application.OpenForms)
            Ensure(!previewType.IsInstanceOfType(open),
                "Review preview: modal окно осталось открытым после закрытия");
        Ensure(f.Visible && f.Enabled && !f.IsDisposed,
            "Review preview: владелец не восстановился после modal окна");
        Ensure(ReferenceEquals(page, Member(view, "_page")) &&
               ReferenceEquals(modelBefore, Member(surface, "SelectionModel")) &&
               ReviewPaneReady(view),
            "Review preview: double-click заменил ready page/trusted layer");
        Ensure(string.Equals(sourcePath,
                   Convert.ToString(Member(Member(view, "_page"), "SourcePath"),
                       System.Globalization.CultureInfo.InvariantCulture),
                   StringComparison.OrdinalIgnoreCase) &&
               Math.Abs(ReviewDouble(view, "ZoomScale") - zoomBefore) < 0.000001 &&
               (Point)Member(view, "ScrollOffset") == offsetBefore,
            "Review preview: modal просмотр изменил source/zoom/offset pane");
        Ensure(ReferenceEquals(result, Field(f, "_result")) &&
               semanticBefore == ReviewResultDetails(result),
            "Review preview: modal просмотр изменил semantic result");
        manifest.AppendLine("double-click-preview=compiled surface event→modal; " +
            "owner/page/caption/raster=ok; review-state=preserved");
    }

    static void AssertReviewCompiledViewStates(Form f, object result,
        StringBuilder manifest)
    {
        object source = Field(f, "_leftSource");
        object page = Member(source, "_page");
        object reviewPage = Member(source, "_reviewPage");
        Ensure(page != null && reviewPage != null,
            "Review states: current page/trusted layer отсутствует");
        IList pairs = Member(result, "Pairs") as IList;
        Ensure(pairs != null && pairs.Count > 0,
            "Review states: fixture pair отсутствует");
        object highlight = BuildReviewHighlight(result, pairs[0], true);

        Type viewType = _app.GetType("ExcelMerger.PdfReviewPageView");
        Type pageType = _app.GetType("ExcelMerger.PdfPageRef");
        Type reviewPageType = _app.GetType("ExcelMerger.PdfReviewPage");
        Type highlightType = _app.GetType("ExcelMerger.PdfReviewHighlight");
        Type positionType = _app.GetType("ExcelMerger.PdfReviewPagePosition");
        Ensure(viewType != null && pageType != null && reviewPageType != null &&
               highlightType != null && positionType != null,
            "Review states: compiled types не найдены");
        MethodInfo showPage = viewType.GetMethod("ShowPage",
            BindingFlags.Instance | BindingFlags.Public, null,
            new[] { pageType, reviewPageType, typeof(long), typeof(string),
                highlightType, positionType }, null);
        MethodInfo showEmpty = viewType.GetMethod("ShowEmpty",
            BindingFlags.Instance | BindingFlags.Public);
        MethodInfo setDropTarget = viewType.GetMethod("SetDropTarget",
            BindingFlags.Instance | BindingFlags.NonPublic);
        Ensure(showPage != null && showEmpty != null && setDropTarget != null,
            "Review states: state API не найден");
        object defaultPosition = Enum.Parse(positionType, "Default");
        long revision = Convert.ToInt64(Member(source, "_targetContentRevision"),
            System.Globalization.CultureInfo.InvariantCulture);

        using (var host = new Form())
        {
            host.ShowInTaskbar = false;
            host.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            host.StartPosition = FormStartPosition.Manual;
            host.Location = new Point(-30000, -30000);
            host.ClientSize = new Size(620, 720);
            var pane = (Control)Activator.CreateInstance(viewType, true);
            pane.Dock = DockStyle.Fill;
            host.Controls.Add(pane);
            host.Show();
            Pump(100);

            showPage.Invoke(pane, new[] { page, reviewPage, (object)revision,
                "Review state acceptance", highlight, defaultPosition });
            Ensure(ReviewViewState(pane) == "Loading",
                "Review states: valid ShowPage не вошёл в Loading");
            AssertReviewNoCanvas(pane, "Loading");
            WaitUntil("Review states: Loading→Ready", delegate
            {
                return ReviewPaneTerminal(pane);
            }, 45000, delegate { return "state=" + ReviewViewState(pane); });
            Ensure(ReviewPaneReady(pane) &&
                   ReferenceEquals(Member(pane, "_reviewPage"), reviewPage),
                "Review states: valid compiled render не опубликовал Ready + trusted layer");
            Control readySurface = Member(pane, "_picture") as Control;
            Ensure(readySurface != null && ReviewBool(readySurface, "HasSelectableText"),
                "Review states: Ready не прикрепил trusted selectable layer");

            object unavailablePage = Activator.CreateInstance(pageType);
            Set(pageType, unavailablePage, "SourcePath", Path.Combine(_out,
                "__review_missing_" + Guid.NewGuid().ToString("N") + ".pdf"));
            pageType.GetField("PageIndex").SetValue(unavailablePage, 0);
            showPage.Invoke(pane, new[] { unavailablePage, null, (object)(revision + 1),
                "Review unavailable acceptance", null, defaultPosition });
            Ensure(ReviewViewState(pane) == "Loading",
                "Review states: недоступная page не вошла в Loading");
            AssertReviewNoCanvas(pane, "Unavailable/Loading");
            WaitUntil("Review states: Loading→Unavailable", delegate
            {
                return ReviewPaneTerminal(pane);
            }, 45000, delegate { return "state=" + ReviewViewState(pane); });
            Ensure(ReviewViewState(pane) == "Unavailable",
                "Review states: missing source не дал Unavailable");
            AssertReviewNoCanvas(pane, "Unavailable");

            showPage.Invoke(pane, new[] { null, null, (object)(revision + 2),
                "Review missing counterpart acceptance", null, defaultPosition });
            Ensure(ReviewViewState(pane) == "MissingCounterpart",
                "Review states: null page не дала MissingCounterpart");
            AssertReviewNoCanvas(pane, "MissingCounterpart");

            showEmpty.Invoke(pane, new object[] { "Review empty acceptance" });
            Ensure(ReviewViewState(pane) == "Empty",
                "Review states: ShowEmpty не дал Empty");
            AssertReviewNoCanvas(pane, "Empty");
            setDropTarget.Invoke(pane, new object[] { true });
            Ensure(ReviewViewState(pane) == "DropTarget",
                "Review states: drop overlay не дал DropTarget");
            AssertReviewNoCanvas(pane, "DropTarget");
            setDropTarget.Invoke(pane, new object[] { false });
            Ensure(ReviewViewState(pane) == "Empty" &&
                   !ReviewBool(pane, "_renderWorker"),
                "Review states: снятие DropTarget не восстановило Empty");
            host.Close();
        }
        Pump(300);
        manifest.AppendLine("compiled-view-states=Loading→Ready, Loading→Unavailable, " +
            "MissingCounterpart, Empty, DropTarget; white-canvas/trusted-layer=absent");
    }

    static string ReviewViewState(object pane)
    {
        return Convert.ToString(Member(pane, "ViewState"),
            System.Globalization.CultureInfo.InvariantCulture);
    }

    static void AssertReviewNoCanvas(object pane, string context)
    {
        var picture = Member(pane, "_picture") as PictureBox;
        Ensure(picture != null && !ReviewBool(pane, "HasVisiblePage") &&
               Member(pane, "_bitmap") == null && Member(pane, "_reviewPage") == null &&
               !picture.Visible && picture.Image == null &&
               !ReviewBool(picture, "HasSelectableText") &&
               Member(picture, "SelectionModel") == null,
            "Review states " + context +
            ": неготовое состояние оставило белый canvas/trusted layer");
    }

    static void AssertReviewCaptureNavigationLifecycle(Form f, ListBox pairs,
        StringBuilder manifest)
    {
        Ensure(pairs.Items.Count >= 2,
            "Review capture: нужны как минимум две viewer rows");
        if (pairs.SelectedIndex != 0)
            pairs.SelectedIndex = 0;
        WaitReviewPanes(f, "capture lifecycle → первая пара");
        Control surface = ReviewSurface(f, true);
        object model = Member(surface, "SelectionModel");
        int start, end;
        FindReviewDraggableRange(surface, model, out start, out end);
        RectangleF word = ReviewWordRectangle(model, start, surface.ClientSize);
        Point point = ReviewRectCenter(word, surface.ClientRectangle);
        Form owner = surface.FindForm();
        Point originalCursor = Cursor.Position;
        const uint MouseEventLeftDown = 0x0002;
        const uint MouseEventLeftUp = 0x0004;
        bool buttonDown = false;
        try
        {
            owner.Activate();
            SetForegroundWindow(owner.Handle);
            surface.Focus();
            Cursor.Position = surface.PointToScreen(point);
            Pump(40);
            mouse_event(MouseEventLeftDown, 0, 0, 0, UIntPtr.Zero);
            buttonDown = true;
            Pump(40);
            Ensure(surface.Capture && ReviewAutoScrollEnabled(surface),
                "Review capture: native mouse-down не запустил capture/timer");
            Ensure(ReviewBool(surface, "HasSelection"),
                "Review capture: native mouse-down не создал selection");

            pairs.SelectedIndex = 1;
            Application.DoEvents();
            Ensure(!surface.Capture && !ReviewAutoScrollEnabled(surface) &&
                   !ReviewBool(surface, "HasSelection") &&
                   Member(surface, "SelectionModel") == null,
                "Review capture: page replacement не остановил capture/timer и stale layer");
        }
        finally
        {
            if (buttonDown)
                mouse_event(MouseEventLeftUp, 0, 0, 0, UIntPtr.Zero);
            Cursor.Position = originalCursor;
            Pump(40);
        }
        WaitReviewPanes(f, "capture lifecycle → вторая пара");
        Ensure(!ReferenceEquals(model, Member(surface, "SelectionModel")) &&
               !surface.Capture && !ReviewAutoScrollEnabled(surface) &&
               !ReviewBool(surface, "HasSelection"),
            "Review capture: новая page layer сохранила старый drag lifecycle");
        pairs.SelectedIndex = 0;
        WaitReviewPanes(f, "capture lifecycle → возврат первой пары");
        manifest.AppendLine("capture-lifecycle=native mouse-down; page replacement=" +
            "capture/timer/selection detached; next layer=clean");
    }

    static bool ReviewAutoScrollEnabled(Control surface)
    {
        var timer = Member(surface, "_autoScrollTimer") as
            System.Windows.Forms.Timer;
        Ensure(timer != null, "Review capture: autoscroll timer не найден");
        return timer.Enabled;
    }

    sealed class ReviewResourceSnapshot
    {
        public uint GdiObjects;
        public uint UserObjects;
        public long PrivateBytes;

        public string ToLine()
        {
            return "gdi=" + GdiObjects + "; user=" + UserObjects +
                "; privateMiB=" + (PrivateBytes / 1048576.0).ToString("F1",
                    System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    static ReviewResourceSnapshot ReviewResources()
    {
        using (Process process = Process.GetCurrentProcess())
        {
            process.Refresh();
            var snapshot = new ReviewResourceSnapshot
            {
                GdiObjects = GetGuiResources(process.Handle, 0),
                UserObjects = GetGuiResources(process.Handle, 1),
                PrivateBytes = process.PrivateMemorySize64
            };
            Ensure(snapshot.GdiObjects > 0 && snapshot.UserObjects > 0 &&
                   snapshot.PrivateBytes > 0,
                "Review lifecycle: Win32 resource snapshot недоступен");
            return snapshot;
        }
    }

    static void CompactReviewResources()
    {
        Pump(250);
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        Pump(250);
    }

    static void AssertReviewRapidLifecycle(Form f, object result, ListBox pairs,
        string semanticBefore, StringBuilder manifest)
    {
        Ensure(pairs.Items.Count == 3,
            "Review lifecycle: root fixture должна иметь три viewer rows");
        for (int warmup = 0; warmup < 3; warmup++)
            RunReviewRapidCycle(f, result, pairs, warmup % 3,
                "warmup " + (warmup + 1));
        CompactReviewResources();
        ReviewResourceSnapshot baseline = ReviewResources();
        ReviewResourceSnapshot peak = new ReviewResourceSnapshot
        {
            GdiObjects = baseline.GdiObjects,
            UserObjects = baseline.UserObjects,
            PrivateBytes = baseline.PrivateBytes
        };

        const int cycles = 12;
        for (int cycle = 0; cycle < cycles; cycle++)
        {
            RunReviewRapidCycle(f, result, pairs, cycle % 3,
                "measured " + (cycle + 1));
            ReviewResourceSnapshot current = ReviewResources();
            peak.GdiObjects = Math.Max(peak.GdiObjects, current.GdiObjects);
            peak.UserObjects = Math.Max(peak.UserObjects, current.UserObjects);
            peak.PrivateBytes = Math.Max(peak.PrivateBytes, current.PrivateBytes);
        }
        RunReviewRapidCycle(f, result, pairs, 0, "final first pair");
        CompactReviewResources();
        ReviewResourceSnapshot after = ReviewResources();

        const long MiB = 1024L * 1024L;
        Ensure(peak.GdiObjects <= baseline.GdiObjects + 24 &&
               after.GdiObjects <= baseline.GdiObjects + 12,
            "Review lifecycle: GDI objects растут без границы; baseline=" +
            baseline.ToLine() + "; peak=" + peak.ToLine() + "; after=" +
            after.ToLine());
        Ensure(peak.UserObjects <= baseline.UserObjects + 12 &&
               after.UserObjects <= baseline.UserObjects + 6,
            "Review lifecycle: USER objects растут без границы; baseline=" +
            baseline.ToLine() + "; peak=" + peak.ToLine() + "; after=" +
            after.ToLine());
        Ensure(peak.PrivateBytes <= baseline.PrivateBytes + 160L * MiB &&
               after.PrivateBytes <= baseline.PrivateBytes + 96L * MiB,
            "Review lifecycle: private memory растёт сверх bounded render budget; baseline=" +
            baseline.ToLine() + "; peak=" + peak.ToLine() + "; after=" +
            after.ToLine());
        Ensure(ReferenceEquals(result, Field(f, "_result")) &&
               semanticBefore == ReviewResultDetails(result),
            "Review lifecycle: rapid switching изменил semantic result");
        Ensure(pairs.SelectedIndex == 0 && ReviewInt(f, "_leftRowIndex") == 0 &&
               ReviewInt(f, "_rightRowIndex") == 0,
            "Review lifecycle: final pair не синхронизирована");
        manifest.AppendLine("rapid-render=3 warmup+" + cycles +
            " measured cycles; each queues 3 pages, publishes final raster+trusted layer");
        manifest.AppendLine("resources[baseline]: " + baseline.ToLine());
        manifest.AppendLine("resources[peak]: " + peak.ToLine());
        manifest.AppendLine("resources[after-gc]: " + after.ToLine() +
            "; bounds=gdi +24/+12, user +12/+6, private +160/+96 MiB");
    }

    static void RunReviewRapidCycle(Form f, object result, ListBox pairs,
        int finalIndex, string context)
    {
        int first = (finalIndex + 1) % 3;
        int second = (finalIndex + 2) % 3;
        pairs.SelectedIndex = first;
        pairs.SelectedIndex = second;
        pairs.SelectedIndex = finalIndex;

        object leftView = Field(f, "_leftSource");
        object rightView = Field(f, "_rightSource");
        Ensure(ReviewViewState(leftView) == "Loading" &&
               ReviewViewState(rightView) == "Loading" &&
               ReviewInt(f, "_leftRowIndex") == finalIndex &&
               ReviewInt(f, "_rightRowIndex") == finalIndex,
            "Review lifecycle " + context +
            ": final rapid request не опубликован как Loading");
        AssertReviewNoCanvas(leftView, "rapid Loading left");
        AssertReviewNoCanvas(rightView, "rapid Loading right");
        WaitReviewPanes(f, "rapid lifecycle → " + context);

        IList resultPairs = Member(result, "Pairs") as IList;
        object expected = resultPairs[finalIndex];
        object leftPage = Member(leftView, "_page");
        object rightPage = Member(rightView, "_page");
        Ensure(leftPage != null && rightPage != null &&
               ReviewInt(leftPage, "PageIndex") == ReviewInt(expected, "LeftPageIndex") &&
               ReviewInt(rightPage, "PageIndex") == ReviewInt(expected, "RightPageIndex"),
            "Review lifecycle " + context +
            ": stale raster/page выиграл гонку у final request");

        Control left = ReviewSurface(f, true);
        Control right = ReviewSurface(f, false);
        Ensure(ReviewBool(left, "HasSelectableText") &&
               ReviewBool(right, "HasSelectableText"),
            "Review lifecycle " + context +
            ": final trusted layers не прикреплены");
        left.Focus();
        Ensure(InvokeReviewSurfaceKey(left, Keys.Control | Keys.A) &&
               ReviewInt(left, "SelectedWordCount") > 0,
            "Review lifecycle " + context + ": Ctrl+A слева не работает");
        right.Focus();
        Ensure(InvokeReviewSurfaceKey(right, Keys.Control | Keys.A) &&
               ReviewInt(right, "SelectedWordCount") > 0,
            "Review lifecycle " + context + ": Ctrl+A справа не работает");
        ReviewClearSelection(left);
        ReviewClearSelection(right);
        Ensure(!left.Capture && !right.Capture &&
               !ReviewAutoScrollEnabled(left) && !ReviewAutoScrollEnabled(right) &&
               !ReviewBool(left, "HasSelection") && !ReviewBool(right, "HasSelection"),
            "Review lifecycle " + context +
            ": selection/capture/timer не вернулись в idle");
    }

    sealed class ReviewPixelBuffer
    {
        public int Width;
        public int Height;
        public byte[] Bytes;
    }

    sealed class ReviewVisualStats
    {
        public int CoveredPixels;
        public int ChangedPixels;
        public int ChangedOutside;
        public int AlphaChanged;
        public int PaperCandidates;
        public int PaperFilled;
        public int DarkCandidates;
        public int DarkPreserved;
        public int EdgeCandidates;
        public int EdgeComposed;
        public int ChromaticCandidates;
        public int ChromaticPreserved;
        public int IntroducedFill;
        public int IntroducedOpposite;
        public int IntroducedFillOutside;

        public string ToLine()
        {
            return "covered=" + CoveredPixels + "; changed=" + ChangedPixels +
                "; changed-outside=" + ChangedOutside + "; alpha-changed=" +
                AlphaChanged + "; paper=" + PaperFilled + "/" + PaperCandidates +
                "; dark-ink=" + DarkPreserved + "/" + DarkCandidates +
                "; antialias=" + EdgeComposed + "/" + EdgeCandidates +
                "; chromatic=" + ChromaticPreserved + "/" + ChromaticCandidates +
                "; introduced-fill=" + IntroducedFill +
                "; introduced-opposite=" + IntroducedOpposite +
                "; introduced-fill-outside=" + IntroducedFillOutside;
        }
    }

    sealed class ReviewHighContrastStats
    {
        public int ChangedPixels;
        public int CoveredPaper;
        public int CoveredPaperUnchanged;
        public int IntroducedDeleteFill;
        public int IntroducedInsertFill;

        public string ToLine()
        {
            return "changed=" + ChangedPixels + "; covered-paper-unchanged=" +
                CoveredPaperUnchanged + "/" + CoveredPaper +
                "; introduced-custom-red=" + IntroducedDeleteFill +
                "; introduced-custom-green=" + IntroducedInsertFill;
        }
    }

    static void AssertReviewVisualPair(Form f, object result, int pairIndex,
        StringBuilder manifest, bool saveHighContrast)
    {
        IList pairs = Member(result, "Pairs") as IList;
        Ensure(pairs != null && pairIndex >= 0 && pairIndex < pairs.Count,
            "Review: нет semantic pair для visual-проверки " + pairIndex);
        Ensure(ReviewInt(f, "_leftRowIndex") == pairIndex &&
               ReviewInt(f, "_rightRowIndex") == pairIndex,
            "Review: visual-проверка пары " + pairIndex +
            " требует синхронных физических страниц");

        object pair = pairs[pairIndex];
        AssertReviewVisualSide(f, result, pair, pairIndex, true, manifest,
            saveHighContrast);
        AssertReviewVisualSide(f, result, pair, pairIndex, false, manifest,
            saveHighContrast);
    }

    static void AssertReviewVisualSide(Form f, object result, object pair, int pairIndex,
        bool leftSide, StringBuilder manifest, bool saveHighContrast)
    {
        string side = leftSide ? "L" : "R";
        object view = Field(f, leftSide ? "_leftSource" : "_rightSource");
        object page = Member(view, "_page");
        Bitmap marked = Member(view, "_bitmap") as Bitmap;
        Ensure(page != null && marked != null && marked.Width > 0 && marked.Height > 0,
            "Review visual[" + pairIndex + "," + side + "]: ready bitmap отсутствует");

        object highlight = BuildReviewHighlight(result, pair, leftSide);
        Color expectedFill = leftSide ? Color.FromArgb(236, 8, 8) :
            Color.FromArgb(27, 233, 26);
        Color oppositeFill = leftSide ? Color.FromArgb(27, 233, 26) :
            Color.FromArgb(236, 8, 8);
        string expectedStyle = leftSide ? "Removed" : "Added";
        string expectedRail = leftSide ? "Left" : "Right";
        Ensure(((Color)Member(highlight, "Color")).ToArgb() == expectedFill.ToArgb() &&
               Convert.ToString(Member(highlight, "Style"),
                   System.Globalization.CultureInfo.InvariantCulture) == expectedStyle &&
               Convert.ToString(Member(highlight, "ChangeBarSide"),
                   System.Globalization.CultureInfo.InvariantCulture) == expectedRail,
            "Review visual[" + pairIndex + "," + side +
            "]: semantic ownership/palette projection нарушена");
        int boxCount = CountObjects(Member(highlight, "Boxes"));
        int markerCount = CountObjects(Member(highlight, "WhitespaceMarkers"));
        if (boxCount == 0 && markerCount == 0)
        {
            manifest.AppendLine("visual[" + pairIndex + "," + side +
                "] unchanged-side=ready; no semantic marks");
            return;
        }
        Ensure(markerCount == 0,
            "Review visual[" + pairIndex + "," + side +
            "]: fixture не должна иметь whitespace marker");

        using (Bitmap raw = RenderReviewRawPage(page))
        {
            Bitmap regenerated = InvokeReviewHighlight(ReviewBitmapCopy(raw),
                highlight, SystemInformation.HighContrast);
            try
            {
                AssertReviewVisualSideRaster(raw, regenerated, highlight, pairIndex, side,
                    expectedFill, oppositeFill, saveHighContrast, manifest);
            }
            finally { regenerated.Dispose(); }
        }
    }

    static void AssertReviewVisualSideRaster(Bitmap raw, Bitmap marked, object highlight,
        int pairIndex, string side, Color expectedFill, Color oppositeFill,
        bool saveHighContrast, StringBuilder manifest)
    {
        IList<RectangleF> rectangles = ReviewHighlightRectangles(highlight,
            raw.Width, raw.Height);
        Ensure(rectangles.Count > 0,
            "Review visual[" + pairIndex + "," + side +
            "]: word-box не отобразился в raster space");
        bool[] coverage = ReviewCoverage(rectangles, raw.Width, raw.Height);

        Bitmap generatedNormal = null;
        Bitmap normal = marked;
        try
        {
            if (SystemInformation.HighContrast)
            {
                generatedNormal = InvokeReviewHighlight(ReviewBitmapCopy(raw),
                    highlight, false);
                normal = generatedNormal;
            }
            ReviewVisualStats stats = MeasureReviewNormal(raw, normal, coverage,
                expectedFill, oppositeFill);
            Ensure(stats.ChangedPixels > 0 && stats.IntroducedFill > 0,
                "Review visual[" + pairIndex + "," + side +
                "]: Word-like fill не появился; " + stats.ToLine());
            Ensure(stats.ChangedOutside <= Math.Max(1024, stats.ChangedPixels / 10) &&
                   stats.IntroducedFillOutside <= 256,
                "Review visual[" + pairIndex + "," + side +
                "]: normal renderer слишком далеко вышел за semantic region; " + stats.ToLine());
            Ensure(stats.IntroducedOpposite == 0,
                "Review visual[" + pairIndex + "," + side +
                "]: появилась заливка противоположной стороны; " + stats.ToLine());
            Ensure(stats.AlphaChanged == 0,
                "Review visual[" + pairIndex + "," + side +
                "]: renderer изменил source alpha; " + stats.ToLine());
            Ensure(stats.PaperCandidates > 0 &&
                   stats.PaperFilled == stats.PaperCandidates,
                "Review visual[" + pairIndex + "," + side +
                "]: нейтральная бумага word-box не получила точный fill; " +
                stats.ToLine());
            Ensure(stats.DarkCandidates > 0 &&
                   stats.DarkPreserved == stats.DarkCandidates,
                "Review visual[" + pairIndex + "," + side +
                "]: тёмный PDF glyph не сохранён; " + stats.ToLine());
            Ensure(stats.EdgeComposed == stats.EdgeCandidates &&
                   stats.ChromaticPreserved == stats.ChromaticCandidates,
                "Review visual[" + pairIndex + "," + side +
                "]: glyph edge/chromatic ink обработан неверно; " + stats.ToLine());
            manifest.AppendLine("visual[" + pairIndex + "," + side +
                "] source=" + (SystemInformation.HighContrast
                    ? "forced-normal-renderer" : "compiled-ui") + "; " +
                stats.ToLine());
        }
        finally
        {
            if (generatedNormal != null)
                generatedNormal.Dispose();
        }

        if (saveHighContrast)
            AssertAndSaveReviewHighContrast(raw, highlight, coverage, side == "L",
                manifest);
    }

    static object BuildReviewHighlight(object result, object pair, bool leftSide)
    {
        MethodInfo build = _app.GetType("ExcelMerger.PdfReviewForm").GetMethod(
            "BuildHighlight", BindingFlags.Static | BindingFlags.NonPublic);
        Ensure(build != null, "Review: BuildHighlight не найден в production assembly");
        return build.Invoke(null, new object[] { result, pair, leftSide });
    }

    static Bitmap RenderReviewRawPage(object page)
    {
        Type rendererType = _app.GetType("ExcelMerger.PdfThumbnailRenderer");
        Ensure(rendererType != null, "Review: PdfThumbnailRenderer не найден");
        object renderer = Activator.CreateInstance(rendererType);
        try
        {
            MethodInfo render = rendererType.GetMethod("Render", new[]
            {
                typeof(string), typeof(int), typeof(int), typeof(int)
            });
            Ensure(render != null, "Review: production Render(path,page,width,maxHeight) не найден");
            Bitmap bitmap = render.Invoke(renderer, new object[]
            {
                Convert.ToString(Member(page, "SourcePath"),
                    System.Globalization.CultureInfo.InvariantCulture),
                ReviewInt(page, "PageIndex"), 1200, 20000
            }) as Bitmap;
            Ensure(bitmap != null, "Review: повторный raw render страницы вернул null");
            int rotation = ReviewInt(page, "Rotation");
            if (rotation != 0)
            {
                MethodInfo flipFor = _app.GetType("ExcelMerger.PageRotation").GetMethod(
                    "FlipFor", BindingFlags.Static | BindingFlags.Public);
                Ensure(flipFor != null, "Review: PageRotation.FlipFor не найден");
                bitmap.RotateFlip((RotateFlipType)flipFor.Invoke(null,
                    new object[] { rotation }));
            }
            return bitmap;
        }
        finally
        {
            IDisposable disposable = renderer as IDisposable;
            if (disposable != null)
                disposable.Dispose();
        }
    }

    static Bitmap InvokeReviewHighlight(Bitmap input, object highlight, bool highContrast)
    {
        MethodInfo draw = _app.GetType("ExcelMerger.PdfReviewPageView").GetMethod(
            "DrawHighlight", BindingFlags.Static | BindingFlags.NonPublic, null,
            new[] { typeof(Bitmap), highlight.GetType(), typeof(bool) }, null);
        Ensure(draw != null, "Review: deterministic DrawHighlight overload не найден");
        Bitmap output = draw.Invoke(null, new object[] { input, highlight, highContrast })
            as Bitmap;
        Ensure(output != null, "Review: DrawHighlight вернул null");
        return output;
    }

    static IList<RectangleF> ReviewHighlightRectangles(object highlight, int width,
        int height)
    {
        MethodInfo method = _app.GetType("ExcelMerger.PdfReviewPageView").GetMethod(
            "HighlightRectangles", BindingFlags.Static | BindingFlags.NonPublic);
        Ensure(method != null, "Review: HighlightRectangles не найден");
        object value = method.Invoke(null, new object[] { highlight, width, height });
        var result = new List<RectangleF>();
        foreach (object item in Objects(value))
            result.Add((RectangleF)item);
        return result;
    }

    static bool[] ReviewCoverage(IList<RectangleF> rectangles, int width, int height)
    {
        var coverage = new bool[checked(width * height)];
        foreach (RectangleF rect in rectangles)
        {
            int left = ReviewPixelFloor(rect.Left, width);
            int top = ReviewPixelFloor(rect.Top, height);
            int right = ReviewPixelCeiling(rect.Right, width);
            int bottom = ReviewPixelCeiling(rect.Bottom, height);
            for (int y = top; y < bottom; y++)
                for (int x = left; x < right; x++)
                    coverage[y * width + x] = true;
        }
        return coverage;
    }

    static int ReviewPixelFloor(float value, int limit)
    {
        if (value <= 0f) return 0;
        if (value >= limit) return limit;
        return (int)Math.Floor(value);
    }

    static int ReviewPixelCeiling(float value, int limit)
    {
        if (value <= 0f) return 0;
        if (value >= limit) return limit;
        return (int)Math.Ceiling(value);
    }

    static ReviewVisualStats MeasureReviewNormal(Bitmap rawBitmap, Bitmap markedBitmap,
        bool[] coverage, Color fill, Color opposite)
    {
        ReviewPixelBuffer raw = ReadReviewPixels(rawBitmap);
        ReviewPixelBuffer marked = ReadReviewPixels(markedBitmap);
        Ensure(raw.Width == marked.Width && raw.Height == marked.Height &&
               coverage.Length == raw.Width * raw.Height,
            "Review: pixel buffers несовместимы");
        var stats = new ReviewVisualStats();
        for (int pixel = 0; pixel < coverage.Length; pixel++)
        {
            int offset = pixel * 4;
            bool inside = coverage[pixel];
            if (inside) stats.CoveredPixels++;
            bool changed = !ReviewPixelEquals(raw.Bytes, marked.Bytes, offset);
            if (changed)
            {
                stats.ChangedPixels++;
                if (!inside) stats.ChangedOutside++;
            }
            if (raw.Bytes[offset + 3] != marked.Bytes[offset + 3])
                stats.AlphaChanged++;

            bool introducedFill = ReviewPixelIs(marked.Bytes, offset, fill) &&
                !ReviewPixelIs(raw.Bytes, offset, fill);
            if (introducedFill)
            {
                stats.IntroducedFill++;
                if (!inside) stats.IntroducedFillOutside++;
            }
            if (ReviewPixelIs(marked.Bytes, offset, opposite) &&
                !ReviewPixelIs(raw.Bytes, offset, opposite))
                stats.IntroducedOpposite++;
            if (!inside)
                continue;

            int blue = raw.Bytes[offset];
            int green = raw.Bytes[offset + 1];
            int red = raw.Bytes[offset + 2];
            int alpha = raw.Bytes[offset + 3];
            int maximum = Math.Max(red, Math.Max(green, blue));
            int minimum = Math.Min(red, Math.Min(green, blue));
            int chroma = maximum - minimum;
            if (alpha == 255 && minimum >= 238)
            {
                stats.PaperCandidates++;
                if (ReviewPixelIs(marked.Bytes, offset, fill))
                    stats.PaperFilled++;
            }
            else if (alpha == 255 && maximum <= 150)
            {
                stats.DarkCandidates++;
                if (!changed) stats.DarkPreserved++;
            }
            else if (alpha == 255 && chroma > 18)
            {
                stats.ChromaticCandidates++;
                if (!changed) stats.ChromaticPreserved++;
            }
            else if (alpha == 255 && maximum > 150 && minimum < 238 && chroma <= 18)
            {
                stats.EdgeCandidates++;
                int luminance = (red * 54 + green * 183 + blue * 19 + 128) >> 8;
                int expectedBlue = ReviewMultiply(fill.B, luminance);
                int expectedGreen = ReviewMultiply(fill.G, luminance);
                int expectedRed = ReviewMultiply(fill.R, luminance);
                if (marked.Bytes[offset] == expectedBlue &&
                    marked.Bytes[offset + 1] == expectedGreen &&
                    marked.Bytes[offset + 2] == expectedRed &&
                    marked.Bytes[offset + 3] == alpha)
                    stats.EdgeComposed++;
            }
        }
        return stats;
    }

    static int ReviewMultiply(int channel, int factor)
    {
        return (channel * factor + 127) / 255;
    }

    static void AssertAndSaveReviewHighContrast(Bitmap raw, object highlight,
        bool[] coverage, bool leftSide, StringBuilder manifest)
    {
        Bitmap input = ReviewBitmapCopy(raw);
        Bitmap output = null;
        try
        {
            output = InvokeReviewHighlight(input, highlight, true);
            ReviewHighContrastStats stats = MeasureReviewHighContrast(raw, output,
                coverage);
            Ensure(stats.ChangedPixels > 0,
                "Review high contrast " + (leftSide ? "left" : "right") +
                ": system outline/pattern не нарисован; " + stats.ToLine());
            Ensure(stats.CoveredPaper > 0 && stats.CoveredPaperUnchanged > 0,
                "Review high contrast " + (leftSide ? "left" : "right") +
                ": word-box ошибочно превращён в сплошную custom fill; " + stats.ToLine());

            Color system = ResolveReviewHighContrastColor(highlight);
            bool systemIsDelete = system.ToArgb() == Color.FromArgb(236, 8, 8).ToArgb();
            bool systemIsInsert = system.ToArgb() == Color.FromArgb(27, 233, 26).ToArgb();
            if (!systemIsDelete)
                Ensure(stats.IntroducedDeleteFill == 0,
                    "Review high contrast: custom red использован вместо system color; " +
                    stats.ToLine());
            if (!systemIsInsert)
                Ensure(stats.IntroducedInsertFill == 0,
                    "Review high contrast: custom green использован вместо system color; " +
                    stats.ToLine());

            string style = Convert.ToString(Member(highlight, "Style"),
                System.Globalization.CultureInfo.InvariantCulture);
            MethodInfo symbolMethod = _app.GetType("ExcelMerger.PdfReviewPageView").GetMethod(
                "HighlightSymbol", BindingFlags.Static | BindingFlags.NonPublic);
            Ensure(symbolMethod != null, "Review: HighlightSymbol не найден");
            string symbol = Convert.ToString(symbolMethod.Invoke(null,
                new[] { Member(highlight, "Style") }),
                System.Globalization.CultureInfo.InvariantCulture);
            Ensure(symbol == (leftSide ? "−" : "+") &&
                   style == (leftSide ? "Removed" : "Added"),
                "Review high contrast: −/+ ownership grammar нарушена");

            string fileName = "review-page-1-high-contrast-" +
                (leftSide ? "left" : "right") + ".png";
            output.Save(Path.Combine(_out, fileName), ImageFormat.Png);
            manifest.AppendLine("high-contrast[" + (leftSide ? "L" : "R") +
                "] system=#" + system.R.ToString("X2") + system.G.ToString("X2") +
                system.B.ToString("X2") + "; symbol=" + symbol + "; " +
                stats.ToLine() + "; file=" + fileName);
        }
        finally
        {
            if (output != null)
                output.Dispose();
            else
                input.Dispose();
        }
    }

    static Color ResolveReviewHighContrastColor(object highlight)
    {
        MethodInfo resolve = _app.GetType("ExcelMerger.PdfReviewPageView").GetMethod(
            "ResolveHighlightColors", BindingFlags.Static | BindingFlags.NonPublic);
        Ensure(resolve != null, "Review: ResolveHighlightColors не найден");
        object[] args = { highlight, true, Color.Empty, Color.Empty };
        resolve.Invoke(null, args);
        Color color = (Color)args[2];
        Ensure(color.ToArgb() == SystemColors.WindowText.ToArgb() ||
               color.ToArgb() == SystemColors.Window.ToArgb(),
            "Review high contrast: цвет должен происходить из Window/WindowText");
        Ensure(((Color)args[3]).ToArgb() == color.ToArgb(),
            "Review high contrast: edge обязан использовать тот же system color");
        return color;
    }

    static ReviewHighContrastStats MeasureReviewHighContrast(Bitmap rawBitmap,
        Bitmap markedBitmap, bool[] coverage)
    {
        ReviewPixelBuffer raw = ReadReviewPixels(rawBitmap);
        ReviewPixelBuffer marked = ReadReviewPixels(markedBitmap);
        Ensure(raw.Width == marked.Width && raw.Height == marked.Height &&
               coverage.Length == raw.Width * raw.Height,
            "Review high contrast: pixel buffers несовместимы");
        var stats = new ReviewHighContrastStats();
        Color delete = Color.FromArgb(236, 8, 8);
        Color insert = Color.FromArgb(27, 233, 26);
        for (int pixel = 0; pixel < coverage.Length; pixel++)
        {
            int offset = pixel * 4;
            bool changed = !ReviewPixelEquals(raw.Bytes, marked.Bytes, offset);
            if (changed) stats.ChangedPixels++;
            if (ReviewPixelIs(marked.Bytes, offset, delete) &&
                !ReviewPixelIs(raw.Bytes, offset, delete))
                stats.IntroducedDeleteFill++;
            if (ReviewPixelIs(marked.Bytes, offset, insert) &&
                !ReviewPixelIs(raw.Bytes, offset, insert))
                stats.IntroducedInsertFill++;
            if (!coverage[pixel]) continue;
            int maximum = Math.Max(raw.Bytes[offset + 2],
                Math.Max(raw.Bytes[offset + 1], raw.Bytes[offset]));
            int minimum = Math.Min(raw.Bytes[offset + 2],
                Math.Min(raw.Bytes[offset + 1], raw.Bytes[offset]));
            if (raw.Bytes[offset + 3] == 255 && minimum >= 238 &&
                maximum - minimum <= 18)
            {
                stats.CoveredPaper++;
                if (!changed) stats.CoveredPaperUnchanged++;
            }
        }
        return stats;
    }

    static ReviewPixelBuffer ReadReviewPixels(Bitmap source)
    {
        using (Bitmap copy = ReviewBitmapCopy(source))
        {
            BitmapData data = null;
            try
            {
                data = copy.LockBits(new Rectangle(0, 0, copy.Width, copy.Height),
                    ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                int rowBytes = checked(copy.Width * 4);
                var bytes = new byte[checked(rowBytes * copy.Height)];
                var row = new byte[rowBytes];
                for (int y = 0; y < copy.Height; y++)
                {
                    Marshal.Copy(IntPtr.Add(data.Scan0, checked(y * data.Stride)),
                        row, 0, rowBytes);
                    Buffer.BlockCopy(row, 0, bytes, y * rowBytes, rowBytes);
                }
                return new ReviewPixelBuffer
                {
                    Width = copy.Width,
                    Height = copy.Height,
                    Bytes = bytes
                };
            }
            finally
            {
                if (data != null)
                    copy.UnlockBits(data);
            }
        }
    }

    static Bitmap ReviewBitmapCopy(Bitmap source)
    {
        Ensure(source != null && source.Width > 0 && source.Height > 0,
            "Review: нельзя копировать пустой bitmap");
        var copy = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
        try
        {
            using (Graphics graphics = Graphics.FromImage(copy))
            {
                graphics.CompositingMode =
                    System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                graphics.DrawImageUnscaled(source, 0, 0);
            }
            return copy;
        }
        catch
        {
            copy.Dispose();
            throw;
        }
    }

    static bool ReviewPixelEquals(byte[] left, byte[] right, int offset)
    {
        return left[offset] == right[offset] &&
            left[offset + 1] == right[offset + 1] &&
            left[offset + 2] == right[offset + 2] &&
            left[offset + 3] == right[offset + 3];
    }

    static bool ReviewPixelIs(byte[] bytes, int offset, Color color)
    {
        return bytes[offset] == color.B && bytes[offset + 1] == color.G &&
            bytes[offset + 2] == color.R && bytes[offset + 3] == color.A;
    }

    sealed class ReviewCopySnapshot
    {
        public string Text;
        public int WordCount;
        public bool UsedFallbackSeparator;

        public string ToLine()
        {
            return "words=" + WordCount + "; chars=" + (Text == null ? 0 : Text.Length) +
                "; fallback=" + UsedFallbackSeparator + "; utf8-sha256=" +
                ReviewTextSha256(Text ?? "");
        }
    }

    sealed class ReviewClipboardEnvironmentException : Exception
    {
        public ReviewClipboardEnvironmentException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }

    static void AssertReviewSelectionClipboard(Form f, StringBuilder manifest)
    {
        Control left = ReviewSurface(f, true);
        Control right = ReviewSurface(f, false);
        AssertReviewSurfaceContract(left, "left");
        AssertReviewSurfaceContract(right, "right");
        ReviewClearSelection(left);
        ReviewClearSelection(right);

        object leftModel = Member(left, "SelectionModel");
        int leftStart, leftEnd;
        FindReviewDraggableRange(left, leftModel, out leftStart, out leftEnd);
        PerformReviewMouseDrag(left, leftModel, leftStart, leftEnd);
        int leftSelected = ReviewInt(left, "SelectedWordCount");
        Ensure(leftSelected == leftEnd - leftStart + 1,
            "Review selection: forward mouse drag выбрал неверный диапазон слева; " +
            "requested=" + leftStart + ".." + leftEnd + "; actual=" +
            ReviewInt(leftModel, "SelectionStart") + ".." +
            ReviewInt(leftModel, "SelectionEnd") + "; count=" + leftSelected);
        Ensure(ReviewBool(left, "HasSelection") && !ReviewBool(right, "HasSelection"),
            "Review selection: левая mouse selection не должна появляться справа");
        ReviewCopySnapshot leftDrag = BuildReviewCopy(left);
        Ensure(leftDrag.WordCount >= 2 && !string.IsNullOrEmpty(leftDrag.Text),
            "Review selection: mouse drag должен выбрать несколько trusted words");
        ReviewShot(f, "review-selection-left", false);
        CopyReviewSelectionAndPaste(left, leftDrag, "left-mouse-drag");
        int leftDragCount = ReviewInt(left, "SelectedWordCount");

        object rightModel = Member(right, "SelectionModel");
        int rightStart, rightEnd;
        FindReviewDraggableRange(right, rightModel, out rightStart, out rightEnd);
        PerformReviewMouseDrag(right, rightModel, rightEnd, rightStart);
        Ensure(ReviewInt(right, "SelectedWordCount") == rightEnd - rightStart + 1,
            "Review selection: reverse mouse drag выбрал неверный диапазон справа");
        Ensure(ReviewInt(left, "SelectedWordCount") == leftDragCount,
            "Review selection: правая mouse selection изменила левую");
        ReviewCopySnapshot rightDrag = BuildReviewCopy(right);
        Ensure(rightDrag.WordCount >= 2 && !string.IsNullOrEmpty(rightDrag.Text),
            "Review selection: reverse drag должен строить canonical copy справа");
        CopyReviewSelectionAndPaste(right, rightDrag, "right-reverse-drag");

        left.Focus();
        Ensure(InvokeReviewSurfaceKey(left, Keys.Control | Keys.A),
            "Review selection: Ctrl+A слева не обработан page surface");
        ReviewCopySnapshot leftAll = BuildReviewCopy(left);
        Ensure(leftAll.WordCount == ReviewInt(leftModel, "Count") &&
               ReviewInt(left, "SelectedWordCount") == leftAll.WordCount,
            "Review selection: Ctrl+A должен выбрать все trusted words слева");
        Ensure(ReviewInt(right, "SelectedWordCount") == rightDrag.WordCount,
            "Review selection: Ctrl+A слева изменил правый диапазон");
        Ensure(ContainsNonAscii(leftAll.Text),
            "Review clipboard: левая fixture должна проверять не-ASCII UnicodeText");
        CopyReviewSelectionAndPaste(left, leftAll, "left-ctrl-a-c");

        right.Focus();
        Ensure(InvokeReviewSurfaceKey(right, Keys.Control | Keys.A),
            "Review selection: Ctrl+A справа не обработан page surface");
        ReviewCopySnapshot rightAll = BuildReviewCopy(right);
        Ensure(rightAll.WordCount == ReviewInt(rightModel, "Count") &&
               ReviewInt(right, "SelectedWordCount") == rightAll.WordCount,
            "Review selection: Ctrl+A должен выбрать все trusted words справа");
        Ensure(ReviewInt(left, "SelectedWordCount") == leftAll.WordCount,
            "Review selection: Ctrl+A справа изменил левый диапазон");
        Ensure(ContainsNonAscii(rightAll.Text),
            "Review clipboard: правая fixture должна проверять не-ASCII UnicodeText");
        CopyReviewSelectionAndPaste(right, rightAll, "right-ctrl-a-c");

        object rightModelBeforeNavigation = Member(right, "SelectionModel");
        TextBox leftPage = (TextBox)Field(f, "_leftPageInput");
        leftPage.Text = "2";
        Ensure(InvokeReviewNavigate(f, true),
            "Review selection: не удалось открыть физическую страницу 2 слева");
        WaitReviewPanes(f, "selection lifecycle → левая страница 2");
        Ensure(!ReviewBool(left, "HasSelection") &&
               !ReferenceEquals(leftModel, Member(left, "SelectionModel")),
            "Review selection: навигация обязана удалить stale left selection/model");
        Ensure(ReviewBool(right, "HasSelection") &&
               ReferenceEquals(rightModelBeforeNavigation, Member(right, "SelectionModel")) &&
               ReviewInt(right, "SelectedWordCount") == rightAll.WordCount,
            "Review selection: левая навигация не должна очищать правую selection");

        leftPage.Text = "1";
        Ensure(InvokeReviewNavigate(f, true),
            "Review selection: не удалось вернуть физическую страницу 1 слева");
        WaitReviewPanes(f, "selection lifecycle → возврат левой страницы 1");
        Ensure(!ReviewBool(left, "HasSelection") && ReviewBool(right, "HasSelection"),
            "Review selection: новая левая page layer должна быть чистой, правая — сохраниться");
        ReviewClearSelection(right);
        Ensure(!ReviewBool(left, "HasSelection") && !ReviewBool(right, "HasSelection"),
            "Review selection: cleanup не очистил независимые диапазоны");
        f.Activate();
        Pump(100);

        manifest.AppendLine("selection[left-drag]: " + leftDrag.ToLine());
        manifest.AppendLine("selection[right-reverse-drag]: " + rightDrag.ToLine());
        manifest.AppendLine("selection[left-ctrl-a-c]: " + leftAll.ToLine());
        manifest.AppendLine("selection[right-ctrl-a-c]: " + rightAll.ToLine());
        manifest.AppendLine("clipboard=DataFormats.UnicodeText+persistent production writer; " +
            "paste=independent multiline TextBox exact; panes=isolated; navigation-clears-stale=ok");
    }

    static Control ReviewSurface(Form f, bool leftSide)
    {
        object view = Field(f, leftSide ? "_leftSource" : "_rightSource");
        Control surface = Member(view, "_picture") as Control;
        Ensure(surface != null, "Review: production page surface отсутствует " +
            (leftSide ? "слева" : "справа"));
        return surface;
    }

    static void AssertReviewSurfaceContract(Control surface, string side)
    {
        Ensure(surface.Visible && surface.Enabled && surface.ClientSize.Width > 0 &&
               surface.ClientSize.Height > 0,
            "Review selection " + side + ": ready surface должна быть видима");
        Ensure(surface is PictureBox && surface.AccessibleRole == AccessibleRole.Document,
            "Review selection " + side +
            ": поверхность должна оставаться PictureBox с ролью Document");
        Ensure(ReviewBool(surface, "HasSelectableText") && surface.TabStop &&
               Member(surface, "SelectionModel") != null,
            "Review selection " + side + ": trusted text layer не прикреплён");
        Ensure(!string.IsNullOrWhiteSpace(surface.AccessibleName) &&
               !string.IsNullOrWhiteSpace(surface.AccessibleDescription) &&
               surface.AccessibleDescription.IndexOf(Text("review.source.interactions"),
                   StringComparison.Ordinal) >= 0,
            "Review selection " + side +
            ": read-only copy interaction не описан accessibility-текстом");
        ContextMenuStrip menu = surface.ContextMenuStrip;
        Ensure(menu != null && menu.Items.Count == 2 &&
               menu.Items[0].Text == Text("common.copy") &&
               menu.Items[1].Text == Text("review.selection.selectAll"),
            "Review selection " + side +
            ": локализованное Copy/Select all context menu нарушено");
    }

    static void FindReviewDraggableRange(Control surface, object model, out int start,
        out int end)
    {
        start = end = -1;
        Size surfaceSize = surface.ClientSize;
        int count = ReviewInt(model, "Count");
        Ensure(count >= 2, "Review selection: для drag нужны хотя бы два trusted words");
        MethodInfo hit = model.GetType().GetMethod("HitTest", new[]
        {
            typeof(PointF), typeof(Size)
        });
        Ensure(hit != null, "Review selection: HitTest не найден");
        Rectangle safeScreen = ReviewSafeDragScreenBounds(surface);
        int search = Math.Min(count, 40);
        for (int i = 0; i < search && start < 0; i++)
        {
            RectangleF first = ReviewWordRectangle(model, i, surfaceSize);
            PointF firstPoint = new PointF(first.Left + first.Width / 2f,
                first.Top + first.Height / 2f);
            if ((int)hit.Invoke(model, new object[] { firstPoint, surfaceSize }) != i ||
                !safeScreen.Contains(surface.PointToScreen(Point.Round(firstPoint))))
                continue;
            for (int j = i + 1; j < search && j <= i + 5; j++)
            {
                RectangleF last = ReviewWordRectangle(model, j, surfaceSize);
                PointF lastPoint = new PointF(last.Left + last.Width / 2f,
                    last.Top + last.Height / 2f);
                if ((int)hit.Invoke(model, new object[] { lastPoint, surfaceSize }) == j &&
                    safeScreen.Contains(surface.PointToScreen(Point.Round(lastPoint))))
                {
                    start = i;
                    end = j;
                    break;
                }
            }
        }
        Ensure(start >= 0 && end > start,
            "Review selection: не найден неперекрывающийся mouse-drag range " +
            "вне зоны autoscroll");
    }

    static Rectangle ReviewSafeDragScreenBounds(Control surface)
    {
        Control viewport = surface.Parent;
        Rectangle visible = viewport == null
            ? surface.RectangleToScreen(surface.ClientRectangle)
            : viewport.RectangleToScreen(viewport.ClientRectangle);
        const int margin = 40;
        if (visible.Width > margin * 2 && visible.Height > margin * 2)
            visible.Inflate(-margin, -margin);
        return visible;
    }

    static RectangleF ReviewWordRectangle(object model, int index, Size surfaceSize)
    {
        MethodInfo method = model.GetType().GetMethod("WordRectangle", new[]
        {
            typeof(int), typeof(Size)
        });
        Ensure(method != null, "Review selection: WordRectangle не найден");
        RectangleF rect = (RectangleF)method.Invoke(model,
            new object[] { index, surfaceSize });
        Ensure(!rect.IsEmpty && rect.Width > 0 && rect.Height > 0,
            "Review selection: trusted word не имеет surface rectangle #" + index);
        return rect;
    }

    static void PerformReviewMouseDrag(Control surface, object model, int from, int to)
    {
        RectangleF first = ReviewWordRectangle(model, from, surface.ClientSize);
        RectangleF last = ReviewWordRectangle(model, to, surface.ClientSize);
        Point start = ReviewRectCenter(first, surface.ClientRectangle);
        Point end = ReviewRectCenter(last, surface.ClientRectangle);
        Form owner = surface.FindForm();
        if (owner != null)
        {
            owner.Activate();
            SetForegroundWindow(owner.Handle);
        }
        surface.Focus();
        Ensure(surface.Focused, "Review selection: page surface не получил focus");

        const uint MouseEventLeftDown = 0x0002;
        const uint MouseEventLeftUp = 0x0004;
        Point originalCursor = Cursor.Position;
        bool buttonDown = false;
        try
        {
            Cursor.Position = surface.PointToScreen(start);
            Pump(50);
            mouse_event(MouseEventLeftDown, 0, 0, 0, UIntPtr.Zero);
            buttonDown = true;
            Pump(50);
            Cursor.Position = surface.PointToScreen(end);
            Pump(100);
            mouse_event(MouseEventLeftUp, 0, 0, 0, UIntPtr.Zero);
            buttonDown = false;
            Pump(100);
        }
        finally
        {
            if (buttonDown)
                mouse_event(MouseEventLeftUp, 0, 0, 0, UIntPtr.Zero);
            Cursor.Position = originalCursor;
            Pump(20);
        }
        Ensure(!surface.Capture,
            "Review selection: mouse-up обязан освободить capture/autoscroll lifecycle");
    }

    static Point ReviewRectCenter(RectangleF rect, Rectangle bounds)
    {
        int x = (int)Math.Round(rect.Left + rect.Width / 2f);
        int y = (int)Math.Round(rect.Top + rect.Height / 2f);
        x = Math.Max(bounds.Left, Math.Min(bounds.Right - 1, x));
        y = Math.Max(bounds.Top, Math.Min(bounds.Bottom - 1, y));
        return new Point(x, y);
    }

    static ReviewCopySnapshot BuildReviewCopy(Control surface)
    {
        object model = Member(surface, "SelectionModel");
        Ensure(model != null, "Review clipboard: selection model отсутствует");
        MethodInfo build = model.GetType().GetMethod("BuildCopyText",
            BindingFlags.Instance | BindingFlags.Public);
        Ensure(build != null, "Review clipboard: BuildCopyText не найден");
        object copy = build.Invoke(model, null);
        return new ReviewCopySnapshot
        {
            Text = Convert.ToString(Member(copy, "Text"),
                System.Globalization.CultureInfo.InvariantCulture),
            WordCount = ReviewInt(copy, "WordCount"),
            UsedFallbackSeparator = ReviewBool(copy, "UsedFallbackSeparator")
        };
    }

    static bool InvokeReviewSurfaceKey(Control surface, Keys key)
    {
        MethodInfo method = surface.GetType().GetMethod("ProcessCmdKey",
            BindingFlags.Instance | BindingFlags.NonPublic);
        Ensure(method != null, "Review selection: ProcessCmdKey не найден");
        object[] args = { new Message(), key };
        return (bool)method.Invoke(surface, args);
    }

    static void ReviewClearSelection(Control surface)
    {
        MethodInfo clear = surface.GetType().GetMethod("ClearSelection",
            BindingFlags.Instance | BindingFlags.NonPublic);
        Ensure(clear != null, "Review selection: ClearSelection не найден");
        clear.Invoke(surface, null);
    }

    static void CopyReviewSelectionAndPaste(Control surface, ReviewCopySnapshot expected,
        string context)
    {
        Ensure(expected != null && expected.WordCount > 0 &&
               !string.IsNullOrEmpty(expected.Text),
            "Review clipboard " + context + ": expected trusted text пуст");
        surface.Focus();
        Ensure(InvokeReviewSurfaceKey(surface, Keys.Control | Keys.C),
            "Review clipboard " + context + ": Ctrl+C не обработан");
        Pump(100);

        string status = Convert.ToString(Member(surface, "InteractionStatus"),
            System.Globalization.CultureInfo.InvariantCulture);
        if (status == Text("review.selection.clipboardUnavailable"))
            throw new ReviewClipboardEnvironmentException(
                "ENVIRONMENTAL CLIPBOARD FAILURE: production writer сообщил busy/unavailable (" +
                context + ")", null);
        string expectedStatus = Text(expected.UsedFallbackSeparator
            ? "review.selection.copiedFallback" : "review.selection.copied");
        Ensure(status == expectedStatus,
            "Review clipboard " + context +
            ": feedback не соответствует fallback provenance");

        string clipboard = ReadReviewClipboardUnicode(context);
        Ensure(string.Equals(clipboard, expected.Text, StringComparison.Ordinal),
            "Review clipboard " + context +
            ": DataFormats.UnicodeText не равен trusted copy builder; expected=" +
            expected.ToLine() + ", actual-chars=" + clipboard.Length +
            ", actual-sha256=" + ReviewTextSha256(clipboard));
        string pasted = PasteReviewClipboardIntoIndependentTextBox();
        Ensure(string.Equals(pasted, expected.Text, StringComparison.Ordinal),
            "Review clipboard " + context +
            ": independent TextBox изменил UnicodeText; expected=" +
            expected.ToLine() + ", pasted-chars=" + pasted.Length +
            ", pasted-sha256=" + ReviewTextSha256(pasted));
    }

    static string ReadReviewClipboardUnicode(string context)
    {
        Exception last = null;
        for (int attempt = 0; attempt < 5; attempt++)
        {
            try
            {
                IDataObject data = Clipboard.GetDataObject();
                Ensure(data != null, "Review clipboard " + context +
                    ": Windows clipboard не вернул IDataObject");
                Ensure(data.GetDataPresent(DataFormats.UnicodeText, false),
                    "Review clipboard " + context +
                    ": отсутствует стандартный DataFormats.UnicodeText");
                object value = data.GetData(DataFormats.UnicodeText, false);
                Ensure(value is string, "Review clipboard " + context +
                    ": UnicodeText имеет неожиданный тип");
                return (string)value;
            }
            catch (ExternalException ex)
            {
                last = ex;
            }
            catch (ThreadStateException ex)
            {
                last = ex;
            }
            if (attempt < 4)
                Thread.Sleep(100);
        }
        throw new ReviewClipboardEnvironmentException(
            "ENVIRONMENTAL CLIPBOARD FAILURE: UnicodeText нельзя прочитать после bounded retry (" +
            context + ")", last);
    }

    static string PasteReviewClipboardIntoIndependentTextBox()
    {
        using (var host = new Form())
        using (var destination = new TextBox())
        {
            host.ShowInTaskbar = false;
            host.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            host.StartPosition = FormStartPosition.Manual;
            host.Location = new Point(-30000, -30000);
            host.ClientSize = new Size(520, 260);
            destination.Multiline = true;
            destination.AcceptsReturn = true;
            destination.AcceptsTab = true;
            destination.MaxLength = 0;
            destination.Dock = DockStyle.Fill;
            host.Controls.Add(destination);
            host.Show();
            destination.Focus();
            destination.Paste();
            Pump(100);
            string value = destination.Text;
            host.Close();
            return value;
        }
    }

    static bool ContainsNonAscii(string value)
    {
        if (value == null) return false;
        foreach (char character in value)
            if (character > 127)
                return true;
        return false;
    }

    static string ReviewTextSha256(string value)
    {
        using (var sha = SHA256.Create())
        {
            byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? ""));
            var text = new StringBuilder(hash.Length * 2);
            foreach (byte item in hash)
                text.Append(item.ToString("x2",
                    System.Globalization.CultureInfo.InvariantCulture));
            return text.ToString();
        }
    }

    static void AssertReviewFixtureResult(object result, bool reverse)
    {
        Ensure(result != null, "Review: production result отсутствует");
        object stats = Member(result, "Stats");
        Ensure(ReviewInt(stats, "PagePairs") == 3,
            "Review: ожидались три пары физических страниц");
        Ensure(ReviewInt(stats, "ChangedPages") == 3,
            "Review: правки должны остаться на трёх страницах");
        Ensure(ReviewInt(stats, "LeftOnlyPages") == 0 &&
               ReviewInt(stats, "RightOnlyPages") == 0,
            "Review: не должно быть ложных односторонних страниц");
        int expectedDeleted = reverse ? 3 : 51;
        int expectedInserted = reverse ? 51 : 3;
        Ensure(ReviewInt(stats, "DeletedWords") == expectedDeleted &&
               ReviewInt(stats, "InsertedWords") == expectedInserted &&
               ReviewInt(stats, "Replacements") == 1,
            "Review: ожидаются deleted=" + expectedDeleted + ", inserted=" +
            expectedInserted + ", replacements=1, фактически deleted=" +
            ReviewInt(stats, "DeletedWords") + ", inserted=" +
            ReviewInt(stats, "InsertedWords") + ", replacements=" +
            ReviewInt(stats, "Replacements"));
        Ensure(ReviewInt(stats, "WhitespaceChanges") == 0 &&
               ReviewInt(stats, "DeletedWhitespaceAtoms") == 0 &&
               ReviewInt(stats, "InsertedWhitespaceAtoms") == 0,
            "Review: геометрия фикстур не должна выдумывать пробельные правки");

        IList pairList = Member(result, "Pairs") as IList;
        Ensure(pairList != null && pairList.Count == 3,
            "Review: result.Pairs должен содержать три строки");
        for (int i = 0; i < pairList.Count; i++)
        {
            object pair = pairList[i];
            Ensure(ReviewInt(pair, "LeftPageIndex") == i &&
                   ReviewInt(pair, "RightPageIndex") == i &&
                   Convert.ToString(Member(pair, "Status"),
                       System.Globalization.CultureInfo.InvariantCulture) == "Changed",
                "Review: viewer row " + i + " обязан быть " + i + "↔" + i + ":Changed");
        }

        string removedKind = reverse ? "Insert" : "Delete";
        string addedKind = reverse ? "Delete" : "Insert";
        Ensure(CountReviewWordAt(result, removedKind, "Оказание", 0,
                   169.5, 451.5) == 1 &&
               CountReviewWordAt(result, removedKind, "колодцев", 0,
                   135.75, 438) == 1,
            "Review: старое описание услуги должно принадлежать стороне 1.pdf");
        Ensure(CountReviewWordAt(result, addedKind, "рпарарпапрапрап", 0,
                   169.5, 451.5) == 1,
            "Review: новый текст должен принадлежать стороне 2.pdf");
        Ensure(CountReviewWordAt(result, removedKind, "Основание:", 0,
                   169.5, 354) == 1,
            "Review: удалённый раздел «Основание:» должен оставаться видимым");
        foreach (double bottom in new[] { 339.75, 326.25, 312.75 })
            Ensure(CountReviewWordAt(result, removedKind, "Коммерческое", 0,
                       187.5, bottom) == 1,
                "Review: строка коммерческого предложения должна оставаться видимой @" + bottom);
        Ensure(CountReviewWordAt(result, "Delete", "2", 1, 418.5, 567) == 1 &&
               CountReviewWordAt(result, "Insert", "2", 1, 418.5, 567) == 1 &&
               CountReviewWordAt(result, "Delete", "3", 2, 418.5, 567) == 1 &&
               CountReviewWordAt(result, "Insert", "3", 2, 418.5, 567) == 1,
            "Review: консервативные номера страниц 2/3 должны остаться двусторонними кандидатами");

        EnsureStableReviewWord(result, reverse, "с.", 249, 57, 249, 108);
        EnsureStableReviewWord(result, reverse, "Б-Глушица", 255.747, 57, 255.747, 108);
        EnsureStableReviewWord(result, reverse, "колодец", 248.251, 48.752, 248.251, 99.752);
        EnsureStableReviewWord(result, reverse, "№53", 274.501, 48.752, 274.501, 99.752);
        EnsureStableReviewWord(result, reverse, "7", 159.75, 44.25, 159.75, 96);
        EnsureStableReviewWord(result, reverse, "Гагарина", 180.749, 44.252, 180.749, 99.752);
        EnsureStableReviewWord(result, reverse, "по", 221.255, 44.252, 221.255, 99.752);
        EnsureStableReviewWord(result, reverse, "по", 259.503, 40.503, 259.503, 91.503);
        EnsureStableReviewWord(result, reverse, "ул.", 268.502, 40.503, 268.502, 91.503);
        EnsureStableReviewWord(result, reverse, "ул.", 185.251, 36.752, 185.251, 91.503);
        EnsureStableReviewWord(result, reverse, "Гагарина,", 195.748, 36.752, 195.748, 91.503);
        EnsureStableReviewWord(result, reverse, "Ленинградской", 246.003, 32.255, 246.003, 84.004);
        EnsureStableReviewWord(result, reverse, "Пугачевская,", 186, 28.5, 186, 84.004);
        EnsureStableReviewWord(result, reverse, "№", 158.25, 288.75, 158.25, 344.25);
    }

    static void EnsureStableReviewWord(object result, bool reverse, string key,
        double earlyLeft, double earlyBottom, double lateLeft, double lateBottom)
    {
        double left = reverse ? lateLeft : earlyLeft;
        double leftBottom = reverse ? lateBottom : earlyBottom;
        double right = reverse ? earlyLeft : lateLeft;
        double rightBottom = reverse ? earlyBottom : lateBottom;
        Ensure(CountReviewWordAt(result, "Delete", key, 0, left, leftBottom) == 0 &&
               CountReviewWordAt(result, "Insert", key, 0, right, rightBottom) == 0,
            "Review: неизменённый физический экземпляр «" + key +
            "» не должен подсвечиваться как правка");
    }

    static int CountReviewWordAt(object result, string kind, string key, int pageIndex,
        double left, double bottom)
    {
        int count = 0;
        foreach (object operation in Objects(Member(result, "Operations")))
        {
            if (!string.Equals(Convert.ToString(Member(operation, "Kind"),
                    System.Globalization.CultureInfo.InvariantCulture), kind,
                    StringComparison.Ordinal))
                continue;
            string side = kind == "Delete" ? "LeftWords" : "RightWords";
            foreach (object word in Objects(Member(operation, side)))
            {
                if (word == null ||
                    !string.Equals(Convert.ToString(Member(word, "Key"),
                        System.Globalization.CultureInfo.InvariantCulture), key,
                        StringComparison.Ordinal) ||
                    ReviewInt(word, "PageIndex") != pageIndex)
                    continue;
                object box = Member(word, "Box");
                if (Math.Abs(ReviewDouble(box, "Left") - left) <= 0.05 &&
                    Math.Abs(ReviewDouble(box, "Bottom") - bottom) <= 0.05)
                    count++;
            }
        }
        return count;
    }

    static void AssertReviewLegendsAndLayout(Form f, string context)
    {
        Label leftLegend = (Label)Field(f, "_leftLegend");
        Label rightLegend = (Label)Field(f, "_rightLegend");
        Ensure(leftLegend.Text == Text("review.legend.removed") &&
               rightLegend.Text == Text("review.legend.added"),
            "Review " + context + ": compact colour legends должны быть локализованы дословно");
        Ensure(leftLegend.Text.IndexOf('−') >= 0 && rightLegend.Text.IndexOf('+') >= 0,
            "Review " + context + ": смысл обязан читаться по −/+ без цвета");
        Ensure(!leftLegend.TabStop && !rightLegend.TabStop,
            "Review " + context + ": декоративные легенды не должны попадать в tab order");
        Ensure(leftLegend.Visible && rightLegend.Visible,
            "Review " + context + ": легенды не должны скрываться");

        Rectangle leftLegendRect = leftLegend.RectangleToScreen(leftLegend.ClientRectangle);
        Rectangle rightLegendRect = rightLegend.RectangleToScreen(rightLegend.ClientRectangle);
        Rectangle client = f.RectangleToScreen(f.ClientRectangle);
        Ensure(client.Contains(leftLegendRect) && client.Contains(rightLegendRect) &&
               leftLegendRect.Width > 20 && rightLegendRect.Width > 20 &&
               !leftLegendRect.IntersectsWith(rightLegendRect),
            "Review " + context + ": легенды должны оставаться внутри своих неперекрывающихся pane");
        AssertReviewLegendPaintLayer(leftLegend,
            (TextBox)Field(f, "_leftPageInput"), context + " left");
        AssertReviewLegendPaintLayer(rightLegend,
            (TextBox)Field(f, "_rightPageInput"), context + " right");

        Control leftPath = (Control)Field(f, "_leftPath");
        Control rightPath = (Control)Field(f, "_rightPath");
        Control swap = (Control)Field(f, "_swap");
        Control compare = (Control)Field(f, "_compare");
        Control summary = (Control)Field(f, "_summary");
        Ensure(!leftPath.Bounds.IntersectsWith(rightPath.Bounds),
            "Review " + context + ": поля ранней и поздней версии не должны пересекаться");
        Ensure(summary.Left >= swap.Right && summary.Right <= compare.Left,
            "Review " + context + ": строка статистики должна оставаться между кнопками");
    }

    static void AssertReviewLegendPaintLayer(Label ownership,
        TextBox pageInput, string context)
    {
        Control legendLayer = ownership.Parent;
        Control navigator = pageInput.Parent;
        Control top = legendLayer == null ? null : legendLayer.Parent;
        Ensure(top != null && navigator != null && ReferenceEquals(top, navigator.Parent),
            "Review " + context + ": легенда и навигатор должны иметь общий layout");

        Rectangle ownershipRect = ownership.RectangleToScreen(ownership.ClientRectangle);
        Rectangle navigatorRect = navigator.RectangleToScreen(navigator.ClientRectangle);
        Ensure(ownershipRect.Height >= ownership.Font.Height &&
               !ownershipRect.IntersectsWith(navigatorRect) && top.Height <= 80,
            "Review " + context + ": компактная легенда не должна закрывать PDF или навигатор");
        Point center = new Point(ownershipRect.Left + ownershipRect.Width / 2,
            ownershipRect.Top + ownershipRect.Height / 2);
        Control topmost = top.GetChildAtPoint(top.PointToClient(center),
            GetChildAtPointSkip.Invisible);
        Ensure(ReferenceEquals(legendLayer, topmost),
            "Review " + context + ": легенда должна быть верхним painted layout-слоем");
    }

    static string ReviewResultDetails(object result)
    {
        var text = new StringBuilder();
        object stats = Member(result, "Stats");
        text.Append("stats: pairs=").Append(ReviewInt(stats, "PagePairs"))
            .Append(" changed=").Append(ReviewInt(stats, "ChangedPages"))
            .Append(" left-only=").Append(ReviewInt(stats, "LeftOnlyPages"))
            .Append(" right-only=").Append(ReviewInt(stats, "RightOnlyPages"))
            .Append(" deleted=").Append(ReviewInt(stats, "DeletedWords"))
            .Append(" inserted=").Append(ReviewInt(stats, "InsertedWords"))
            .Append(" replacements=").Append(ReviewInt(stats, "Replacements"))
            .Append(" whitespace=").Append(ReviewInt(stats, "WhitespaceChanges"))
            .Append(" deleted-whitespace-atoms=").Append(ReviewInt(stats,
                "DeletedWhitespaceAtoms"))
            .Append(" inserted-whitespace-atoms=").Append(ReviewInt(stats,
                "InsertedWhitespaceAtoms")).AppendLine();

        int pairIndex = 0;
        foreach (object pair in Objects(Member(result, "Pairs")))
        {
            text.Append("pair[").Append(pairIndex++).Append("]=")
                .Append(ReviewInt(pair, "LeftPageIndex")).Append("↔")
                .Append(ReviewInt(pair, "RightPageIndex")).Append(':')
                .Append(Member(pair, "Status")).AppendLine();
        }

        int operationIndex = 0;
        foreach (object operation in Objects(Member(result, "Operations")))
        {
            text.Append("op[").Append(operationIndex++).Append("]=")
                .Append(Member(operation, "Kind")).Append('/')
                .Append(Member(operation, "MatchKind"))
                .Append(" matches=").Append(CountObjects(Member(operation, "Matches")))
                .Append(" left=");
            AppendReviewWords(text, Member(operation, "LeftWords"));
            text.Append(" right=");
            AppendReviewWords(text, Member(operation, "RightWords"));
            text.AppendLine();
        }

        MethodInfo build = _app.GetType("ExcelMerger.PdfReviewForm").GetMethod(
            "BuildHighlight", BindingFlags.Static | BindingFlags.NonPublic);
        int highlightPair = 0;
        foreach (object pair in Objects(Member(result, "Pairs")))
        {
            foreach (bool leftSide in new[] { true, false })
            {
                object highlight = build.Invoke(null, new object[] { result, pair, leftSide });
                text.Append("highlight[").Append(highlightPair).Append(leftSide ? ",L" : ",R")
                    .Append("]: boxes=").Append(CountObjects(Member(highlight, "Boxes")))
                    .Append(" whitespace=").Append(CountObjects(Member(highlight,
                        "WhitespaceMarkers")))
                    .Append(" style=").Append(Member(highlight, "Style"))
                    .Append(" rail=").Append(Member(highlight, "ChangeBarSide"))
                    .AppendLine();
            }
            highlightPair++;
        }
        return text.ToString();
    }

    static void AppendReviewWords(StringBuilder text, object words)
    {
        text.Append('[');
        bool first = true;
        foreach (object word in Objects(words))
        {
            if (!first) text.Append(';');
            first = false;
            object box = Member(word, "Box");
            text.Append(EscapeManifest(Convert.ToString(Member(word, "Key"),
                    System.Globalization.CultureInfo.InvariantCulture)))
                .Append("@p").Append(ReviewInt(word, "PageIndex"))
                .Append('(').Append(Invariant(ReviewDouble(box, "Left"))).Append(',')
                .Append(Invariant(ReviewDouble(box, "Bottom"))).Append(',')
                .Append(Invariant(ReviewDouble(box, "Right"))).Append(',')
                .Append(Invariant(ReviewDouble(box, "Top"))).Append(')');
        }
        text.Append(']');
    }

    static string EscapeManifest(string value)
    {
        return (value ?? "").Replace("\\", "\\\\").Replace("\r", "\\r")
            .Replace("\n", "\\n").Replace("\t", "\\t")
            .Replace(";", "\\;");
    }

    static string Invariant(double value)
    {
        return value.ToString("0.###",
            System.Globalization.CultureInfo.InvariantCulture);
    }

    static IEnumerable<object> Objects(object source)
    {
        IEnumerable values = source as IEnumerable;
        if (values == null)
            yield break;
        foreach (object value in values)
            yield return value;
    }

    static int CountObjects(object source)
    {
        ICollection collection = source as ICollection;
        if (collection != null)
            return collection.Count;
        int count = 0;
        foreach (object ignored in Objects(source)) count++;
        return count;
    }

    static bool InvokeReviewKey(Form f, Keys key)
    {
        MethodInfo method = f.GetType().GetMethod("ProcessCmdKey",
            BindingFlags.Instance | BindingFlags.NonPublic);
        object[] args = { new Message(), key };
        return (bool)method.Invoke(f, args);
    }

    static bool InvokeReviewNavigate(Form f, bool leftSide)
    {
        MethodInfo method = f.GetType().GetMethod("NavigateToPhysicalPage",
            BindingFlags.Instance | BindingFlags.NonPublic, null,
            new[] { typeof(bool) }, null);
        return (bool)method.Invoke(f, new object[] { leftSide });
    }

    static bool InvokeReviewWheel(Form f, Point screenPoint, int delta,
        bool controlDown)
    {
        MethodInfo method = f.GetType().GetMethod("RouteWheel",
            BindingFlags.Instance | BindingFlags.NonPublic, null,
            new[] { typeof(Point), typeof(int), typeof(bool) }, null);
        return (bool)method.Invoke(f, new object[] { screenPoint, delta, controlDown });
    }

    static bool ReviewWorking(Form f)
    {
        return (bool)Member(f, "Working");
    }

    static bool ReviewBool(object target, string name)
    {
        return Convert.ToBoolean(Member(target, name),
            System.Globalization.CultureInfo.InvariantCulture);
    }

    static int ReviewInt(object target, string name)
    {
        return Convert.ToInt32(Member(target, name),
            System.Globalization.CultureInfo.InvariantCulture);
    }

    static double ReviewDouble(object target, string name)
    {
        return Convert.ToDouble(Member(target, name),
            System.Globalization.CultureInfo.InvariantCulture);
    }

    static object Member(object target, string name)
    {
        if (target == null)
            throw new Exception("нет объекта для члена " + name);
        Type type = target.GetType();
        while (type != null)
        {
            FieldInfo field = type.GetField(name, BindingFlags.Instance |
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.DeclaredOnly);
            if (field != null)
                return field.GetValue(target);
            PropertyInfo property = type.GetProperty(name, BindingFlags.Instance |
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.DeclaredOnly);
            if (property != null)
                return property.GetValue(target, null);
            type = type.BaseType;
        }
        throw new Exception("нет члена " + name + " в " + target.GetType().Name);
    }

    static string ReviewState(Form f)
    {
        try
        {
            object left = Field(f, "_leftSource");
            object right = Field(f, "_rightSource");
            return "working=" + ReviewWorking(f) +
                ", sourceChecking=" + ReviewBool(f, "_leftSourceChecking") + "/" +
                    ReviewBool(f, "_rightSourceChecking") +
                ", compare=" + ((Button)Field(f, "_compare")).Enabled +
                ", result=" + (Field(f, "_result") != null) +
                ", panes=" + Member(left, "ViewState") + "/" +
                    Member(right, "ViewState") +
                ", workers=" + ReviewBool(left, "_renderWorker") + "/" +
                    ReviewBool(right, "_renderWorker") +
                ", rows=" + ReviewInt(f, "_leftRowIndex") + "/" +
                    ReviewInt(f, "_rightRowIndex");
        }
        catch (Exception ex)
        {
            return "state unavailable: " + RootCause(ex).Message;
        }
    }

    static void WaitUntil(string description, Func<bool> ready, int timeoutMs,
        Func<string> diagnostics)
    {
        Stopwatch timer = Stopwatch.StartNew();
        while (timer.ElapsedMilliseconds < timeoutMs)
        {
            Application.DoEvents();
            if (ready())
                return;
            Thread.Sleep(25);
        }
        Application.DoEvents();
        if (ready())
            return;
        throw new TimeoutException(description + " не завершено за " + timeoutMs +
            " мс; " + (diagnostics == null ? "" : diagnostics()));
    }

    static void Ensure(bool condition, string message)
    {
        if (!condition)
            throw new Exception(message);
    }

    static Exception RootCause(Exception error)
    {
        Exception current = error;
        while (current is TargetInvocationException && current.InnerException != null)
            current = current.InnerException;
        return current;
    }

    static Point ScreenCenter(Control control)
    {
        Rectangle bounds = control.RectangleToScreen(control.ClientRectangle);
        return new Point(bounds.Left + bounds.Width / 2,
            bounds.Top + bounds.Height / 2);
    }

    static void ReviewShot(Form f, string name, bool grayscale)
    {
        Shot(f, name);
        if (grayscale)
            MakeGrayscale(Path.Combine(_out, name + ".png"),
                Path.Combine(_out, name + "-grayscale.png"));
    }

    static void MakeGrayscale(string sourcePath, string destinationPath)
    {
        using (var source = new Bitmap(sourcePath))
        using (var gray = new Bitmap(source.Width, source.Height,
            PixelFormat.Format32bppArgb))
        using (Graphics graphics = Graphics.FromImage(gray))
        using (var attributes = new ImageAttributes())
        {
            var matrix = new ColorMatrix(new[]
            {
                new[] { 0.299f, 0.299f, 0.299f, 0f, 0f },
                new[] { 0.587f, 0.587f, 0.587f, 0f, 0f },
                new[] { 0.114f, 0.114f, 0.114f, 0f, 0f },
                new[] { 0f, 0f, 0f, 1f, 0f },
                new[] { 0f, 0f, 0f, 0f, 1f }
            });
            attributes.SetColorMatrix(matrix);
            graphics.DrawImage(source, new Rectangle(0, 0, gray.Width, gray.Height),
                0, 0, source.Width, source.Height, GraphicsUnit.Pixel, attributes);
            gray.Save(destinationPath, ImageFormat.Png);
        }
        Say(Path.GetFileName(destinationPath) + " -> grayscale");
    }
    static void WhatsNew()
    {
        Type type = _app.GetType("ExcelMerger.WhatsNewForm");
        ConstructorInfo ctor = type.GetConstructor(BindingFlags.NonPublic |
            BindingFlags.Instance, null, new[] { typeof(string) }, null);
        var form = (Form)ctor.Invoke(new object[] { "1.18.5" });
        Place(form, 0, 0);
        Shot(form, "whatsnew");

        var support = (LinkLabel)Field(form, "_supportLink");
        support.Links[0].LinkData = null;
        support.Focus();
        support.GetType().GetMethod("OnLinkClicked", BindingFlags.Instance |
            BindingFlags.NonPublic).Invoke(support,
                new object[] { new LinkLabelLinkClickedEventArgs(support.Links[0]) });
        Pump(500);
        Shot(form, "whatsnew-support");
        Kill(form);
    }

    /// <summary>«Настройки» — общие для всей программы, с 1.18.0 отдельным окном.</summary>
    static void Settings()
    {
        Form f = New("ExcelMerger.SettingsForm");
        Place(f, 0, 0);
        Shot(f, "settings");
        Kill(f);
    }

    static void Preview()
    {
        Type refType = _app.GetType("ExcelMerger.PdfPageRef");
        object page = Activator.CreateInstance(refType);
        refType.GetField("SourcePath").SetValue(page, Doc);
        refType.GetField("PageIndex").SetValue(page, 2);

        Type t = _app.GetType("ExcelMerger.PagePreviewForm");
        ConstructorInfo ctor = t.GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance,
            null, new[] { refType, typeof(string), typeof(Size), typeof(Action<int>) }, null);
        var f = (Form)ctor.Invoke(new object[] { page, string.Format(Text("preview.title"), 3), new Size(900, 1000), null });
        Place(f, 820, 940);
        Pump(2500); // страница рисуется в фоне
        Shot(f, "preview");
        Kill(f);
    }

    static void Metadata()
    {
        Type meta = _app.GetType("ExcelMerger.PdfMetadata");
        object current = Activator.CreateInstance(meta);
        // Заполнители — на языке набора: в английском руководстве русские «Иванов И. И.»
        // и «Демонстрация возможностей» читаются как чужой скриншот, а не как пример.
        Set(meta, current, "Title", _en ? "Sample document" : "Образец документа");
        Set(meta, current, "Author", _en ? "John Smith" : "Иванов И. И.");
        Set(meta, current, "Subject", _en ? "Feature demonstration" : "Демонстрация возможностей");
        Set(meta, current, "Keywords", _en ? "sample, demonstration, PDF" : "образец, демонстрация, PDF");

        // С 1.17.9 окно называет файл, который правит, — второй аргумент конструктора.
        Type t = _app.GetType("ExcelMerger.MetadataForm");
        ConstructorInfo ctor = t.GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance,
            null, new[] { meta, typeof(string) }, null);
        // Имя — у того же образца, что показан в остальных окнах, чтобы снимки не расходились.
        var f = (Form)ctor.Invoke(new object[] { current, Path.GetFileName(Doc) });
        Place(f, 0, 0);
        Shot(f, "metadata");
        Kill(f);
    }

    static void Set(Type t, object o, string name, string value)
    {
        FieldInfo fi = t.GetField(name);
        if (fi != null) { fi.SetValue(o, value); return; }
        t.GetProperty(name).SetValue(o, value, null);
    }

    static void About()
    {
        Form f = New("ExcelMerger.AboutForm");
        Place(f, 0, 0);
        Shot(f, "about");
        Kill(f);
    }

    static void Stats()
    {
        Form f = New("ExcelMerger.StatsForm");
        Place(f, 0, 0);
        Shot(f, "stats");
        Kill(f);
    }

    static void HelpMerge() { Message("help-merge", Text("hub.pdf.name"), Text("menu.howTo"), Text("pdf.help.body")); }
    static void HelpExcel() { Message("help-excel", Text("hub.excel.name"), Text("menu.howTo"), Text("excel.help.body")); }
    static void HelpSplit() { Message("help-split", Text("hub.split.name"), Text("menu.howTo"), Text("split.help.body")); }
    static void HelpOcr() { Message("help-ocr", Text("hub.ocr.name"), Text("menu.howTo"), Text("ocr.help.body")); }

    /// <summary>Шпаргалка клавиш — ровно тот текст, что собирает само приложение.</summary>
    static void Shortcuts()
    {
        Type baseType = _app.GetType("ExcelMerger.PdfToolFormBase");
        MethodInfo build = baseType.GetMethod("BuildShortcuts", BindingFlags.NonPublic | BindingFlags.Static);
        Func<string, string> t = Text;
        var body = (string)build.Invoke(null, new object[] { true, true, t });
        Message("shortcuts", Text("menu.shortcuts"), Text("shortcuts.title"), body);
    }

    /// <summary>Диалог сообщения приложения с заданным содержимым.</summary>
    static void Message(string name, string title, string header, string body)
    {
        Type t = _app.GetType("ExcelMerger.MessageForm");
        Type kind = t.GetNestedType("Kind");
        ConstructorInfo ctor = t.GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null,
            new[] { kind, typeof(string), typeof(string), typeof(string), typeof(bool), typeof(string), typeof(string), typeof(string) }, null);
        var f = (Form)ctor.Invoke(new object[] { Enum.ToObject(kind, 0), title, header, body, false, null, null, null });
        Place(f, 0, 0);
        Shot(f, name);
        Kill(f);
    }

    static void GoToPage()
    {
        Type t = _app.GetType("ExcelMerger.NumberPromptDialog");
        ConstructorInfo ctor = t.GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null,
            new[] { typeof(string), typeof(string), typeof(string), typeof(int), typeof(int), typeof(int) }, null);
        var f = (Form)ctor.Invoke(new object[] { Text("grid.goto.title"), string.Format(Text("grid.goto.prompt"), 8), Text("grid.goto.ok"), 1, 8, 3 });
        Place(f, 0, 0);
        Shot(f, "goto");
        Kill(f);
    }

    // ---------- вспомогательное ----------

    static string Text(string key)
    {
        return (string)_loc.GetMethod("T").Invoke(null, new object[] { key });
    }

    static Form New(string typeName)
    {
        var f = (Form)Activator.CreateInstance(_app.GetType(typeName), true);
        return f;
    }

    static Form Tool(string typeName)
    {
        Type t = _app.GetType(typeName);
        ConstructorInfo ctor = t.GetConstructor(new[] { typeof(Action) });
        return (Form)(ctor != null ? ctor.Invoke(new object[] { null }) : Activator.CreateInstance(t, true));
    }

    /// <summary>Окно инструмента с загруженными документами-образцами.</summary>
    static Form Loaded(string typeName, string[] files) { return Loaded(typeName, files, 1000, 800); }

    /// <summary>То же, но с заданным размером окна: 0×0 — оставить размер, который окно берёт само.</summary>
    static Form Loaded(string typeName, string[] files, int w, int h)
    {
        Form f = Tool(typeName);
        Place(f, w, h);
        Type acceptor = _app.GetType("ExcelMerger.IFileAcceptor");
        acceptor.GetMethod("AcceptFiles").Invoke(f, new object[] { files });
        Pump(3500); // разбор PDF и отрисовка миниатюр идут в фоне
        return f;
    }

    static void Place(Form f, int w, int h)
    {
        f.StartPosition = FormStartPosition.Manual;
        f.Location = new Point(30, 30);
        if (w > 0 && h > 0)
            f.Size = new Size(w, h);
        else if (f is Form)
        {
            // A manual screenshot must exercise the same initial layout as a user opening the
            // tool, not the stale location/size inherited from another scenario.
            f.StartPosition = FormStartPosition.CenterScreen;
        }
        f.Show();
        f.Activate();
        SetForegroundWindow(f.Handle);
        Pump(700);
    }

    static void Kill(Form f)
    {
        try { f.Close(); f.Dispose(); } catch { }
        Pump(300);
    }

    static void Pump(int ms)
    {
        int step = 40;
        for (int i = 0; i < ms; i += step)
        {
            Application.DoEvents();
            System.Threading.Thread.Sleep(step);
        }
    }

    static object Field(Form f, string name)
    {
        FieldInfo fi = f.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);
        if (fi == null) throw new Exception("нет поля " + name + " в " + f.GetType().Name);
        return fi.GetValue(f);
    }

    static void SetText(Form f, string field, string value)
    {
        ((Control)Field(f, field)).Text = value;
    }

    static Control FindByType(Control root, string typeName)
    {
        foreach (Control c in root.Controls)
        {
            if (c.GetType().Name == typeName) return c;
            Control nested = FindByType(c, typeName);
            if (nested != null) return nested;
        }
        return null;
    }

    static void Select(Control grid, params int[] indices)
    {
        MethodInfo one = grid.GetType().GetMethod("SelectIndex");
        MethodInfo range = grid.GetType().GetMethod("SelectRange");
        if (indices.Length > 1 && range != null)
            range.Invoke(grid, new object[] { indices[0], indices.Length });
        else if (one != null)
            one.Invoke(grid, new object[] { indices[0] });
        Pump(300);
    }

    static void OpenMainMenu(Form f)
    {
        var strip = (MenuStrip)FindByType(f, "MenuStrip");
        var root = (ToolStripMenuItem)strip.Items[0];
        root.ShowDropDown();
        Pump(600);
        _popup = root.DropDown.Bounds;
    }

    static void ShowMenu(ContextMenuStrip menu, Control at, Point where)
    {
        if (menu == null) throw new Exception("у контрола нет контекстного меню");
        menu.Show(at, where);
        Pump(600);
        _popup = menu.Bounds;
    }

    /// <summary>Раскрыть список и запомнить его границы: у ComboBox список — окно системы, своих Bounds у него нет.</summary>
    static void DropDown(ComboBox combo)
    {
        Form owner = combo.FindForm();
        owner.TopMost = true;
        owner.Activate();
        SetForegroundWindow(owner.Handle);
        Pump(400);
        combo.Focus();
        combo.DroppedDown = true;
        Pump(500);
        Rectangle box = combo.RectangleToScreen(combo.ClientRectangle);
        int shown = Math.Min(combo.Items.Count, combo.MaxDropDownItems);
        _popup = new Rectangle(box.Left, box.Bottom, box.Width, shown * combo.ItemHeight + 6);
    }

    // ---------- захват ----------

    static void Shot(Form f, string name)
    {
        string png = Path.Combine(_out, name + ".png");
        using (var bmp = new Bitmap(f.Width, f.Height))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                IntPtr hdc = g.GetHdc();
                PrintWindow(f.Handle, hdc, 0);
                g.ReleaseHdc(hdc);
            }
            bmp.Save(png, ImageFormat.Png);
        }
        Say(name + " -> " + f.Width + "x" + f.Height);
    }

    /// <summary>Снимок С ЭКРАНА: захватывает и выпадающие меню (они — отдельные окна).</summary>
    static void ScreenShot(Form f, string name) { ScreenShot(f, name, true); }

    /// <summary>
    /// Снимок окна С РАСКРЫТЫМ МЕНЮ. Меню — отдельное окно Windows, в PrintWindow формы оно не
    /// попадает, поэтому его приходится брать с экрана. Но брать С ЭКРАНА ВСЁ нельзя: за краем
    /// окна тогда оказывается рабочий стол снимающего — чужие окна, папки, имена файлов.
    /// Поэтому окно печатаем через PrintWindow, а с экрана вырезаем ТОЛЬКО прямоугольник меню и
    /// накладываем его на своё место. Подложка при этом перестаёт быть единственной защитой:
    /// даже если её перекрыли, в кадр не попадёт ни одного постороннего пикселя.
    /// </summary>
    static void ScreenShot(Form f, string name, bool activate)
    {
        if (activate)
        {
            f.TopMost = true;
            f.Activate();
            SetForegroundWindow(f.Handle);
            Pump(500);
        }
        Rectangle form = f.Bounds;
        Rectangle shot = _popup.IsEmpty ? form : Rectangle.Union(form, _popup);
        string png = Path.Combine(_out, name + ".png");
        using (var bmp = new Bitmap(shot.Width, shot.Height))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                // Фон кадра — тот же нейтральный цвет, что у подложки: за окном не должно быть
                // ни рабочего стола, ни чёрных полей.
                g.Clear(Color.FromArgb(232, 234, 238));
                using (var window = new Bitmap(form.Width, form.Height))
                {
                    using (Graphics wg = Graphics.FromImage(window))
                    {
                        IntPtr hdc = wg.GetHdc();
                        PrintWindow(f.Handle, hdc, 0);
                        wg.ReleaseHdc(hdc);
                    }
                    g.DrawImage(window, form.X - shot.X, form.Y - shot.Y);
                }
                if (!_popup.IsEmpty)
                {
                    // Меню — единственное, что берётся с экрана, и ровно по своим границам.
                    Rectangle menu = _popup;
                    menu.Intersect(Screen.FromControl(f).Bounds);
                    using (var popup = new Bitmap(menu.Width, menu.Height))
                    {
                        using (Graphics pg = Graphics.FromImage(popup))
                            pg.CopyFromScreen(menu.Location, Point.Empty, menu.Size);
                        g.DrawImage(popup, menu.X - shot.X, menu.Y - shot.Y);
                    }
                }
            }
            bmp.Save(png, ImageFormat.Png);
        }
        f.TopMost = false;
        _popup = Rectangle.Empty;
        Say(name + " -> window+menu " + shot.Width + "x" + shot.Height);
    }
}
