using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// Снимки РЕАЛЬНЫХ окон приложения для инструкции. Грузим собранный exe как сборку,
// строим окна, наполняем их документами-образцами и печатаем окно в PNG.
// Окна с раскрытым меню снимаем с экрана: выпадающее меню — отдельное окно Windows,
// в PrintWindow формы оно не попадает.
static class Shots
{
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);

    static Assembly _app;
    static Type _loc, _lang;
    static string _out, _samples;
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
        _app = Assembly.LoadFrom(args[0]);
        _out = args[1];
        _samples = args[2];
        // Язык интерфейса на снимках: в английском руководстве русские окна выглядели бы
        // так же неуместно, как английские в русском. По умолчанию — русский, как было.
        string lang = args.Length > 3 ? args[3] : "ru";
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
        foreach (KeyValuePair<string, Action> s in Scenarios())
        {
            if (only.Count > 0 && !only.Contains(s.Key)) continue;
            try { s.Value(); }
            catch (Exception ex) { Say(s.Key + " -> ERR " + ex.Message); }
        }
        Console.WriteLine(string.Join(Environment.NewLine, _log.ToArray()));
        return 0;
    }

    static void Say(string line) { _log.Add(line); }

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

    /// <summary>Раздел PDF: четыре инструмента внутри раздела.</summary>
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
        var button = (Button)Field(f, "_btnImages");
        button.PerformClick();
        Pump(700);
        var menu = (ContextMenuStrip)Field(f, "_dpiMenu");
        _popup = menu.Bounds;
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
        Set(meta, current, "Title", "Образец документа");
        Set(meta, current, "Author", "Иванов И. И.");
        Set(meta, current, "Subject", "Демонстрация возможностей");
        Set(meta, current, "Keywords", "образец, демонстрация, PDF");

        // С 1.17.9 окно называет файл, который правит, — второй аргумент конструктора.
        Type t = _app.GetType("ExcelMerger.MetadataForm");
        ConstructorInfo ctor = t.GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance,
            null, new[] { meta, typeof(string) }, null);
        var f = (Form)ctor.Invoke(new object[] { current, "Образец документа.pdf" });
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
        if (w > 0 && h > 0) f.Size = new Size(w, h);
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
