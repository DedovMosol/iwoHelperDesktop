using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Корпус разметки (этап 0 плана распознавания): сотни настоящих PDF, нарисованных из
    /// моделей документов, проходят настоящий разбор, и модель разметки сверяется с
    /// исходной. Известные дефекты записаны в <see cref="CorpusKnown"/> со ссылкой на
    /// audit.md и точным числом документов: и регрессия, и исправление меняют это число.
    /// </summary>
    internal static partial class UnitTests
    {
        /// <summary>Вариантов на семейство. Детерминированно: тот же набор на любой машине.</summary>
        private const int CorpusPerFamily = 120;

        private static List<CCase> _corpus;

        private static List<CCase> Corpus()
        {
            if (_corpus == null)
                _corpus = CorpusCases.All(CorpusPerFamily);
            return _corpus;
        }

        /// <summary>Нарисовать документ корпуса во временную папку.</summary>
        private static CRendered RenderCase(CCase c, TempDir dir, out string path)
        {
            path = dir.File("corpus_" + c.Id.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".pdf");
            return CorpusRenderer.Write(c.Doc, path, dir.Path);
        }

        /// <summary>
        /// Сверка корпуса с реестром <see cref="CorpusKnown"/>: сколько документов каждого
        /// семейства показали каждый вид расхождения. Строго в обе стороны: вид без записи,
        /// запись без проявлений и любое изменившееся число проваливают тест.
        /// </summary>
        private sealed class CorpusLedger
        {
            private const int Examples = 3;
            private readonly string _output;
            // «семейство␟вид» → число документов, случаев и первые примеры.
            private readonly Dictionary<string, int> _docs = new Dictionary<string, int>(StringComparer.Ordinal);
            private readonly Dictionary<string, int> _cases = new Dictionary<string, int>(StringComparer.Ordinal);
            private readonly Dictionary<string, List<string>> _examples = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            public readonly List<string> Report = new List<string>();
            public readonly List<string> Registry = new List<string>();

            public CorpusLedger(string output) { _output = output; }

            private static string Key(string family, string aspect) { return family + "\u001F" + aspect; }

            /// <summary>Находки одного документа (CorpusCheck уже оставляет по одной на вид).</summary>
            public void Reconcile(CCase c, List<CFinding> findings)
            {
                foreach (string aspect in findings.Select(f => f.Aspect).Distinct(StringComparer.Ordinal))
                {
                    string key = Key(c.Family, aspect);
                    int n; _docs.TryGetValue(key, out n); _docs[key] = n + 1;
                    int k; _cases.TryGetValue(key, out k); _cases[key] = k + findings.Where(f => f.Aspect == aspect).Sum(f => f.Count);
                    List<string> list;
                    if (!_examples.TryGetValue(key, out list)) _examples[key] = list = new List<string>();
                    if (list.Count < Examples)
                        list.Add(c.Name + " — " + findings.First(f => f.Aspect == aspect));
                }
            }

            public void Finish()
            {
                var known = new HashSet<string>(StringComparer.Ordinal);
                foreach (CKnown k in CorpusKnown.All.Where(k => k.Output == _output))
                {
                    string key = Key(k.Family, k.Aspect);
                    known.Add(key);
                    int n, cases;
                    _docs.TryGetValue(key, out n);
                    _cases.TryGetValue(key, out cases);
                    if (n == 0)
                        Report.Add(k.Family + " — " + k.Id + " «" + k.Aspect + "» больше не воспроизводится (было ×" + k.Docs +
                            "): исправлен? Снимите запись в Corpus.Known.cs");
                    else if (n != k.Docs || cases != k.Cases)
                        Report.Add(k.Family + " — " + k.Id + " «" + k.Aspect + "»: документов ×" + n + ", случаев ×" + cases +
                            " вместо ×" + k.Docs + " и ×" + k.Cases + (n > k.Docs || cases > k.Cases
                                ? " (регрессия), напр. " + _examples[key][0] : " (стало лучше: обновите числа)"));
                }
                foreach (KeyValuePair<string, int> p in _docs.OrderBy(p => p.Key, StringComparer.Ordinal))
                {
                    string[] parts = p.Key.Split('\u001F');
                    // Подсказка ID для новой строки: та же пара «семейство, вид», иначе тот же вид.
                    string id = CorpusKnown.All.Where(k => k.Output == _output && k.Family == parts[0] && k.Aspect == parts[1])
                        .Concat(CorpusKnown.All.Where(k => k.Output == _output && k.Aspect == parts[1]))
                        .Select(k => k.Id).FirstOrDefault() ?? "??";
                    Registry.Add("K(" + char.ToUpperInvariant(_output[0]) + _output.Substring(1) + ", \"" + parts[0] + "\", \"" + parts[1] +
                        "\", \"" + id + "\", " + p.Value + ", " + _cases[p.Key] + "),");
                    if (known.Contains(p.Key))
                        continue;
                    Report.Add(parts[0] + " — неизвестное «" + parts[1] + "» ×" + p.Value + ":");
                    foreach (string e in _examples[p.Key])
                        Report.Add("    " + e);
                }
            }

            /// <summary>Сводка известных дефектов: «ID вид ×документов» по всем семействам.</summary>
            public string Stats()
            {
                var stats = new SortedDictionary<string, int>(StringComparer.Ordinal);
                foreach (CKnown k in CorpusKnown.All.Where(k => k.Output == _output))
                {
                    int n; _docs.TryGetValue(Key(k.Family, k.Aspect), out n);
                    string s = k.Id + " " + k.Aspect;
                    int t; stats.TryGetValue(s, out t); stats[s] = t + n;
                }
                return string.Join(", ", stats.Select(p => p.Key + " ×" + p.Value).ToArray());
            }
        }

        /// <summary>
        /// IWO_CORPUS_DUMP=начала имён через «|» — разобранная модель этих документов в вывод
        /// (для разбора расхождений: абзацы с выключкой и списком, таблицы, картинки).
        /// </summary>
        private static void CorpusDump(CCase c, List<PdfPageText> pages)
        {
            string filter = Environment.GetEnvironmentVariable("IWO_CORPUS_DUMP");
            if (string.IsNullOrEmpty(filter) || !filter.Split('|').Any(f => f.Length > 0 && c.Name.StartsWith(f)))
                return;
            Console.WriteLine("      === " + c.Name);
            foreach (PdfPageText page in pages)
            {
                Console.WriteLine("      --- стр. " + (page.PageIndex + 1) + ": таблиц " + page.Tables.Count + ", картинок " + page.Images.Count);
                foreach (PageBlocks.PageItem item in PageBlocks.CoalesceRowBands(PageBlocks.OrderedItems(page)))
                {
                    var blocks = new List<PageBlocks.Block>();
                    if (item.IsBand) foreach (List<PageBlocks.Block> col in item.Columns) blocks.AddRange(col);
                    else blocks.Add(item.Single);
                    foreach (PageBlocks.Block b in blocks)
                    {
                        if (b.Paragraph != null)
                            Console.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture,
                                "      P{0} {1} {2}{3} top {4:0.#} left {5:0.#} right {6:0.#} lines {7}: {8}", item.IsBand ? "[band]" : "",
                                b.Paragraph.Alignment, b.Paragraph.ListKind, b.Paragraph.ListLevel, b.Paragraph.TopPt,
                                b.Paragraph.LeftPt, b.Paragraph.RightPt, b.Paragraph.LineCount, b.Paragraph.Text));
                        else if (b.Table != null)
                            Console.WriteLine("      T " + b.Table.Rows.Count + "×" + b.Table.ColumnCount + (b.Table.Borderless ? " без линовки" : "") +
                                ": " + string.Join(" | ", b.Table.Rows.SelectMany(r => r.Cells).Take(8)
                                    .Select(x => string.Join(" ", x.Paragraphs.Select(p => p.Text).ToArray())).ToArray()));
                        else if (b.Image != null)
                            Console.WriteLine("      I " + b.Image.WidthPt + "×" + b.Image.HeightPt);
                    }
                }
            }
        }

        private static string CorpusReport(string what, CorpusLedger ledger, int docs, long ms)
        {
            ledger.Finish();
            List<string> report = ledger.Report;
            bool full = Environment.GetEnvironmentVariable("IWO_CORPUS_FULL") == "1";
            var sb = new StringBuilder();
            sb.Append(what).Append(": ").Append(report.Count).Append(" строк расхождений на ").Append(docs).Append(" документах");
            // IWO_CORPUS_FULL=1 — весь отчёт и готовые строки реестра, иначе первые 40 строк.
            int limit = full ? int.MaxValue : 40;
            for (int i = 0; i < report.Count && i < limit; i++)
                sb.Append("\n    ").Append(report[i]);
            if (report.Count > limit)
                sb.Append("\n    … и ещё ").Append(report.Count - limit);
            Console.WriteLine("      " + what + ": " + docs + " документов за " + ms + " мс; известные дефекты: " + ledger.Stats());
            if (full)
                foreach (string line in ledger.Registry)
                    Console.WriteLine("      REG " + line);
            string knownPath = Environment.GetEnvironmentVariable("IWO_CORPUS_WRITE_KNOWN");
            if (!string.IsNullOrEmpty(knownPath) && ledger.Registry.Count > 0)
                File.AppendAllLines(knownPath, ledger.Registry, Encoding.UTF8);
            return sb.ToString();
        }

        /// <summary>
        /// Live two-column PDF (variant 3: Letter/Georgia-class). Left column words
        /// must all precede the right column — the dump of #3 still zigzagged (RC-14).
        /// </summary>
        [Test("Корпус: две колонки #3 — левая колонка целиком раньше правой (RC-14)")]
        private static void CorpusTwoColumnReadingOrder()
        {
            EmbeddedAssemblies.Ensure();
            CCase c = CorpusCases.All(4).Find(x => x.Family == "две колонки" && x.Variant.Index == 3);
            AssertTrue(c != null, "variant 3 exists");
            using (var dir = new TempDir("iwo_twocol_"))
            {
                string path;
                RenderCase(c, dir, out path);
                List<PdfPageText> pages = PdfTextExtract.Extract(path);
                AssertTrue(pages.Count >= 1 && pages[0].Paragraphs.Count >= 3, "title + columns");
                var flow = new List<OcrParagraph>();
                foreach (PageBlocks.PageItem item in PageBlocks.CoalesceRowBands(PageBlocks.OrderedItems(pages[0])))
                {
                    if (item.IsBand)
                        foreach (List<PageBlocks.Block> col in item.Columns)
                            foreach (PageBlocks.Block b in col)
                                if (b.Paragraph != null) flow.Add(b.Paragraph);
                    else if (item.Single != null && item.Single.Paragraph != null)
                        flow.Add(item.Single.Paragraph);
                }
                var lefts = new List<string>();
                int firstRight = -1, lastLeft = -1;
                for (int i = 0; i < flow.Count; i++)
                {
                    lefts.Add(i + ":" + flow[i].LeftPt.ToString("0") + "x" + flow[i].LineCount);
                    if (flow[i].LeftPt < 120)
                        lastLeft = i;
                    else if (flow[i].LeftPt > 250 && firstRight < 0)
                        firstRight = i;
                }
                var raw = new List<string>();
                foreach (OcrParagraph p in pages[0].Paragraphs)
                    raw.Add(string.Format(System.Globalization.CultureInfo.InvariantCulture,
                        "L{0:G} R{1:G} T{2:G} B{3:G}", p.LeftPt, p.RightPt, p.TopPt, p.BottomPt));
                string dump = "raw " + string.Join(",", raw.ToArray()) + " flow " + string.Join(" ", lefts.ToArray());
                var clone = new PdfPageText();
                clone.Paragraphs.AddRange(pages[0].Paragraphs);
                var items = PageBlocks.OrderedItems(pages[0]);
                var cloned = PageBlocks.OrderedItems(clone);
                dump += " items=" + items.Count + " clone=" + cloned.Count
                    + " tbl=" + (pages[0].Tables == null ? 0 : pages[0].Tables.Count)
                    + " img=" + (pages[0].Images == null ? 0 : pages[0].Images.Count);
                for (int i = 0; i < items.Count; i++)
                    dump += items[i].IsBand ? " B" + items[i].Columns.Count : " S";
                AssertTrue(lastLeft >= 0 && firstRight >= 0, "both columns present: " + dump);
                AssertTrue(firstRight > lastLeft, "Word/PPTX flow right after whole left: " + dump);
            }
        }

        [Test("Корпус: все варианты «две колонки» — левая колонка раньше правой (RC-14)")]
        private static void CorpusTwoColumnAllVariantsReadingOrder()
        {
            EmbeddedAssemblies.Ensure();
            var fails = new List<string>();
            using (var dir = new TempDir("iwo_twocol_all_"))
                foreach (CCase c in CorpusCases.All(CorpusPerFamily))
                {
                    if (c.Family != "две колонки")
                        continue;
                    string path;
                    RenderCase(c, dir, out path);
                    List<PdfPageText> pages = PdfTextExtract.Extract(path);
                    if (pages.Count == 0) continue;
                    var flow = new List<OcrParagraph>();
                    foreach (PageBlocks.PageItem item in PageBlocks.CoalesceRowBands(PageBlocks.OrderedItems(pages[0])))
                    {
                        if (item.IsBand)
                            foreach (List<PageBlocks.Block> col in item.Columns)
                                foreach (PageBlocks.Block b in col)
                                    if (b.Paragraph != null) flow.Add(b.Paragraph);
                        else if (item.Single != null && item.Single.Paragraph != null)
                            flow.Add(item.Single.Paragraph);
                    }
                    double mid = pages[0].WidthPt > 0 ? pages[0].WidthPt / 2.0 : 300;
                    int firstRight = -1, lastLeft = -1;
                    for (int i = 0; i < flow.Count; i++)
                    {
                        if (flow[i].Alignment == OcrAlignment.Center)
                            continue;
                        if (flow[i].LeftPt < mid)
                            lastLeft = i;
                        else if (firstRight < 0)
                            firstRight = i;
                    }
                    if (lastLeft >= 0 && firstRight >= 0 && firstRight <= lastLeft)
                        fails.Add("#" + c.Variant.Index + " " + c.Variant + " lastL=" + lastLeft + " firstR=" + firstRight);
                    if (fails.Count >= 8)
                        break;
                }
            AssertTrue(fails.Count == 0, "zigzag: " + string.Join("; ", fails.ToArray()));
        }

        [Test("Корпус: римский список #4 — маркеры на месте (RC-3)")]
        private static void CorpusLetterListMarkers()
        {
            EmbeddedAssemblies.Ensure();
            CCase c = CorpusCases.All(10).Find(x => x.Family.StartsWith("списки с буквами", StringComparison.Ordinal) && x.Variant.Index == 4);
            AssertTrue(c != null, "variant 4");
            using (var dir = new TempDir("iwo_letterlist_"))
            {
                string path;
                RenderCase(c, dir, out path);
                List<PdfPageText> pages = PdfTextExtract.Extract(path);
                int numbered = 0;
                var kinds = new List<string>();
                foreach (PdfPageText p in pages)
                    foreach (OcrParagraph para in p.Paragraphs)
                    {
                        kinds.Add(string.Format(System.Globalization.CultureInfo.InvariantCulture,
                            "{0} L{1:0} T{2:0} {3}",
                            para.ListKind == ListKind.None ? "t" : "n" + para.ListNumber,
                            para.LeftPt, para.TopPt,
                            para.Text.Length > 36 ? para.Text.Substring(0, 36) : para.Text));
                        if (para.ListKind == ListKind.Numbered) numbered++;
                    }
                AssertTrue(numbered >= 6, "letters+roman+compound: " + string.Join(" | ", kinds.ToArray()));
            }
        }

        /// <summary>
        /// Три колонки: после заголовка поток — колонка 0 целиком, затем 1, затем 2.
        /// Построчный L/M/R — RC-14 (XY не держал вертикальный канал на всю высоту).
        /// </summary>
        [Test("Корпус: три колонки #0 и Courier — колонки целиком, не зигзаг (RC-14)")]
        private static void CorpusThreeColumnReadingOrder()
        {
            EmbeddedAssemblies.Ensure();
            var fails = new List<string>();
            using (var dir = new TempDir("iwo_threecol_"))
                foreach (int idx in new[] { 0, 29 })
                {
                    CCase c = CorpusCases.All(30).Find(x => x.Family == "три колонки" && x.Variant.Index == idx);
                    AssertTrue(c != null, "variant " + idx);
                    string path;
                    RenderCase(c, dir, out path);
                    List<PdfPageText> pages = PdfTextExtract.Extract(path);
                    AssertTrue(pages.Count >= 1, "page");
                    var flow = new List<OcrParagraph>();
                    foreach (PageBlocks.PageItem item in PageBlocks.CoalesceRowBands(PageBlocks.OrderedItems(pages[0])))
                    {
                        if (item.IsBand)
                            foreach (List<PageBlocks.Block> col in item.Columns)
                                foreach (PageBlocks.Block b in col)
                                    if (b.Paragraph != null) flow.Add(b.Paragraph);
                        else if (item.Single != null && item.Single.Paragraph != null)
                            flow.Add(item.Single.Paragraph);
                    }
                    var body = new List<OcrParagraph>();
                    foreach (OcrParagraph p in flow)
                        if (p.Alignment != OcrAlignment.Center)
                            body.Add(p);
                    var dump = new List<string>();
                    dump.Add("n=" + body.Count + " font=" + c.Variant.Font);
                    foreach (OcrParagraph p in body)
                        dump.Add(string.Format(System.Globalization.CultureInfo.InvariantCulture,
                            "L{0:0.#} R{1:0.#} T{2:0.#} lines {3}", p.LeftPt, p.RightPt, p.TopPt, p.LineCount));
                    if (body.Count < 3)
                    {
                        fails.Add("#" + idx + " мало абзацев " + string.Join(" ", dump.ToArray()));
                        continue;
                    }
                    double minL = body[0].LeftPt, maxL = body[0].LeftPt;
                    foreach (OcrParagraph p in body)
                    {
                        if (p.LeftPt < minL) minL = p.LeftPt;
                        if (p.LeftPt > maxL) maxL = p.LeftPt;
                    }
                    double span = maxL - minL;
                    if (span < 40)
                    {
                        fails.Add("#" + idx + " нет трёх левых краёв: " + string.Join(" ", dump.ToArray()));
                        continue;
                    }
                    double t1 = minL + span / 3.0, t2 = minL + 2.0 * span / 3.0;
                    int last0 = -1, first1 = -1, last1 = -1, first2 = -1;
                    for (int i = 0; i < body.Count; i++)
                    {
                        int col = body[i].LeftPt < t1 ? 0 : (body[i].LeftPt < t2 ? 1 : 2);
                        if (col == 0) last0 = i;
                        else if (col == 1)
                        {
                            if (first1 < 0) first1 = i;
                            last1 = i;
                        }
                        else if (first2 < 0) first2 = i;
                    }
                    if (first1 < 0 || first2 < 0 || last0 < 0)
                        fails.Add("#" + idx + " нет трёх колонок: " + string.Join(" ", dump.ToArray()));
                    else if (!(last0 < first1 && last1 < first2))
                        fails.Add("#" + idx + " зигзаг last0=" + last0 + " first1=" + first1 +
                            " last1=" + last1 + " first2=" + first2 + " " + string.Join(" ", dump.ToArray()));
                }
            AssertTrue(fails.Count == 0, string.Join(" ;; ", fails.ToArray()));
        }

        [Test("Корпус: генератор рисует ровно те слова, что описаны в модели")]
        private static void CorpusGeneratorDrawsTheModel()
        {
            EmbeddedAssemblies.Ensure();
            var report = new List<string>();
            int pages = 0, words = 0;
            using (var dir = new TempDir("iwo_corpus_gen_"))
                foreach (CCase c in Corpus())
                {
                    string path;
                    CRendered drawn = RenderCase(c, dir, out path);
                    pages += drawn.Pages;
                    var want = new Dictionary<string, int>(StringComparer.Ordinal);
                    Action<string> add = delegate(string text)
                    {
                        foreach (string w in text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            int n; want.TryGetValue(w, out n); want[w] = n + 1;
                        }
                    };
                    foreach (CDrawn d in drawn.Drawn)
                    {
                        var p = d.Block as CParagraph;
                        if (p != null)
                        {
                            foreach (CRun r in p.Runs) add(r.Text);
                            if (p.Marker != null) add(p.Marker);
                        }
                        var t = d.Block as CTable;
                        if (t != null)
                            foreach (CCell[] row in t.Rows)
                                foreach (CCell cell in row)
                                    if (cell != null) foreach (CRun r in cell.Runs) add(r.Text);
                        if (d.Text != null) add(d.Text);
                    }
                    string rotatedLetters;
                    var got = CorpusRawWords(path, out rotatedLetters);
                    words += got.Values.Sum();
                    // Повёрнутую строку PdfPig собирает в слова по-своему (часто по букве),
                    // поэтому её сверяем по буквам, а не по словам.
                    string wantRotated = string.Concat(drawn.Drawn.Where(d => d.Block is CRotated)
                        .Select(d => ((CRotated)d.Block).Text.Replace(" ", "")).ToArray());
                    if (Sorted(wantRotated) != Sorted(rotatedLetters))
                        report.Add(c.Name + ": повёрнутые буквы «" + rotatedLetters + "» вместо «" + wantRotated + "»");
                    var problems = new List<string>();
                    foreach (KeyValuePair<string, int> p in want)
                    {
                        int n; got.TryGetValue(p.Key, out n);
                        if (n != p.Value) problems.Add("«" + p.Key + "» ×" + n + " вместо ×" + p.Value);
                    }
                    foreach (KeyValuePair<string, int> p in got)
                        if (!want.ContainsKey(p.Key)) problems.Add("лишнее «" + p.Key + "»");
                    if (problems.Count > 0)
                        report.Add(c.Name + ": " + string.Join(", ", problems.Take(5).ToArray()));
                }
            Console.WriteLine("      корпус: " + Corpus().Count + " документов, " + pages + " страниц, " + words + " слов");
            AssertTrue(report.Count == 0, "генератор корпуса: " + string.Join("\n    ", report.Take(30).ToArray()));
        }

        /// <summary>
        /// Слова файла, как их видит PdfPig без фильтров разбора: горизонтальные — словами,
        /// буквы прочих направлений — одной строкой.
        /// </summary>
        private static Dictionary<string, int> CorpusRawWords(string path, out string rotated)
        {
            var result = new Dictionary<string, int>(StringComparer.Ordinal);
            var other = new StringBuilder();
            using (UglyToad.PdfPig.PdfDocument doc = PdfPageProbe.OpenPig(path))
                foreach (UglyToad.PdfPig.Content.Page page in doc.GetPages())
                {
                    foreach (UglyToad.PdfPig.Content.Letter l in page.Letters)
                        if (l.TextOrientation != UglyToad.PdfPig.Content.TextOrientation.Horizontal && l.Value != " ")
                            other.Append(l.Value);
                    foreach (UglyToad.PdfPig.Content.Word w in page.GetWords())
                    {
                        if (w.Letters.Count > 0 && w.Letters[0].TextOrientation != UglyToad.PdfPig.Content.TextOrientation.Horizontal)
                            continue;
                        int n; result.TryGetValue(w.Text, out n); result[w.Text] = n + 1;
                    }
                }
            rotated = other.ToString();
            return result;
        }

        private static string Sorted(string s)
        {
            char[] chars = s.ToCharArray();
            Array.Sort(chars);
            return new string(chars);
        }

        [Test("Корпус DOCX: модель разметки совпадает с документом (известные дефекты — строго)")]
        private static void CorpusDocumentModel()
        {
            var ledger = new CorpusLedger(CorpusKnown.Docx);
            Stopwatch watch = Stopwatch.StartNew();
            using (var dir = new TempDir("iwo_corpus_doc_"))
                foreach (CCase c in Corpus())
                {
                    string path;
                    CRendered drawn = RenderCase(c, dir, out path);
                    List<PdfPageText> pages = PdfTextExtract.Extract(path);
                    CorpusDump(c, pages);
                    ledger.Reconcile(c, CorpusCheck.Document(drawn, pages));
                }
            string summary = CorpusReport("корпус DOCX", ledger, Corpus().Count, watch.ElapsedMilliseconds);
            AssertTrue(ledger.Report.Count == 0, summary);
        }

        [Test("Корпус PPTX: каждое слово на слайде один раз и с форматом (известные дефекты — строго)")]
        private static void CorpusSlides()
        {
            var ledger = new CorpusLedger(CorpusKnown.Pptx);
            Stopwatch watch = Stopwatch.StartNew();
            using (var dir = new TempDir("iwo_corpus_pptx_"))
                foreach (CCase c in Corpus())
                {
                    string path;
                    CRendered drawn = RenderCase(c, dir, out path);
                    List<PdfPageText> pages = PdfTextExtract.Extract(path, null, null, null, PageLayoutMode.Slide);
                    OoxmlPackage package = PptxWriter.BuildPackage(pages, PptxStamp);
                    ledger.Reconcile(c, CorpusCheck.Slides(drawn, package, PptxValidationProblems, dir.Path));
                }
            string summary = CorpusReport("корпус PPTX", ledger, Corpus().Count, watch.ElapsedMilliseconds);
            AssertTrue(ledger.Report.Count == 0, summary);
        }

        /// <summary>
        /// Документы корпуса для живого Word: по одному на семейство, номер варианта сдвигается
        /// с семейством — так в выборку попадают все шрифты, кегли, языки и форматы листа.
        /// </summary>
        private static List<CCase> CorpusWordSample()
        {
            return Corpus().Where(c => c.Variant.Index == (c.Id / CorpusPerFamily) % CorpusPerFamily).ToList();
        }

        [Test("Корпус DOCX вживую: Word записывает ровно модель разметки (живой Word)")]
        private static void CorpusWordWriter()
        {
            AuditBRequireWord();
            var ledger = new CorpusLedger(CorpusKnown.Word);
            List<CCase> sample = CorpusWordSample();
            Stopwatch watch = Stopwatch.StartNew();
            using (var dir = new TempDir("iwo_corpus_word_"))
                foreach (CCase c in sample)
                {
                    string path;
                    RenderCase(c, dir, out path);
                    List<PdfPageText> pages = PdfTextExtract.Extract(path);
                    string docx = Path.ChangeExtension(path, ".docx");
                    AuditBOfficeSta(180000, delegate { WordDocxWriter.Write(pages, docx); });
                    ledger.Reconcile(c, CorpusDocx.Compare(pages, docx));
                }
            string summary = CorpusReport("корпус Word", ledger, sample.Count, watch.ElapsedMilliseconds);
            AssertTrue(ledger.Report.Count == 0, summary);
        }

        [Test("Корпус: реестр известных дефектов цел и ссылается на записи audit.md")]
        private static void CorpusKnownDefectsAreAudited()
        {
            string root = RepoRoot();
            AssertTrue(root != null, "корень репозитория найден");
            string audit = File.ReadAllText(Path.Combine(root, "audit.md"));
            var families = new HashSet<string>(Corpus().Select(c => c.Family), StringComparer.Ordinal);
            var problems = new SortedSet<string>(StringComparer.Ordinal);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (CKnown k in CorpusKnown.All)
            {
                if (!audit.Contains("| " + k.Id + " |"))
                    problems.Add("нет в audit.md: " + k.Id);
                if (k.Output != CorpusKnown.Docx && k.Output != CorpusKnown.Pptx && k.Output != CorpusKnown.Word)
                    problems.Add("неизвестный вывод «" + k.Output + "»");
                if (!families.Contains(k.Family))
                    problems.Add("нет семейства «" + k.Family + "»");
                if (k.Docs < 1 || k.Docs > CorpusPerFamily || k.Cases < k.Docs)
                    problems.Add(k.Family + " «" + k.Aspect + "»: документов ×" + k.Docs + " вне 1…" + CorpusPerFamily +
                        " или случаев ×" + k.Cases + " меньше документов");
                if (!seen.Add(k.Output + "\u001F" + k.Family + "\u001F" + k.Aspect))
                    problems.Add("повтор: " + k.Output + " " + k.Family + " «" + k.Aspect + "»");
            }
            AssertTrue(problems.Count == 0, string.Join("; ", problems.ToArray()));
        }

        /// <summary>
        /// IWO_SHAPE=&lt;dir&gt; — layout counts only (pages/tables/paras/lists/images).
        /// No file names and no text: local probes of private samples stay private.
        /// </summary>
        [Test("Корпус: счётчики разметки каталога IWO_SHAPE (без текста)")]
        private static void CorpusShapeCounts()
        {
            string dir = Environment.GetEnvironmentVariable("IWO_SHAPE");
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
                return;
            EmbeddedAssemblies.Ensure();
            int i = 0;
            foreach (string pdf in Directory.GetFiles(dir, "*.pdf"))
            {
                List<PdfPageText> pages = PdfTextExtract.Extract(pdf);
                int tables = 0, paras = 0, lists = 0, images = 0;
                foreach (PdfPageText p in pages)
                {
                    if (p.Tables != null) tables += p.Tables.Count;
                    if (p.Paragraphs != null)
                    {
                        paras += p.Paragraphs.Count;
                        foreach (OcrParagraph para in p.Paragraphs)
                            if (para.ListKind != ListKind.None) lists++;
                    }
                    if (p.Images != null) images += p.Images.Count;
                }
                Console.WriteLine("      shape #" + i + ": pages " + pages.Count +
                    " tables " + tables + " paras " + paras + " lists " + lists + " images " + images);
                i++;
            }
        }

        /// <summary>
        /// IWO_SHOTS=&lt;папка&gt; — страница 1 выбранных семейств в PNG + текстовый разбор.
        /// Без переменной — пустой успех (CI не пишет файлы).
        /// </summary>
        [Test("Корпус: снимки страниц, если задан IWO_SHOTS")]
        private static void CorpusWriteShots()
        {
            string dir = Environment.GetEnvironmentVariable("IWO_SHOTS");
            if (string.IsNullOrEmpty(dir))
                return;
            Directory.CreateDirectory(dir);
            EmbeddedAssemblies.Ensure();
            string[] want =
            {
                "две колонки",
                "три колонки",
                "маркированный список в два уровня",
                "нумерованный список в два уровня",
                "таблица без линовки «метка — значение»",
                "таблица с линовкой",
                "маркеры Word: Symbol, Courier «o», Wingdings",
                "белый текст на тёмной плашке",
                "белый текст на светлой цветной плашке",
                "полужирные, курсивные и цветные слова",
                "подчёркнутые фразы",
                "зачёркнутые фразы",
                "списки с буквами, римскими и составными номерами",
                "разный кегль в абзаце",
                "выключка: влево, по центру, вправо, по ширине",
                "таблица с переносом строк в ячейках",
                "жирный курсив цвет черта и зачёркивание вместе",
                "список с цветом жирным чертой и зачёркиванием",
                "нумерованный список с цветом жирным и чертой",
                "скобочная нумерация",
                "нумерация в три уровня",
                "таблица и пункты списка"
            };
            int n = 0;
            using (var tmp = new TempDir("iwo_shots_"))
                foreach (string family in want)
                {
                    CCase c = Corpus().FirstOrDefault(x => x.Family == family && x.Variant.Index == 0);
                    AssertTrue(c != null, "семейство есть: " + family);
                    string pdf;
                    RenderCase(c, tmp, out pdf);
                    using (BudgetedBitmap bmp = PageRasterizer.RenderPage(pdf, 1, c.Doc.PageW, c.Doc.PageH))
                    {
                        AssertTrue(bmp != null && bmp.Bitmap != null, "растр страницы: " + family);
                        string slug = n.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".png";
                        bmp.Bitmap.Save(Path.Combine(dir, slug), ImageFormat.Png);
                    }
                    var dump = new StringBuilder();
                    dump.AppendLine(c.Name);
                    AppendShotDump(dump, "document", PdfTextExtract.Extract(pdf));
                    AppendShotDump(dump, "slide", PdfTextExtract.Extract(pdf, null, null, null, PageLayoutMode.Slide));
                    File.WriteAllText(Path.Combine(dir, n.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".txt"),
                        dump.ToString(), Encoding.UTF8);
                    n++;
                }
        }

        private static void AppendShotDump(StringBuilder dump, string label, List<PdfPageText> pages)
        {
            dump.AppendLine("== " + label + " ==");
            if (pages == null || pages.Count == 0)
            {
                dump.AppendLine("(empty)");
                return;
            }
            PdfPageText page = pages[0];
            foreach (OcrParagraph p in page.Paragraphs)
            {
                dump.AppendLine((p.ListKind == ListKind.None ? "P" : "L" + (int)p.ListKind)
                    + " " + p.Alignment + " lv" + p.ListLevel + " " + p.Text);
                if (p.Runs == null) continue;
                foreach (OcrRun r in p.Runs)
                {
                    dump.Append("  r");
                    if (r.Bold) dump.Append(" B");
                    if (r.Italic) dump.Append(" I");
                    if (r.Underline) dump.Append(" U");
                    if (r.Strike) dump.Append(" S");
                    if (r.ColorArgb != 0) dump.Append(" #").Append(r.ColorArgb.ToString("X6"));
                    dump.Append(" ").AppendLine(r.Text);
                }
            }
            dump.AppendLine("tables " + page.Tables.Count);
        }
    }
}
