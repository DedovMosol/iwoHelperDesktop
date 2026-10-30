using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        private const int CorpusPerFamily = 12;

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
                                "      P{0} {1} {2}{3} top {4:0.#} left {5:0.#} lines {6}: {7}", item.IsBand ? "[band]" : "",
                                b.Paragraph.Alignment, b.Paragraph.ListKind, b.Paragraph.ListLevel, b.Paragraph.TopPt,
                                b.Paragraph.LeftPt, b.Paragraph.LineCount, b.Paragraph.Text));
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
            return sb.ToString();
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
    }
}
