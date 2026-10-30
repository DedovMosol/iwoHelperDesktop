using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    /// <summary>Расхождение разбора с документом: вид (для реестра известных дефектов) и пример.</summary>
    internal sealed class CFinding
    {
        public readonly string Aspect;
        public readonly string Detail;
        /// <summary>Сколько раз вид встретился в документе (слов, абзацев, страниц — по месту проверки).</summary>
        public readonly int Count;
        public CFinding(string aspect, string detail) : this(aspect, detail, 1) { }
        public CFinding(string aspect, string detail, int count) { Aspect = aspect; Detail = detail; Count = count; }
        public override string ToString() { return Aspect + ": " + Detail; }
    }

    /// <summary>
    /// Сверка разобранного PDF с моделью, из которой он нарисован. Сравнивается МОДЕЛЬ
    /// разметки — абзацы, их границы и выключка, списки и их уровни, ячейки таблиц, формат
    /// каждого слова, — а не байты .docx: так проверка не зависит от Word и идёт на CI.
    /// </summary>
    internal static class CorpusCheck
    {
        /// <summary>Слово с форматом: из модели документа или из разобранного абзаца.</summary>
        internal sealed class Tok
        {
            public string Text;
            public bool Bold, Italic, Underline, Strike;
            public int Color;
            public double Size;
            public string Font;
            public int Para; // номер абзаца в потоке страницы
        }

        // ---------- ожидаемое ----------

        /// <summary>Ожидаемый абзац страницы.</summary>
        private sealed class EPara
        {
            public List<Tok> Tokens;
            public CParagraph Spec;    // null — колонтитул или строка скана
            public int Lines;
            public bool Centered;      // колонтитул
            public bool LooseBounds;   // строки скана: где кончается абзац, по равному шагу не сказать
            public string Label;       // для сообщений
        }

        private static List<Tok> Tokens(IEnumerable<CRun> runs)
        {
            var result = new List<Tok>();
            foreach (CRun run in runs)
                foreach (string part in run.Text.Split(' '))
                    if (part.Length > 0)
                        result.Add(new Tok
                        {
                            Text = part, Bold = run.Style.Bold, Italic = run.Style.Italic,
                            Underline = run.Style.Underline, Strike = run.Style.Strike,
                            Color = run.Style.Color, Size = run.Style.Size, Font = run.Style.Font
                        });
            return result;
        }

        private static string Short(IList<Tok> tokens, int max = 6)
        {
            var parts = new List<string>();
            for (int i = 0; i < tokens.Count && i < max; i++)
                parts.Add(tokens[i].Text);
            return "«" + string.Join(" ", parts.ToArray()) + (tokens.Count > max ? " …" : "") + "»";
        }

        private static void Expected(CRendered drawn, int page, List<EPara> paras, List<CTable> tables,
            List<Tok> rotated, out bool scan)
        {
            scan = false;
            foreach (CDrawn d in drawn.Drawn)
            {
                if (d.Page != page)
                    continue;
                var p = d.Block as CParagraph;
                if (p != null)
                {
                    paras.Add(new EPara { Tokens = Tokens(p.Runs), Spec = p, Lines = d.Lines, Label = "абзац" });
                    continue;
                }
                var t = d.Block as CTable;
                if (t != null)
                {
                    tables.Add(t);
                    continue;
                }
                var r = d.Block as CRotated;
                if (r != null)
                {
                    rotated.AddRange(Tokens(new[] { new CRun(r.Text, r.Style) }));
                    continue;
                }
                if (d.Block is CScan)
                {
                    // Строки скана стоят с равным шагом — для читателя это один абзац.
                    List<Tok> line = Tokens(new[] { new CRun(d.Text, d.Style) });
                    if (scan)
                        paras[paras.Count - 1].Tokens.AddRange(line);
                    else
                        paras.Add(new EPara { Tokens = line, Lines = 1, LooseBounds = true, Label = "текст скана" });
                    scan = true;
                    continue;
                }
                if (d.Text != null)
                    paras.Add(new EPara
                    {
                        Tokens = Tokens(new[] { new CRun(d.Text, d.Style) }),
                        Lines = 1, Centered = d.Running, Label = "колонтитул"
                    });
            }
        }

        // ---------- разобранное ----------

        private sealed class APara
        {
            public OcrParagraph Source;
            public List<Tok> Tokens;
            public bool MarkerSeen;
        }

        /// <summary>
        /// Слова абзаца с форматом по знакам: граница рана может прийтись на середину слова,
        /// формат слова — формат его первого знака. Маркер списка снимается: распознанный —
        /// по ListContentStart, нераспознанный — если первое слово совпадает с нарисованным
        /// маркером (иначе потерянный список проявлялся бы и как «лишнее слово»).
        /// </summary>
        private static APara Actual(OcrParagraph p, HashSet<string> markers)
        {
            var text = new StringBuilder();
            var owners = new List<OcrRun>();
            foreach (OcrRun run in p.Runs)
            {
                string t = run.Text ?? "";
                text.Append(t);
                for (int i = 0; i < t.Length; i++)
                    owners.Add(run);
            }
            string s = text.ToString();
            int start = p.ListKind != ListKind.None ? Math.Min(p.ListContentStart, s.Length) : 0;
            var result = new APara { Source = p, Tokens = new List<Tok>(), MarkerSeen = start > 0 };
            int i0 = start;
            bool first = true;
            while (i0 < s.Length)
            {
                while (i0 < s.Length && char.IsWhiteSpace(s[i0])) i0++;
                if (i0 >= s.Length) break;
                int j = i0;
                while (j < s.Length && !char.IsWhiteSpace(s[j])) j++;
                string word = s.Substring(i0, j - i0);
                if (first && start == 0 && markers.Contains(word))
                    result.MarkerSeen = true;
                else
                    AddWords(result.Tokens, s, i0, j, owners);
                first = false;
                i0 = j;
            }
            return result;
        }

        /// <summary>
        /// Слово разобранного абзаца. Прочерк, вставленный вместо линии (AddRuleWords), бывает
        /// склеен со словами («основание____перечень»): он отделяется своим словом, чтобы
        /// расхождение звалось прочерком, а не потерей двух слов и появлением третьего.
        /// </summary>
        private static void AddWords(List<Tok> into, string s, int from, int to, List<OcrRun> owners)
        {
            int i = from;
            while (i < to)
            {
                int j = i;
                bool underscores = s[i] == '_';
                while (j < to && (s[j] == '_') == underscores)
                    j++;
                OcrRun r = owners[i];
                into.Add(new Tok
                {
                    Text = s.Substring(i, j - i), Bold = r.Bold, Italic = r.Italic, Underline = r.Underline,
                    Color = r.ColorArgb, Size = r.FontSizePt, Font = r.FontName
                });
                i = j;
            }
        }

        private static void ActualItems(PdfPageText page, HashSet<string> markers, List<APara> paras,
            List<OcrTable> tables, List<OcrImage> images)
        {
            foreach (PageBlocks.PageItem item in PageBlocks.CoalesceRowBands(PageBlocks.OrderedItems(page)))
            {
                var blocks = new List<PageBlocks.Block>();
                if (item.IsBand)
                    foreach (List<PageBlocks.Block> column in item.Columns)
                        blocks.AddRange(column);
                else
                    blocks.Add(item.Single);
                foreach (PageBlocks.Block b in blocks)
                {
                    if (b.Paragraph != null)
                        paras.Add(Actual(b.Paragraph, markers));
                    else if (b.Table != null)
                        tables.Add(b.Table);
                    else if (b.Image != null)
                        images.Add(b.Image);
                }
            }
        }

        // ---------- сверка ----------

        /// <summary>Сверка всех страниц документа (режим Document — то, что получает Word).</summary>
        public static List<CFinding> Document(CRendered drawn, List<PdfPageText> pages)
        {
            var findings = new List<CFinding>();
            if (pages.Count != drawn.Pages)
            {
                findings.Add(new CFinding("pages", pages.Count + " страниц вместо " + drawn.Pages));
                return findings;
            }
            for (int p = 0; p < pages.Count; p++)
                Page(drawn, p, pages[p], findings);
            return Dedup(findings);
        }

        /// <summary>
        /// Один вид — одна запись на документ (иначе отчёт тонет в повторах), но число повторов
        /// сохраняется в <see cref="CFinding.Count"/>: реестр сверяет и его, чтобы регрессия
        /// внутри уже затронутого документа не пряталась за той же записью.
        /// </summary>
        private static List<CFinding> Dedup(List<CFinding> findings)
        {
            var seen = new Dictionary<string, int>(StringComparer.Ordinal);   // записей вида
            var total = new Dictionary<string, int>(StringComparer.Ordinal);  // случаев вида
            var result = new List<CFinding>();
            foreach (CFinding f in findings)
            {
                int n, t;
                total.TryGetValue(f.Aspect, out t);
                total[f.Aspect] = t + f.Count;
                if (seen.TryGetValue(f.Aspect, out n))
                {
                    seen[f.Aspect] = n + 1;
                    continue;
                }
                seen[f.Aspect] = 1;
                result.Add(f);
            }
            for (int i = 0; i < result.Count; i++)
            {
                CFinding f = result[i];
                result[i] = new CFinding(f.Aspect, seen[f.Aspect] > 1 ? f.Detail + " (и ещё " + (seen[f.Aspect] - 1) + ")" : f.Detail,
                    total[f.Aspect]);
            }
            return result;
        }

        private static void Page(CRendered drawn, int index, PdfPageText page, List<CFinding> findings)
        {
            var eParas = new List<EPara>();
            var eTables = new List<CTable>();
            var rotated = new List<Tok>();
            bool scan;
            Expected(drawn, index, eParas, eTables, rotated, out scan);
            var aParas = new List<APara>();
            var aTables = new List<OcrTable>();
            var images = new List<OcrImage>();
            ActualItems(page, drawn.Markers, aParas, aTables, images);
            string where = "стр. " + (index + 1);

            CompareParagraphs(eParas, aParas, aTables, where, findings);
            CompareTables(eTables, aTables, where, findings);

            if (rotated.Count > 0)
            {
                var all = new HashSet<string>(StringComparer.Ordinal);
                foreach (PdfWord w in page.Words)
                    all.Add(w.Text);
                foreach (Tok t in rotated)
                    if (!all.Contains(t.Text))
                    {
                        findings.Add(new CFinding("rotated-lost", where + ": повёрнутая надпись " + Short(rotated) + " пропала"));
                        break;
                    }
            }
            if (scan && images.Count > 0 && eParas.Count > 0)
                findings.Add(new CFinding("scan-duplicate", where + ": текстовый слой скана выведен и текстом, и картинкой"));
            else if (!scan && images.Count > 0)
                findings.Add(new CFinding("image-extra", where + ": " + images.Count + " картинок там, где их не рисовали"));
        }

        // ---------- абзацы ----------

        private static void CompareParagraphs(List<EPara> eParas, List<APara> aParas, List<OcrTable> aTables,
            string where, List<CFinding> findings)
        {
            var et = new List<Tok>();
            var eStart = new List<int>();
            for (int i = 0; i < eParas.Count; i++)
            {
                eStart.Add(et.Count);
                foreach (Tok t in eParas[i].Tokens) { t.Para = i; et.Add(t); }
            }
            var at = new List<Tok>();
            for (int i = 0; i < aParas.Count; i++)
                foreach (Tok t in aParas[i].Tokens) { t.Para = i; at.Add(t); }

            int[] map = Align(et, at); // индекс разобранного слова для каждого ожидаемого, -1 — нет
            var lost = new List<Tok>();
            var used = new bool[at.Count];
            for (int i = 0; i < et.Count; i++)
            {
                if (map[i] < 0) lost.Add(et[i]);
                else used[map[i]] = true;
            }
            var extra = new List<Tok>();
            var filler = new List<Tok>();
            for (int j = 0; j < at.Count; j++)
                if (!used[j])
                {
                    if (at[j].Text.Trim('_').Length == 0) filler.Add(at[j]);
                    else extra.Add(at[j]);
                }
            // Слова абзацев, оказавшиеся в ячейках таблицы, — не потеря, а ложная таблица:
            // сетка без линовки забрала обычный текст.
            if (lost.Count > 0 && aTables.Count > 0)
            {
                var inTables = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (OcrTable table in aTables)
                    foreach (OcrTableRow row in table.Rows)
                        foreach (OcrTableCell cell in row.Cells)
                            foreach (string w in CellText(cell).Split(' '))
                            {
                                int n; inTables.TryGetValue(w, out n); inTables[w] = n + 1;
                            }
                var swallowed = new List<Tok>();
                for (int k = lost.Count - 1; k >= 0; k--)
                {
                    int n;
                    if (inTables.TryGetValue(lost[k].Text, out n) && n > 0)
                    {
                        inTables[lost[k].Text] = n - 1;
                        swallowed.Insert(0, lost[k]);
                        lost.RemoveAt(k);
                    }
                }
                if (swallowed.Count > 0)
                    findings.Add(new CFinding("into-table", where + ": " + swallowed.Count + " слов абзацев ушли в ячейки таблицы: " + Short(swallowed), swallowed.Count));
            }
            if (lost.Count > 0)
            {
                // Те же слова, но в другом порядке, — не потеря, а порядок чтения.
                var a = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (Tok t in extra) { int n; a.TryGetValue(t.Text, out n); a[t.Text] = n + 1; }
                bool reordered = lost.Count == extra.Count;
                foreach (Tok t in lost)
                {
                    int n;
                    if (!a.TryGetValue(t.Text, out n) || n == 0) { reordered = false; break; }
                    a[t.Text] = n - 1;
                }
                if (reordered)
                    findings.Add(new CFinding("order", where + ": порядок чтения нарушен у " + lost.Count + " слов, первое " + Short(lost, 3), lost.Count));
                else
                    findings.Add(new CFinding("lost", where + ": потеряно " + lost.Count + " слов: " + Short(lost), lost.Count));
            }
            if (extra.Count > 0 && lost.Count != extra.Count)
                findings.Add(new CFinding("extra", where + ": лишние слова " + Short(extra), extra.Count));
            if (filler.Count > 0)
                findings.Add(new CFinding("filler", where + ": прочерки «" + filler[0].Text + "» вместо линии", filler.Count));

            // Границы абзацев: ожидаемый абзац целиком в ОДНОМ разобранном и только он.
            var exact = new Dictionary<int, int>(); // ожидаемый → разобранный
            for (int i = 0; i < eParas.Count; i++)
            {
                int from = eStart[i], to = from + eParas[i].Tokens.Count;
                if (to == from)
                    continue;
                var owners = new HashSet<int>();
                bool complete = true;
                for (int k = from; k < to; k++)
                {
                    if (map[k] < 0) { complete = false; continue; }
                    owners.Add(at[map[k]].Para);
                }
                if (owners.Count == 0 || eParas[i].LooseBounds)
                    continue;
                // Начало абзаца — там, где стоит его первое слово: у дроблёного или слитого
                // абзаца список (маркер, вид, уровень) всё равно проверяется по этой части.
                int head = map[from] >= 0 ? at[map[from]].Para : -1;
                bool startsThere = head >= 0 && aParas[head].Tokens.Count > 0 && ReferenceEquals(aParas[head].Tokens[0], at[map[from]]);
                if (owners.Count > 1)
                    findings.Add(new CFinding("split", where + ": " + eParas[i].Label + " " + Short(eParas[i].Tokens) +
                        " разбит на " + owners.Count));
                else if (complete && WordCount(aParas[head]) != eParas[i].Tokens.Count)
                    findings.Add(new CFinding("merge", where + ": " + eParas[i].Label + " " + Short(eParas[i].Tokens) +
                        " слит с соседним"));
                else if (complete)
                {
                    exact[i] = head;
                    continue;
                }
                if (startsThere)
                    CompareList(eParas[i], aParas[head], where, findings);
            }

            foreach (KeyValuePair<int, int> pair in exact)
                CompareParagraph(eParas[pair.Key], aParas[pair.Value], where, findings);

            for (int i = 0; i < et.Count; i++)
                if (map[i] >= 0)
                    CompareToken(et[i], at[map[i]], where, findings);
        }

        /// <summary>Слов в разобранном абзаце, без прочерков-заполнителей (о них — свой вид).</summary>
        private static int WordCount(APara a)
        {
            int n = 0;
            foreach (Tok t in a.Tokens)
                if (t.Text.Trim('_').Length > 0)
                    n++;
            return n;
        }

        private static void CompareParagraph(EPara e, APara a, string where, List<CFinding> findings)
        {
            OcrParagraph p = a.Source;
            string name = where + ": " + e.Label + " " + Short(e.Tokens);
            if (e.Centered)
            {
                if (p.Alignment != OcrAlignment.Center)
                    findings.Add(new CFinding("align-center", name + " — " + p.Alignment + " вместо Center"));
                return;
            }
            if (e.Spec == null)
                return;
            CParagraph s = e.Spec;
            CompareList(e, a, where, findings);
            if (s.List == ListKind.None)
            {
                if (s.Align == CAlign.Center && p.Alignment != OcrAlignment.Center)
                    findings.Add(new CFinding("align-center", name + " — " + p.Alignment + " вместо Center"));
                else if (s.Align == CAlign.Right)
                    findings.Add(new CFinding("align-right", name + " — " + p.Alignment + " вместо Right (в модели нет Right)"));
                else if ((s.Align == CAlign.Justify || s.Align == CAlign.Distribute) && e.Lines >= 2 && p.Alignment != OcrAlignment.Justify)
                    findings.Add(new CFinding("align-justify", name + " — " + p.Alignment + " вместо Justify"));
                else if (s.Align == CAlign.Left && e.Lines >= 2 && p.Alignment != OcrAlignment.Left)
                    findings.Add(new CFinding("align-left", name + " — " + p.Alignment + " вместо Left"));
            }
        }

        /// <summary>Список: обычный абзац не стал пунктом; пункт — вид, маркер, уровень, номер.</summary>
        private static void CompareList(EPara e, APara a, string where, List<CFinding> findings)
        {
            OcrParagraph p = a.Source;
            CParagraph s = e.Spec;
            if (s == null)
                return;
            string name = where + ": " + e.Label + " " + Short(e.Tokens);
            if (s.List == ListKind.None)
            {
                if (p.ListKind != ListKind.None)
                    findings.Add(new CFinding("list-false", name + " стал пунктом списка (" + p.ListKind + ")"));
                return;
            }
            if (p.ListKind == ListKind.None)
            {
                // Центрированный абзац список не ищет вовсе (OcrLayout): это другая причина,
                // чем незнакомый маркер, и у неё свой вид.
                if (a.MarkerSeen && p.Alignment == OcrAlignment.Center)
                    findings.Add(new CFinding("list-centered", name + " счёлся центрированным и потерял список"));
                else
                    findings.Add(new CFinding(a.MarkerSeen ? "list-kind" : "marker-lost",
                        name + (a.MarkerSeen ? " не распознан как пункт (маркер «" + s.Marker + "»)"
                            : " — маркер «" + Visible(s.Marker) + "» пропал")));
                return;
            }
            if (p.ListKind != s.List)
                findings.Add(new CFinding("list-kind", name + " — " + p.ListKind + " вместо " + s.List));
            if (p.ListLevel != s.Level)
                findings.Add(new CFinding("list-level", name + " — уровень " + p.ListLevel + " вместо " + s.Level));
            if (s.List == ListKind.Numbered && p.ListKind == ListKind.Numbered && p.ListNumber != s.Number)
                findings.Add(new CFinding("list-number", name + " — номер " + p.ListNumber + " вместо " + s.Number));
        }

        private static string Visible(string marker)
        {
            var sb = new StringBuilder();
            foreach (char c in marker ?? "")
                sb.Append(c >= 0xE000 && c <= 0xF8FF ? "U+" + ((int)c).ToString("X4") : c.ToString());
            return sb.ToString();
        }

        private static string FontKey(string font)
        {
            return (font ?? "").Replace(" ", "").ToLowerInvariant();
        }

        private static void CompareToken(Tok e, Tok a, string where, List<CFinding> findings)
        {
            string w = where + ": «" + e.Text + "»";
            if (e.Bold != a.Bold) findings.Add(new CFinding("bold", w + (e.Bold ? " потерял полужирный" : " стал полужирным")));
            if (e.Italic != a.Italic) findings.Add(new CFinding("italic", w + (e.Italic ? " потерял курсив" : " стал курсивом")));
            if (e.Underline != a.Underline) findings.Add(new CFinding("underline", w + (e.Underline ? " потерял подчёркивание" : " стал подчёркнутым")));
            if (e.Strike) findings.Add(new CFinding("strike", w + " — зачёркивания в модели нет"));
            if (e.Color != a.Color) findings.Add(new CFinding("color", w + " — цвет " + a.Color.ToString("X6") + " вместо " + e.Color.ToString("X6")));
            if (Math.Abs(e.Size - a.Size) > 0.3) findings.Add(new CFinding("size", w + " — кегль " + a.Size.ToString("0.##", CultureInfo.InvariantCulture) + " вместо " + e.Size.ToString(CultureInfo.InvariantCulture)));
            if (FontKey(e.Font) != FontKey(a.Font)) findings.Add(new CFinding("font", w + " — шрифт «" + a.Font + "» вместо «" + e.Font + "»"));
        }

        /// <summary>
        /// Сопоставление ожидаемых и разобранных слов: наибольшая общая подпоследовательность
        /// по тексту. Общие начало и конец снимаются сразу (в норме совпадает почти всё), на
        /// середину — динамика с ограничением размера: на очень большом расхождении слова
        /// сопоставляются жадно, и потерь это не скрывает.
        /// </summary>
        internal static int[] Align(List<Tok> e, List<Tok> a)
        {
            var map = new int[e.Count];
            for (int i = 0; i < map.Length; i++) map[i] = -1;
            int lo = 0;
            while (lo < e.Count && lo < a.Count && e[lo].Text == a[lo].Text) { map[lo] = lo; lo++; }
            int he = e.Count - 1, ha = a.Count - 1;
            while (he >= lo && ha >= lo && e[he].Text == a[ha].Text) { map[he] = ha; he--; ha--; }
            int n = he - lo + 1, m = ha - lo + 1;
            if (n <= 0 || m <= 0)
                return map;
            if ((long)n * m > 4000000)
            {
                int j = lo;
                for (int i = lo; i <= he; i++)
                    for (int k = j; k <= ha; k++)
                        if (e[i].Text == a[k].Text) { map[i] = k; j = k + 1; break; }
                return map;
            }
            var dp = new int[n + 1, m + 1];
            for (int i = n - 1; i >= 0; i--)
                for (int k = m - 1; k >= 0; k--)
                    dp[i, k] = e[lo + i].Text == a[lo + k].Text ? dp[i + 1, k + 1] + 1 : Math.Max(dp[i + 1, k], dp[i, k + 1]);
            int x = 0, y = 0;
            while (x < n && y < m)
            {
                if (e[lo + x].Text == a[lo + y].Text) { map[lo + x] = lo + y; x++; y++; }
                else if (dp[x + 1, y] >= dp[x, y + 1]) x++;
                else y++;
            }
            return map;
        }

        // ---------- таблицы ----------

        private static string CellText(OcrTableCell cell)
        {
            var parts = new List<string>();
            foreach (OcrParagraph p in cell.Paragraphs)
                parts.Add(p.Text);
            return Norm(string.Join(" ", parts.ToArray()));
        }

        private static string Norm(string s)
        {
            return string.Join(" ", (s ?? "").Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries));
        }

        private static void CompareTables(List<CTable> expected, List<OcrTable> actual, string where, List<CFinding> findings)
        {
            if (actual.Count > expected.Count)
                findings.Add(new CFinding("table-extra", where + ": таблиц " + actual.Count + " вместо " + expected.Count +
                    (actual.Any(t => t.Borderless) ? " (лишняя — без линовки)" : "")));
            else if (actual.Count < expected.Count)
                findings.Add(new CFinding("table-lost", where + ": таблиц " + actual.Count + " вместо " + expected.Count));
            if (actual.Count != expected.Count)
                return; // какая с какой — не сказать: ячейки сверяются только при равном числе
            for (int t = 0; t < expected.Count && t < actual.Count; t++)
            {
                CTable e = expected[t];
                OcrTable a = actual[t];
                string name = where + ", таблица " + (t + 1);
                int cols = e.ColWidths.Length;
                if (a.Rows.Count != e.Rows.Count || a.ColumnCount != cols)
                {
                    findings.Add(new CFinding("table-shape", name + ": " + a.Rows.Count + "×" + a.ColumnCount +
                        " вместо " + e.Rows.Count + "×" + cols));
                    continue;
                }
                if (a.Borderless == e.Ruled)
                    findings.Add(new CFinding("table-ruling", name + (e.Ruled ? ": линованная стала безлиновочной" : ": безлиновочная стала линованной")));
                for (int r = 0; r < e.Rows.Count; r++)
                    for (int c = 0; c < cols; c++)
                    {
                        CCell ec = e.Rows[r][c];
                        if (ec == null)
                            continue;
                        OcrTableCell ac = a.Rows[r].Cells[c];
                        string want = Norm(string.Join(" ", ec.Runs.Select(x => x.Text).ToArray()));
                        string got = CellText(ac);
                        if (want != got)
                            findings.Add(new CFinding("table-cell", name + " [" + r + "," + c + "]: «" + got + "» вместо «" + want + "»"));
                        if (ac.ColSpan != ec.ColSpan || ac.RowSpan != ec.RowSpan)
                            findings.Add(new CFinding("table-span", name + " [" + r + "," + c + "]: объединение " +
                                ac.ColSpan + "×" + ac.RowSpan + " вместо " + ec.ColSpan + "×" + ec.RowSpan));
                    }
            }
        }

        // ---------- слайды ----------

        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        /// <summary>
        /// Сверка .pptx (режим Slide, без подложки): каждое нарисованное слово ровно один раз,
        /// с тем же форматом; повёрнутые надписи и маркеры списков на месте; текстовый слой
        /// скана невидим. Пакет проверяется официальным валидатором Open XML.
        /// </summary>
        public static List<CFinding> Slides(CRendered drawn, OoxmlPackage package, Func<string, List<string>> validate,
            string tempDir)
        {
            var findings = new List<CFinding>();
            string path = Path.Combine(tempDir, "slides_" + Guid.NewGuid().ToString("N") + ".pptx");
            package.Save(path);
            foreach (string problem in validate(path))
            {
                findings.Add(new CFinding("pptx-invalid", problem));
                break;
            }
            for (int page = 0; page < drawn.Pages; page++)
            {
                var expected = new List<Tok>();
                var markers = new List<string>();
                bool scan = false;
                foreach (CDrawn d in drawn.Drawn)
                {
                    if (d.Page != page) continue;
                    var p = d.Block as CParagraph;
                    if (p != null)
                    {
                        expected.AddRange(Tokens(p.Runs));
                        if (p.Marker != null) markers.Add(p.Marker);
                    }
                    var t = d.Block as CTable;
                    if (t != null)
                        foreach (CCell[] row in t.Rows)
                            foreach (CCell cell in row)
                                if (cell != null) expected.AddRange(Tokens(cell.Runs));
                    var r = d.Block as CRotated;
                    if (r != null) expected.AddRange(Tokens(new[] { new CRun(r.Text, r.Style) }));
                    if (d.Block is CScan) scan = true;
                    if (d.Text != null) expected.AddRange(Tokens(new[] { new CRun(d.Text, d.Style) }));
                }
                SlidePage(drawn, package, page, expected, markers, scan, findings);
            }
            return Dedup(findings);
        }

        private static void SlidePage(CRendered drawn, OoxmlPackage package, int page, List<Tok> expected,
            List<string> markers, bool scan, List<CFinding> findings)
        {
            string where = "слайд " + (page + 1);
            XDocument xml = SlideXml(package, page + 1);
            var actual = new List<Tok>();
            var seenMarkers = new List<string>();
            bool visibleOcr = false;
            foreach (XElement run in xml.Descendants(A + "r"))
            {
                XElement rPr = run.Element(A + "rPr");
                string text = (string)run.Element(A + "t") ?? "";
                bool bold = rPr != null && (string)rPr.Attribute("b") == "1";
                bool italic = rPr != null && (string)rPr.Attribute("i") == "1";
                bool underline = rPr != null && rPr.Attribute("u") != null && (string)rPr.Attribute("u") != "none";
                bool noFill = rPr != null && rPr.Element(A + "noFill") != null;
                foreach (string word in text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (drawn.Markers.Contains(word)) { seenMarkers.Add(word); continue; }
                    actual.Add(new Tok { Text = word, Bold = bold, Italic = italic, Underline = underline });
                    if (scan && !noFill) visibleOcr = true;
                }
            }
            var want = Count(expected, t => t.Text);
            var got = Count(actual, t => t.Text);
            var lost = Diff(want, got);
            var extra = Diff(got, want);
            if (lost.Count > 0) findings.Add(new CFinding("pptx-lost", where + ": потеряны " + string.Join(" ", lost.Take(6).ToArray()), lost.Count));
            if (extra.Count > 0) findings.Add(new CFinding("pptx-extra", where + ": лишние " + string.Join(" ", extra.Take(6).ToArray()), extra.Count));
            foreach (string m in markers)
                if (!seenMarkers.Remove(m))
                {
                    findings.Add(new CFinding("pptx-marker-lost", where + ": маркер «" + Visible(m) + "» пропал"));
                    break;
                }
            CompareFormat("pptx-bold", "полужирный", expected, actual, t => t.Bold, where, findings);
            CompareFormat("pptx-italic", "курсив", expected, actual, t => t.Italic, where, findings);
            CompareFormat("pptx-underline", "подчёркивание", expected, actual, t => t.Underline, where, findings);
            if (visibleOcr)
                findings.Add(new CFinding("pptx-ocr-visible", where + ": невидимый текстовый слой скана стал видимым текстом"));
        }

        private static void CompareFormat(string aspect, string what, List<Tok> expected, List<Tok> actual,
            Func<Tok, bool> flag, string where, List<CFinding> findings)
        {
            var want = Count(expected.Where(flag), t => t.Text);
            var got = Count(actual.Where(flag), t => t.Text);
            List<string> lost = Diff(want, got), extra = Diff(got, want);
            if (lost.Count > 0 || extra.Count > 0)
                findings.Add(new CFinding(aspect, where + ": " + what + (lost.Count > 0 ? " потерян у " + string.Join(" ", lost.Take(4).ToArray()) : "")
                    + (extra.Count > 0 ? " лишний у " + string.Join(" ", extra.Take(4).ToArray()) : ""), lost.Count + extra.Count));
        }

        private static Dictionary<string, int> Count(IEnumerable<Tok> tokens, Func<Tok, string> key)
        {
            var d = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (Tok t in tokens) { int n; d.TryGetValue(key(t), out n); d[key(t)] = n + 1; }
            return d;
        }

        private static List<string> Diff(Dictionary<string, int> a, Dictionary<string, int> b)
        {
            var result = new List<string>();
            foreach (KeyValuePair<string, int> p in a)
            {
                int n;
                b.TryGetValue(p.Key, out n);
                for (int i = n; i < p.Value; i++)
                    result.Add(p.Key);
            }
            return result;
        }

        private static XDocument SlideXml(OoxmlPackage package, int slide)
        {
            using (var ms = new MemoryStream(package.ToArray()))
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
            using (Stream s = zip.GetEntry("ppt/slides/slide" + slide + ".xml").Open())
            using (var reader = new StreamReader(s, Encoding.UTF8))
                return XDocument.Parse(reader.ReadToEnd());
        }
    }
}
