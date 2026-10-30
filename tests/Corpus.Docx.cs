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
    /// <summary>
    /// Сверка .docx, записанного Word (<see cref="WordDocxWriter"/>), с моделью разметки, из которой
    /// он записан. Это вторая половина проверки DOCX: распознавание сверяется с моделью документа
    /// в <see cref="CorpusCheck.Document"/>, а здесь — что писатель перенёс модель в Word без потерь:
    /// абзацы и их текст, выключку, списки (вид и уровень), таблицы, формат каждого слова.
    /// Правила писателя (что меняется намеренно) повторены явно: в ячейке выключка по ширине
    /// становится влево и списков нет, маркер распознанного пункта снимается — Word ставит свой.
    /// </summary>
    internal static class CorpusDocx
    {
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace Dml = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        /// <summary>
        /// Узел внутри плавающей фигуры (надписи): w:drawing, w:pict или mc:AlternateContent. Текст
        /// надписи лежит в якорном абзаце потока, но к нему не относится; к тому же Word пишет
        /// его дважды — в mc:Choice (DrawingML) и в mc:Fallback (VML).
        /// </summary>
        private static bool InFloatingShape(XElement e)
        {
            return e.Ancestors().Any(a => a.Name == W + "drawing" || a.Name == W + "pict" || a.Name == Mc + "AlternateContent");
        }

        /// <summary>Абзац для сверки: текст, выключка в терминах w:jc, список и слова с форматом.</summary>
        private sealed class WPara
        {
            public string Text;
            public string Align;        // left | center | both | right
            public ListKind List;
            public int Level;
            public List<CorpusCheck.Tok> Words;
        }

        // ---------- ожидаемое: модель так, как её пишет WordDocxWriter ----------

        private static string AlignOf(OcrAlignment a, bool inCell)
        {
            switch (a)
            {
                case OcrAlignment.Center: return "center";
                case OcrAlignment.Justify: return inCell ? "left" : "both";
                case OcrAlignment.Right: return "right";
                default: return "left";
            }
        }

        private static WPara Expected(OcrParagraph p, bool inCell)
        {
            bool list = !inCell && p.ListKind != ListKind.None;
            int skip = list ? p.ListContentStart : 0;
            var chars = new StringBuilder();
            var owners = new List<OcrRun>();
            foreach (OcrRun run in p.Runs)
            {
                string t = run.Text ?? "";
                for (int i = 0; i < t.Length; i++)
                {
                    if (skip > 0) { skip--; continue; }
                    chars.Append(t[i]);
                    owners.Add(run);
                }
            }
            var words = new List<CorpusCheck.Tok>();
            ForEachWord(chars.ToString(), delegate(string word, int at)
            {
                OcrRun r = owners[at];
                words.Add(new CorpusCheck.Tok
                {
                    Text = word, Bold = r.Bold, Italic = r.Italic, Underline = r.Underline,
                    Color = r.ColorArgb & 0xFFFFFF,
                    // Word хранит кегль в полупунктах: 11,1 пишется как 11.
                    Size = Math.Round(FontResolver.ClampSizePt(r.FontSizePt) * 2) / 2,
                    Font = FontResolver.Resolve(r.FontName, r.Text)
                });
            });
            return new WPara
            {
                Text = CorpusCheck.Norm(chars.ToString()), Align = AlignOf(p.Alignment, inCell),
                List = list ? p.ListKind : ListKind.None, Level = p.ListLevel, Words = words
            };
        }

        private static void ForEachWord(string s, Action<string, int> word)
        {
            int i = 0;
            while (i < s.Length)
            {
                while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
                if (i >= s.Length) break;
                int j = i;
                while (j < s.Length && !char.IsWhiteSpace(s[j])) j++;
                word(s.Substring(i, j - i), i);
                i = j;
            }
        }

        /// <summary>
        /// Модель всех страниц в порядке записи: тот же обход, что в WordDocxWriter.Write. Таблицы
        /// идут в список tables по порядку появления w:tbl; полоса колонок — тоже w:tbl (null:
        /// её сетка — приём вёрстки, а не таблица источника, и сверяются только абзацы в ней).
        /// </summary>
        private static List<WPara> Expected(List<PdfPageText> pages, List<CorpusCheck.CGrid> tables)
        {
            var result = new List<WPara>();
            foreach (PdfPageText page in pages)
                foreach (PageBlocks.PageItem item in PageBlocks.CoalesceRowBands(PageBlocks.OrderedItems(page)))
                {
                    if (item.IsBand)
                    {
                        tables.Add(null);
                        foreach (List<PageBlocks.Block> column in item.Columns)
                            foreach (PageBlocks.Block b in column)
                                if (b.Paragraph != null)
                                    result.Add(Expected(b.Paragraph, true));
                        continue;
                    }
                    PageBlocks.Block blk = item.Single;
                    if (blk.Paragraph != null)
                        result.Add(Expected(blk.Paragraph, false));
                    else if (blk.Table != null && blk.Table.Rows.Count > 0 && blk.Table.ColumnCount > 0)
                    {
                        tables.Add(CorpusCheck.CGrid.Of(blk.Table));
                        foreach (OcrTableRow row in blk.Table.Rows)
                            foreach (OcrTableCell cell in row.Cells)
                                if (!cell.Covered && cell.Paragraphs != null)
                                    foreach (OcrParagraph p in cell.Paragraphs)
                                        result.Add(Expected(p, true));
                    }
                }
            result.RemoveAll(p => p.Text.Length == 0);
            return result;
        }

        // ---------- разобранное: .docx, как его прочтёт Word ----------

        /// <summary>Части пакета и разрешение формата по цепочке «ран → стиль → стиль абзаца → умолчания».</summary>
        private sealed class Package
        {
            public XDocument Document, Styles, Numbering, Theme;

            public Package(string path)
            {
                using (FileStream file = File.OpenRead(path))
                using (var zip = new ZipArchive(file, ZipArchiveMode.Read))
                {
                    Document = Part(zip, "word/document.xml");
                    Styles = Part(zip, "word/styles.xml");
                    Numbering = Part(zip, "word/numbering.xml");
                    Theme = Part(zip, "word/theme/theme1.xml");
                }
            }

            private static XDocument Part(ZipArchive zip, string name)
            {
                ZipArchiveEntry e = zip.GetEntry(name);
                if (e == null) return null;
                using (Stream s = e.Open())
                    return XDocument.Load(s);
            }

            private XElement Style(string id)
            {
                return Styles == null || id == null ? null
                    : Styles.Root.Elements(W + "style").FirstOrDefault(s => (string)s.Attribute(W + "styleId") == id);
            }

            /// <summary>Стиль и его предки по basedOn (не глубже 10 — защита от цикла).</summary>
            private IEnumerable<XElement> Chain(string id)
            {
                for (int depth = 0; depth < 10 && id != null; depth++)
                {
                    XElement s = Style(id);
                    if (s == null) yield break;
                    yield return s;
                    XElement based = s.Element(W + "basedOn");
                    id = based != null ? (string)based.Attribute(W + "val") : null;
                }
            }

            private string DefaultParagraphStyle()
            {
                if (Styles == null) return null;
                XElement s = Styles.Root.Elements(W + "style").FirstOrDefault(x =>
                    (string)x.Attribute(W + "type") == "paragraph" && (string)x.Attribute(W + "default") == "1");
                return s != null ? (string)s.Attribute(W + "styleId") : null;
            }

            private static string Val(XElement props, string name)
            {
                XElement e = props != null ? props.Element(W + name) : null;
                return e == null ? null : ((string)e.Attribute(W + "val") ?? "");
            }

            private string ParagraphStyle(XElement p)
            {
                return Val(p.Element(W + "pPr"), "pStyle") ?? DefaultParagraphStyle();
            }

            /// <summary>Свойство абзаца (w:pPr/name/@val) с учётом стиля абзаца; null — нигде не задано.</summary>
            public string ParagraphProp(XElement p, string name)
            {
                string v = Val(p.Element(W + "pPr"), name);
                if (v != null) return v;
                foreach (XElement s in Chain(ParagraphStyle(p)))
                {
                    v = Val(s.Element(W + "pPr"), name);
                    if (v != null) return v;
                }
                return null;
            }

            /// <summary>Элемент свойства рана с учётом стиля рана, стиля абзаца и умолчаний документа.</summary>
            public XElement RunProp(XElement r, XElement p, string name)
            {
                XElement rPr = r.Element(W + "rPr");
                XElement e = rPr != null ? rPr.Element(W + name) : null;
                if (e != null) return e;
                string rStyle = Val(rPr, "rStyle");
                foreach (XElement s in Chain(rStyle).Concat(Chain(ParagraphStyle(p))))
                {
                    XElement sr = s.Element(W + "rPr");
                    e = sr != null ? sr.Element(W + name) : null;
                    if (e != null) return e;
                }
                XElement defaults = Styles == null ? null : Styles.Root.Element(W + "docDefaults");
                XElement dr = defaults != null && defaults.Element(W + "rPrDefault") != null
                    ? defaults.Element(W + "rPrDefault").Element(W + "rPr") : null;
                return dr != null ? dr.Element(W + name) : null;
            }

            /// <summary>Шрифт латиницы: явное имя или шрифт темы (minorHAnsi/majorHAnsi).</summary>
            public string Font(XElement rFonts)
            {
                if (rFonts == null) return null;
                string ascii = (string)rFonts.Attribute(W + "ascii");
                if (ascii != null) return ascii;
                string theme = (string)rFonts.Attribute(W + "asciiTheme");
                if (theme == null || Theme == null) return null;
                XElement scheme = Theme.Descendants(Dml + (theme.StartsWith("major", StringComparison.Ordinal) ? "majorFont" : "minorFont")).FirstOrDefault();
                XElement latin = scheme != null ? scheme.Element(Dml + "latin") : null;
                return latin != null ? (string)latin.Attribute("typeface") : null;
            }

            /// <summary>Вид нумерации уровня ilvl списка numId (bullet, decimal, lowerLetter…).</summary>
            public string NumFormat(string numId, int ilvl)
            {
                if (Numbering == null || numId == null) return null;
                XElement num = Numbering.Root.Elements(W + "num").FirstOrDefault(n => (string)n.Attribute(W + "numId") == numId);
                if (num == null) return null;
                string abs = Val(num, "abstractNumId");
                XElement an = Numbering.Root.Elements(W + "abstractNum").FirstOrDefault(a => (string)a.Attribute(W + "abstractNumId") == abs);
                XElement lvl = an == null ? null
                    : an.Elements(W + "lvl").FirstOrDefault(l => (string)l.Attribute(W + "ilvl") == ilvl.ToString(CultureInfo.InvariantCulture));
                return Val(lvl, "numFmt");
            }
        }

        private static bool On(XElement toggle)
        {
            if (toggle == null) return false;
            string v = (string)toggle.Attribute(W + "val");
            return v == null || v == "1" || v == "true" || v == "on";
        }

        private static WPara Actual(Package pkg, XElement p)
        {
            var chars = new StringBuilder();
            var owners = new List<XElement>();
            bool floating = InFloatingShape(p);
            foreach (XElement e in p.Descendants())
            {
                if (!floating && InFloatingShape(e))
                    continue; // текст надписи — не текст якорного абзаца
                if (e.Name == W + "t") { chars.Append(e.Value); for (int i = 0; i < e.Value.Length; i++) owners.Add(e.Parent); }
                else if (e.Name == W + "tab" || e.Name == W + "br") { chars.Append(' '); owners.Add(e.Parent); }
            }
            var words = new List<CorpusCheck.Tok>();
            ForEachWord(chars.ToString(), delegate(string word, int at)
            {
                XElement r = owners[at];
                XElement sz = pkg.RunProp(r, p, "sz"), color = pkg.RunProp(r, p, "color"), u = pkg.RunProp(r, p, "u");
                string colorVal = color != null ? (string)color.Attribute(W + "val") : null;
                string uVal = u != null ? (string)u.Attribute(W + "val") : null;
                words.Add(new CorpusCheck.Tok
                {
                    Text = word,
                    Bold = On(pkg.RunProp(r, p, "b")),
                    Italic = On(pkg.RunProp(r, p, "i")),
                    Underline = uVal != null && uVal != "none",
                    Color = colorVal == null || colorVal == "auto" ? 0 : int.Parse(colorVal, NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                    Size = sz != null ? (double)sz.Attribute(W + "val") / 2 : 0,
                    Font = pkg.Font(pkg.RunProp(r, p, "rFonts"))
                });
            });
            string jc = pkg.ParagraphProp(p, "jc");
            XElement numPr = p.Element(W + "pPr") != null ? p.Element(W + "pPr").Element(W + "numPr") : null;
            var result = new WPara
            {
                Text = CorpusCheck.Norm(chars.ToString()),
                Align = jc == null || jc == "start" ? "left" : (jc == "end" ? "right" : jc),
                Words = words
            };
            if (numPr != null)
            {
                XElement ilvl = numPr.Element(W + "ilvl"), numId = numPr.Element(W + "numId");
                result.Level = ilvl != null ? (int)ilvl.Attribute(W + "val") : 0;
                string id = numId != null ? (string)numId.Attribute(W + "val") : null;
                // numId 0 — явное «без списка» (так Word снимает унаследованную нумерацию).
                if (id != null && id != "0")
                    result.List = pkg.NumFormat(id, result.Level) == "bullet" ? ListKind.Bulleted : ListKind.Numbered;
            }
            return result;
        }

        /// <summary>Таблица Word как сетка: gridSpan растягивает ячейку, vMerge продолжает верхнюю.</summary>
        private static CorpusCheck.CGrid Grid(Package pkg, XElement tbl)
        {
            List<XElement> rows = tbl.Elements(W + "tr").ToList();
            XElement gridEl = tbl.Element(W + "tblGrid");
            int cols = gridEl != null ? gridEl.Elements(W + "gridCol").Count() : 0;
            var g = new CorpusCheck.CGrid(rows.Count, cols);
            XElement borders = tbl.Element(W + "tblPr") != null ? tbl.Element(W + "tblPr").Element(W + "tblBorders") : null;
            g.Ruled = borders != null && borders.Elements().Any(b => (string)b.Attribute(W + "val") != "none" && (string)b.Attribute(W + "val") != "nil");
            for (int r = 0; r < rows.Count; r++)
            {
                int c = 0;
                foreach (XElement tc in rows[r].Elements(W + "tc"))
                {
                    if (c >= cols) break;
                    XElement tcPr = tc.Element(W + "tcPr");
                    XElement span = tcPr != null ? tcPr.Element(W + "gridSpan") : null;
                    XElement vMerge = tcPr != null ? tcPr.Element(W + "vMerge") : null;
                    int colSpan = span != null ? (int)span.Attribute(W + "val") : 1;
                    bool continues = vMerge != null && (string)vMerge.Attribute(W + "val") != "restart";
                    for (int k = c; k < c + colSpan && k < cols; k++)
                        g.Text[r, k] = null;
                    if (!continues)
                    {
                        g.Text[r, c] = string.Join(" ", tc.Elements(W + "p").Select(p => Actual(pkg, p).Text).Where(t => t.Length > 0).ToArray());
                        g.ColSpan[r, c] = colSpan;
                        if (vMerge != null)
                            g.RowSpan[r, c] = 1 + rows.Skip(r + 1).TakeWhile(next => Continues(next, c)).Count();
                    }
                    c += colSpan;
                }
            }
            return g;
        }

        /// <summary>Продолжает ли ячейка строки row в колонке col объединение сверху.</summary>
        private static bool Continues(XElement row, int col)
        {
            int c = 0;
            foreach (XElement tc in row.Elements(W + "tc"))
            {
                XElement tcPr = tc.Element(W + "tcPr");
                XElement span = tcPr != null ? tcPr.Element(W + "gridSpan") : null;
                if (c == col)
                {
                    XElement vMerge = tcPr != null ? tcPr.Element(W + "vMerge") : null;
                    return vMerge != null && (string)vMerge.Attribute(W + "val") != "restart";
                }
                c += span != null ? (int)span.Attribute(W + "val") : 1;
                if (c > col) return false;
            }
            return false;
        }

        // ---------- сверка ----------

        /// <summary>Сверить записанный .docx с моделью pages. Виды — с приставкой «word-».</summary>
        public static List<CFinding> Compare(List<PdfPageText> pages, string docxPath)
        {
            var findings = new List<CFinding>();
            var eTables = new List<CorpusCheck.CGrid>();
            List<WPara> expected = Expected(pages, eTables);
            var pkg = new Package(docxPath);
            XElement body = pkg.Document.Root.Element(W + "body");
            List<WPara> actual = body.Descendants(W + "p").Where(p => !InFloatingShape(p))
                .Select(p => Actual(pkg, p)).Where(p => p.Text.Length > 0).ToList();
            CompareRotated(pages, body, findings);

            // Абзацы — по тексту целиком: LCS тем же сопоставлением, что у слов корпуса.
            List<CorpusCheck.Tok> et = expected.Select(p => new CorpusCheck.Tok { Text = p.Text }).ToList();
            List<CorpusCheck.Tok> at = actual.Select(p => new CorpusCheck.Tok { Text = p.Text }).ToList();
            int[] map = CorpusCheck.Align(et, at);
            var used = new bool[actual.Count];
            for (int i = 0; i < expected.Count; i++)
            {
                if (map[i] < 0)
                {
                    findings.Add(new CFinding("word-lost", "абзац «" + Head(expected[i].Text) + "» не найден в .docx"));
                    continue;
                }
                used[map[i]] = true;
                ComparePara(expected[i], actual[map[i]], findings);
            }
            for (int j = 0; j < actual.Count; j++)
                if (!used[j])
                    findings.Add(new CFinding("word-extra", "лишний абзац «" + Head(actual[j].Text) + "»"));

            // Таблицы — по порядку w:tbl верхнего уровня; полосы колонок (null) пропускаются.
            List<XElement> tbls = body.Descendants(W + "tbl").Where(t => !t.Ancestors(W + "tbl").Any()).ToList();
            if (tbls.Count != eTables.Count)
                findings.Add(new CFinding("word-table-count", "таблиц " + tbls.Count + " вместо " + eTables.Count));
            else
            {
                var e = new List<CorpusCheck.CGrid>();
                var a = new List<CorpusCheck.CGrid>();
                for (int i = 0; i < eTables.Count; i++)
                    if (eTables[i] != null) { e.Add(eTables[i]); a.Add(Grid(pkg, tbls[i])); }
                CorpusCheck.CompareGrids("word-", e, a, "docx", findings);
            }
            return CorpusCheck.Dedup(findings);
        }

        /// <summary>
        /// Надписи под углом (WP-3): каждая надпись модели — ровно одна плавающая надпись Word
        /// (только mc:Choice — копию mc:Fallback не считаем) тем же текстом и с тем же поворотом
        /// по часовой (a:xfrm/@rot в 60000-х градуса, ±1°).
        /// </summary>
        private static void CompareRotated(List<PdfPageText> pages, XElement body, List<CFinding> findings)
        {
            var actual = new List<KeyValuePair<string, string>>();
            foreach (XElement box in body.Descendants(W + "txbxContent").Where(b => !b.Ancestors(Mc + "Fallback").Any()))
            {
                string text = CorpusCheck.Norm(string.Join(" ", box.Descendants(W + "t").Select(t => t.Value).ToArray()));
                XElement shape = box.Ancestors().FirstOrDefault(a => a.Name.LocalName == "wsp");
                XElement bodyPr = shape == null ? null : shape.Elements().FirstOrDefault(e => e.Name.LocalName == "bodyPr");
                actual.Add(new KeyValuePair<string, string>(text, bodyPr == null ? null : (string)bodyPr.Attribute("vert")));
            }
            foreach (PdfPageText page in pages)
                foreach (OcrRotatedText t in page.RotatedTexts)
                {
                    string text = CorpusCheck.Norm(t.Text);
                    int i = actual.FindIndex(a => a.Key == text);
                    if (i < 0)
                    {
                        findings.Add(new CFinding("word-rotated-lost", "надпись «" + Head(text) + "» не найдена в .docx"));
                        continue;
                    }
                    // DrawingML: vert270 — снизу вверх (наш угол 90°), vert — сверху вниз (270°);
                    // прочие углы Word через COM не ставит — надпись горизонтальная.
                    double angle = RotatedText.Normalize(t.AngleDeg);
                    string want = Math.Abs(angle - 90) <= 2 ? "vert270" : (Math.Abs(angle - 270) <= 2 ? "vert" : null);
                    string got = actual[i].Value == "horz" ? null : actual[i].Value;
                    if (got != want)
                        findings.Add(new CFinding("word-rotated-angle", "надпись «" + Head(text) + "»: vert=" + (got ?? "—") + " вместо " + (want ?? "—")));
                    actual.RemoveAt(i);
                }
            foreach (KeyValuePair<string, string> extra in actual)
                findings.Add(new CFinding("word-rotated-extra", "лишняя надпись «" + Head(extra.Key) + "»"));
        }

        private static string Head(string s)
        {
            return s.Length > 50 ? s.Substring(0, 50) + "…" : s;
        }

        private static void ComparePara(WPara e, WPara a, List<CFinding> findings)
        {
            string name = "«" + Head(e.Text) + "»";
            if (e.Align != a.Align)
                findings.Add(new CFinding("word-align", name + ": " + a.Align + " вместо " + e.Align));
            if (e.List != a.List)
                findings.Add(new CFinding("word-list", name + ": список " + a.List + " вместо " + e.List));
            else if (e.List != ListKind.None && e.Level != a.Level)
                findings.Add(new CFinding("word-list-level", name + ": уровень " + a.Level + " вместо " + e.Level));
            if (e.Words.Count != a.Words.Count)
                return; // текст совпал после нормализации — счёт слов тот же; иначе формат не сопоставить
            for (int i = 0; i < e.Words.Count; i++)
            {
                CorpusCheck.Tok x = e.Words[i], y = a.Words[i];
                string w = "«" + x.Text + "»";
                if (x.Bold != y.Bold) findings.Add(new CFinding("word-bold", w + (x.Bold ? " потерял полужирный" : " стал полужирным")));
                if (x.Italic != y.Italic) findings.Add(new CFinding("word-italic", w + (x.Italic ? " потерял курсив" : " стал курсивом")));
                if (x.Underline != y.Underline) findings.Add(new CFinding("word-underline", w + (x.Underline ? " потерял подчёркивание" : " стал подчёркнутым")));
                if (x.Color != y.Color) findings.Add(new CFinding("word-color", w + " — цвет " + y.Color.ToString("X6") + " вместо " + x.Color.ToString("X6")));
                if (Math.Abs(x.Size - y.Size) > 0.01) findings.Add(new CFinding("word-size", w + " — кегль " + y.Size.ToString(CultureInfo.InvariantCulture) + " вместо " + x.Size.ToString(CultureInfo.InvariantCulture)));
                if (CorpusCheck.FontKey(x.Font) != CorpusCheck.FontKey(y.Font)) findings.Add(new CFinding("word-font", w + " — шрифт «" + y.Font + "» вместо «" + x.Font + "»"));
            }
        }
    }
}
