using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Чистая логика сравнения: нормализация, сопоставление страниц, текстовый diff,
    /// статистика и сворачивание длинных неизменённых фрагментов. Ни PDF, ни UI здесь нет.
    /// </summary>
    internal static class PdfReviewDiff
    {
        private const double SimilarityThreshold = 0.24;
        private const double GapCost = 0.58;

        public static string Normalize(string text)
        {
            if (string.IsNullOrEmpty(text))
                return "";
            string source = text.Normalize(NormalizationForm.FormC)
                .Replace("\r\n", "\n").Replace('\r', '\n').Replace(' ', ' ');
            var result = new StringBuilder(source.Length);
            int lineStart = 0;
            bool spaces = false;
            for (int i = 0; i < source.Length; i++)
            {
                char c = source[i];
                if (c == '\n')
                {
                    while (result.Length > lineStart && result[result.Length - 1] == ' ')
                        result.Length--;
                    result.Append('\n');
                    lineStart = result.Length;
                    spaces = false;
                }
                else if (c == '\t')
                {
                    result.Append('\t');
                    spaces = false;
                }
                else if (c == ' ' || c == '\v')
                {
                    if (!spaces)
                        result.Append(' ');
                    spaces = true;
                }
                else
                {
                    result.Append(c);
                    spaces = false;
                }
            }
            while (result.Length > lineStart && result[result.Length - 1] == ' ')
                result.Length--;
            return result.ToString();
        }

        public static string Fingerprint(string normalized, double widthPt, double heightPt)
        {
            string text = normalized ?? "";
            ulong hash = 1469598103934665603UL;
            byte[] bytes = Encoding.UTF8.GetBytes(text);
            foreach (byte b in bytes)
            {
                hash ^= b;
                hash *= 1099511628211UL;
            }
            return hash.ToString("X16", CultureInfo.InvariantCulture) + ":" +
                Math.Round(widthPt, 1).ToString(CultureInfo.InvariantCulture) + "x" +
                Math.Round(heightPt, 1).ToString(CultureInfo.InvariantCulture);
        }

        public static PdfReviewResult Compare(PdfReviewDocument left, PdfReviewDocument right,
            PdfReviewLimits limits)
        {
            if (left == null) throw new ArgumentNullException("left");
            if (right == null) throw new ArgumentNullException("right");
            if (limits == null) limits = PdfReviewLimits.Default();
            var result = new PdfReviewResult { Left = left, Right = right };
            foreach (PdfReviewPagePair pair in Align(left.Pages, right.Pages, limits))
                result.Pairs.Add(pair);
            result.Stats = Statistics(result);
            return result;
        }

        /// <summary>
        /// Глобальное выравнивание последовательностей: точное совпадение почти бесплатно,
        /// похожие соседние страницы спариваются, непохожие становятся удалением/вставкой.
        /// Переставленная далеко страница не угадывается молча как прежнее место.
        /// </summary>
        public static List<PdfReviewPagePair> Align(IList<PdfReviewPage> left,
            IList<PdfReviewPage> right, PdfReviewLimits limits)
        {
            int n = left == null ? 0 : left.Count;
            int m = right == null ? 0 : right.Count;
            if ((long)(n + 1) * (m + 1) > limits.MaxDiffCells)
                throw new PdfReviewException(PdfReviewFailure.TooLarge, null,
                    Loc.T("review.err.tooLarge"));

            var cost = new double[n + 1, m + 1];
            var move = new byte[n + 1, m + 1]; // 1 pair, 2 left-only, 3 right-only
            // Множества слов каждой страницы считаем ОДИН раз, а не в каждой из «n·m»
            // ячеек выравнивания: иначе два документа по 500 страниц делали бы
            // полмиллиона лишних токенизаций одного и того же текста.
            var leftSets = new HashSet<string>[n];
            for (int i = 0; i < n; i++)
                leftSets[i] = WordSet(left[i].NormalizedText);
            var rightSets = new HashSet<string>[m];
            for (int j = 0; j < m; j++)
                rightSets[j] = WordSet(right[j].NormalizedText);
            for (int i = n - 1; i >= 0; i--)
            {
                cost[i, m] = cost[i + 1, m] + GapCost;
                move[i, m] = 2;
            }
            for (int j = m - 1; j >= 0; j--)
            {
                cost[n, j] = cost[n, j + 1] + GapCost;
                move[n, j] = 3;
            }
            for (int i = n - 1; i >= 0; i--)
            {
                for (int j = m - 1; j >= 0; j--)
                {
                    double sim = Similarity(left[i], right[j], leftSets[i], rightSets[j]);
                    double pairCost = sim >= SimilarityThreshold
                        ? cost[i + 1, j + 1] + (1.0 - sim)
                        : double.MaxValue;
                    double dropLeft = cost[i + 1, j] + GapCost;
                    double dropRight = cost[i, j + 1] + GapCost;
                    if (pairCost <= dropLeft && pairCost <= dropRight)
                    {
                        cost[i, j] = pairCost;
                        move[i, j] = 1;
                    }
                    else if (dropLeft <= dropRight)
                    {
                        cost[i, j] = dropLeft;
                        move[i, j] = 2;
                    }
                    else
                    {
                        cost[i, j] = dropRight;
                        move[i, j] = 3;
                    }
                }
            }

            var result = new List<PdfReviewPagePair>();
            int li = 0, ri = 0;
            while (li < n || ri < m)
            {
                byte step = move[li, ri];
                if (step == 1)
                {
                    result.Add(Pair(left[li], right[ri], limits));
                    li++; ri++;
                }
                else if (step == 2 || ri >= m)
                {
                    result.Add(new PdfReviewPagePair
                    {
                        LeftPageIndex = left[li].PageIndex,
                        RightPageIndex = -1,
                        Status = PdfReviewPairStatus.LeftOnly,
                        Similarity = 0
                    });
                    li++;
                }
                else
                {
                    result.Add(new PdfReviewPagePair
                    {
                        LeftPageIndex = -1,
                        RightPageIndex = right[ri].PageIndex,
                        Status = PdfReviewPairStatus.RightOnly,
                        Similarity = 0
                    });
                    ri++;
                }
            }
            return result;
        }

        public static PdfReviewPagePair Pair(PdfReviewPage left, PdfReviewPage right,
            PdfReviewLimits limits)
        {
            var pair = new PdfReviewPagePair
            {
                LeftPageIndex = left == null ? -1 : left.PageIndex,
                RightPageIndex = right == null ? -1 : right.PageIndex
            };
            if (left == null)
            {
                pair.Status = PdfReviewPairStatus.RightOnly;
                return pair;
            }
            if (right == null)
            {
                pair.Status = PdfReviewPairStatus.LeftOnly;
                return pair;
            }
            pair.Similarity = Similarity(left, right);
            if (string.Equals(left.NormalizedText, right.NormalizedText, StringComparison.Ordinal))
            {
                pair.Status = PdfReviewPairStatus.Unchanged;
                pair.Operations.Add(new PdfReviewDiffOp
                {
                    Kind = PdfReviewDiffKind.Equal,
                    Text = left.NormalizedText ?? ""
                });
                return pair;
            }
            pair.Status = PdfReviewPairStatus.Changed;
            foreach (PdfReviewDiffOp op in DiffText(left.NormalizedText, right.NormalizedText, limits))
                pair.Operations.Add(op);
            return pair;
        }

        /// <summary>
        /// Ручная пара заменяет прежние пары обеих страниц и встаёт на более раннее из их мест.
        /// Остальной порядок не меняется. one-to-one гарантируется удалением прежних владельцев.
        /// </summary>
        public static List<PdfReviewPagePair> ApplyManualPair(IList<PdfReviewPagePair> pairs,
            PdfReviewDocument left, PdfReviewDocument right, int leftPageIndex, int rightPageIndex,
            PdfReviewLimits limits)
        {
            PdfReviewPage lp = PageAt(left, leftPageIndex);
            PdfReviewPage rp = PageAt(right, rightPageIndex);
            if (lp == null || rp == null)
                throw new ArgumentOutOfRangeException();
            var result = new List<PdfReviewPagePair>();
            int insertAt = pairs == null ? 0 : pairs.Count;
            if (pairs != null)
            {
                for (int i = 0; i < pairs.Count; i++)
                {
                    PdfReviewPagePair p = pairs[i];
                    if (p.LeftPageIndex == leftPageIndex || p.RightPageIndex == rightPageIndex)
                    {
                        if (result.Count < insertAt) insertAt = result.Count;
                        continue;
                    }
                    result.Add(p);
                }
            }
            if (insertAt < 0 || insertAt > result.Count) insertAt = result.Count;
            result.Insert(insertAt, Pair(lp, rp, limits));
            return result;
        }

        public static List<PdfReviewDiffOp> DiffText(string left, string right, PdfReviewLimits limits)
        {
            string a = left ?? "", b = right ?? "";
            if (a == b)
                return new List<PdfReviewDiffOp>
                {
                    new PdfReviewDiffOp { Kind = PdfReviewDiffKind.Equal, Text = a }
                };
            List<string> at = Tokens(a);
            List<string> bt = Tokens(b);
            if (at.Count + bt.Count > limits.MaxTokens ||
                (long)(at.Count + 1) * (bt.Count + 1) > limits.MaxDiffCells)
                return WholeChange(a, b);
            return DiffTokens(at, bt);
        }

        private static List<PdfReviewDiffOp> DiffTokens(IList<string> left, IList<string> right)
        {
            int n = left.Count, m = right.Count;
            var dp = new int[n + 1, m + 1];
            for (int i = n - 1; i >= 0; i--)
                for (int j = m - 1; j >= 0; j--)
                    dp[i, j] = left[i] == right[j]
                        ? dp[i + 1, j + 1] + 1
                        : Math.Max(dp[i + 1, j], dp[i, j + 1]);
            var ops = new List<PdfReviewDiffOp>();
            int li = 0, ri = 0;
            while (li < n || ri < m)
            {
                if (li < n && ri < m && left[li] == right[ri])
                {
                    Add(ops, PdfReviewDiffKind.Equal, left[li]); li++; ri++;
                }
                else if (ri >= m || (li < n && dp[li + 1, ri] >= dp[li, ri + 1]))
                {
                    Add(ops, PdfReviewDiffKind.Delete, left[li]); li++;
                }
                else
                {
                    Add(ops, PdfReviewDiffKind.Insert, right[ri]); ri++;
                }
            }
            return ops;
        }

        public static List<PdfReviewDiffOp> Collapse(IList<PdfReviewDiffOp> operations,
            int threshold = 800, int keep = 220)
        {
            var result = new List<PdfReviewDiffOp>();
            if (operations == null) return result;
            foreach (PdfReviewDiffOp op in operations)
            {
                string text = op.Text ?? "";
                if (op.Kind != PdfReviewDiffKind.Equal || text.Length <= threshold || keep < 1)
                {
                    result.Add(Copy(op));
                    continue;
                }
                int lead = Math.Min(keep, text.Length / 2);
                int tail = Math.Min(keep, text.Length - lead);
                int hidden = text.Length - lead - tail;
                result.Add(new PdfReviewDiffOp { Kind = PdfReviewDiffKind.Equal, Text = text.Substring(0, lead) });
                result.Add(new PdfReviewDiffOp
                {
                    Kind = PdfReviewDiffKind.Equal,
                    Text = "\n… " + hidden.ToString(CultureInfo.InvariantCulture) + " …\n",
                    Collapsed = true,
                    HiddenCharacters = hidden
                });
                result.Add(new PdfReviewDiffOp { Kind = PdfReviewDiffKind.Equal, Text = text.Substring(text.Length - tail) });
            }
            return result;
        }

        public static PdfReviewStats Statistics(PdfReviewResult result)
        {
            var stats = new PdfReviewStats();
            if (result == null) return stats;
            stats.PagePairs = result.Pairs.Count;
            int totalChars = (result.Left == null ? 0 : result.Left.CharacterCount) +
                             (result.Right == null ? 0 : result.Right.CharacterCount);
            foreach (PdfReviewPagePair pair in result.Pairs)
            {
                if (pair.Status == PdfReviewPairStatus.Changed) stats.ChangedPages++;
                if (pair.Status == PdfReviewPairStatus.LeftOnly) stats.LeftOnlyPages++;
                if (pair.Status == PdfReviewPairStatus.RightOnly) stats.RightOnlyPages++;
                int deletedWords = 0, insertedWords = 0;
                foreach (PdfReviewDiffOp op in pair.Operations)
                {
                    if (op.Kind == PdfReviewDiffKind.Delete)
                    {
                        stats.DeletedCharacters += CodePoints(op.Text);
                        deletedWords += WordCount(op.Text);
                    }
                    else if (op.Kind == PdfReviewDiffKind.Insert)
                    {
                        stats.InsertedCharacters += CodePoints(op.Text);
                        insertedWords += WordCount(op.Text);
                    }
                }
                stats.DeletedWords += deletedWords;
                stats.InsertedWords += insertedWords;
                stats.Replacements += Math.Min(deletedWords, insertedWords);
            }
            int changed = stats.DeletedCharacters + stats.InsertedCharacters;
            stats.ChangedPercent = totalChars <= 0 ? 0 :
                Math.Min(100, Math.Max(changed > 0 ? 1 : 0,
                    (int)Math.Round(100.0 * changed / totalChars)));
            return stats;
        }

        /// <summary>
        /// Похожесть двух страниц: 0.9 — пересечение слов, 0.1 — совпадение размеров.
        /// Одиночные вызовы (пара уже выбрана) считают множества сами; выравнивание
        /// подаёт предвычисленные, чтобы не токенизировать страницы в каждой ячейке.
        /// </summary>
        private static double Similarity(PdfReviewPage a, PdfReviewPage b)
        {
            if (a == null || b == null) return 0;
            return Similarity(a, b, WordSet(a.NormalizedText), WordSet(b.NormalizedText));
        }

        private static double Similarity(PdfReviewPage a, PdfReviewPage b,
            HashSet<string> aw, HashSet<string> bw)
        {
            if (a == null || b == null) return 0;
            if (string.Equals(a.NormalizedText, b.NormalizedText, StringComparison.Ordinal)) return 1;
            int common = 0;
            foreach (string word in aw)
                if (bw.Contains(word)) common++;
            int union = aw.Count + bw.Count - common;
            double words = union == 0 ? 0 : (double)common / union;
            double dw = Math.Abs(a.WidthPt - b.WidthPt), dh = Math.Abs(a.HeightPt - b.HeightPt);
            double dims = dw <= Math.Max(4, 0.02 * Math.Max(a.WidthPt, b.WidthPt)) &&
                          dh <= Math.Max(4, 0.02 * Math.Max(a.HeightPt, b.HeightPt)) ? 1 : 0;
            return 0.9 * words + 0.1 * dims;
        }

        private static HashSet<string> WordSet(string text)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string token in Tokens(text ?? ""))
                if (IsWord(token)) set.Add(token);
            return set;
        }

        private static List<string> Tokens(string text)
        {
            var result = new List<string>();
            if (string.IsNullOrEmpty(text)) return result;
            var current = new StringBuilder();
            int kind = -1;
            for (int i = 0; i < text.Length;)
            {
                int length = char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]) ? 2 : 1;
                string unit = text.Substring(i, length);
                UnicodeCategory category = char.GetUnicodeCategory(text, i);
                int nextKind = char.IsWhiteSpace(text, i) ? 0 :
                    category == UnicodeCategory.UppercaseLetter || category == UnicodeCategory.LowercaseLetter ||
                    category == UnicodeCategory.TitlecaseLetter || category == UnicodeCategory.OtherLetter ||
                    category == UnicodeCategory.DecimalDigitNumber ? 1 : 2;
                if (kind != nextKind || nextKind == 2)
                {
                    if (current.Length > 0) { result.Add(current.ToString()); current.Length = 0; }
                    kind = nextKind;
                }
                current.Append(unit);
                if (nextKind == 2) { result.Add(current.ToString()); current.Length = 0; kind = -1; }
                i += length;
            }
            if (current.Length > 0) result.Add(current.ToString());
            return result;
        }

        private static bool IsWord(string token)
        {
            return !string.IsNullOrEmpty(token) && char.IsLetterOrDigit(token, 0);
        }

        private static int WordCount(string text)
        {
            int count = 0;
            foreach (string token in Tokens(text ?? ""))
                if (IsWord(token)) count++;
            return count;
        }

        /// <summary>
        /// Число кодовых точек (не char-ов): суррогатная пара считается одним символом.
        /// Единая точка подсчёта для сервиса (лимит знаков) и диффа (статистика изменений).
        /// </summary>
        internal static int CodePoints(string text)
        {
            if (string.IsNullOrEmpty(text)) return 0;
            int count = 0;
            for (int i = 0; i < text.Length; i++, count++)
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1])) i++;
            return count;
        }

        private static void Add(List<PdfReviewDiffOp> ops, PdfReviewDiffKind kind, string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            if (ops.Count > 0 && ops[ops.Count - 1].Kind == kind && !ops[ops.Count - 1].Collapsed)
                ops[ops.Count - 1].Text += text;
            else
                ops.Add(new PdfReviewDiffOp { Kind = kind, Text = text });
        }

        private static List<PdfReviewDiffOp> WholeChange(string left, string right)
        {
            var result = new List<PdfReviewDiffOp>();
            Add(result, PdfReviewDiffKind.Delete, left);
            Add(result, PdfReviewDiffKind.Insert, right);
            return result;
        }

        private static PdfReviewDiffOp Copy(PdfReviewDiffOp op)
        {
            return new PdfReviewDiffOp
            {
                Kind = op.Kind,
                Text = op.Text,
                Collapsed = op.Collapsed,
                HiddenCharacters = op.HiddenCharacters
            };
        }

        private static PdfReviewPage PageAt(PdfReviewDocument doc, int pageIndex)
        {
            if (doc == null) return null;
            foreach (PdfReviewPage page in doc.Pages)
                if (page.PageIndex == pageIndex) return page;
            return null;
        }
    }
}
