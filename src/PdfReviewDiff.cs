using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Чистая логика сравнения: нормализация, сопоставление страниц, ворд-дифф,
    /// статистика. Ни PDF, ни UI здесь нет.
    ///
    /// Дифф идёт по СЛОВАМ страницы в порядке чтения, а не по сырому тексту: перепад
    /// переносов строк или пробелов при том же тексте не даёт ложных правок, а страница
    /// в тысячи слов не упирается в потолок матрицы (старая посимвольная версия на любой
    /// полной странице включала «изменено всё»). Операции ворд-диффа — единственный
    /// источник и подсветки, и счётчиков: что нарисовано, то и посчитано.
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
                .Replace("\r\n", "\n").Replace('\r', '\n').Replace('\u00A0', ' ');
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
                leftSets[i] = WordSet(left[i]);
            var rightSets = new HashSet<string>[m];
            for (int j = 0; j < m; j++)
                rightSets[j] = WordSet(right[j]);
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

        /// <summary>
        /// Пара страниц: статус и ворд-дифф. Идентичные последовательности слов — один
        /// Equal без всякой матрицы; изменённые — точные операции по словам.
        /// </summary>
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
            if (WordsEqual(left.Words, right.Words))
            {
                pair.Status = PdfReviewPairStatus.Unchanged;
                var equal = new PdfReviewWordOp { Kind = PdfReviewDiffKind.Equal };
                equal.Words.AddRange(left.Words);
                pair.Operations.Add(equal);
                return pair;
            }
            pair.Status = PdfReviewPairStatus.Changed;
            foreach (PdfReviewWordOp op in DiffWords(left.Words, right.Words, limits))
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

        /// <summary>Одинаковы ли две последовательности слов (по ключам).</summary>
        internal static bool WordsEqual(IList<PdfReviewWord> left, IList<PdfReviewWord> right)
        {
            int n = left == null ? 0 : left.Count;
            int m = right == null ? 0 : right.Count;
            if (n != m) return false;
            for (int i = 0; i < n; i++)
                if (WordKey(left[i]) != WordKey(right[i]))
                    return false;
            return true;
        }

        /// <summary>Ключ сравнения слова: текст без обрамляющих пробелов.</summary>
        internal static string WordKey(PdfReviewWord word)
        {
            return word == null || word.Key == null ? "" : word.Key;
        }

        /// <summary>
        /// Ворд-дифф: общий префикс/суффикс — сразу Equal, середина — ЛСД-матрица, если она
        /// помещается в потолок; не помещается — середина делится пополам и дифф продолжается
        /// рекурсивно (граница раздела может слегка сместить правку, но «изменено всё» не
        /// получится никогда). Пустые входы дают один пустой Equal.
        /// </summary>
        public static List<PdfReviewWordOp> DiffWords(IList<PdfReviewWord> left,
            IList<PdfReviewWord> right, PdfReviewLimits limits)
        {
            var ops = new List<PdfReviewWordOp>();
            DiffRange(left ?? EmptyWords, 0, left == null ? 0 : left.Count,
                right ?? EmptyWords, 0, right == null ? 0 : right.Count, limits, ops);
            if (ops.Count == 0)
                ops.Add(new PdfReviewWordOp { Kind = PdfReviewDiffKind.Equal });
            return ops;
        }

        private static readonly PdfReviewWord[] EmptyWords = new PdfReviewWord[0];

        private static void DiffRange(IList<PdfReviewWord> a, int aStart, int aEnd,
            IList<PdfReviewWord> b, int bStart, int bEnd, PdfReviewLimits limits,
            List<PdfReviewWordOp> ops)
        {
            // Общий префикс: одинаковые слова с начала — общее без всякой матрицы.
            while (aStart < aEnd && bStart < bEnd && WordKey(a[aStart]) == WordKey(b[bStart]))
            {
                Append(ops, PdfReviewDiffKind.Equal, a[aStart]);
                aStart++; bStart++;
            }
            // Общий суффикс: запоминаем, выводим ПОСЛЕ середины.
            int suffix = 0;
            while (aEnd > aStart && bEnd > bStart && WordKey(a[aEnd - 1]) == WordKey(b[bEnd - 1]))
            {
                suffix++;
                aEnd--; bEnd--;
            }

            int n = aEnd - aStart, m = bEnd - bStart;
            if (n == 0 && m == 0)
            {
                // середина пуста
            }
            else if (n == 0)
            {
                for (int j = bStart; j < bEnd; j++) Append(ops, PdfReviewDiffKind.Insert, b[j]);
            }
            else if (m == 0)
            {
                for (int i = aStart; i < aEnd; i++) Append(ops, PdfReviewDiffKind.Delete, a[i]);
            }
            else if (n == 1)
            {
                // Одно слово слева против длинной правой стороны: делить пополам нечего
                // (рекурсия не уменьшала бы вход) — линейный поиск слова в правой
                // последовательности, O(m).
                int j = bStart;
                while (j < bEnd && WordKey(b[j]) != WordKey(a[aStart])) j++;
                for (int k = bStart; k < j; k++) Append(ops, PdfReviewDiffKind.Insert, b[k]);
                if (j < bEnd)
                {
                    Append(ops, PdfReviewDiffKind.Equal, a[aStart]);
                    for (int k = j + 1; k < bEnd; k++) Append(ops, PdfReviewDiffKind.Insert, b[k]);
                }
                else
                {
                    Append(ops, PdfReviewDiffKind.Delete, a[aStart]);
                }
            }
            else if (m == 1)
            {
                // Зеркальный случай: одно слово справа.
                int i = aStart;
                while (i < aEnd && WordKey(a[i]) != WordKey(b[bStart])) i++;
                for (int k = aStart; k < i; k++) Append(ops, PdfReviewDiffKind.Delete, a[k]);
                if (i < aEnd)
                {
                    Append(ops, PdfReviewDiffKind.Equal, b[bStart]);
                    for (int k = i + 1; k < aEnd; k++) Append(ops, PdfReviewDiffKind.Delete, a[k]);
                }
                else
                {
                    Append(ops, PdfReviewDiffKind.Insert, b[bStart]);
                }
            }
            else if ((long)(n + 1) * (m + 1) <= limits.MaxDiffCells)
            {
                LcsCore(a, aStart, aEnd, b, bStart, bEnd, ops);
            }
            else
            {
                // Потолок матрицы: делим обе середины пополам (право пропорционально длине
                // левой) и диффим половины независимо. Рекурсия конечна: при n, m >= 2
                // каждая половина строго меньше входа.
                int aMid = aStart + n / 2;
                int bMid = bStart + (int)((long)m * (aMid - aStart) / n);
                DiffRange(a, aStart, aMid, b, bStart, bMid, limits, ops);
                DiffRange(a, aMid, aEnd, b, bMid, bEnd, limits, ops);
            }

            for (int s = 0; s < suffix; s++)
                Append(ops, PdfReviewDiffKind.Equal, a[aEnd + s]);
        }

        /// <summary>Точный ЛСД-дифф середины, которая помещается в потолок матрицы.</summary>
        private static void LcsCore(IList<PdfReviewWord> a, int aStart, int aEnd,
            IList<PdfReviewWord> b, int bStart, int bEnd, List<PdfReviewWordOp> ops)
        {
            int n = aEnd - aStart, m = bEnd - bStart;
            var dp = new int[n + 1, m + 1];
            for (int i = n - 1; i >= 0; i--)
                for (int j = m - 1; j >= 0; j--)
                    dp[i, j] = WordKey(a[aStart + i]) == WordKey(b[bStart + j])
                        ? dp[i + 1, j + 1] + 1
                        : Math.Max(dp[i + 1, j], dp[i, j + 1]);
            int li = 0, ri = 0;
            while (li < n || ri < m)
            {
                if (li < n && ri < m && WordKey(a[aStart + li]) == WordKey(b[bStart + ri]))
                {
                    Append(ops, PdfReviewDiffKind.Equal, a[aStart + li]); li++; ri++;
                }
                else if (ri >= m || (li < n && dp[li + 1, ri] >= dp[li, ri + 1]))
                {
                    Append(ops, PdfReviewDiffKind.Delete, a[aStart + li]); li++;
                }
                else
                {
                    Append(ops, PdfReviewDiffKind.Insert, b[bStart + ri]); ri++;
                }
            }
        }

        /// <summary>Дописывает слово в последнюю операцию того же вида или начинает новую.</summary>
        private static void Append(List<PdfReviewWordOp> ops, PdfReviewDiffKind kind, PdfReviewWord word)
        {
            if (ops.Count > 0 && ops[ops.Count - 1].Kind == kind)
            {
                ops[ops.Count - 1].Words.Add(word);
                return;
            }
            var op = new PdfReviewWordOp { Kind = kind };
            op.Words.Add(word);
            ops.Add(op);
        }

        /// <summary>
        /// Счётчики — из тех же ворд-операций, что и подсветка (один источник правды).
        /// Страницы без пары дают все свои слова как удалённые/добавленные.
        /// </summary>
        public static PdfReviewStats Statistics(PdfReviewResult result)
        {
            var stats = new PdfReviewStats();
            if (result == null) return stats;
            stats.PagePairs = result.Pairs.Count;
            foreach (PdfReviewPagePair pair in result.Pairs)
            {
                if (pair.Status == PdfReviewPairStatus.Changed) stats.ChangedPages++;
                if (pair.Status == PdfReviewPairStatus.LeftOnly)
                {
                    stats.LeftOnlyPages++;
                    stats.DeletedWords += PageWords(result.Left, pair.LeftPageIndex);
                }
                if (pair.Status == PdfReviewPairStatus.RightOnly)
                {
                    stats.RightOnlyPages++;
                    stats.InsertedWords += PageWords(result.Right, pair.RightPageIndex);
                }
                int deleted = 0, inserted = 0;
                foreach (PdfReviewWordOp op in pair.Operations)
                {
                    if (op.Kind == PdfReviewDiffKind.Delete) deleted += op.Words.Count;
                    else if (op.Kind == PdfReviewDiffKind.Insert) inserted += op.Words.Count;
                }
                stats.DeletedWords += deleted;
                stats.InsertedWords += inserted;
                stats.Replacements += Math.Min(deleted, inserted);
            }
            int total = (result.Left == null ? 0 : result.Left.WordCount) +
                        (result.Right == null ? 0 : result.Right.WordCount);
            int changed = stats.DeletedWords + stats.InsertedWords;
            stats.ChangedPercent = total <= 0 ? 0 :
                Math.Min(100, Math.Max(changed > 0 ? 1 : 0,
                    (int)Math.Round(100.0 * changed / total)));
            return stats;
        }

        private static int PageWords(PdfReviewDocument doc, int pageIndex)
        {
            PdfReviewPage page = PageAt(doc, pageIndex);
            return page == null ? 0 : page.Words.Count;
        }

        /// <summary>
        /// Похожесть двух страниц: 0.9 — пересечение слов, 0.1 — совпадение размеров.
        /// Слова берутся из готового списка страницы; пустой список (страницы, собранные
        /// вне сервиса) токенизируется из нормализованного текста.
        /// </summary>
        private static double Similarity(PdfReviewPage a, PdfReviewPage b)
        {
            if (a == null || b == null) return 0;
            return Similarity(a, b, WordSet(a), WordSet(b));
        }

        private static double Similarity(PdfReviewPage a, PdfReviewPage b,
            HashSet<string> aw, HashSet<string> bw)
        {
            if (a == null || b == null) return 0;
            if (WordsEqual(a.Words, b.Words)) return 1;
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

        private static HashSet<string> WordSet(PdfReviewPage page)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (page == null) return set;
            if (page.Words.Count > 0)
            {
                foreach (PdfReviewWord word in page.Words)
                    set.Add(WordKey(word));
                return set;
            }
            foreach (string token in Tokens(page.NormalizedText))
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

        /// <summary>
        /// Число кодовых точек (не char-ов): суррогатная пара считается одним символом.
        /// Единая точка подсчёта для сервиса (лимит знаков).
        /// </summary>
        internal static int CodePoints(string text)
        {
            if (string.IsNullOrEmpty(text)) return 0;
            int count = 0;
            for (int i = 0; i < text.Length; i++, count++)
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1])) i++;
            return count;
        }

        internal static PdfReviewPage PageAt(PdfReviewDocument doc, int pageIndex)
        {
            if (doc == null) return null;
            foreach (PdfReviewPage page in doc.Pages)
                if (page.PageIndex == pageIndex) return page;
            return null;
        }
    }
}
