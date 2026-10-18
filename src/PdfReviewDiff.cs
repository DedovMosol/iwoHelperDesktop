using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Чистая логика сравнения: document-wide ворд-дифф, сопоставление физических страниц
    /// для viewer и обратная проекция кандидатов. Ни PDF, ни UI здесь нет.
    ///
    /// Семантика строится один раз по сквозным последовательностям канонических видимых слов.
    /// Физическая граница страницы не участвует в ключе diff: каждое слово лишь сохраняет
    /// страницу-владельца, чтобы Delete/Insert можно было подсветить в исходном документе.
    /// </summary>
    internal static class PdfReviewDiff
    {
        private const double SimilarityThreshold = 0.24;
        private const double GapCost = 0.58;
        private const double AlignmentTieEpsilon = 1e-12;
        private const int CancellationPollMask = 255;
        private const int MaxDiffRecursionDepth = 128;
        // Лист точной LCS-матрицы (~256×256 слов). Матрица оплачивает n·m при любом числе
        // правок, а Myers — O((N+M)·D); большие диапазоны делит бисекция, а матрица
        // остаётся там, где её вторичное предпочтение страниц дёшево.
        private const int MaxMatrixLeafCells = 65536;
        // Знаменатель доли бюджета Compare (1/8), недоступной словесным стадиям и
        // гарантированной сравнению пробелов.
        private const int WhitespaceWorkShare = 8;
        // Знаменатель доли остатка (1/2) на одну бисекцию Myers. При равномерных правках
        // верхняя бисекция стоит около половины всей рекурсии (~D² из ~2·D²/2), поэтому
        // если ей мало половины, точный diff не уложился бы и целиком; её провал не
        // должен обнулять бюджет — диапазон делится якорями (DiffAnchoredRange).
        private const int BisectWorkShare = 2;

        /// <summary>
        /// Контекст принадлежит одному Compare/DiffWords: никакого глобального состояния,
        /// wall-clock timeout или гонки между сравнениями. Work расходуется на необязательное
        /// семантическое доказательство, которое может разрастись; обязательные линейные проходы
        /// извлечения и проекции лишь опрашивают отмену.
        /// </summary>
        private sealed class DiffContext
        {
            private readonly Func<bool> _cancelled;
            private readonly long _initialWork;
            private long _remainingWork;
            private int _pollCounter;

            public DiffContext(PdfReviewLimits limits, Func<bool> cancelled)
                : this((limits ?? PdfReviewLimits.Default()).MaxDiffWork, cancelled)
            {
            }

            public DiffContext(long work, Func<bool> cancelled)
            {
                _cancelled = cancelled;
                _remainingWork = Math.Max(0L, work);
                _initialWork = _remainingWork;
            }

            public bool WorkExhausted { get; private set; }

            /// <summary>
            /// Часть diff разбита эвристическими якорями, потому что точной бисекции не
            /// хватило её доли: результат полон, но минимальность не доказана. В отличие
            /// от WorkExhausted остаток бюджета сохраняется для остальных диапазонов.
            /// </summary>
            public bool Approximated { get; private set; }

            public void MarkApproximated()
            {
                Approximated = true;
            }

            /// <summary>
            /// Дочерний бюджет одной попытки: её провал исчерпывает только выданную долю,
            /// а не весь остаток. Израсходованное списывает с родителя Settle.
            /// </summary>
            public DiffContext Slice(long units)
            {
                return new DiffContext(Math.Min(units, _remainingWork), _cancelled);
            }

            public void Settle(DiffContext slice)
            {
                _remainingWork -= slice._initialWork - slice._remainingWork;
            }

            /// <summary>Неизрасходованный остаток; после исчерпания он равен нулю.</summary>
            public long RemainingWork { get { return _remainingWork; } }

            public void ThrowIfCancellation()
            {
                Cancellation.ThrowIf(_cancelled);
            }

            public void PollCancellation()
            {
                unchecked { _pollCounter++; }
                if ((_pollCounter & CancellationPollMask) == 0)
                    Cancellation.ThrowIf(_cancelled);
            }

            public bool TryReserve(long units)
            {
                PollCancellation();
                if (WorkExhausted)
                    return false;
                if (units < 0)
                {
                    // Отрицательная «работа» означает ошибку в расчёте, а не бесплатный
                    // semantic proof. Консервативно закрываем все необязательные стадии.
                    ExhaustWork();
                    return false;
                }
                if (units == 0)
                    return true;
                if (units > _remainingWork)
                {
                    WorkExhausted = true;
                    _remainingWork = 0;
                    return false;
                }
                _remainingWork -= units;
                return true;
            }

            public bool TryStep()
            {
                return TryReserve(1);
            }

            public bool TryReserveSort(int count)
            {
                if (count <= 1)
                    return TryReserve(count);
                int levels = 0;
                int remaining = count - 1;
                while (remaining > 0)
                {
                    levels++;
                    remaining >>= 1;
                }
                return TryReserve((long)count * (levels + 1));
            }

            public void ExhaustWork()
            {
                WorkExhausted = true;
                _remainingWork = 0;
            }
        }

        /// <summary>
        /// Владение слов страницей применяется одной короткой транзакцией после полного
        /// обхода ОБОИХ документов. До commit отмена/ошибка не меняет модель вызывающего;
        /// после commit любой незавершённый Compare восстанавливает точные исходные индексы.
        /// </summary>
        private sealed class PageOwnershipTransaction
        {
            private sealed class Entry
            {
                public PdfReviewWord Word;
                public PdfReviewPage Page;
                public int OriginalPageIndex;
                public int TargetPageIndex;
            }

            private sealed class WordReferenceComparer : IEqualityComparer<PdfReviewWord>
            {
                public bool Equals(PdfReviewWord left, PdfReviewWord right)
                {
                    return ReferenceEquals(left, right);
                }

                public int GetHashCode(PdfReviewWord word)
                {
                    return word == null ? 0 :
                        System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(word);
                }
            }

            private readonly Dictionary<PdfReviewWord, Entry> _byWord =
                new Dictionary<PdfReviewWord, Entry>(new WordReferenceComparer());
            private readonly List<Entry> _entries = new List<Entry>();

            public void Stage(PdfReviewWord word, PdfReviewPage page)
            {
                if (word == null || page == null)
                    return;
                Entry existing;
                if (_byWord.TryGetValue(word, out existing))
                {
                    // Даже одинаковый числовой PageIndex не делает две разные физические
                    // страницы одним владельцем. Такое повторное использование объекта
                    // разрушило бы side/page provenance, поэтому молча выбирать нельзя.
                    if (!ReferenceEquals(existing.Page, page) ||
                        existing.TargetPageIndex != page.PageIndex)
                        throw new InvalidOperationException(
                            "PdfReviewWord is present on more than one physical page.");
                    return;
                }

                var entry = new Entry
                {
                    Word = word,
                    Page = page,
                    OriginalPageIndex = word.PageIndex,
                    TargetPageIndex = page.PageIndex
                };
                _byWord.Add(word, entry);
                _entries.Add(entry);
            }

            public void Apply()
            {
                // Сначала проверяем весь snapshot, затем присваиваем без отмены и allocation:
                // callback не может увидеть половину применённой транзакции.
                foreach (Entry entry in _entries)
                    if (entry.Page.PageIndex != entry.TargetPageIndex)
                        throw new InvalidOperationException(
                            "PdfReviewPage.PageIndex changed during comparison.");
                // Документы сервиса лежат в общем LRU и уже имеют верного владельца: слово,
                // чьё значение не меняется, не переписываем ни здесь, ни при откате, чтобы
                // параллельный Compare того же кэша не видел чужих записей.
                foreach (Entry entry in _entries)
                    if (entry.OriginalPageIndex != entry.TargetPageIndex)
                        entry.Word.PageIndex = entry.TargetPageIndex;
            }

            public void Restore()
            {
                // Rollback намеренно не опрашивает отмену: восстановление входа обязательно.
                foreach (Entry entry in _entries)
                    if (entry.OriginalPageIndex != entry.TargetPageIndex)
                        entry.Word.PageIndex = entry.OriginalPageIndex;
            }
        }

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
            return Compare(left, right, limits, null);
        }

        public static PdfReviewResult Compare(PdfReviewDocument left, PdfReviewDocument right,
            PdfReviewLimits limits, Func<bool> cancelled)
        {
            if (left == null) throw new ArgumentNullException("left");
            if (right == null) throw new ArgumentNullException("right");
            if (limits == null) limits = PdfReviewLimits.Default();

            // Строки viewer — лишь вторичное предпочтение diff, поэтому у выравнивания свой
            // бюджет: сопоставление сотен страниц не должно съедать доказательство слов.
            var alignment = new DiffContext(limits, cancelled);
            alignment.ThrowIfCancellation();
            var result = new PdfReviewResult { Left = left, Right = right };
            foreach (PdfReviewPagePair pair in Align(left.Pages, right.Pages, limits, alignment))
                result.Pairs.Add(pair);
            PdfReviewStage incomplete = IncompleteStage(alignment,
                PdfReviewStage.PageAlignment);

            var ownership = new PageOwnershipTransaction();
            List<PdfReviewWord> leftWords = Flatten(left, alignment, ownership);
            List<PdfReviewWord> rightWords = Flatten(right, alignment, ownership);
            alignment.ThrowIfCancellation();

            // Сравнению пробелов гарантирована доля бюджета: исчерпание word diff на одном
            // участке не должно молча снимать проверку пробелов между доказанными словами.
            long work = limits.DiffWorkFor((long)leftWords.Count + rightWords.Count);
            long whitespaceWork = work / WhitespaceWorkShare;
            var context = new DiffContext(work - whitespaceWork, cancelled);

            bool completed = false;
            try
            {
                ownership.Apply();
                result.ReplaceOperations(DiffWords(leftWords, rightWords, limits,
                    PreferredPageMatches(result.Pairs), context));
                incomplete |= IncompleteStage(context, PdfReviewStage.WordDiff);
                // Якорное разбиение бюджет не обнуляет, но минимальность участка не доказана.
                if (context.Approximated)
                    incomplete |= PdfReviewStage.WordDiff;
                var pages = new ReviewPageLookup(result);
                var evidence = new WhitespaceEvidenceCache();
                if (!context.WorkExhausted)
                    ReconcileOrderArtifacts(result, pages, context);
                incomplete |= IncompleteStage(context, PdfReviewStage.OrderReconciliation);
                if (!context.WorkExhausted)
                    ClassifySplitJoin(result, pages, evidence, context);
                incomplete |= IncompleteStage(context, PdfReviewStage.SplitJoin);
                var whitespace = new DiffContext(
                    whitespaceWork + context.RemainingWork, cancelled);
                CompareWhitespace(result, limits, evidence, whitespace);
                incomplete |= IncompleteStage(whitespace, PdfReviewStage.Whitespace);
                // Флаг и список стадий выставляются до проекции: окно читает их вместе
                // с готовым snapshot и обязано пояснить, что часть сравнена грубо.
                result.IncompleteStages = incomplete;
                result.WorkExhausted = incomplete != PdfReviewStage.None;
                Project(result, context);
                context.ThrowIfCancellation();
                completed = true;
                return result;
            }
            finally
            {
                if (!completed)
                    ownership.Restore();
            }
        }

        private static PdfReviewStage IncompleteStage(DiffContext context,
            PdfReviewStage stage)
        {
            return context.WorkExhausted ? stage : PdfReviewStage.None;
        }

        /// <summary>
        /// Сквозная последовательность документа в физическом порядке страниц. Владельца
        /// восстанавливаем только после полного безопасного обхода: это сохраняет инвариант
        /// для тестовых и старых моделей, не оставляя частичного состояния при ошибке.
        /// </summary>
        internal static List<PdfReviewWord> Flatten(PdfReviewDocument document)
        {
            var ownership = new PageOwnershipTransaction();
            List<PdfReviewWord> words = Flatten(document, null, ownership);
            bool completed = false;
            try
            {
                ownership.Apply();
                completed = true;
                return words;
            }
            finally
            {
                if (!completed)
                    ownership.Restore();
            }
        }

        private static List<PdfReviewWord> Flatten(PdfReviewDocument document,
            DiffContext context, PageOwnershipTransaction ownership)
        {
            var words = new List<PdfReviewWord>();
            if (document == null)
                return words;
            foreach (PdfReviewPage page in document.Pages)
            {
                if (context != null)
                    context.PollCancellation();
                if (page == null)
                    continue;
                foreach (PdfReviewWord word in page.Words)
                {
                    if (context != null)
                        context.PollCancellation();
                    if (word == null)
                        continue;
                    ownership.Stage(word, page);
                    words.Add(word);
                }
            }
            return words;
        }

        /// <summary>
        /// Физическое сопоставление не меняет длину/состав LCS, но разрешает неоднозначность
        /// повторных колонтитулов: среди одинаково оптимальных сквозных выравниваний предпочитаем
        /// слова из страниц, которые viewer уже сопоставил друг с другом.
        /// </summary>
        private sealed class PreferredPageAlignment
        {
            private readonly Dictionary<int, int> _rightByLeft =
                new Dictionary<int, int>();
            private readonly Dictionary<int, int> _leftByRight =
                new Dictionary<int, int>();

            public void Add(int leftPageIndex, int rightPageIndex)
            {
                _rightByLeft[leftPageIndex] = rightPageIndex;
                _leftByRight[rightPageIndex] = leftPageIndex;
            }

            public bool IsPreferred(PdfReviewWord left, PdfReviewWord right)
            {
                if (left == null || right == null)
                    return false;
                int rightPageIndex;
                return _rightByLeft.TryGetValue(left.PageIndex, out rightPageIndex) &&
                    rightPageIndex == right.PageIndex;
            }

            public bool CanTrim(PdfReviewWord left, PdfReviewWord right)
            {
                if (left == null || right == null)
                    return true;
                int rightPageIndex;
                bool leftPaired = _rightByLeft.TryGetValue(left.PageIndex,
                    out rightPageIndex);
                int leftPageIndex;
                bool rightPaired = _leftByRight.TryGetValue(right.PageIndex,
                    out leftPageIndex);
                return leftPaired
                    ? rightPageIndex == right.PageIndex
                    : !rightPaired;
            }

            public PreferredPageAlignment Reverse()
            {
                var reversed = new PreferredPageAlignment();
                foreach (KeyValuePair<int, int> pair in _rightByLeft)
                    reversed.Add(pair.Value, pair.Key);
                return reversed;
            }
        }

        private static PreferredPageAlignment PreferredPageMatches(
            IList<PdfReviewPagePair> pairs)
        {
            var result = new PreferredPageAlignment();
            if (pairs == null)
                return result;
            foreach (PdfReviewPagePair pair in pairs)
                if (pair != null && pair.LeftPageIndex >= 0 && pair.RightPageIndex >= 0)
                    result.Add(pair.LeftPageIndex, pair.RightPageIndex);
            return result;
        }

        /// <summary>
        /// Снимок строк viewer и страниц обоих документов на одну стадию или проекцию.
        /// Явная пара и страница-владелец запрашиваются во внутренних циклах сверки и
        /// проекции; линейный поиск делал бы их стоимость O(строк × кандидатов). Строки и
        /// страницы ограничены MaxPages, поэтому снимок строится без опроса отмены, а каждый
        /// запрос пары по-прежнему оплачивается единицей семантической работы.
        /// </summary>
        private sealed class ReviewPageLookup
        {
            private readonly Dictionary<long, int> _pairCounts = new Dictionary<long, int>();
            private readonly Dictionary<int, PdfReviewPage> _leftPages;
            private readonly Dictionary<int, PdfReviewPage> _rightPages;

            public ReviewPageLookup(PdfReviewResult result)
            {
                if (result != null)
                {
                    foreach (PdfReviewPagePair pair in result.Pairs)
                    {
                        if (pair == null)
                            continue;
                        long key = PairKey(pair.LeftPageIndex, pair.RightPageIndex);
                        int count;
                        _pairCounts.TryGetValue(key, out count);
                        _pairCounts[key] = count + 1;
                    }
                }
                _leftPages = PagesByIndex(result == null ? null : result.Left);
                _rightPages = PagesByIndex(result == null ? null : result.Right);
            }

            /// <summary>Ровно одна строка viewer сопоставляет эти физические страницы.</summary>
            public bool IsExplicitPair(int leftPage, int rightPage, DiffContext context)
            {
                int count;
                return context.TryStep() &&
                    _pairCounts.TryGetValue(PairKey(leftPage, rightPage), out count) &&
                    count == 1;
            }

            /// <summary>Страница-владелец как у PageAt: при повторе индекса — первая.</summary>
            public PdfReviewPage Page(bool left, int pageIndex)
            {
                PdfReviewPage page;
                return (left ? _leftPages : _rightPages).TryGetValue(pageIndex, out page)
                    ? page : null;
            }

            private static long PairKey(int leftPage, int rightPage)
            {
                return ((long)leftPage << 32) | (uint)rightPage;
            }

            private static Dictionary<int, PdfReviewPage> PagesByIndex(
                PdfReviewDocument document)
            {
                var pages = new Dictionary<int, PdfReviewPage>();
                if (document == null)
                    return pages;
                foreach (PdfReviewPage page in document.Pages)
                    if (page != null && !pages.ContainsKey(page.PageIndex))
                        pages.Add(page.PageIndex, page);
                return pages;
            }
        }

        /// <summary>
        /// Проверяет размеры прямоугольной DP-матрицы до любого сложения в int и до
        /// резервирования/выделения памяти. Пользовательский предел может только уменьшить
        /// безопасный абсолютный потолок; выходы заполняются лишь для допустимой матрицы.
        /// </summary>
        internal static bool TryGetMatrixSize(int firstCount, int secondCount,
            int configuredMaxCells, out int rows, out int columns, out long cells)
        {
            rows = 0;
            columns = 0;
            cells = 0;
            if (firstCount < 0 || secondCount < 0 || configuredMaxCells <= 0)
                return false;

            long rowCount = (long)firstCount + 1L;
            long columnCount = (long)secondCount + 1L;
            if (rowCount > int.MaxValue || columnCount > int.MaxValue)
                return false;

            // Максимальный возможный результат здесь равен 2^62 и помещается в Int64.
            long total = rowCount * columnCount;
            long allowed = Math.Min((long)configuredMaxCells,
                PdfReviewLimits.AbsoluteMaxDiffCells);
            if (total <= 0 || total > allowed)
                return false;

            rows = (int)rowCount;
            columns = (int)columnCount;
            cells = total;
            return true;
        }

        /// <summary>
        /// Глобальное выравнивание последовательностей: точное совпадение почти бесплатно,
        /// похожие соседние страницы спариваются, непохожие становятся удалением/вставкой.
        /// Переставленная далеко страница не угадывается молча как прежнее место.
        /// </summary>
        public static List<PdfReviewPagePair> Align(IList<PdfReviewPage> left,
            IList<PdfReviewPage> right, PdfReviewLimits limits)
        {
            if (limits == null) limits = PdfReviewLimits.Default();
            var context = new DiffContext(limits, null);
            return Align(left, right, limits, context);
        }

        private static List<PdfReviewPagePair> Align(IList<PdfReviewPage> left,
            IList<PdfReviewPage> right, PdfReviewLimits limits, DiffContext context)
        {
            if (limits == null) limits = PdfReviewLimits.Default();
            context.ThrowIfCancellation();
            int n = left == null ? 0 : left.Count;
            int m = right == null ? 0 : right.Count;
            int rows;
            int columns;
            long cells;
            if (!TryGetMatrixSize(n, m, limits.MaxDiffCells,
                    out rows, out columns, out cells))
                throw new PdfReviewException(PdfReviewFailure.TooLarge, null,
                    Loc.T("review.err.tooLarge"));

            var cost = new double[rows, columns];
            var pairedWords = new int[rows, columns];
            var move = new byte[rows, columns]; // 1 pair, 2 left-only, 3 right-only
            // Множества слов каждой страницы считаем ОДИН раз, а не в каждой из «n·m»
            // ячеек выравнивания: иначе два документа по 500 страниц делали бы
            // полмиллиона лишних токенизаций одного и того же текста.
            var leftSets = new HashSet<string>[n];
            var leftOrderKeys = new string[n];
            for (int i = 0; i < n; i++)
            {
                context.PollCancellation();
                leftSets[i] = WordSet(left[i]);
                leftOrderKeys[i] = PageOrderKey(left[i]);
            }
            var rightSets = new HashSet<string>[m];
            var rightOrderKeys = new string[m];
            for (int j = 0; j < m; j++)
            {
                context.PollCancellation();
                rightSets[j] = WordSet(right[j]);
                rightOrderKeys[j] = PageOrderKey(right[j]);
            }

            // Вся потенциально комбинаторная работа резервируется до заполнения DP:
            // частичная матрица не должна порождать кажущееся page-соответствие.
            double[,] similarity;
            if (!TryAlignmentSimilarities(left, right, leftSets, rightSets,
                    leftOrderKeys, rightOrderKeys, context, out similarity))
                return UnalignedPages(left, right, context);

            for (int i = n - 1; i >= 0; i--)
            {
                context.PollCancellation();
                cost[i, m] = cost[i + 1, m] + GapCost;
                move[i, m] = 2;
            }
            for (int j = m - 1; j >= 0; j--)
            {
                context.PollCancellation();
                cost[n, j] = cost[n, j + 1] + GapCost;
                move[n, j] = 3;
            }
            for (int i = n - 1; i >= 0; i--)
            {
                for (int j = m - 1; j >= 0; j--)
                {
                    context.PollCancellation();
                    double sim = similarity[i, j];
                    double pairCost = sim >= SimilarityThreshold
                        ? cost[i + 1, j + 1] + (1.0 - sim)
                        : double.MaxValue;
                    int pairWords = sim >= SimilarityThreshold
                        ? pairedWords[i + 1, j + 1] + PagePairWeight(left[i], right[j])
                        : int.MinValue;
                    double dropLeft = cost[i + 1, j] + GapCost;
                    int dropLeftWords = pairedWords[i + 1, j];
                    double dropRight = cost[i, j + 1] + GapCost;
                    int dropRightWords = pairedWords[i, j + 1];

                    double bestCost = pairCost;
                    int bestWords = pairWords;
                    byte bestMove = 1;
                    if (BetterAlignment(dropLeft, dropLeftWords, bestCost, bestWords))
                    {
                        bestCost = dropLeft;
                        bestWords = dropLeftWords;
                        bestMove = 2;
                    }
                    if (BetterAlignment(dropRight, dropRightWords, bestCost, bestWords) ||
                        (SameAlignment(dropRight, dropRightWords, bestCost, bestWords) &&
                         bestMove == 2 && ComparePageIdentity(leftOrderKeys[i], left[i],
                             rightOrderKeys[j], right[j]) > 0))
                    {
                        bestCost = dropRight;
                        bestWords = dropRightWords;
                        bestMove = 3;
                    }
                    cost[i, j] = bestCost;
                    pairedWords[i, j] = bestWords;
                    move[i, j] = bestMove;
                }
            }

            var result = new List<PdfReviewPagePair>();
            int li = 0, ri = 0;
            while (li < n || ri < m)
            {
                context.PollCancellation();
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
        /// Сходство всех пар страниц для DP без перебора n·m·(слов на странице). DP берёт
        /// пару только начиная с SimilarityThreshold, а Similarity при заданных размерах
        /// множеств монотонна по числу общих слов. Поэтому общие слова считаем
        /// инвертированным индексом лишь по «редким» ключам (совместных вхождений не больше
        /// n+m), а каждый «частый» ключ получает номер бита в маске страницы. Пары, которым
        /// даже оценка сверху common + min(частых_i, частых_j) не даёт порога, отсеиваются
        /// сразу, остальным общее число частых ключей считается точно: popcount(Fᵢ & Fⱼ)
        /// за ⌈K/64⌉ машинных слов. Служебные слова есть почти на каждой странице, и
        /// поштучная проверка стоила бы n·m·|F| (на 500 страницах больше всего бюджета), а
        /// частых ключей K < Σвхождений/(2√(n+m)), так что маски короткие. Отсечённые пары
        /// заведомо ниже порога, остальные получают то же значение, так что DP видит те же
        /// числа, что при полном переборе. Это фильтрация кандидатов set-similarity join:
        /// Bayardo, Ma, Srikant, «Scaling Up All Pairs Similarity Search», WWW 2007.
        /// </summary>
        private static bool TryAlignmentSimilarities(IList<PdfReviewPage> left,
            IList<PdfReviewPage> right, HashSet<string>[] leftSets,
            HashSet<string>[] rightSets, string[] leftKeys, string[] rightKeys,
            DiffContext context, out double[,] similarity)
        {
            similarity = null;
            int n = leftSets.Length;
            int m = rightSets.Length;
            Dictionary<string, int> frequentBits = FrequentKeyBits(
                PageFrequency(leftSets, context), PageFrequency(rightSets, context),
                (long)n + m, context);
            int maskWords = (frequentBits.Count + 63) / 64;

            var rightPagesByRareKey =
                new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
            var rightFrequent = new int[m];
            var rightMasks = new ulong[m][];
            for (int j = 0; j < m; j++)
            {
                rightMasks[j] = new ulong[maskWords];
                foreach (string key in rightSets[j])
                {
                    context.PollCancellation();
                    int bit;
                    if (frequentBits.TryGetValue(key, out bit))
                    {
                        rightMasks[j][bit >> 6] |= 1UL << (bit & 63);
                        rightFrequent[j]++;
                        continue;
                    }
                    List<int> pages;
                    if (!rightPagesByRareKey.TryGetValue(key, out pages))
                    {
                        pages = new List<int>();
                        rightPagesByRareKey.Add(key, pages);
                    }
                    pages.Add(j);
                }
            }
            var leftFrequent = new int[n];
            var leftMasks = new ulong[n][];
            long rareWork = 0;
            for (int i = 0; i < n; i++)
            {
                leftMasks[i] = new ulong[maskWords];
                foreach (string key in leftSets[i])
                {
                    context.PollCancellation();
                    int bit;
                    List<int> pages;
                    if (frequentBits.TryGetValue(key, out bit))
                    {
                        leftMasks[i][bit >> 6] |= 1UL << (bit & 63);
                        leftFrequent[i]++;
                    }
                    else if (rightPagesByRareKey.TryGetValue(key, out pages))
                    {
                        rareWork = SaturatingAdd(rareWork, pages.Count);
                    }
                }
            }
            // Подсчёт редких совпадений плюс по единице на ячейку: оценка пары и сам DP.
            if (!context.TryReserve(SaturatingAdd(rareWork, (long)n * m)))
                return false;

            var common = new int[n, m];
            for (int i = 0; i < n; i++)
            {
                foreach (string key in leftSets[i])
                {
                    // Частых ключей в индексе редких нет: класс ключа общий для сторон.
                    List<int> pages;
                    if (!rightPagesByRareKey.TryGetValue(key, out pages))
                        continue;
                    foreach (int j in pages)
                    {
                        context.PollCancellation();
                        common[i, j]++;
                    }
                }
            }

            var values = new double[n, m];
            var verify = new List<long>();
            long verifyWork = 0;
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < m; j++)
                {
                    context.PollCancellation();
                    if (left[i] == null || right[j] == null)
                        continue;
                    if (string.Equals(leftKeys[i], rightKeys[j], StringComparison.Ordinal))
                    {
                        values[i, j] = 1;
                        continue;
                    }
                    int frequent = Math.Min(leftFrequent[i], rightFrequent[j]);
                    double bound = Similarity(left[i], right[j], leftSets[i].Count,
                        rightSets[j].Count, common[i, j] + frequent);
                    if (bound < SimilarityThreshold)
                        continue;
                    if (frequent == 0)
                    {
                        values[i, j] = bound;
                        continue;
                    }
                    verify.Add((long)i * m + j);
                    verifyWork = SaturatingAdd(verifyWork, maskWords);
                }
            }
            if (!context.TryReserve(verifyWork))
                return false;
            foreach (long cell in verify)
            {
                context.PollCancellation();
                int i = (int)(cell / m);
                int j = (int)(cell % m);
                values[i, j] = Similarity(left[i], right[j], leftSets[i].Count,
                    rightSets[j].Count,
                    common[i, j] + SharedBitCount(leftMasks[i], rightMasks[j]));
            }
            similarity = values;
            return true;
        }

        private static Dictionary<string, int> PageFrequency(HashSet<string>[] sets,
            DiffContext context)
        {
            var frequency = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (HashSet<string> set in sets)
            {
                foreach (string key in set)
                {
                    context.PollCancellation();
                    int count;
                    frequency.TryGetValue(key, out count);
                    frequency[key] = count + 1;
                }
            }
            return frequency;
        }

        /// <summary>
        /// Номера бит «частых» ключей: встречаются в большем числе пар страниц, чем n+m.
        /// Произведение частот симметрично, поэтому класс ключа один на обеих сторонах, а
        /// частый ключ есть на обеих. Каждый ключ классифицируется один раз.
        /// </summary>
        private static Dictionary<string, int> FrequentKeyBits(
            Dictionary<string, int> leftFrequency, Dictionary<string, int> rightFrequency,
            long rarePairs, DiffContext context)
        {
            var bits = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, int> entry in leftFrequency)
            {
                context.PollCancellation();
                int rightCount;
                if (rightFrequency.TryGetValue(entry.Key, out rightCount) &&
                    (long)entry.Value * rightCount > rarePairs)
                    bits.Add(entry.Key, bits.Count);
            }
            return bits;
        }

        /// <summary>
        /// popcount(left & right) для масок одной длины. BitOperations.PopCount есть только
        /// в .NET Core 3.0+, поэтому — параллельный (SWAR) подсчёт из «Bit Twiddling Hacks»
        /// С. Андерсона («Counting bits set, in parallel»), расширенный до 64 бит. Чистая
        /// целочисленная арифметика даёт одинаковый результат на x86 и x64.
        /// </summary>
        private static int SharedBitCount(ulong[] left, ulong[] right)
        {
            int count = 0;
            for (int w = 0; w < left.Length; w++)
            {
                ulong v = left[w] & right[w];
                v -= (v >> 1) & 0x5555555555555555UL;
                v = (v & 0x3333333333333333UL) + ((v >> 2) & 0x3333333333333333UL);
                v = (v + (v >> 4)) & 0x0F0F0F0F0F0F0F0FUL;
                count += (int)(unchecked(v * 0x0101010101010101UL) >> 56);
            }
            return count;
        }

        private static long SaturatingAdd(long value, long addition)
        {
            if (addition <= 0)
                return value;
            return value > long.MaxValue - addition ? long.MaxValue : value + addition;
        }

        private static List<PdfReviewPagePair> UnalignedPages(
            IList<PdfReviewPage> left, IList<PdfReviewPage> right, DiffContext context)
        {
            var result = new List<PdfReviewPagePair>();
            if (left != null)
            {
                foreach (PdfReviewPage page in left)
                {
                    context.PollCancellation();
                    result.Add(new PdfReviewPagePair
                    {
                        LeftPageIndex = page == null ? -1 : page.PageIndex,
                        RightPageIndex = -1,
                        Status = PdfReviewPairStatus.LeftOnly,
                        Similarity = 0
                    });
                }
            }
            if (right != null)
            {
                foreach (PdfReviewPage page in right)
                {
                    context.PollCancellation();
                    result.Add(new PdfReviewPagePair
                    {
                        LeftPageIndex = -1,
                        RightPageIndex = page == null ? -1 : page.PageIndex,
                        Status = PdfReviewPairStatus.RightOnly,
                        Similarity = 0
                    });
                }
            }
            return result;
        }

        private static bool BetterAlignment(double candidateCost, int candidateWords,
            double currentCost, int currentWords)
        {
            if (!SameCost(candidateCost, currentCost))
                return candidateCost < currentCost;
            return candidateWords > currentWords;
        }

        private static bool SameAlignment(double candidateCost, int candidateWords,
            double currentCost, int currentWords)
        {
            return SameCost(candidateCost, currentCost) && candidateWords == currentWords;
        }

        private static bool SameCost(double left, double right)
        {
            return left == right || Math.Abs(left - right) <= AlignmentTieEpsilon;
        }

        private static int PagePairWeight(PdfReviewPage left, PdfReviewPage right)
        {
            int leftWords = left == null || left.Words == null ? 0 : left.Words.Count;
            int rightWords = right == null || right.Words == null ? 0 : right.Words.Count;
            return Math.Min(leftWords, rightWords);
        }

        /// <summary>
        /// Канонический ключ страницы для симметричного tie-break. Длины не дают разным
        /// границам слов слиться в один и тот же ключ; форматирование и геометрия слов
        /// в содержательную часть не входят.
        /// </summary>
        private static string PageOrderKey(PdfReviewPage page)
        {
            if (page == null)
                return "";
            if (page.Words != null && page.Words.Count > 0)
            {
                var key = new StringBuilder();
                foreach (PdfReviewWord word in page.Words)
                {
                    string value = WordKey(word);
                    key.Append(value.Length.ToString(CultureInfo.InvariantCulture))
                        .Append(':').Append(value);
                }
                return key.ToString();
            }
            return Normalize(page.NormalizedText ?? page.Text ?? "");
        }

        /// <summary>
        /// При полностью равных DP-путях обе стороны используют один канонический порядок.
        /// Это не меняет стоимость и pairing, а лишь выбирает тот же page-anchor при A↔B.
        /// Размеры — последний физический tie-break для страниц с одинаковым текстом.
        /// </summary>
        private static int ComparePageIdentity(string leftKey, PdfReviewPage left,
            string rightKey, PdfReviewPage right)
        {
            int comparison = string.CompareOrdinal(leftKey ?? "", rightKey ?? "");
            if (comparison != 0)
                return comparison;
            comparison = PageViewWidth(left).CompareTo(PageViewWidth(right));
            if (comparison != 0)
                return comparison;
            return PageViewHeight(left).CompareTo(PageViewHeight(right));
        }

        private static double PageViewWidth(PdfReviewPage page)
        {
            if (page == null) return 0;
            return page.ViewWidthPt > 0 ? page.ViewWidthPt : page.WidthPt;
        }

        private static double PageViewHeight(PdfReviewPage page)
        {
            if (page == null) return 0;
            return page.ViewHeightPt > 0 ? page.ViewHeightPt : page.HeightPt;
        }

        /// <summary>
        /// Создать только строку viewer. Канонические операции принадлежат результату целиком;
        /// статус этой строки окончательно устанавливает Project(...).
        /// </summary>
        public static PdfReviewPagePair Pair(PdfReviewPage left, PdfReviewPage right,
            PdfReviewLimits limits)
        {
            var pair = new PdfReviewPagePair
            {
                LeftPageIndex = left == null ? -1 : left.PageIndex,
                RightPageIndex = right == null ? -1 : right.PageIndex,
                Similarity = Similarity(left, right)
            };
            if (left == null)
                pair.Status = PdfReviewPairStatus.RightOnly;
            else if (right == null)
                pair.Status = PdfReviewPairStatus.LeftOnly;
            else
                pair.Status = WordsEqual(left.Words, right.Words)
                    ? PdfReviewPairStatus.Unchanged : PdfReviewPairStatus.Changed;
            return pair;
        }

        /// <summary>
        /// Ворд-дифф: общий префикс/суффикс — сразу Equal, середина — ЛСД-матрица, если она
        /// помещается в потолок; не помещается — середина делится пополам и дифф продолжается
        /// рекурсивно, а если бисекции не хватает её доли бюджета — делится якорями
        /// уникальных слов. Операции всегда владеют словами своей стороны.
        /// </summary>
        public static List<PdfReviewWordOp> DiffWords(IList<PdfReviewWord> left,
            IList<PdfReviewWord> right, PdfReviewLimits limits)
        {
            return DiffWords(left, right, limits, null);
        }

        internal static List<PdfReviewWordOp> DiffWords(IList<PdfReviewWord> left,
            IList<PdfReviewWord> right, PdfReviewLimits limits, Func<bool> cancelled)
        {
            if (limits == null) limits = PdfReviewLimits.Default();
            var context = new DiffContext(limits, cancelled);
            context.ThrowIfCancellation();
            List<PdfReviewWordOp> result = DiffWords(left, right, limits, null, context);
            context.ThrowIfCancellation();
            return result;
        }

        private static List<PdfReviewWordOp> DiffWords(IList<PdfReviewWord> left,
            IList<PdfReviewWord> right, PdfReviewLimits limits,
            PreferredPageAlignment preferredPages, DiffContext context)
        {
            if (limits == null) limits = PdfReviewLimits.Default();
            IList<PdfReviewWord> sourceLeft = left ?? EmptyWords;
            IList<PdfReviewWord> sourceRight = right ?? EmptyWords;

            // LCS может иметь несколько одинаково длинных решений. Ориентируем обе версии
            // по одному каноническому порядку ключей, поэтому A↔B и B↔A запускают и матрицу,
            // и Myers над одной и той же парой объектов, а затем лишь меняют владельцев сторон.
            bool swapped = CompareWordSequences(sourceLeft, sourceRight, context) > 0;
            IList<PdfReviewWord> a = swapped ? sourceRight : sourceLeft;
            IList<PdfReviewWord> b = swapped ? sourceLeft : sourceRight;
            PreferredPageAlignment canonicalPages = swapped && preferredPages != null
                ? preferredPages.Reverse() : preferredPages;

            var canonical = new List<PdfReviewWordOp>();
            DiffRange(a, 0, a.Count, b, 0, b.Count, limits, canonical,
                canonicalPages, context, 0);
            if (canonicalPages != null)
                canonical = ShiftMatchesToPreferredPages(canonical, canonicalPages, context);
            if (canonical.Count == 0)
                canonical.Add(new PdfReviewWordOp { Kind = PdfReviewDiffKind.Equal });

            List<PdfReviewWordOp> oriented = swapped
                ? MirrorRawOperations(canonical, context) : canonical;
            return NormalizeRawChangeHunks(oriented, context);
        }

        private static int CompareWordSequences(IList<PdfReviewWord> left,
            IList<PdfReviewWord> right, DiffContext context)
        {
            int common = Math.Min(left == null ? 0 : left.Count,
                right == null ? 0 : right.Count);
            for (int i = 0; i < common; i++)
            {
                context.PollCancellation();
                int comparison = string.CompareOrdinal(WordKey(left[i]), WordKey(right[i]));
                if (comparison != 0)
                    return comparison;
            }
            int leftCount = left == null ? 0 : left.Count;
            int rightCount = right == null ? 0 : right.Count;
            return leftCount.CompareTo(rightCount);
        }

        /// <summary>Сырой результат DiffRange по одному слову.</summary>
        private struct RawDiffEntry
        {
            public PdfReviewDiffKind Kind;
            public PdfReviewWord Left;
            public PdfReviewWord Right;
        }

        /// <summary>
        /// Предпочтение страниц видит только матричный лист; срезание краёв без
        /// предпочтений и Myers сопоставляют повторное слово жадно, и колонтитул вставленной
        /// страницы «забирает» совпадение соседней. Длина LCS от этого не меняется, а
        /// страница-владелец совпадения — меняется. Проход переносит Equal(l, r) без
        /// предпочтения на r' с тем же ключом из смежного блока одних вставок (или на l'
        /// из блока одних удалений), если viewer сопоставил страницы l и r'. Обмен ролями
        /// двух элементов сохраняет порядок обеих сторон и число Equal, а число
        /// предпочтённых совпадений только растёт. Без единого сдвига операции не
        /// перестраиваются.
        /// </summary>
        private static List<PdfReviewWordOp> ShiftMatchesToPreferredPages(
            List<PdfReviewWordOp> ops, PreferredPageAlignment preferredPages,
            DiffContext context)
        {
            var entries = new List<RawDiffEntry>();
            foreach (PdfReviewWordOp op in ops)
            {
                context.PollCancellation();
                if (op.Kind == PdfReviewDiffKind.Equal)
                    foreach (PdfReviewWordMatch match in op.Matches)
                        entries.Add(new RawDiffEntry
                        {
                            Kind = PdfReviewDiffKind.Equal,
                            Left = match.Left,
                            Right = match.Right
                        });
                else if (op.Kind == PdfReviewDiffKind.Delete)
                    foreach (PdfReviewWord word in op.LeftWords)
                        entries.Add(new RawDiffEntry { Kind = op.Kind, Left = word });
                else
                    foreach (PdfReviewWord word in op.RightWords)
                        entries.Add(new RawDiffEntry { Kind = op.Kind, Right = word });
            }
            bool forward = ShiftMatches(entries, preferredPages, 1, context);
            bool backward = ShiftMatches(entries, preferredPages, -1, context);
            if (!forward && !backward)
                return ops;

            // DiffRange порождает только точные совпадения ключей.
            var result = new List<PdfReviewWordOp>();
            foreach (RawDiffEntry entry in entries)
            {
                context.PollCancellation();
                if (entry.Kind == PdfReviewDiffKind.Equal)
                    AppendEqual(result, entry.Left, entry.Right, PdfReviewMatchKind.Exact);
                else
                    Append(result, entry.Kind, entry.Kind == PdfReviewDiffKind.Delete
                        ? entry.Left : entry.Right);
            }
            return result;
        }

        /// <summary>
        /// direction = +1 переносит совпадение вперёд, в однородный блок сразу за ним,
        /// -1 — назад. Обход идёт против направления: блок за текущим элементом уже
        /// окончателен, и многословный колонтитул переезжает целиком, слово за словом.
        /// Кандидат ищется от дальнего края блока, поэтому совпадение примыкает к соседнему
        /// Equal, а следующий поиск начинается сразу перед ним: каждый элемент блока
        /// просматривается не больше одного раза, и проход линеен.
        /// </summary>
        private static bool ShiftMatches(List<RawDiffEntry> entries,
            PreferredPageAlignment preferredPages, int direction, DiffContext context)
        {
            bool shifted = false;
            int start = direction > 0 ? entries.Count - 1 : 0;
            // Блок за элементом i: entries[i + direction .. far] одного вида kind;
            // far == i означает пустой блок.
            int far = start;
            PdfReviewDiffKind kind = PdfReviewDiffKind.Equal;
            for (int i = start; i >= 0 && i < entries.Count; i -= direction)
            {
                context.PollCancellation();
                RawDiffEntry entry = entries[i];
                if (entry.Kind == PdfReviewDiffKind.Equal && far != i &&
                    !preferredPages.IsPreferred(entry.Left, entry.Right))
                {
                    int target = PreferredShiftTarget(entries, i, far, kind, direction,
                        preferredPages, context);
                    if (target >= 0)
                    {
                        RawDiffEntry other = entries[target];
                        bool inserts = kind == PdfReviewDiffKind.Insert;
                        entries[target] = new RawDiffEntry
                        {
                            Kind = PdfReviewDiffKind.Equal,
                            Left = inserts ? entry.Left : other.Left,
                            Right = inserts ? other.Right : entry.Right
                        };
                        entries[i] = new RawDiffEntry
                        {
                            Kind = kind,
                            Left = inserts ? null : entry.Left,
                            Right = inserts ? entry.Right : null
                        };
                        shifted = true;
                        far = target - direction;
                        continue;
                    }
                }
                if (entry.Kind == PdfReviewDiffKind.Equal)
                {
                    far = i - direction;
                }
                else if (far == i || entry.Kind != kind)
                {
                    far = i;
                    kind = entry.Kind;
                }
            }
            return shifted;
        }

        private static int PreferredShiftTarget(List<RawDiffEntry> entries, int index,
            int far, PdfReviewDiffKind kind, int direction,
            PreferredPageAlignment preferredPages, DiffContext context)
        {
            RawDiffEntry match = entries[index];
            bool inserts = kind == PdfReviewDiffKind.Insert;
            string key = WordKey(inserts ? match.Right : match.Left);
            for (int j = far; j != index; j -= direction)
            {
                context.PollCancellation();
                RawDiffEntry candidate = entries[j];
                if (inserts
                    ? WordKey(candidate.Right) == key &&
                      preferredPages.IsPreferred(match.Left, candidate.Right)
                    : WordKey(candidate.Left) == key &&
                      preferredPages.IsPreferred(candidate.Left, match.Right))
                    return j;
            }
            return -1;
        }

        private static List<PdfReviewWordOp> MirrorRawOperations(
            IList<PdfReviewWordOp> source, DiffContext context)
        {
            var result = new List<PdfReviewWordOp>();
            foreach (PdfReviewWordOp op in source)
            {
                context.PollCancellation();
                if (op == null)
                    continue;
                if (op.Kind == PdfReviewDiffKind.Delete)
                {
                    foreach (PdfReviewWord word in op.LeftWords)
                    {
                        context.PollCancellation();
                        Append(result, PdfReviewDiffKind.Insert, word);
                    }
                }
                else if (op.Kind == PdfReviewDiffKind.Insert)
                {
                    foreach (PdfReviewWord word in op.RightWords)
                    {
                        context.PollCancellation();
                        Append(result, PdfReviewDiffKind.Delete, word);
                    }
                }
                else if (op.Matches.Count > 0)
                {
                    foreach (PdfReviewWordMatch match in op.Matches)
                    {
                        context.PollCancellation();
                        if (match != null)
                            AppendEqual(result, match.Right, match.Left, match.Kind);
                    }
                }
                else
                {
                    result.Add(new PdfReviewWordOp
                    {
                        Kind = PdfReviewDiffKind.Equal,
                        MatchKind = op.MatchKind
                    });
                }
            }
            return result;
        }

        private static List<PdfReviewWordOp> NormalizeRawChangeHunks(
            IList<PdfReviewWordOp> source, DiffContext context)
        {
            var result = new List<PdfReviewWordOp>();
            var deleted = new List<PdfReviewWord>();
            var inserted = new List<PdfReviewWord>();
            foreach (PdfReviewWordOp op in source)
            {
                context.PollCancellation();
                if (op != null && op.Kind == PdfReviewDiffKind.Delete)
                {
                    deleted.AddRange(op.LeftWords);
                    continue;
                }
                if (op != null && op.Kind == PdfReviewDiffKind.Insert)
                {
                    inserted.AddRange(op.RightWords);
                    continue;
                }
                AppendRawChanges(result, deleted, inserted, context);
                deleted.Clear();
                inserted.Clear();
                AppendOperation(result, op);
            }
            AppendRawChanges(result, deleted, inserted, context);
            if (result.Count == 0)
                result.Add(new PdfReviewWordOp { Kind = PdfReviewDiffKind.Equal });
            return result;
        }

        private static void AppendRawChanges(List<PdfReviewWordOp> operations,
            IList<PdfReviewWord> deleted, IList<PdfReviewWord> inserted,
            DiffContext context)
        {
            foreach (PdfReviewWord word in deleted)
            {
                context.PollCancellation();
                Append(operations, PdfReviewDiffKind.Delete, word);
            }
            foreach (PdfReviewWord word in inserted)
            {
                context.PollCancellation();
                Append(operations, PdfReviewDiffKind.Insert, word);
            }
        }

        private static readonly PdfReviewWord[] EmptyWords = new PdfReviewWord[0];

        private static void DiffRange(IList<PdfReviewWord> a, int aStart, int aEnd,
            IList<PdfReviewWord> b, int bStart, int bEnd, PdfReviewLimits limits,
            List<PdfReviewWordOp> ops, PreferredPageAlignment preferredPages,
            DiffContext context, int recursionDepth)
        {
            context.PollCancellation();
            if (context.WorkExhausted)
            {
                AppendUnprovenRange(a, aStart, aEnd, b, bStart, bEnd, ops, context);
                return;
            }

            int suffix = TrimEqualEnds(a, ref aStart, ref aEnd, b, ref bStart, ref bEnd,
                preferredPages, ops, context);
            // Матрица стоит n·m независимо от числа правок, поэтому только для малых листьев.
            int leafCells = Math.Min(limits.MaxDiffCells, MaxMatrixLeafCells);
            int matrixRows;
            int matrixColumns;
            long matrixCells;
            bool fitsMatrix = TryGetMatrixSize(aEnd - aStart, bEnd - bStart, leafCells,
                out matrixRows, out matrixColumns, out matrixCells);
            if (!fitsMatrix && preferredPages != null &&
                aEnd - aStart > 1 && bEnd - bStart > 1)
            {
                // Myers предпочтений страниц не видит, а бисекция корректна только без общих
                // краёв (diff-match-patch срезает их в diff_main до diff_bisect). Оставленный
                // CanTrim край дал бы вырожденную точку D=0/D=1 в конце диапазона, и весь
                // хвост ушёл бы в сырые Delete/Insert. Общий край всегда входит в некоторый
                // LCS, так что срезание без предпочтений длину LCS не меняет.
                suffix += TrimEqualEnds(a, ref aStart, ref aEnd, b, ref bStart, ref bEnd,
                    null, ops, context);
                fitsMatrix = TryGetMatrixSize(aEnd - aStart, bEnd - bStart, leafCells,
                    out matrixRows, out matrixColumns, out matrixCells);
            }

            int n = aEnd - aStart;
            int m = bEnd - bStart;
            if (n == 0 && m == 0)
            {
                // Середина пуста.
            }
            else if (n == 0)
            {
                for (int j = bStart; j < bEnd; j++)
                {
                    context.PollCancellation();
                    Append(ops, PdfReviewDiffKind.Insert, b[j]);
                }
            }
            else if (m == 0)
            {
                for (int i = aStart; i < aEnd; i++)
                {
                    context.PollCancellation();
                    Append(ops, PdfReviewDiffKind.Delete, a[i]);
                }
            }
            else if (n == 1)
            {
                if (!context.TryReserve(m))
                {
                    AppendUnprovenRange(a, aStart, aEnd, b, bStart, bEnd, ops, context);
                }
                else
                {
                    int firstMatch = -1;
                    int preferredMatch = -1;
                    for (int j = bStart; j < bEnd; j++)
                    {
                        context.PollCancellation();
                        if (WordKey(b[j]) != WordKey(a[aStart]))
                            continue;
                        if (firstMatch < 0)
                            firstMatch = j;
                        if (preferredPages != null &&
                            preferredPages.IsPreferred(a[aStart], b[j]))
                        {
                            preferredMatch = j;
                            break;
                        }
                        if (preferredPages == null)
                            break;
                    }
                    int match = preferredMatch >= 0 ? preferredMatch : firstMatch;
                    if (match >= 0)
                    {
                        for (int k = bStart; k < match; k++)
                        {
                            context.PollCancellation();
                            Append(ops, PdfReviewDiffKind.Insert, b[k]);
                        }
                        AppendEqual(ops, a[aStart], b[match], PdfReviewMatchKind.Exact);
                        for (int k = match + 1; k < bEnd; k++)
                        {
                            context.PollCancellation();
                            Append(ops, PdfReviewDiffKind.Insert, b[k]);
                        }
                    }
                    else
                    {
                        Append(ops, PdfReviewDiffKind.Delete, a[aStart]);
                        for (int k = bStart; k < bEnd; k++)
                        {
                            context.PollCancellation();
                            Append(ops, PdfReviewDiffKind.Insert, b[k]);
                        }
                    }
                }
            }
            else if (m == 1)
            {
                if (!context.TryReserve(n))
                {
                    AppendUnprovenRange(a, aStart, aEnd, b, bStart, bEnd, ops, context);
                }
                else
                {
                    int firstMatch = -1;
                    int preferredMatch = -1;
                    for (int i = aStart; i < aEnd; i++)
                    {
                        context.PollCancellation();
                        if (WordKey(a[i]) != WordKey(b[bStart]))
                            continue;
                        if (firstMatch < 0)
                            firstMatch = i;
                        if (preferredPages != null &&
                            preferredPages.IsPreferred(a[i], b[bStart]))
                        {
                            preferredMatch = i;
                            break;
                        }
                        if (preferredPages == null)
                            break;
                    }
                    int match = preferredMatch >= 0 ? preferredMatch : firstMatch;
                    for (int k = aStart; k < (match >= 0 ? match : aEnd); k++)
                    {
                        context.PollCancellation();
                        Append(ops, PdfReviewDiffKind.Delete, a[k]);
                    }
                    if (match >= 0)
                    {
                        AppendEqual(ops, a[match], b[bStart], PdfReviewMatchKind.Exact);
                        for (int k = match + 1; k < aEnd; k++)
                        {
                            context.PollCancellation();
                            Append(ops, PdfReviewDiffKind.Delete, a[k]);
                        }
                    }
                    else
                    {
                        Append(ops, PdfReviewDiffKind.Insert, b[bStart]);
                    }
                }
            }
            else if (fitsMatrix)
            {
                if (!LcsCore(a, aStart, aEnd, b, bStart, bEnd, ops,
                        preferredPages, context, matrixRows, matrixColumns, matrixCells))
                    AppendUnprovenRange(a, aStart, aEnd, b, bStart, bEnd, ops, context);
            }
            else if (recursionDepth >= MaxDiffRecursionDepth)
            {
                context.ExhaustWork();
                AppendUnprovenRange(a, aStart, aEnd, b, bStart, bEnd, ops, context);
            }
            else if (context.Approximated)
            {
                // Бисекция уже не уложилась в свою долю, значит точный diff целиком не
                // укладывается в бюджет: следующие большие диапазоны сразу делятся якорями,
                // а не сжигают по половине остатка, оставляя хвост документа без бюджета.
                DiffAnchoredRange(a, aStart, aEnd, b, bStart, bEnd, limits, ops,
                    preferredPages, context, recursionDepth);
            }
            else
            {
                DiffContext attempt = context.Slice(context.RemainingWork / BisectWorkShare);
                int aMid, bMid;
                bool bisected = TryBisect(a, aStart, aEnd, b, bStart, bEnd,
                    attempt, out aMid, out bMid);
                context.Settle(attempt);
                if (attempt.WorkExhausted)
                {
                    DiffAnchoredRange(a, aStart, aEnd, b, bStart, bEnd, limits, ops,
                        preferredPages, context, recursionDepth);
                }
                else if (bisected &&
                    !(aMid == aStart && bMid == bStart) &&
                    !(aMid == aEnd && bMid == bEnd))
                {
                    DiffRange(a, aStart, aMid, b, bStart, bMid, limits, ops,
                        preferredPages, context, recursionDepth + 1);
                    DiffRange(a, aMid, aEnd, b, bMid, bEnd, limits, ops,
                        preferredPages, context, recursionDepth + 1);
                }
                else
                {
                    AppendRawRange(a, aStart, aEnd, b, bStart, bEnd, ops, context);
                }
            }

            AppendEqualSuffix(a, aEnd, b, bEnd, suffix, ops, context);
        }

        /// <summary>
        /// Срезает общий префикс и суффикс равных ключей: префикс сразу дописывается как
        /// Equal, длина суффикса возвращается, чтобы вызывающий дописал его после середины.
        /// Общий край всегда входит в некоторый LCS, поэтому срезание не укорачивает LCS и,
        /// будучи линейным точным сравнением, не расходует бюджет. preferredPages лишь не
        /// даёт срезать пару страниц, которую viewer не сопоставил: её решит матрица.
        /// </summary>
        private static int TrimEqualEnds(IList<PdfReviewWord> a, ref int aStart,
            ref int aEnd, IList<PdfReviewWord> b, ref int bStart, ref int bEnd,
            PreferredPageAlignment preferredPages, List<PdfReviewWordOp> ops,
            DiffContext context)
        {
            while (aStart < aEnd && bStart < bEnd)
            {
                context.PollCancellation();
                if (WordKey(a[aStart]) != WordKey(b[bStart]) ||
                    (preferredPages != null &&
                     !preferredPages.CanTrim(a[aStart], b[bStart])))
                    break;
                AppendEqual(ops, a[aStart], b[bStart], PdfReviewMatchKind.Exact);
                aStart++;
                bStart++;
            }

            int suffix = 0;
            while (aEnd > aStart && bEnd > bStart)
            {
                context.PollCancellation();
                if (WordKey(a[aEnd - 1]) != WordKey(b[bEnd - 1]) ||
                    (preferredPages != null &&
                     !preferredPages.CanTrim(a[aEnd - 1], b[bEnd - 1])))
                    break;
                suffix++;
                aEnd--;
                bEnd--;
            }
            return suffix;
        }

        private static void AppendEqualSuffix(IList<PdfReviewWord> a, int aEnd,
            IList<PdfReviewWord> b, int bEnd, int suffix, List<PdfReviewWordOp> ops,
            DiffContext context)
        {
            for (int s = 0; s < suffix; s++)
            {
                context.PollCancellation();
                AppendEqual(ops, a[aEnd + s], b[bEnd + s], PdfReviewMatchKind.Exact);
            }
        }

        /// <summary>
        /// Диапазон, на доказательство которого бюджета не хватило. Линейное точное
        /// срезание общих краёв остаётся и после исчерпания: каждое слово срезается не
        /// больше одного раза за весь diff. Консервативной парой Delete/Insert становится
        /// только непроверенная середина, а не весь диапазон.
        /// </summary>
        private static void AppendUnprovenRange(IList<PdfReviewWord> a, int aStart,
            int aEnd, IList<PdfReviewWord> b, int bStart, int bEnd,
            List<PdfReviewWordOp> ops, DiffContext context)
        {
            int suffix = TrimEqualEnds(a, ref aStart, ref aEnd, b, ref bStart, ref bEnd,
                null, ops, context);
            AppendRawRange(a, aStart, aEnd, b, bStart, bEnd, ops, context);
            AppendEqualSuffix(a, aEnd, b, bEnd, suffix, ops, context);
        }

        private static void AppendRawRange(IList<PdfReviewWord> a, int aStart,
            int aEnd, IList<PdfReviewWord> b, int bStart, int bEnd,
            List<PdfReviewWordOp> ops, DiffContext context)
        {
            for (int i = aStart; i < aEnd; i++)
            {
                context.PollCancellation();
                Append(ops, PdfReviewDiffKind.Delete, a[i]);
            }
            for (int j = bStart; j < bEnd; j++)
            {
                context.PollCancellation();
                Append(ops, PdfReviewDiffKind.Insert, b[j]);
            }
        }

        /// <summary>
        /// Запасное разбиение диапазона, когда точной бисекции не хватило её доли бюджета.
        /// Якоря — слова, чей ключ встречается в диапазоне ровно один раз с каждой стороны;
        /// из них берётся наибольшая цепочка, возрастающая на обеих сторонах (patience diff
        /// Б. Коэна, он же git diff --patience), а промежутки между якорями снова сравнивает
        /// DiffRange. Уникальное слово почти всегда входит в оптимальное выравнивание, но
        /// не обязано, поэтому результат помечается Approximated. Без якорей диапазон
        /// остаётся непроверенным, как при исчерпании.
        /// </summary>
        private static void DiffAnchoredRange(IList<PdfReviewWord> a, int aStart, int aEnd,
            IList<PdfReviewWord> b, int bStart, int bEnd, PdfReviewLimits limits,
            List<PdfReviewWordOp> ops, PreferredPageAlignment preferredPages,
            DiffContext context, int recursionDepth)
        {
            context.MarkApproximated();
            List<int> anchorsA;
            List<int> anchorsB;
            if (!TryUniqueAnchors(a, aStart, aEnd, b, bStart, bEnd, context,
                    out anchorsA, out anchorsB) || anchorsA.Count == 0)
            {
                AppendUnprovenRange(a, aStart, aEnd, b, bStart, bEnd, ops, context);
                return;
            }
            int ai = aStart;
            int bi = bStart;
            for (int k = 0; k < anchorsA.Count; k++)
            {
                DiffRange(a, ai, anchorsA[k], b, bi, anchorsB[k], limits, ops,
                    preferredPages, context, recursionDepth + 1);
                AppendEqual(ops, a[anchorsA[k]], b[anchorsB[k]], PdfReviewMatchKind.Exact);
                ai = anchorsA[k] + 1;
                bi = anchorsB[k] + 1;
            }
            DiffRange(a, ai, aEnd, b, bi, bEnd, limits, ops, preferredPages, context,
                recursionDepth + 1);
        }

        /// <summary>
        /// Позиции якорей patience diff в порядке документа. Кандидаты (ключ уникален на
        /// обеих сторонах) перечисляются по a, и среди их позиций в b ищется наибольшая
        /// возрастающая подпоследовательность сортировкой «пасьянсом» за O(U log U):
        /// tails[k] — кандидат с наименьшей позицией в b, которым кончается цепочка длины
        /// k+1 (Aldous, Diaconis, «Longest increasing subsequences: from patience sorting
        /// to the Baik-Deift-Johansson theorem», Bull. AMS 36, 1999). Подсчёт ключей стоит
        /// по единице на слово, LIS — как сортировка кандидатов; false — бюджета не хватило.
        /// </summary>
        private static bool TryUniqueAnchors(IList<PdfReviewWord> a, int aStart, int aEnd,
            IList<PdfReviewWord> b, int bStart, int bEnd, DiffContext context,
            out List<int> anchorsA, out List<int> anchorsB)
        {
            anchorsA = null;
            anchorsB = null;
            if (!context.TryReserve((long)(aEnd - aStart) + (bEnd - bStart)))
                return false;
            Dictionary<string, int> uniqueA = UniquePositions(a, aStart, aEnd, context);
            Dictionary<string, int> uniqueB = UniquePositions(b, bStart, bEnd, context);
            var candidatesA = new List<int>();
            var candidatesB = new List<int>();
            for (int i = aStart; i < aEnd; i++)
            {
                context.PollCancellation();
                string key = WordKey(a[i]);
                int j;
                if (uniqueA[key] == i && uniqueB.TryGetValue(key, out j) && j >= 0)
                {
                    candidatesA.Add(i);
                    candidatesB.Add(j);
                }
            }
            if (!context.TryReserveSort(candidatesB.Count))
                return false;

            var tails = new List<int>();
            var previous = new int[candidatesB.Count];
            for (int c = 0; c < candidatesB.Count; c++)
            {
                context.PollCancellation();
                int low = 0;
                int high = tails.Count;
                while (low < high)
                {
                    int middle = (low + high) >> 1;
                    if (candidatesB[tails[middle]] < candidatesB[c])
                        low = middle + 1;
                    else
                        high = middle;
                }
                previous[c] = low > 0 ? tails[low - 1] : -1;
                if (low == tails.Count)
                    tails.Add(c);
                else
                    tails[low] = c;
            }

            var chain = new int[tails.Count];
            for (int k = tails.Count - 1, c = tails.Count == 0 ? -1 : tails[k]; k >= 0;
                k--, c = previous[c])
                chain[k] = c;
            anchorsA = new List<int>(chain.Length);
            anchorsB = new List<int>(chain.Length);
            foreach (int c in chain)
            {
                anchorsA.Add(candidatesA[c]);
                anchorsB.Add(candidatesB[c]);
            }
            return true;
        }

        /// <summary>Позиция единственного вхождения ключа в диапазоне; -1 при повторе.</summary>
        private static Dictionary<string, int> UniquePositions(IList<PdfReviewWord> words,
            int start, int end, DiffContext context)
        {
            var positions = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int i = start; i < end; i++)
            {
                context.PollCancellation();
                string key = WordKey(words[i]);
                positions[key] = positions.ContainsKey(key) ? -1 : i;
            }
            return positions;
        }

        /// <summary>
        /// Linear-space Myers bisect для диапазона, который не помещается в LCS-матрицу.
        /// Встреча прямого и обратного фронтов даёт точку оптимального shortest-edit-path:
        /// вставка в начале не сдвигает весь оставшийся документ, а память остаётся O(n+m).
        /// </summary>
        private static bool TryBisect(IList<PdfReviewWord> a, int aStart, int aEnd,
            IList<PdfReviewWord> b, int bStart, int bEnd, DiffContext context,
            out int aMid, out int bMid)
        {
            return TryBisectCore(a, aStart, aEnd, b, bStart, bEnd,
                delegate(PdfReviewWord left, PdfReviewWord right)
                {
                    return WordKey(left) == WordKey(right);
                }, context, out aMid, out bMid);
        }

        private static bool TryBisectCore<T>(IList<T> a, int aStart, int aEnd,
            IList<T> b, int bStart, int bEnd, Func<T, T, bool> equal,
            DiffContext context, out int aMid, out int bMid)
        {
            aMid = aStart;
            bMid = bStart;
            int n = aEnd - aStart;
            int m = bEnd - bStart;
            long maxDLong = ((long)n + m + 1L) / 2L;
            if (maxDLong <= 0 || equal == null)
                return false;
            if (maxDLong > (int.MaxValue - 1L) / 2L)
            {
                if (context != null)
                    context.ExhaustWork();
                return false;
            }
            int maxD = (int)maxDLong;
            int offset = maxD;
            int length = 2 * maxD + 1;
            if (context != null && !context.TryReserve(length))
                return false;
            var forward = new int[length];
            var reverse = new int[length];
            for (int i = 0; i < length; i++)
            {
                if (context != null)
                    context.PollCancellation();
                forward[i] = reverse[i] = -1;
            }
            forward[offset + 1] = 0;
            reverse[offset + 1] = 0;

            int delta = n - m;
            bool overlapOnForward = (delta & 1) != 0;
            int forwardStart = 0, forwardEnd = 0;
            int reverseStart = 0, reverseEnd = 0;

            for (int d = 0; d < maxD; d++)
            {
                for (int k = -d + forwardStart; k <= d - forwardEnd; k += 2)
                {
                    if (context != null && !context.TryStep())
                        return false;
                    int index = offset + k;
                    int x = k == -d || (k != d && forward[index - 1] < forward[index + 1])
                        ? forward[index + 1] : forward[index - 1] + 1;
                    int y = x - k;
                    while (x < n && y < m)
                    {
                        if (context != null && !context.TryStep())
                            return false;
                        if (!equal(a[aStart + x], b[bStart + y]))
                            break;
                        x++;
                        y++;
                    }
                    forward[index] = x;
                    if (x > n)
                    {
                        forwardEnd += 2;
                    }
                    else if (y > m)
                    {
                        forwardStart += 2;
                    }
                    else if (overlapOnForward)
                    {
                        int reverseK = delta - k;
                        int reverseIndex = offset + reverseK;
                        if (reverseIndex >= 0 && reverseIndex < length &&
                            reverse[reverseIndex] >= 0 &&
                            x + reverse[reverseIndex] >= n)
                        {
                            aMid = aStart + x;
                            bMid = bStart + y;
                            return true;
                        }
                    }
                }

                for (int k = -d + reverseStart; k <= d - reverseEnd; k += 2)
                {
                    if (context != null && !context.TryStep())
                        return false;
                    int index = offset + k;
                    int x = k == -d || (k != d && reverse[index - 1] < reverse[index + 1])
                        ? reverse[index + 1] : reverse[index - 1] + 1;
                    int y = x - k;
                    while (x < n && y < m)
                    {
                        if (context != null && !context.TryStep())
                            return false;
                        if (!equal(a[aEnd - x - 1], b[bEnd - y - 1]))
                            break;
                        x++;
                        y++;
                    }
                    reverse[index] = x;
                    if (x > n)
                    {
                        reverseEnd += 2;
                    }
                    else if (y > m)
                    {
                        reverseStart += 2;
                    }
                    else if (!overlapOnForward)
                    {
                        int forwardK = delta - k;
                        int forwardIndex = offset + forwardK;
                        if (forwardIndex >= 0 && forwardIndex < length &&
                            forward[forwardIndex] >= 0 &&
                            forward[forwardIndex] + x >= n)
                        {
                            int forwardX = forward[forwardIndex];
                            aMid = aStart + forwardX;
                            bMid = bStart + forwardX - forwardK;
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Точный ЛСД-дифф середины, которая помещается в потолок матрицы. Длина LCS
        /// всегда первична; совпадения страниц служат только вторичным критерием среди
        /// путей той же длины.
        /// </summary>
        private static bool LcsCore(IList<PdfReviewWord> a, int aStart, int aEnd,
            IList<PdfReviewWord> b, int bStart, int bEnd, List<PdfReviewWordOp> ops,
            PreferredPageAlignment preferredPages, DiffContext context,
            int rows, int columns, long cells)
        {
            int n = aEnd - aStart;
            int m = bEnd - bStart;
            long work = SaturatingAdd(cells, (long)n + m);
            if (!context.TryReserve(work))
                return false;

            var dp = new int[rows, columns];
            int[,] preference = preferredPages == null ? null : new int[rows, columns];
            for (int i = n - 1; i >= 0; i--)
            {
                for (int j = m - 1; j >= 0; j--)
                {
                    context.PollCancellation();
                    int bestLength = dp[i + 1, j];
                    int bestPreference = preference == null ? 0 : preference[i + 1, j];

                    int insertLength = dp[i, j + 1];
                    int insertPreference = preference == null ? 0 : preference[i, j + 1];
                    if (insertLength > bestLength ||
                        (insertLength == bestLength && insertPreference > bestPreference))
                    {
                        bestLength = insertLength;
                        bestPreference = insertPreference;
                    }

                    if (WordKey(a[aStart + i]) == WordKey(b[bStart + j]))
                    {
                        int matchLength = dp[i + 1, j + 1] + 1;
                        int matchPreference = preference == null ? 0 :
                            preference[i + 1, j + 1] +
                            (preferredPages.IsPreferred(a[aStart + i], b[bStart + j])
                                ? 1 : 0);
                        if (matchLength > bestLength ||
                            (matchLength == bestLength && matchPreference > bestPreference))
                        {
                            bestLength = matchLength;
                            bestPreference = matchPreference;
                        }
                    }

                    dp[i, j] = bestLength;
                    if (preference != null)
                        preference[i, j] = bestPreference;
                }
            }

            int li = 0;
            int ri = 0;
            while (li < n || ri < m)
            {
                context.PollCancellation();
                int currentPreference = preference == null ? 0 : preference[li, ri];
                int matchPreference = 0;
                bool canMatch = li < n && ri < m &&
                    WordKey(a[aStart + li]) == WordKey(b[bStart + ri]);
                if (canMatch && preference != null)
                {
                    matchPreference = preference[li + 1, ri + 1] +
                        (preferredPages.IsPreferred(a[aStart + li], b[bStart + ri])
                            ? 1 : 0);
                }
                if (canMatch && dp[li, ri] == dp[li + 1, ri + 1] + 1 &&
                    currentPreference == matchPreference)
                {
                    AppendEqual(ops, a[aStart + li], b[bStart + ri],
                        PdfReviewMatchKind.Exact);
                    li++;
                    ri++;
                }
                else
                {
                    int deletePreference = li < n && preference != null
                        ? preference[li + 1, ri] : 0;
                    if (ri >= m || (li < n && dp[li, ri] == dp[li + 1, ri] &&
                        currentPreference == deletePreference))
                    {
                        Append(ops, PdfReviewDiffKind.Delete, a[aStart + li]);
                        li++;
                    }
                    else
                    {
                        Append(ops, PdfReviewDiffKind.Insert, b[bStart + ri]);
                        ri++;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Ручное сопоставление меняет только строки viewer. Целевые страницы образуют пару,
        /// а их прежние контрагенты остаются one-sided, поэтому ни одна физическая страница
        /// не исчезает из представления и сквозной diff не пересчитывается.
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
                    PdfReviewPagePair source = pairs[i];
                    bool removeLeft = source.LeftPageIndex == leftPageIndex;
                    bool removeRight = source.RightPageIndex == rightPageIndex;
                    if (!removeLeft && !removeRight)
                    {
                        result.Add(source);
                        continue;
                    }

                    if (insertAt > result.Count)
                        insertAt = result.Count;
                    int remainingLeft = removeLeft ? -1 : source.LeftPageIndex;
                    int remainingRight = removeRight ? -1 : source.RightPageIndex;
                    if (remainingLeft >= 0 || remainingRight >= 0)
                    {
                        result.Add(new PdfReviewPagePair
                        {
                            LeftPageIndex = remainingLeft,
                            RightPageIndex = remainingRight,
                            Status = remainingLeft >= 0
                                ? PdfReviewPairStatus.LeftOnly : PdfReviewPairStatus.RightOnly,
                            Similarity = 0
                        });
                    }
                }
            }
            if (insertAt < 0 || insertAt > result.Count)
                insertAt = result.Count;
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

        /// <summary>Добавить одно одностороннее слово, не теряя владельца.</summary>
        internal static void Append(List<PdfReviewWordOp> ops, PdfReviewDiffKind kind,
            PdfReviewWord word)
        {
            if (kind == PdfReviewDiffKind.Equal)
            {
                // Совместимость только для искусственных вызовов: production diff всегда
                // передаёт фактический объект каждой стороны через AppendEqual.
                AppendEqual(ops, word, word, PdfReviewMatchKind.Exact);
                return;
            }

            PdfReviewWordOp op = LastCompatible(ops, kind, PdfReviewMatchKind.None);
            if (op == null)
            {
                op = new PdfReviewWordOp { Kind = kind };
                ops.Add(op);
            }
            if (kind == PdfReviewDiffKind.Delete)
                op.LeftWords.Add(word);
            else
                op.RightWords.Add(word);
        }

        /// <summary>Добавить реальную пару Equal с её происхождением.</summary>
        internal static void AppendEqual(List<PdfReviewWordOp> ops, PdfReviewWord left,
            PdfReviewWord right, PdfReviewMatchKind matchKind)
        {
            if (matchKind == PdfReviewMatchKind.None)
                throw new ArgumentException("Equal match must have provenance.", "matchKind");
            PdfReviewWordOp op = LastCompatible(ops, PdfReviewDiffKind.Equal, matchKind);
            if (op == null)
            {
                op = new PdfReviewWordOp
                {
                    Kind = PdfReviewDiffKind.Equal,
                    MatchKind = matchKind
                };
                ops.Add(op);
            }
            op.LeftWords.Add(left);
            op.RightWords.Add(right);
            op.Matches.Add(new PdfReviewWordMatch
            {
                Left = left,
                Right = right,
                Kind = matchKind
            });
        }

        /// <summary>
        /// Добавить уже сформированную операцию. Списки сторон и явные связи копируются
        /// целиком; метод используется всеми post-diff перестройками.
        /// </summary>
        internal static void AppendOperation(List<PdfReviewWordOp> ops,
            PdfReviewWordOp source)
        {
            if (source == null)
                return;
            PdfReviewWordOp target = LastCompatible(ops, source.Kind, source.MatchKind);
            if (target == null)
            {
                target = new PdfReviewWordOp
                {
                    Kind = source.Kind,
                    MatchKind = source.MatchKind
                };
                ops.Add(target);
            }
            target.LeftWords.AddRange(source.LeftWords);
            target.RightWords.AddRange(source.RightWords);
            target.SplitJoinLeftBoundary = source.SplitJoinLeftBoundary;
            target.SplitJoinRightBoundary = source.SplitJoinRightBoundary;
            foreach (PdfReviewWordMatch match in source.Matches)
                if (match != null)
                    target.Matches.Add(new PdfReviewWordMatch
                    {
                        Left = match.Left,
                        Right = match.Right,
                        Kind = match.Kind
                    });
        }

        private static PdfReviewWordOp LastCompatible(List<PdfReviewWordOp> ops,
            PdfReviewDiffKind kind, PdfReviewMatchKind matchKind)
        {
            if (ops == null)
                throw new ArgumentNullException("ops");
            if (ops.Count == 0 ||
                (kind == PdfReviewDiffKind.Equal &&
                 (matchKind == PdfReviewMatchKind.SplitJoin ||
                  matchKind == PdfReviewMatchKind.MixedOrder)))
                return null;
            PdfReviewWordOp last = ops[ops.Count - 1];
            if (last == null || last.Kind != kind)
                return null;
            return kind != PdfReviewDiffKind.Equal || last.MatchKind == matchKind
                ? last : null;
        }

        private const int ReconcileMaxWordsPerSide = 20;
        private const int ReconcileMaxItemSpan = 48;
        private const int ReconcileAnchorNeighborhood = 32;
        private const int ReconcileMinInternalBlockCluster = 3;
        private const double ReconcileMaxWordAreaShare = 0.08;
        private const double ReconcileMaxBoundsAreaShare = 0.55;
        private const double ReconcileMaxWitnessAxisShare = 0.25;
        private const double ReconcilePageMarginShare = 0.08;
        private const double ReconcileInternalRegistrationHeightShare = 1.0;
        private const double ReconcileMinPageScale = 0.80;
        private const double ReconcileMaxPageScale = 1.25;

        private sealed class ReconcileItem
        {
            public PdfReviewDiffKind Kind;
            public PdfReviewMatchKind MatchKind;
            public PdfReviewWord Left;
            public PdfReviewWord Right;
            public PdfReviewWordOp Barrier;

            public bool IsExact
            {
                get
                {
                    return Barrier == null && Kind == PdfReviewDiffKind.Equal &&
                        MatchKind == PdfReviewMatchKind.Exact && Left != null && Right != null;
                }
            }
        }

        private sealed class ReconcileRegion
        {
            public int Start;
            public int End;
            public int LeftPage;
            public int RightPage;
        }

        private sealed class RepeatedExactRepair
        {
            public ReconcileItem Exact;
            public PdfReviewWord Left;
            public PdfReviewWord Right;
            public PdfReviewWord ReleasedLeft;
            public PdfReviewWord ReleasedRight;
        }

        private sealed class ReconcileNode
        {
            public PdfReviewWord Left;
            public PdfReviewWord Right;
            public PdfReviewMatchKind MatchKind;
            public int LeftOrder = -1;
            public int RightOrder = -1;
            public int OriginalOrder = int.MaxValue;
            public readonly List<int> Edges = new List<int>();
        }

        private struct ReconcileTopoEntry : IComparable<ReconcileTopoEntry>
        {
            public int Order;
            public int Component;

            public int CompareTo(ReconcileTopoEntry other)
            {
                int comparison = Order.CompareTo(other.Order);
                return comparison != 0 ? comparison : Component.CompareTo(other.Component);
            }
        }

        /// <summary>
        /// Узкая semantic-cleanup после единственного глобального diff. Сначала она объединяет
        /// пересекающиеся Delete/Insert seeds одной физической page-pair: в таблице одна
        /// extraction-order перестановка закономерно даёт несколько частичных LCS-окон. Затем
        /// весь малый компонент обязан иметь одинаковое NFC-мультимножество, единственное
        /// двустороннее геометрическое matching и согласованное соответствие trusted blocks.
        /// При перестройке SCC сохраняют точный порядок обеих документов. Exact-пара другой
        /// физической page-pair является границей кандидата: она не может быть поглощена
        /// локальным циклом. Любая неоднозначность оставляет исходные Delete/Insert без изменений.
        /// </summary>
        internal static void ReconcileOrderArtifacts(PdfReviewResult result)
        {
            ReconcileOrderArtifacts(result, PdfReviewLimits.Default(), null);
        }

        internal static void ReconcileOrderArtifacts(PdfReviewResult result,
            PdfReviewLimits limits, Func<bool> cancelled)
        {
            if (limits == null)
                limits = PdfReviewLimits.Default();
            ReconcileOrderArtifacts(result, new ReviewPageLookup(result),
                new DiffContext(limits, cancelled));
        }

        private static void ReconcileOrderArtifacts(PdfReviewResult result,
            ReviewPageLookup pages, DiffContext context)
        {
            context.ThrowIfCancellation();
            if (result == null || result.Left == null || result.Right == null ||
                result.Operations.Count == 0)
                return;
            List<ReconcileItem> items;
            if (!TryReconcileItems(result.Operations, context, out items) ||
                items.Count == 0)
                return;

            List<ReconcileRegion> regions = ReconcileSeedRegions(pages, items, context);
            if (context.WorkExhausted)
                return;

            var rightByLeft = new Dictionary<PdfReviewWord, PdfReviewWord>();
            var leftByRight = new Dictionary<PdfReviewWord, PdfReviewWord>();
            var releasedLeft = new HashSet<PdfReviewWord>();
            var releasedRight = new HashSet<PdfReviewWord>();
            foreach (ReconcileRegion region in regions)
            {
                if (!context.TryStep())
                    return;
                Dictionary<PdfReviewWord, PdfReviewWord> component;
                if (!TryReconcileRegion(pages, items, region, context, out component))
                {
                    if (context.WorkExhausted)
                        return;
                    continue;
                }
                bool conflict = false;
                foreach (KeyValuePair<PdfReviewWord, PdfReviewWord> pair in component)
                {
                    if (!context.TryStep())
                        return;
                    PdfReviewWord existing;
                    if (pair.Key == null || pair.Value == null ||
                        !pages.IsExplicitPair(pair.Key.PageIndex,
                            pair.Value.PageIndex, context) ||
                        (rightByLeft.TryGetValue(pair.Key, out existing) &&
                         !ReferenceEquals(existing, pair.Value)) ||
                        (leftByRight.TryGetValue(pair.Value, out existing) &&
                         !ReferenceEquals(existing, pair.Key)))
                    {
                        conflict = true;
                        break;
                    }
                }
                if (context.WorkExhausted)
                    return;
                if (conflict)
                    continue;
                foreach (KeyValuePair<PdfReviewWord, PdfReviewWord> pair in component)
                {
                    if (!context.TryStep())
                        return;
                    rightByLeft[pair.Key] = pair.Value;
                    leftByRight[pair.Value] = pair.Key;
                }
            }
            if (!TryAddRepeatedExactRepairs(pages, items, rightByLeft, leftByRight,
                    releasedLeft, releasedRight, context) || context.WorkExhausted ||
                rightByLeft.Count == 0)
                return;

            List<PdfReviewWordOp> rebuilt;
            if (!TryRebuildReconciledOperations(result.Operations, items, rightByLeft,
                leftByRight, releasedLeft, releasedRight, context, out rebuilt) ||
                context.WorkExhausted || !context.TryReserve(rebuilt.Count))
                return;
            context.ThrowIfCancellation();
            result.ReplaceOperations(rebuilt);
        }

        private static List<ReconcileRegion> ReconcileSeedRegions(ReviewPageLookup pages,
            IList<ReconcileItem> items, DiffContext context)
        {
            var seeds = new List<ReconcileRegion>();
            for (int i = 0; i < items.Count; i++)
            {
                if (!context.TryStep())
                    return seeds;
                ReconcileItem deletion = items[i];
                if (deletion.Barrier != null || deletion.Kind != PdfReviewDiffKind.Delete ||
                    deletion.Left == null || deletion.Left.PageIndex < 0)
                    continue;
                int firstCandidate = Math.Max(0, i - ReconcileMaxItemSpan);
                int lastCandidate = Math.Min(items.Count - 1,
                    i + ReconcileMaxItemSpan);
                for (int j = firstCandidate; j <= lastCandidate; j++)
                {
                    if (!context.TryStep())
                        return seeds;
                    ReconcileItem insertion = items[j];
                    if (insertion.Barrier != null ||
                        insertion.Kind != PdfReviewDiffKind.Insert || insertion.Right == null ||
                        insertion.Right.PageIndex < 0 ||
                        WordKey(deletion.Left) != WordKey(insertion.Right) ||
                        !pages.IsExplicitPair(deletion.Left.PageIndex,
                            insertion.Right.PageIndex, context))
                        continue;
                    int first = Math.Min(i, j);
                    int last = Math.Max(i, j);
                    seeds.Add(new ReconcileRegion
                    {
                        Start = first,
                        End = last,
                        LeftPage = deletion.Left.PageIndex,
                        RightPage = insertion.Right.PageIndex
                    });
                }
            }
            if (context.WorkExhausted || seeds.Count == 0)
                return seeds;
            if (!context.TryReserveSort(seeds.Count))
                return seeds;
            seeds.Sort(delegate(ReconcileRegion a, ReconcileRegion b)
            {
                int comparison = a.LeftPage.CompareTo(b.LeftPage);
                if (comparison != 0) return comparison;
                comparison = a.RightPage.CompareTo(b.RightPage);
                if (comparison != 0) return comparison;
                comparison = a.Start.CompareTo(b.Start);
                return comparison != 0 ? comparison : a.End.CompareTo(b.End);
            });

            var merged = new List<ReconcileRegion>();
            foreach (ReconcileRegion seed in seeds)
            {
                if (!context.TryStep())
                    return merged;
                ReconcileRegion last = merged.Count == 0 ? null : merged[merged.Count - 1];
                // В зеркальном diff соседние половины одного цикла могут закончиться
                // Delete и сразу начаться Insert без общего item. Склеиваем только
                // касающиеся seed того же физического page-pair; любой Exact/barrier
                // между ними по-прежнему оставляет хотя бы один элемент разрыва.
                if (last != null && last.LeftPage == seed.LeftPage &&
                    last.RightPage == seed.RightPage && seed.Start <= last.End + 1)
                {
                    last.End = Math.Max(last.End, seed.End);
                }
                else
                {
                    merged.Add(new ReconcileRegion
                    {
                        Start = seed.Start,
                        End = seed.End,
                        LeftPage = seed.LeftPage,
                        RightPage = seed.RightPage
                    });
                }
            }
            return merged;
        }

        /// <summary>
        /// LCS может связать позднюю копию повторяющегося ключа с физически удалённой
        /// ранней копией, оставив настоящий counterpart как Delete/Insert. Перепривязка
        /// допустима только когда две независимые Exact-пары того же trusted block-pair
        /// задают единственную локальную трансляцию. Старая Exact-связь при этом обязана
        /// противоречить этой трансляции и соответствию блоков. Неоднозначные предложения
        /// полностью отбрасываются; реально лишняя старая копия остаётся изменением.
        /// </summary>
        private static bool TryAddRepeatedExactRepairs(ReviewPageLookup pages,
            IList<ReconcileItem> items,
            IDictionary<PdfReviewWord, PdfReviewWord> rightByLeft,
            IDictionary<PdfReviewWord, PdfReviewWord> leftByRight,
            ISet<PdfReviewWord> releasedLeft,
            ISet<PdfReviewWord> releasedRight, DiffContext context)
        {
            if (pages == null || items == null || rightByLeft == null ||
                leftByRight == null || releasedLeft == null ||
                releasedRight == null)
                return false;

            var proposals = new List<RepeatedExactRepair>();
            for (int exactIndex = 0; exactIndex < items.Count; exactIndex++)
            {
                if (!context.TryStep())
                    return false;
                ReconcileItem exact = items[exactIndex];
                if (exact == null || !exact.IsExact)
                    continue;
                int first = Math.Max(0, exactIndex - ReconcileMaxItemSpan);
                int last = Math.Min(items.Count - 1,
                    exactIndex + ReconcileMaxItemSpan);
                for (int changedIndex = first; changedIndex <= last; changedIndex++)
                {
                    if (!context.TryStep())
                        return false;
                    ReconcileItem changed = items[changedIndex];
                    bool deletion = changed != null && changed.Barrier == null &&
                        changed.Kind == PdfReviewDiffKind.Delete && changed.Left != null;
                    bool insertion = changed != null && changed.Barrier == null &&
                        changed.Kind == PdfReviewDiffKind.Insert && changed.Right != null;
                    if (!deletion && !insertion)
                        continue;
                    PdfReviewWord changedWord = deletion ? changed.Left : changed.Right;
                    if (WordKey(changedWord) != WordKey(exact.Left) ||
                        !RepeatedExactRepairCorroborated(pages, items, exactIndex,
                            changedIndex, deletion, context))
                    {
                        if (context.WorkExhausted)
                            return false;
                        continue;
                    }
                    proposals.Add(new RepeatedExactRepair
                    {
                        Exact = exact,
                        Left = deletion ? changed.Left : exact.Left,
                        Right = deletion ? exact.Right : changed.Right,
                        ReleasedLeft = deletion ? exact.Left : null,
                        ReleasedRight = deletion ? null : exact.Right
                    });
                }
            }
            if (proposals.Count == 0)
                return !context.WorkExhausted;

            var exactUses = new Dictionary<ReconcileItem, int>();
            var leftUses = new Dictionary<PdfReviewWord, int>();
            var rightUses = new Dictionary<PdfReviewWord, int>();
            foreach (RepeatedExactRepair repair in proposals)
            {
                if (!context.TryStep())
                    return false;
                IncrementReconcileUse(exactUses, repair.Exact);
                IncrementReconcileUse(leftUses, repair.Left);
                IncrementReconcileUse(rightUses, repair.Right);
                if (repair.ReleasedLeft != null)
                    IncrementReconcileUse(leftUses, repair.ReleasedLeft);
                if (repair.ReleasedRight != null)
                    IncrementReconcileUse(rightUses, repair.ReleasedRight);
            }

            foreach (RepeatedExactRepair repair in proposals)
            {
                if (!context.TryStep())
                    return false;
                if (!UniqueReconcileUse(exactUses, repair.Exact) ||
                    !UniqueReconcileUse(leftUses, repair.Left) ||
                    !UniqueReconcileUse(rightUses, repair.Right) ||
                    !UniqueReconcileUse(leftUses, repair.ReleasedLeft) ||
                    !UniqueReconcileUse(rightUses, repair.ReleasedRight) ||
                    rightByLeft.ContainsKey(repair.Left) ||
                    leftByRight.ContainsKey(repair.Right) ||
                    (repair.ReleasedLeft != null &&
                     (rightByLeft.ContainsKey(repair.ReleasedLeft) ||
                      releasedLeft.Contains(repair.ReleasedLeft))) ||
                    (repair.ReleasedRight != null &&
                     (leftByRight.ContainsKey(repair.ReleasedRight) ||
                      releasedRight.Contains(repair.ReleasedRight))) ||
                    releasedLeft.Contains(repair.Left) ||
                    releasedRight.Contains(repair.Right))
                    continue;

                rightByLeft.Add(repair.Left, repair.Right);
                leftByRight.Add(repair.Right, repair.Left);
                if (repair.ReleasedLeft != null)
                    releasedLeft.Add(repair.ReleasedLeft);
                if (repair.ReleasedRight != null)
                    releasedRight.Add(repair.ReleasedRight);
            }
            return !context.WorkExhausted;
        }

        private static bool RepeatedExactRepairCorroborated(ReviewPageLookup pages,
            IList<ReconcileItem> items, int exactIndex, int changedIndex,
            bool deletion, DiffContext context)
        {
            if (pages == null || items == null || exactIndex < 0 ||
                exactIndex >= items.Count || changedIndex < 0 ||
                changedIndex >= items.Count || exactIndex == changedIndex ||
                Math.Abs((long)exactIndex - changedIndex) > ReconcileMaxItemSpan)
                return false;
            ReconcileItem exact = items[exactIndex];
            ReconcileItem changed = items[changedIndex];
            if (exact == null || !exact.IsExact || changed == null ||
                changed.Barrier != null)
                return false;

            PdfReviewWord left = deletion ? changed.Left : exact.Left;
            PdfReviewWord right = deletion ? exact.Right : changed.Right;
            if ((deletion && (changed.Kind != PdfReviewDiffKind.Delete ||
                    changed.Left == null)) ||
                (!deletion && (changed.Kind != PdfReviewDiffKind.Insert ||
                    changed.Right == null)) || left == null || right == null ||
                left.PageIndex != exact.Left.PageIndex ||
                right.PageIndex != exact.Right.PageIndex ||
                left.BlockId < 0 || right.BlockId < 0 ||
                exact.Left.BlockId < 0 || exact.Right.BlockId < 0 ||
                WordKey(left).Length == 0 || WordKey(left) != WordKey(right) ||
                WordKey(exact.Left) != WordKey(exact.Right) ||
                (deletion && left.BlockId == exact.Left.BlockId) ||
                (!deletion && right.BlockId == exact.Right.BlockId) ||
                !ValidReviewBox(left.Box) || !ValidReviewBox(right.Box) ||
                !ValidReviewBox(exact.Left.Box) ||
                !ValidReviewBox(exact.Right.Box) ||
                !pages.IsExplicitPair(left.PageIndex, right.PageIndex, context))
                return false;

            PdfReviewPage leftPage = pages.Page(true, left.PageIndex);
            PdfReviewPage rightPage = pages.Page(false, right.PageIndex);
            double scaleX, scaleY, reverseScaleX, reverseScaleY;
            if (leftPage == null || rightPage == null ||
                !TrySymmetricPageScale(leftPage, rightPage, out scaleX, out scaleY,
                    out reverseScaleX, out reverseScaleY) ||
                !StructuredGeometryShape(left, right, scaleX, scaleY,
                    reverseScaleX, reverseScaleY))
                return false;

            int leftKeyCount, rightKeyCount;
            if (!TryBlockKeyCount(leftPage, left.BlockId, WordKey(left), context,
                    out leftKeyCount) ||
                !TryBlockKeyCount(rightPage, right.BlockId, WordKey(right), context,
                    out rightKeyCount) || leftKeyCount != 1 || rightKeyCount != 1)
                return false;

            int first, last;
            if (!TryRepeatedExactLocalBounds(items, changedIndex, left.PageIndex,
                    right.PageIndex, context, out first, out last) ||
                exactIndex < first || exactIndex > last)
                return false;

            var clusterLeft = new List<PdfReviewWord> { left };
            var clusterRight = new List<PdfReviewWord> { right };
            var witnessKeys = new HashSet<string>(StringComparer.Ordinal);
            for (int i = first; i <= last; i++)
            {
                if (!context.TryStep())
                    return false;
                ReconcileItem witness = items[i];
                if (i == exactIndex || witness == null || !witness.IsExact ||
                    witness.Left.PageIndex != left.PageIndex ||
                    witness.Right.PageIndex != right.PageIndex ||
                    witness.Left.BlockId != left.BlockId ||
                    witness.Right.BlockId != right.BlockId ||
                    WordKey(witness.Left).Length == 0 ||
                    WordKey(witness.Left) != WordKey(witness.Right) ||
                    !StableReconcileAnchor(items, i, context) ||
                    !StructuredGeometryShape(witness.Left, witness.Right, scaleX,
                        scaleY, reverseScaleX, reverseScaleY) ||
                    !TranslationsConsistent(left, right, witness.Left,
                        witness.Right, scaleX, scaleY, reverseScaleX,
                        reverseScaleY))
                {
                    if (context.WorkExhausted)
                        return false;
                    continue;
                }
                string witnessKey = WordKey(witness.Left);
                if (!TryBlockKeyCount(leftPage, left.BlockId, witnessKey, context,
                        out leftKeyCount) ||
                    !TryBlockKeyCount(rightPage, right.BlockId, witnessKey,
                        context, out rightKeyCount))
                    return false;
                if (leftKeyCount != 1 || rightKeyCount != 1 ||
                    !witnessKeys.Add(witnessKey))
                    continue;
                clusterLeft.Add(witness.Left);
                clusterRight.Add(witness.Right);
                if (clusterLeft.Count > ReconcileMaxWordsPerSide)
                    return false;
            }

            return !context.WorkExhausted &&
                clusterLeft.Count >= ReconcileMinInternalBlockCluster &&
                witnessKeys.Count >= ReconcileMinInternalBlockCluster - 1 &&
                !TranslationsConsistent(left, right, exact.Left, exact.Right,
                    scaleX, scaleY, reverseScaleX, reverseScaleY) &&
                WindowGeometryBounded(clusterLeft, leftPage, context) &&
                WindowGeometryBounded(clusterRight, rightPage, context);
        }

        private static bool TryRepeatedExactLocalBounds(IList<ReconcileItem> items,
            int center, int leftPage, int rightPage, DiffContext context,
            out int first, out int last)
        {
            first = last = center;
            if (items == null || center < 0 || center >= items.Count ||
                !RepeatedExactPairItem(items[center], leftPage, rightPage))
                return false;
            int minimum = Math.Max(0, center - ReconcileMaxItemSpan);
            for (int i = center - 1; i >= minimum; i--)
            {
                if (!context.TryStep())
                    return false;
                if (!RepeatedExactPairItem(items[i], leftPage, rightPage))
                    break;
                first = i;
            }
            int maximum = Math.Min(items.Count - 1,
                center + ReconcileMaxItemSpan);
            for (int i = center + 1; i <= maximum; i++)
            {
                if (!context.TryStep())
                    return false;
                if (!RepeatedExactPairItem(items[i], leftPage, rightPage))
                    break;
                last = i;
            }
            return !context.WorkExhausted;
        }

        private static bool RepeatedExactPairItem(ReconcileItem item,
            int leftPage, int rightPage)
        {
            return item != null && item.Barrier == null &&
                (item.Left != null || item.Right != null) &&
                (item.Left == null || item.Left.PageIndex == leftPage) &&
                (item.Right == null || item.Right.PageIndex == rightPage);
        }

        private static bool TryBlockKeyCount(PdfReviewPage page, int blockId,
            string key, DiffContext context, out int count)
        {
            count = 0;
            if (page == null || page.Words == null || blockId < 0 ||
                string.IsNullOrEmpty(key))
                return false;
            foreach (PdfReviewWord word in page.Words)
            {
                if (!context.TryStep())
                    return false;
                if (word != null && word.BlockId == blockId &&
                    WordKey(word) == key)
                    count++;
            }
            return !context.WorkExhausted;
        }

        private static void IncrementReconcileUse<TKey>(
            IDictionary<TKey, int> counts, TKey key)
        {
            int count;
            counts.TryGetValue(key, out count);
            counts[key] = count + 1;
        }

        private static bool UniqueReconcileUse<TKey>(
            IDictionary<TKey, int> counts, TKey key)
        {
            if (object.ReferenceEquals(key, null))
                return true;
            int count;
            return counts.TryGetValue(key, out count) && count == 1;
        }

        private static bool TryReconcileRegion(ReviewPageLookup pages,
            IList<ReconcileItem> items, ReconcileRegion region, DiffContext context,
            out Dictionary<PdfReviewWord, PdfReviewWord> rightByLeft)
        {
            rightByLeft = null;
            if (region == null || region.Start < 0 || region.End < region.Start ||
                region.End >= items.Count || region.End - region.Start > ReconcileMaxItemSpan)
                return false;
            var left = new List<PdfReviewWord>();
            var right = new List<PdfReviewWord>();
            var exactRightByLeft = new Dictionary<PdfReviewWord, PdfReviewWord>();
            bool hasDelete = false, hasInsert = false;
            for (int i = region.Start; i <= region.End; i++)
            {
                if (!context.TryStep())
                    return false;
                ReconcileItem item = items[i];
                if (item.Barrier != null)
                    return false;
                if (item.Kind == PdfReviewDiffKind.Delete && item.Left != null &&
                    item.Left.PageIndex == region.LeftPage)
                {
                    left.Add(item.Left);
                    hasDelete = true;
                }
                else if (item.Kind == PdfReviewDiffKind.Insert && item.Right != null &&
                    item.Right.PageIndex == region.RightPage)
                {
                    right.Add(item.Right);
                    hasInsert = true;
                }
                else if (item.IsExact)
                {
                    // Двусторонняя Exact-связь другой физической page-pair задаёт
                    // независимый порядок документов. Локальный цикл не вправе
                    // перестраивать её или искать соответствие «сквозь» неё.
                    if (item.Left.PageIndex != region.LeftPage ||
                        item.Right.PageIndex != region.RightPage)
                        return false;
                    left.Add(item.Left);
                    right.Add(item.Right);
                    exactRightByLeft.Add(item.Left, item.Right);
                }
            }
            if (!hasDelete || !hasInsert || left.Count == 0 ||
                left.Count != right.Count || left.Count > ReconcileMaxWordsPerSide ||
                !SameKeyMultiset(left, right, context))
                return false;

            PdfReviewPage leftOwner = pages.Page(true, region.LeftPage);
            PdfReviewPage rightOwner = pages.Page(false, region.RightPage);
            if (leftOwner == null || rightOwner == null ||
                !WindowGeometryBounded(left, leftOwner, context) ||
                !WindowGeometryBounded(right, rightOwner, context))
                return false;

            int[] matching;
            if (!UniqueRegionMatching(items, region, left, right, leftOwner, rightOwner,
                    context, out matching) ||
                !ConsistentRegionBlocks(items, region, left, right, matching,
                    leftOwner, rightOwner, context))
                return false;
            rightByLeft = new Dictionary<PdfReviewWord, PdfReviewWord>();
            for (int i = 0; i < left.Count; i++)
            {
                if (!context.TryStep())
                    return false;
                PdfReviewWord exactRight;
                PdfReviewWord matchedRight = right[matching[i]];
                if (exactRightByLeft.TryGetValue(left[i], out exactRight) &&
                    ReferenceEquals(exactRight, matchedRight))
                    continue;
                rightByLeft.Add(left[i], matchedRight);
            }
            return rightByLeft.Count > 0;
        }

        private static bool TryReconcileItems(IList<PdfReviewWordOp> operations,
            DiffContext context, out List<ReconcileItem> items)
        {
            items = new List<ReconcileItem>();
            foreach (PdfReviewWordOp op in operations)
            {
                if (!context.TryStep())
                    return false;
                if (op == null)
                    return false;
                if (op.Kind == PdfReviewDiffKind.Delete)
                {
                    if (op.RightWords.Count != 0)
                        return false;
                    foreach (PdfReviewWord word in op.LeftWords)
                    {
                        if (!context.TryStep())
                            return false;
                        items.Add(new ReconcileItem
                        {
                            Kind = PdfReviewDiffKind.Delete,
                            Left = word
                        });
                    }
                }
                else if (op.Kind == PdfReviewDiffKind.Insert)
                {
                    if (op.LeftWords.Count != 0)
                        return false;
                    foreach (PdfReviewWord word in op.RightWords)
                    {
                        if (!context.TryStep())
                            return false;
                        items.Add(new ReconcileItem
                        {
                            Kind = PdfReviewDiffKind.Insert,
                            Right = word
                        });
                    }
                }
                else if (op.MatchKind == PdfReviewMatchKind.Exact &&
                    op.LeftWords.Count == op.RightWords.Count &&
                    op.Matches.Count == op.LeftWords.Count && op.LeftWords.Count > 0)
                {
                    for (int i = 0; i < op.LeftWords.Count; i++)
                    {
                        if (!context.TryStep())
                            return false;
                        PdfReviewWordMatch match = op.Matches[i];
                        if (match == null || match.Kind != PdfReviewMatchKind.Exact ||
                            !object.ReferenceEquals(match.Left, op.LeftWords[i]) ||
                            !object.ReferenceEquals(match.Right, op.RightWords[i]))
                            return false;
                        items.Add(new ReconcileItem
                        {
                            Kind = PdfReviewDiffKind.Equal,
                            MatchKind = PdfReviewMatchKind.Exact,
                            Left = match.Left,
                            Right = match.Right
                        });
                    }
                }
                else
                {
                    // Raster/group Equal и пустой sentinel являются непрозрачной границей.
                    items.Add(new ReconcileItem
                    {
                        Kind = PdfReviewDiffKind.Equal,
                        MatchKind = op.MatchKind,
                        Barrier = op
                    });
                }
            }
            return true;
        }

        /// <summary>
        /// Два устойчивых Exact-свидетеля задают локальную регистрацию. Обычно они находятся
        /// снаружи окна. У края физической страницы один свидетель может оказаться внутри
        /// перестановки; тогда он допустим только как неизменённая фактическая пара и позже
        /// обязан остаться тем же ребром unique matching.
        /// </summary>
        private static bool TryRegionAnchors(IList<ReconcileItem> items,
            ReconcileRegion region, DiffContext context, out ReconcileItem before,
            out ReconcileItem after)
        {
            before = null;
            after = null;
            if (items == null || region == null || region.Start < 0 ||
                region.End < region.Start || region.End >= items.Count)
                return false;

            int first = Math.Max(0, region.Start - ReconcileAnchorNeighborhood);
            for (int i = region.Start - 1; i >= first; i--)
            {
                if (!context.TryStep())
                    return false;
                ReconcileItem item = items[i];
                if (item.Barrier != null || (item.IsExact &&
                    !SameReconcilePagePair(item, region)))
                    break;
                if (SameReconcilePagePair(item, region) &&
                    StableReconcileAnchor(items, i, context))
                {
                    before = item;
                    break;
                }
            }
            int last = Math.Min(items.Count - 1,
                region.End + ReconcileAnchorNeighborhood);
            for (int i = region.End + 1; i <= last; i++)
            {
                if (!context.TryStep())
                    return false;
                ReconcileItem item = items[i];
                if (item.Barrier != null || (item.IsExact &&
                    !SameReconcilePagePair(item, region)))
                    break;
                if (SameReconcilePagePair(item, region) &&
                    StableReconcileAnchor(items, i, context))
                {
                    after = item;
                    break;
                }
            }

            // У последнего/первого extraction-order цикла страницы внешнего anchor
            // закономерно нет. Берём ближайшую устойчивую Exact-пару внутри окна, но
            // никогда не заменяем ею уже найденное двустороннее окружение.
            if (before == null)
                for (int i = region.Start; i <= region.End; i++)
                {
                    if (!context.TryStep())
                        return false;
                    if (SameReconcilePagePair(items[i], region) &&
                        StableReconcileAnchor(items, i, context) &&
                        !SameReconcileAnchor(items[i], after))
                    {
                        before = items[i];
                        break;
                    }
                }
            if (after == null)
                for (int i = region.End; i >= region.Start; i--)
                {
                    if (!context.TryStep())
                        return false;
                    if (SameReconcilePagePair(items[i], region) &&
                        StableReconcileAnchor(items, i, context) &&
                        !SameReconcileAnchor(items[i], before))
                    {
                        after = items[i];
                        break;
                    }
                }
            return !context.WorkExhausted && before != null && after != null &&
                !SameReconcileAnchor(before, after);
        }

        private static bool SameReconcilePagePair(ReconcileItem item,
            ReconcileRegion region)
        {
            return item != null && region != null && item.IsExact &&
                item.Left.PageIndex == region.LeftPage &&
                item.Right.PageIndex == region.RightPage;
        }

        private static bool SameReconcileAnchor(ReconcileItem first,
            ReconcileItem second)
        {
            return first != null && second != null &&
                ReferenceEquals(first.Left, second.Left) &&
                ReferenceEquals(first.Right, second.Right);
        }

        private static ReconcileItem ReverseReconcileAnchor(ReconcileItem item)
        {
            return item == null ? null : new ReconcileItem
            {
                Kind = PdfReviewDiffKind.Equal,
                MatchKind = PdfReviewMatchKind.Exact,
                Left = item.Right,
                Right = item.Left
            };
        }

        private static bool TryRegionRegistrations(IList<ReconcileItem> items,
            ReconcileRegion region, PdfReviewPage leftPage, PdfReviewPage rightPage,
            DiffContext context, out ReconcileItem before, out ReconcileItem after,
            out double scaleX, out double scaleY, out double dx, out double dy,
            out double reverseScaleX, out double reverseScaleY,
            out double reverseDx, out double reverseDy)
        {
            before = null;
            after = null;
            scaleX = scaleY = reverseScaleX = reverseScaleY = 1;
            dx = dy = reverseDx = reverseDy = 0;
            if (!TryRegionAnchors(items, region, context, out before, out after) ||
                !TryLocalRegistration(before, after, leftPage, rightPage,
                    out scaleX, out scaleY, out dx, out dy))
                return false;
            ReconcileItem reverseBefore = ReverseReconcileAnchor(before);
            ReconcileItem reverseAfter = ReverseReconcileAnchor(after);
            return TryLocalRegistration(reverseBefore, reverseAfter, rightPage, leftPage,
                out reverseScaleX, out reverseScaleY, out reverseDx, out reverseDy);
        }

        private static bool UniqueRegionMatching(IList<ReconcileItem> items,
            ReconcileRegion region, IList<PdfReviewWord> left,
            IList<PdfReviewWord> right, PdfReviewPage leftPage,
            PdfReviewPage rightPage, DiffContext context, out int[] rightByLeft)
        {
            rightByLeft = null;
            double scaleX, scaleY, reverseScaleX, reverseScaleY;
            if (!TrySymmetricPageScale(leftPage, rightPage, out scaleX, out scaleY,
                out reverseScaleX, out reverseScaleY))
                return false;

            ReconcileItem before, after;
            double dx, dy, reverseDx, reverseDy;
            bool hasRegistration = TryRegionRegistrations(items, region, leftPage,
                rightPage, context, out before, out after, out scaleX, out scaleY,
                out dx, out dy, out reverseScaleX, out reverseScaleY,
                out reverseDx, out reverseDy);
            if (context.WorkExhausted)
                return false;
            if (!hasRegistration)
                dx = dy = reverseDx = reverseDy = 0;

            int count = left.Count;
            var edges = new List<int>[count];
            for (int i = 0; i < count; i++)
            {
                if (!context.TryStep())
                    return false;
                edges[i] = new List<int>();
                for (int j = 0; j < count; j++)
                {
                    if (!context.TryStep())
                        return false;
                    if (WordKey(left[i]) == WordKey(right[j]) &&
                        ReconcileGeometryEdge(items, region, left[i], right[j],
                            leftPage, rightPage, hasRegistration, scaleX, scaleY,
                            dx, dy, reverseScaleX, reverseScaleY, reverseDx,
                            reverseDy, context))
                        edges[i].Add(j);
                    if (context.WorkExhausted)
                        return false;
                }
                if (edges[i].Count == 0)
                    return false;
            }

            int[] forward;
            if (!UniqueReconcileMatching(edges, count, context, out forward))
                return false;
            var reverseEdges = new List<int>[count];
            for (int i = 0; i < count; i++)
            {
                if (!context.TryStep())
                    return false;
                reverseEdges[i] = new List<int>();
            }
            for (int i = 0; i < count; i++)
            {
                if (!context.TryStep())
                    return false;
                foreach (int rightIndex in edges[i])
                {
                    if (!context.TryStep())
                        return false;
                    reverseEdges[rightIndex].Add(i);
                }
            }
            int[] reverse;
            if (!UniqueReconcileMatching(reverseEdges, count, context, out reverse))
                return false;
            for (int i = 0; i < forward.Length; i++)
            {
                if (!context.TryStep())
                    return false;
                int rightIndex = forward[i];
                if (rightIndex < 0 || rightIndex >= reverse.Length ||
                    reverse[rightIndex] != i)
                    return false;
            }
            if (!ReconcileAnchorPreserved(before, left, right, forward, context) ||
                !ReconcileAnchorPreserved(after, left, right, forward, context) ||
                !ReconcileMatchingCorroborated(items, region, left, right,
                    forward, leftPage, rightPage, hasRegistration, scaleX, scaleY,
                    dx, dy, reverseScaleX, reverseScaleY, reverseDx, reverseDy,
                    context))
                return false;
            rightByLeft = forward;
            return true;
        }

        private static bool ReconcileAnchorPreserved(ReconcileItem anchor,
            IList<PdfReviewWord> left, IList<PdfReviewWord> right,
            int[] rightByLeft, DiffContext context)
        {
            if (anchor == null || left == null || right == null || rightByLeft == null)
                return false;
            int leftIndex = -1;
            for (int i = 0; i < left.Count; i++)
            {
                if (!context.TryStep())
                    return false;
                if (ReferenceEquals(left[i], anchor.Left))
                {
                    leftIndex = i;
                    break;
                }
            }
            if (leftIndex < 0)
                return true;
            int rightIndex = rightByLeft[leftIndex];
            return rightIndex >= 0 && rightIndex < right.Count &&
                ReferenceEquals(right[rightIndex], anchor.Right);
        }

        private static bool ReconcileGeometryEdge(IList<ReconcileItem> items,
            ReconcileRegion region, PdfReviewWord left, PdfReviewWord right,
            PdfReviewPage leftPage, PdfReviewPage rightPage, bool hasRegistration,
            double scaleX, double scaleY, double dx, double dy,
            double reverseScaleX, double reverseScaleY, double reverseDx,
            double reverseDy, DiffContext context)
        {
            if (left == null || right == null || !ValidReviewBox(left.Box) ||
                !ValidReviewBox(right.Box))
                return false;
            if (hasRegistration && GeometryColocated(left.Box, right.Box,
                    scaleX, scaleY, dx, dy) &&
                GeometryColocated(right.Box, left.Box, reverseScaleX,
                    reverseScaleY, reverseDx, reverseDy))
                return true;
            if (FixedPageMarginEdge(left, right, leftPage, rightPage, scaleX,
                    scaleY, reverseScaleX, reverseScaleY, context))
                return true;
            if (context.WorkExhausted || !StructuredGeometryShape(left, right,
                    scaleX, scaleY, reverseScaleX, reverseScaleY))
                return false;
            int witnesses;
            if (!TryReconcileWitnessCount(items, region, left, right, leftPage,
                    rightPage, scaleX, scaleY, reverseScaleX, reverseScaleY,
                    context, out witnesses))
                return false;
            if (witnesses > 0)
                return true;
            return InternalBlockClusterCorroborates(items, region, left, right,
                hasRegistration, scaleX, scaleY, dx, dy, reverseScaleX,
                reverseScaleY, reverseDx, reverseDy, context);
        }

        private static bool ReconcileMatchingCorroborated(
            IList<ReconcileItem> items, ReconcileRegion region,
            IList<PdfReviewWord> left, IList<PdfReviewWord> right,
            int[] rightByLeft, PdfReviewPage leftPage, PdfReviewPage rightPage,
            bool hasRegistration, double scaleX, double scaleY, double dx,
            double dy, double reverseScaleX, double reverseScaleY,
            double reverseDx, double reverseDy, DiffContext context)
        {
            for (int i = 0; i < left.Count; i++)
            {
                if (!context.TryStep())
                    return false;
                PdfReviewWord candidateLeft = left[i];
                PdfReviewWord candidateRight = right[rightByLeft[i]];
                if (hasRegistration && GeometryColocated(candidateLeft.Box,
                        candidateRight.Box, scaleX, scaleY, dx, dy) &&
                    GeometryColocated(candidateRight.Box, candidateLeft.Box,
                        reverseScaleX, reverseScaleY, reverseDx, reverseDy))
                    continue;
                if (FixedPageMarginEdge(candidateLeft, candidateRight, leftPage,
                    rightPage, scaleX, scaleY, reverseScaleX, reverseScaleY,
                    context))
                    continue;
                if (context.WorkExhausted)
                    return false;
                int witnesses;
                if (!TryReconcileWitnessCount(items, region, candidateLeft,
                        candidateRight, leftPage, rightPage, scaleX, scaleY,
                        reverseScaleX, reverseScaleY, context, out witnesses))
                    return false;
                if (witnesses == 0)
                {
                    if (!InternalBlockClusterCorroborates(items, region,
                            candidateLeft, candidateRight, hasRegistration,
                            scaleX, scaleY, dx, dy, reverseScaleX, reverseScaleY,
                            reverseDx, reverseDy, context))
                        return false;
                    continue;
                }
                int translationCluster;
                if (witnesses == 1 &&
                    (!TryStructuredTranslationClusterSize(left, right,
                        rightByLeft, i, scaleX, scaleY, reverseScaleX,
                        reverseScaleY, context, out translationCluster) ||
                     translationCluster < 2))
                    return false;
            }
            return true;
        }

        /// <summary>
        /// В соседних строках одной формы локальный сдвиг может немного отличаться от
        /// регистрации по Exact-якорям. Такой край допустим только при готовой регистрации,
        /// рядом с ней и внутри не менее трёх взаимно-однозначных слов одного trusted
        /// block-pair. Это не расширяет правило для одиночного физического перемещения.
        /// </summary>
        private static bool InternalBlockClusterCorroborates(
            IList<ReconcileItem> items, ReconcileRegion region,
            PdfReviewWord left, PdfReviewWord right, bool hasRegistration,
            double scaleX, double scaleY, double dx, double dy,
            double reverseScaleX, double reverseScaleY, double reverseDx,
            double reverseDy, DiffContext context)
        {
            int clusterSize;
            return hasRegistration && TranslationNearRegistration(left, right,
                    scaleX, scaleY, dx, dy, reverseScaleX, reverseScaleY,
                    reverseDx, reverseDy) &&
                TryReconcileInternalBlockClusterSize(items, region, left, right,
                    scaleX, scaleY, reverseScaleX, reverseScaleY, context,
                    out clusterSize) &&
                clusterSize >= ReconcileMinInternalBlockCluster;
        }

        private static bool TranslationNearRegistration(PdfReviewWord left,
            PdfReviewWord right, double scaleX, double scaleY, double dx,
            double dy, double reverseScaleX, double reverseScaleY,
            double reverseDx, double reverseDy)
        {
            if (left == null || right == null || !ValidReviewBox(left.Box) ||
                !ValidReviewBox(right.Box))
                return false;
            double actualDx = CenterX(right.Box) - CenterX(left.Box) * scaleX;
            double actualDy = CenterY(right.Box) - CenterY(left.Box) * scaleY;
            double actualReverseDx = CenterX(left.Box) -
                CenterX(right.Box) * reverseScaleX;
            double actualReverseDy = CenterY(left.Box) -
                CenterY(right.Box) * reverseScaleY;
            double height = Math.Max(BoxHeight(left.Box), BoxHeight(right.Box));
            double tolerance = Math.Max(3.0,
                ReconcileInternalRegistrationHeightShare * height);
            return Math.Abs(actualDx - dx) <= tolerance &&
                Math.Abs(actualDy - dy) <= tolerance &&
                Math.Abs(actualReverseDx - reverseDx) <= tolerance &&
                Math.Abs(actualReverseDy - reverseDy) <= tolerance;
        }

        private static bool TryReconcileInternalBlockClusterSize(
            IList<ReconcileItem> items, ReconcileRegion region,
            PdfReviewWord left, PdfReviewWord right, double scaleX,
            double scaleY, double reverseScaleX, double reverseScaleY,
            DiffContext context, out int clusterSize)
        {
            clusterSize = 0;
            if (items == null || region == null || left == null || right == null ||
                region.Start < 0 || region.End < region.Start ||
                region.End >= items.Count || left.BlockId < 0 || right.BlockId < 0)
                return true;

            var leftCandidates = new List<PdfReviewWord>();
            var rightCandidates = new List<PdfReviewWord>();
            bool containsLeft = false, containsRight = false;
            for (int i = region.Start; i <= region.End; i++)
            {
                if (!context.TryStep())
                    return false;
                ReconcileItem item = items[i];
                if (item.Kind == PdfReviewDiffKind.Delete && item.Left != null &&
                    item.Left.PageIndex == region.LeftPage &&
                    item.Left.BlockId == left.BlockId)
                {
                    leftCandidates.Add(item.Left);
                    containsLeft |= ReferenceEquals(item.Left, left);
                }
                else if (item.Kind == PdfReviewDiffKind.Insert && item.Right != null &&
                    item.Right.PageIndex == region.RightPage &&
                    item.Right.BlockId == right.BlockId)
                {
                    rightCandidates.Add(item.Right);
                    containsRight |= ReferenceEquals(item.Right, right);
                }
            }
            if (!containsLeft || !containsRight ||
                leftCandidates.Count < ReconcileMinInternalBlockCluster ||
                rightCandidates.Count < ReconcileMinInternalBlockCluster)
                return true;

            int count = 0;
            bool candidatePairIsUnique = false;
            foreach (PdfReviewWord otherLeft in leftCandidates)
            {
                if (!context.TryStep())
                    return false;
                PdfReviewWord uniqueRight = null;
                int rightOptions = 0;
                foreach (PdfReviewWord otherRight in rightCandidates)
                {
                    if (!context.TryStep())
                        return false;
                    if (InternalBlockPairCompatible(left, right, otherLeft,
                            otherRight, scaleX, scaleY, reverseScaleX,
                            reverseScaleY))
                    {
                        uniqueRight = otherRight;
                        rightOptions++;
                    }
                }
                if (rightOptions != 1)
                    continue;

                PdfReviewWord uniqueLeft = null;
                int leftOptions = 0;
                foreach (PdfReviewWord reverseLeft in leftCandidates)
                {
                    if (!context.TryStep())
                        return false;
                    if (InternalBlockPairCompatible(left, right, reverseLeft,
                            uniqueRight, scaleX, scaleY, reverseScaleX,
                            reverseScaleY))
                    {
                        uniqueLeft = reverseLeft;
                        leftOptions++;
                    }
                }
                if (leftOptions != 1 || !ReferenceEquals(uniqueLeft, otherLeft))
                    continue;
                if (ReferenceEquals(otherLeft, left))
                {
                    if (!ReferenceEquals(uniqueRight, right))
                        return true;
                    candidatePairIsUnique = true;
                }
                else if (ReferenceEquals(uniqueRight, right))
                {
                    return true;
                }
                count++;
            }
            clusterSize = candidatePairIsUnique ? count : 0;
            return true;
        }

        private static bool InternalBlockPairCompatible(PdfReviewWord clusterLeft,
            PdfReviewWord clusterRight, PdfReviewWord left, PdfReviewWord right,
            double scaleX, double scaleY, double reverseScaleX,
            double reverseScaleY)
        {
            return clusterLeft != null && clusterRight != null && left != null &&
                right != null && left.BlockId == clusterLeft.BlockId &&
                right.BlockId == clusterRight.BlockId &&
                WordKey(left) == WordKey(right) &&
                StructuredGeometryShape(left, right, scaleX, scaleY,
                    reverseScaleX, reverseScaleY) &&
                TranslationsConsistent(clusterLeft, clusterRight, left, right,
                    scaleX, scaleY, reverseScaleX, reverseScaleY);
        }

        private static bool TryReconcileWitnessCount(IList<ReconcileItem> items,
            ReconcileRegion region, PdfReviewWord left, PdfReviewWord right,
            PdfReviewPage leftPage, PdfReviewPage rightPage, double scaleX,
            double scaleY, double reverseScaleX, double reverseScaleY,
            DiffContext context, out int count)
        {
            count = 0;
            if (left.BlockId < 0 || right.BlockId < 0)
                return true;
            int first = Math.Max(0, region.Start - ReconcileAnchorNeighborhood);
            int last = Math.Min(items.Count - 1,
                region.End + ReconcileAnchorNeighborhood);
            for (int i = first; i <= last; i++)
            {
                if (!context.TryStep())
                    return false;
                if (i >= region.Start && i <= region.End)
                    continue;
                ReconcileItem witness = items[i];
                if (!witness.IsExact || witness.Left.PageIndex != region.LeftPage ||
                    witness.Right.PageIndex != region.RightPage ||
                    witness.Left.BlockId < 0 || witness.Right.BlockId < 0 ||
                    !StableReconcileAnchor(items, i, context) ||
                    !ReconcilePathClear(items, region, i, context) ||
                    !BlocksCanCoexist(left, right, witness.Left, witness.Right) ||
                    !WitnessNearby(left, right, witness.Left, witness.Right,
                        leftPage, rightPage) ||
                    !StructuredGeometryShape(witness.Left, witness.Right, scaleX,
                        scaleY, reverseScaleX, reverseScaleY) ||
                    !TranslationsConsistent(left, right, witness.Left,
                        witness.Right, scaleX, scaleY, reverseScaleX,
                        reverseScaleY))
                    continue;
                count++;
            }
            return !context.WorkExhausted;
        }

        private static bool ReconcilePathClear(IList<ReconcileItem> items,
            ReconcileRegion region, int witnessIndex, DiffContext context)
        {
            int first = witnessIndex < region.Start ? witnessIndex + 1 : region.End + 1;
            int last = witnessIndex < region.Start ? region.Start - 1 : witnessIndex - 1;
            for (int i = first; i <= last; i++)
            {
                if (!context.TryStep())
                    return false;
                ReconcileItem item = items[i];
                if (item.Barrier != null || (item.IsExact &&
                    (item.Left.PageIndex != region.LeftPage ||
                     item.Right.PageIndex != region.RightPage)))
                    return false;
            }
            return true;
        }

        private static bool TryStructuredTranslationClusterSize(
            IList<PdfReviewWord> left, IList<PdfReviewWord> right,
            int[] rightByLeft, int candidateIndex, double scaleX,
            double scaleY, double reverseScaleX, double reverseScaleY,
            DiffContext context, out int clusterSize)
        {
            clusterSize = 0;
            if (left == null || right == null || rightByLeft == null ||
                rightByLeft.Length != left.Count || candidateIndex < 0 ||
                candidateIndex >= left.Count)
                return true;
            int candidateRightIndex = rightByLeft[candidateIndex];
            if (candidateRightIndex < 0 || candidateRightIndex >= right.Count)
                return true;
            PdfReviewWord candidateLeft = left[candidateIndex];
            PdfReviewWord candidateRight = right[candidateRightIndex];
            int count = 0;
            for (int i = 0; i < left.Count; i++)
            {
                if (!context.TryStep())
                    return false;
                int rightIndex = rightByLeft[i];
                if (rightIndex < 0 || rightIndex >= right.Count)
                    return true;
                PdfReviewWord otherLeft = left[i];
                PdfReviewWord otherRight = right[rightIndex];
                if (otherLeft.BlockId >= 0 && otherRight.BlockId >= 0 &&
                    BlocksCanCoexist(candidateLeft, candidateRight, otherLeft,
                        otherRight) &&
                    StructuredGeometryShape(otherLeft, otherRight, scaleX,
                        scaleY, reverseScaleX, reverseScaleY) &&
                    TranslationsConsistent(candidateLeft, candidateRight,
                        otherLeft, otherRight, scaleX, scaleY, reverseScaleX,
                        reverseScaleY))
                    count++;
            }
            clusterSize = count;
            return true;
        }

        private static bool StructuredGeometryShape(PdfReviewWord left,
            PdfReviewWord right, double scaleX, double scaleY,
            double reverseScaleX, double reverseScaleY)
        {
            if (left == null || right == null || left.BlockId < 0 ||
                right.BlockId < 0 || !ValidReviewBox(left.Box) ||
                !ValidReviewBox(right.Box))
                return false;
            double dx = CenterX(right.Box) - CenterX(left.Box) * scaleX;
            double dy = CenterY(right.Box) - CenterY(left.Box) * scaleY;
            double reverseDx = CenterX(left.Box) -
                CenterX(right.Box) * reverseScaleX;
            double reverseDy = CenterY(left.Box) -
                CenterY(right.Box) * reverseScaleY;
            return GeometryColocated(left.Box, right.Box, scaleX, scaleY, dx, dy) &&
                GeometryColocated(right.Box, left.Box, reverseScaleX,
                    reverseScaleY, reverseDx, reverseDy);
        }

        private static bool TranslationsConsistent(PdfReviewWord left,
            PdfReviewWord right, PdfReviewWord witnessLeft,
            PdfReviewWord witnessRight, double scaleX, double scaleY,
            double reverseScaleX, double reverseScaleY)
        {
            double dx = CenterX(right.Box) - CenterX(left.Box) * scaleX;
            double dy = CenterY(right.Box) - CenterY(left.Box) * scaleY;
            double witnessDx = CenterX(witnessRight.Box) -
                CenterX(witnessLeft.Box) * scaleX;
            double witnessDy = CenterY(witnessRight.Box) -
                CenterY(witnessLeft.Box) * scaleY;
            double reverseDx = CenterX(left.Box) -
                CenterX(right.Box) * reverseScaleX;
            double reverseDy = CenterY(left.Box) -
                CenterY(right.Box) * reverseScaleY;
            double witnessReverseDx = CenterX(witnessLeft.Box) -
                CenterX(witnessRight.Box) * reverseScaleX;
            double witnessReverseDy = CenterY(witnessLeft.Box) -
                CenterY(witnessRight.Box) * reverseScaleY;
            double height = Math.Max(Math.Max(BoxHeight(left.Box),
                BoxHeight(right.Box)), Math.Max(BoxHeight(witnessLeft.Box),
                BoxHeight(witnessRight.Box)));
            double tolerance = Math.Max(2.5, 0.35 * height);
            return Math.Abs(dx - witnessDx) <= tolerance &&
                Math.Abs(dy - witnessDy) <= tolerance &&
                Math.Abs(reverseDx - witnessReverseDx) <= tolerance &&
                Math.Abs(reverseDy - witnessReverseDy) <= tolerance;
        }

        private static bool WitnessNearby(PdfReviewWord left,
            PdfReviewWord right, PdfReviewWord witnessLeft,
            PdfReviewWord witnessRight, PdfReviewPage leftPage,
            PdfReviewPage rightPage)
        {
            double leftWidth = PageViewWidth(leftPage);
            double leftHeight = PageViewHeight(leftPage);
            double rightWidth = PageViewWidth(rightPage);
            double rightHeight = PageViewHeight(rightPage);
            return Math.Abs(CenterX(left.Box) - CenterX(witnessLeft.Box)) <=
                    ReconcileMaxWitnessAxisShare * leftWidth &&
                Math.Abs(CenterY(left.Box) - CenterY(witnessLeft.Box)) <=
                    ReconcileMaxWitnessAxisShare * leftHeight &&
                Math.Abs(CenterX(right.Box) - CenterX(witnessRight.Box)) <=
                    ReconcileMaxWitnessAxisShare * rightWidth &&
                Math.Abs(CenterY(right.Box) - CenterY(witnessRight.Box)) <=
                    ReconcileMaxWitnessAxisShare * rightHeight;
        }

        private static bool BlocksCanCoexist(PdfReviewWord left,
            PdfReviewWord right, PdfReviewWord otherLeft,
            PdfReviewWord otherRight)
        {
            return left != null && right != null && otherLeft != null &&
                otherRight != null && left.BlockId >= 0 && right.BlockId >= 0 &&
                otherLeft.BlockId >= 0 && otherRight.BlockId >= 0 &&
                (left.BlockId != otherLeft.BlockId ||
                 right.BlockId == otherRight.BlockId) &&
                (right.BlockId != otherRight.BlockId ||
                 left.BlockId == otherLeft.BlockId);
        }

        private static bool FixedPageMarginEdge(PdfReviewWord left,
            PdfReviewWord right, PdfReviewPage leftPage, PdfReviewPage rightPage,
            double scaleX, double scaleY, double reverseScaleX,
            double reverseScaleY, DiffContext context)
        {
            int leftKeyCount, rightKeyCount;
            if (left == null || right == null || left.BlockId >= 0 ||
                right.BlockId >= 0 || !OuterPageMargin(left, leftPage) ||
                !OuterPageMargin(right, rightPage) ||
                !TryPageMarginKeyCount(leftPage, WordKey(left), context,
                    out leftKeyCount) || leftKeyCount != 1 ||
                !TryPageMarginKeyCount(rightPage, WordKey(right), context,
                    out rightKeyCount) || rightKeyCount != 1)
                return false;
            return !context.WorkExhausted &&
                GeometryColocated(left.Box, right.Box, scaleX, scaleY, 0, 0) &&
                GeometryColocated(right.Box, left.Box, reverseScaleX,
                    reverseScaleY, 0, 0);
        }

        private static bool OuterPageMargin(PdfReviewWord word, PdfReviewPage page)
        {
            double height = PageViewHeight(page);
            if (word == null || !ValidReviewBox(word.Box) || !FinitePositive(height))
                return false;
            double center = CenterY(word.Box);
            return center <= ReconcilePageMarginShare * height ||
                center >= (1.0 - ReconcilePageMarginShare) * height;
        }

        private static bool TryPageMarginKeyCount(PdfReviewPage page, string key,
            DiffContext context, out int count)
        {
            count = 0;
            if (page == null || page.Words == null)
                return true;
            foreach (PdfReviewWord word in page.Words)
            {
                if (!context.TryStep())
                {
                    count = 0;
                    return false;
                }
                if (word != null && WordKey(word) == key &&
                    OuterPageMargin(word, page))
                    count++;
            }
            return true;
        }

        private static bool TrySymmetricPageScale(PdfReviewPage leftPage,
            PdfReviewPage rightPage, out double scaleX, out double scaleY,
            out double reverseScaleX, out double reverseScaleY)
        {
            scaleX = scaleY = reverseScaleX = reverseScaleY = 1;
            double leftWidth = PageViewWidth(leftPage);
            double leftHeight = PageViewHeight(leftPage);
            double rightWidth = PageViewWidth(rightPage);
            double rightHeight = PageViewHeight(rightPage);
            if (!FinitePositive(leftWidth) || !FinitePositive(leftHeight) ||
                !FinitePositive(rightWidth) || !FinitePositive(rightHeight))
                return false;
            scaleX = rightWidth / leftWidth;
            scaleY = rightHeight / leftHeight;
            reverseScaleX = leftWidth / rightWidth;
            reverseScaleY = leftHeight / rightHeight;
            return scaleX >= ReconcileMinPageScale &&
                scaleX <= ReconcileMaxPageScale &&
                scaleY >= ReconcileMinPageScale &&
                scaleY <= ReconcileMaxPageScale &&
                reverseScaleX >= ReconcileMinPageScale &&
                reverseScaleX <= ReconcileMaxPageScale &&
                reverseScaleY >= ReconcileMinPageScale &&
                reverseScaleY <= ReconcileMaxPageScale;
        }

        private static bool UniqueReconcileMatching(IList<List<int>> edges,
            int rightCount, DiffContext context, out int[] rightByLeft)
        {
            rightByLeft = null;
            int count = edges == null ? 0 : edges.Count;
            if (count == 0 || rightCount != count)
                return false;
            var order = new int[count];
            for (int i = 0; i < count; i++)
            {
                if (!context.TryStep())
                    return false;
                if (edges[i] == null || edges[i].Count == 0)
                    return false;
                order[i] = i;
            }
            if (!context.TryReserveSort(order.Length))
                return false;
            Array.Sort(order, delegate(int a, int b)
            {
                int comparison = edges[a].Count.CompareTo(edges[b].Count);
                return comparison != 0 ? comparison : a.CompareTo(b);
            });
            var used = new bool[rightCount];
            var current = new int[count];
            for (int i = 0; i < count; i++)
            {
                if (!context.TryStep())
                    return false;
                current[i] = -1;
            }
            int solutions = 0;
            int[] unique = null;
            CountGeometryMatchings(0, order, edges, used, current, context,
                ref solutions, ref unique);
            if (context.WorkExhausted || solutions != 1)
                return false;
            rightByLeft = unique;
            return true;
        }

        private static bool ConsistentRegionBlocks(IList<ReconcileItem> items,
            ReconcileRegion region, IList<PdfReviewWord> left,
            IList<PdfReviewWord> right, int[] rightByLeft,
            PdfReviewPage leftPage, PdfReviewPage rightPage, DiffContext context)
        {
            if (leftPage == null || rightPage == null)
                return false;
            ReconcileItem before, after;
            if (!TryRegionAnchors(items, region, context, out before, out after))
                return false;
            return ConsistentBlockCorrespondence(before, after, left, right,
                rightByLeft, context);
        }

        private static bool TryRebuildReconciledOperations(
            IList<PdfReviewWordOp> operations, IList<ReconcileItem> items,
            IDictionary<PdfReviewWord, PdfReviewWord> rightByLeft,
            IDictionary<PdfReviewWord, PdfReviewWord> leftByRight,
            ISet<PdfReviewWord> releasedLeft,
            ISet<PdfReviewWord> releasedRight,
            DiffContext context, out List<PdfReviewWordOp> rebuilt)
        {
            rebuilt = null;
            if (operations == null || items == null || rightByLeft == null ||
                leftByRight == null || releasedLeft == null ||
                releasedRight == null || rightByLeft.Count == 0 ||
                rightByLeft.Count != leftByRight.Count)
                return false;

            var leftItem = new Dictionary<PdfReviewWord, int>();
            var rightItem = new Dictionary<PdfReviewWord, int>();
            int foundReleasedLeft = 0, foundReleasedRight = 0;
            for (int i = 0; i < items.Count; i++)
            {
                if (!context.TryStep())
                    return false;
                ReconcileItem item = items[i];
                if (item == null)
                    return false;
                if (item.Left != null)
                    leftItem[item.Left] = i;
                if (item.Right != null)
                    rightItem[item.Right] = i;
                bool isReleasedLeft = item.Left != null &&
                    releasedLeft.Contains(item.Left);
                bool isReleasedRight = item.Right != null &&
                    releasedRight.Contains(item.Right);
                if (isReleasedLeft)
                    foundReleasedLeft++;
                if (isReleasedRight)
                    foundReleasedRight++;
                if (!item.IsExact)
                {
                    if (isReleasedLeft || isReleasedRight)
                        return false;
                    continue;
                }
                bool leftMapped = rightByLeft.ContainsKey(item.Left);
                bool rightMapped = leftByRight.ContainsKey(item.Right);
                if (leftMapped == rightMapped)
                {
                    if (isReleasedLeft || isReleasedRight)
                        return false;
                }
                else if (leftMapped)
                {
                    if (!isReleasedRight || isReleasedLeft)
                        return false;
                }
                else if (!isReleasedLeft || isReleasedRight)
                {
                    return false;
                }
            }
            if (foundReleasedLeft != releasedLeft.Count ||
                foundReleasedRight != releasedRight.Count)
                return false;
            var ranges = new List<ReconcileRegion>();
            foreach (KeyValuePair<PdfReviewWord, PdfReviewWord> pair in rightByLeft)
            {
                if (!context.TryStep())
                    return false;
                int leftIndex, rightIndex;
                PdfReviewWord reverse;
                ReconcileItem leftSource;
                ReconcileItem rightSource;
                if (!leftItem.TryGetValue(pair.Key, out leftIndex) ||
                    !rightItem.TryGetValue(pair.Value, out rightIndex) ||
                    !leftByRight.TryGetValue(pair.Value, out reverse) ||
                    !ReferenceEquals(reverse, pair.Key) ||
                    WordKey(pair.Key) != WordKey(pair.Value) ||
                    (leftSource = items[leftIndex]).Left == null ||
                    (rightSource = items[rightIndex]).Right == null ||
                    (leftSource.Kind != PdfReviewDiffKind.Delete &&
                     !leftSource.IsExact) ||
                    (rightSource.Kind != PdfReviewDiffKind.Insert &&
                     !rightSource.IsExact))
                    return false;
                ranges.Add(new ReconcileRegion
                {
                    Start = Math.Min(leftIndex, rightIndex),
                    End = Math.Max(leftIndex, rightIndex)
                });
            }
            if (ranges.Count == 0 || !context.TryReserveSort(ranges.Count))
                return false;
            ranges.Sort(delegate(ReconcileRegion a, ReconcileRegion b)
            {
                int comparison = a.Start.CompareTo(b.Start);
                return comparison != 0 ? comparison : a.End.CompareTo(b.End);
            });
            var merged = new List<ReconcileRegion>();
            foreach (ReconcileRegion range in ranges)
            {
                if (!context.TryStep())
                    return false;
                ReconcileRegion last = merged.Count == 0 ? null : merged[merged.Count - 1];
                if (last != null && range.Start <= last.End)
                    last.End = Math.Max(last.End, range.End);
                else
                    merged.Add(new ReconcileRegion { Start = range.Start, End = range.End });
            }

            var output = new List<PdfReviewWordOp>();
            int rangeIndex = 0;
            for (int i = 0; i < items.Count;)
            {
                if (!context.TryStep())
                    return false;
                ReconcileRegion range = rangeIndex < merged.Count
                    ? merged[rangeIndex] : null;
                if (range != null && i == range.Start)
                {
                    if (range.End - range.Start > ReconcileMaxItemSpan ||
                        !TryRebuildReconcileSegment(items, range.Start, range.End,
                            rightByLeft, leftByRight, releasedLeft, releasedRight,
                            output, context))
                        return false;
                    i = range.End + 1;
                    rangeIndex++;
                }
                else
                {
                    if (!TryAppendReconcileItem(output, items[i], context))
                        return false;
                    i++;
                }
            }
            if (rangeIndex != merged.Count)
                return false;
            if (output.Count == 0 || context.WorkExhausted)
                return false;
            rebuilt = output;
            return true;
        }

        private static bool TryRebuildReconcileSegment(IList<ReconcileItem> items,
            int start, int end,
            IDictionary<PdfReviewWord, PdfReviewWord> rightByLeft,
            IDictionary<PdfReviewWord, PdfReviewWord> leftByRight,
            ISet<PdfReviewWord> releasedLeft,
            ISet<PdfReviewWord> releasedRight,
            List<PdfReviewWordOp> rebuilt, DiffContext context)
        {
            var indexByLeft = new Dictionary<PdfReviewWord, int>();
            var indexByRight = new Dictionary<PdfReviewWord, int>();
            var leftOrder = new Dictionary<PdfReviewWord, int>();
            var rightOrder = new Dictionary<PdfReviewWord, int>();
            int nextLeft = 0, nextRight = 0;
            for (int i = start; i <= end; i++)
            {
                if (!context.TryStep())
                    return false;
                ReconcileItem item = items[i];
                if (item == null || item.Barrier != null)
                    return false;
                if (item.Left != null)
                {
                    indexByLeft[item.Left] = i;
                    leftOrder[item.Left] = nextLeft++;
                }
                if (item.Right != null)
                {
                    indexByRight[item.Right] = i;
                    rightOrder[item.Right] = nextRight++;
                }
            }

            var nodes = new List<ReconcileNode>();
            var nodeByLeft = new Dictionary<PdfReviewWord, int>();
            var nodeByRight = new Dictionary<PdfReviewWord, int>();
            foreach (KeyValuePair<PdfReviewWord, PdfReviewWord> pair in rightByLeft)
            {
                if (!context.TryStep())
                    return false;
                int leftIndex, rightIndex;
                PdfReviewWord reverse;
                if (!indexByLeft.TryGetValue(pair.Key, out leftIndex) ||
                    !indexByRight.TryGetValue(pair.Value, out rightIndex))
                    continue;
                ReconcileItem leftSource = items[leftIndex];
                ReconcileItem rightSource = items[rightIndex];
                if ((leftSource.Kind != PdfReviewDiffKind.Delete &&
                     !leftSource.IsExact) ||
                    (rightSource.Kind != PdfReviewDiffKind.Insert &&
                     !rightSource.IsExact) ||
                    !leftByRight.TryGetValue(pair.Value, out reverse) ||
                    !ReferenceEquals(reverse, pair.Key))
                    return false;
                int nodeIndex = nodes.Count;
                nodes.Add(new ReconcileNode
                {
                    Left = pair.Key,
                    Right = pair.Value,
                    MatchKind = PdfReviewMatchKind.ReconciledOrder,
                    LeftOrder = leftOrder[pair.Key],
                    RightOrder = rightOrder[pair.Value],
                    OriginalOrder = Math.Min(leftIndex, rightIndex)
                });
                nodeByLeft.Add(pair.Key, nodeIndex);
                nodeByRight.Add(pair.Value, nodeIndex);
            }
            for (int i = start; i <= end; i++)
            {
                if (!context.TryStep())
                    return false;
                ReconcileItem item = items[i];
                if (item.Kind == PdfReviewDiffKind.Delete && item.Left != null)
                {
                    if (rightByLeft.ContainsKey(item.Left))
                    {
                        if (!nodeByLeft.ContainsKey(item.Left))
                            return false;
                        continue;
                    }
                    int nodeIndex = nodes.Count;
                    nodes.Add(new ReconcileNode
                    {
                        Left = item.Left,
                        MatchKind = PdfReviewMatchKind.None,
                        LeftOrder = leftOrder[item.Left],
                        OriginalOrder = i
                    });
                    nodeByLeft.Add(item.Left, nodeIndex);
                }
                else if (item.Kind == PdfReviewDiffKind.Insert && item.Right != null)
                {
                    if (leftByRight.ContainsKey(item.Right))
                    {
                        if (!nodeByRight.ContainsKey(item.Right))
                            return false;
                        continue;
                    }
                    int nodeIndex = nodes.Count;
                    nodes.Add(new ReconcileNode
                    {
                        Right = item.Right,
                        MatchKind = PdfReviewMatchKind.None,
                        RightOrder = rightOrder[item.Right],
                        OriginalOrder = i
                    });
                    nodeByRight.Add(item.Right, nodeIndex);
                }
                else if (item.IsExact)
                {
                    bool leftMapped = rightByLeft.ContainsKey(item.Left);
                    bool rightMapped = leftByRight.ContainsKey(item.Right);
                    bool releaseLeft = releasedLeft.Contains(item.Left);
                    bool releaseRight = releasedRight.Contains(item.Right);
                    if (leftMapped && rightMapped)
                    {
                        if (releaseLeft || releaseRight ||
                            !nodeByLeft.ContainsKey(item.Left) ||
                            !nodeByRight.ContainsKey(item.Right))
                            return false;
                        continue;
                    }
                    if (rightMapped)
                    {
                        if (leftMapped || !releaseLeft || releaseRight ||
                            nodeByLeft.ContainsKey(item.Left) ||
                            !nodeByRight.ContainsKey(item.Right))
                            return false;
                        int nodeIndex = nodes.Count;
                        nodes.Add(new ReconcileNode
                        {
                            Left = item.Left,
                            MatchKind = PdfReviewMatchKind.None,
                            LeftOrder = leftOrder[item.Left],
                            OriginalOrder = i
                        });
                        nodeByLeft.Add(item.Left, nodeIndex);
                        continue;
                    }
                    if (leftMapped)
                    {
                        if (!releaseRight || releaseLeft ||
                            !nodeByLeft.ContainsKey(item.Left) ||
                            nodeByRight.ContainsKey(item.Right))
                            return false;
                        int nodeIndex = nodes.Count;
                        nodes.Add(new ReconcileNode
                        {
                            Right = item.Right,
                            MatchKind = PdfReviewMatchKind.None,
                            RightOrder = rightOrder[item.Right],
                            OriginalOrder = i
                        });
                        nodeByRight.Add(item.Right, nodeIndex);
                        continue;
                    }
                    if (releaseLeft || releaseRight)
                        return false;
                    int exactNodeIndex = nodes.Count;
                    nodes.Add(new ReconcileNode
                    {
                        Left = item.Left,
                        Right = item.Right,
                        MatchKind = PdfReviewMatchKind.Exact,
                        LeftOrder = leftOrder[item.Left],
                        RightOrder = rightOrder[item.Right],
                        OriginalOrder = i
                    });
                    nodeByLeft.Add(item.Left, exactNodeIndex);
                    nodeByRight.Add(item.Right, exactNodeIndex);
                }
                else
                {
                    return false;
                }
            }
            if (nodes.Count == 0)
                return false;

            if (!AddReconcileOrderEdges(nodes, true, context) ||
                !AddReconcileOrderEdges(nodes, false, context))
                return false;
            int[] componentByNode;
            List<List<int>> components;
            if (!ReconcileStrongComponents(nodes, context, out componentByNode,
                    out components) || components.Count == 0)
                return false;

            var componentEdges = new HashSet<int>[components.Count];
            var indegree = new int[components.Count];
            var componentOrder = new int[components.Count];
            for (int i = 0; i < components.Count; i++)
            {
                if (!context.TryStep())
                    return false;
                componentEdges[i] = new HashSet<int>();
                componentOrder[i] = int.MaxValue;
                foreach (int nodeIndex in components[i])
                {
                    if (!context.TryStep() || nodeIndex < 0 ||
                        nodeIndex >= nodes.Count)
                        return false;
                    componentOrder[i] = Math.Min(componentOrder[i],
                        nodes[nodeIndex].OriginalOrder);
                }
            }
            for (int i = 0; i < nodes.Count; i++)
            {
                if (!context.TryStep())
                    return false;
                int from = componentByNode[i];
                if (from < 0 || from >= components.Count)
                    return false;
                foreach (int edge in nodes[i].Edges)
                {
                    if (!context.TryStep() || edge < 0 || edge >= nodes.Count)
                        return false;
                    int to = componentByNode[edge];
                    if (to < 0 || to >= components.Count)
                        return false;
                    if (from != to && componentEdges[from].Add(to))
                        indegree[to]++;
                }
            }

            if (!context.TryReserveSort(components.Count) ||
                !context.TryReserveSort(components.Count))
                return false;
            var ready = new SortedSet<ReconcileTopoEntry>();
            for (int i = 0; i < components.Count; i++)
            {
                if (!context.TryStep())
                    return false;
                if (indegree[i] == 0)
                    ready.Add(new ReconcileTopoEntry
                    {
                        Order = componentOrder[i],
                        Component = i
                    });
            }
            int emitted = 0;
            while (ready.Count > 0)
            {
                if (!context.TryStep())
                    return false;
                ReconcileTopoEntry entry = ready.Min;
                ready.Remove(entry);
                if (!AppendReconcileComponent(rebuilt, nodes,
                    components[entry.Component], context))
                    return false;
                emitted++;
                foreach (int next in componentEdges[entry.Component])
                {
                    if (!context.TryStep())
                        return false;
                    indegree[next]--;
                    if (indegree[next] == 0)
                        ready.Add(new ReconcileTopoEntry
                        {
                            Order = componentOrder[next],
                            Component = next
                        });
                }
            }
            return !context.WorkExhausted && emitted == components.Count;
        }

        private static bool AddReconcileOrderEdges(IList<ReconcileNode> nodes,
            bool leftSide, DiffContext context)
        {
            var order = new List<int>();
            for (int i = 0; i < nodes.Count; i++)
            {
                if (!context.TryStep())
                    return false;
                if ((leftSide ? nodes[i].LeftOrder : nodes[i].RightOrder) >= 0)
                    order.Add(i);
            }
            if (!context.TryReserveSort(order.Count))
                return false;
            order.Sort(delegate(int a, int b)
            {
                int first = leftSide ? nodes[a].LeftOrder : nodes[a].RightOrder;
                int second = leftSide ? nodes[b].LeftOrder : nodes[b].RightOrder;
                int comparison = first.CompareTo(second);
                return comparison != 0 ? comparison : a.CompareTo(b);
            });
            for (int i = 1; i < order.Count; i++)
            {
                if (!context.TryStep())
                    return false;
                List<int> edges = nodes[order[i - 1]].Edges;
                if (!context.TryReserve(edges.Count))
                    return false;
                if (!edges.Contains(order[i]))
                    edges.Add(order[i]);
            }
            return !context.WorkExhausted;
        }

        private static bool ReconcileStrongComponents(IList<ReconcileNode> nodes,
            DiffContext context, out int[] componentByNode,
            out List<List<int>> components)
        {
            componentByNode = null;
            components = null;
            if (nodes == null)
                return false;
            int count = nodes.Count;
            var index = new int[count];
            var low = new int[count];
            var onStack = new bool[count];
            for (int i = 0; i < count; i++)
            {
                if (!context.TryStep())
                    return false;
                index[i] = -1;
            }
            var stack = new Stack<int>();
            int nextIndex = 0;
            var found = new List<List<int>>();
            for (int i = 0; i < count; i++)
            {
                if (!context.TryStep())
                    return false;
                if (index[i] < 0 && !VisitReconcileNode(i, nodes, index, low,
                        onStack, stack, ref nextIndex, found, context))
                    return false;
            }
            var byNode = new int[count];
            for (int component = 0; component < found.Count; component++)
            {
                if (!context.TryStep())
                    return false;
                foreach (int node in found[component])
                {
                    if (!context.TryStep() || node < 0 || node >= count)
                        return false;
                    byNode[node] = component;
                }
            }
            if (context.WorkExhausted)
                return false;
            componentByNode = byNode;
            components = found;
            return true;
        }

        private static bool VisitReconcileNode(int node,
            IList<ReconcileNode> nodes, int[] index, int[] low, bool[] onStack,
            Stack<int> stack, ref int nextIndex, List<List<int>> components,
            DiffContext context)
        {
            if (!context.TryStep() || node < 0 || node >= nodes.Count)
                return false;
            index[node] = low[node] = nextIndex++;
            stack.Push(node);
            onStack[node] = true;
            foreach (int edge in nodes[node].Edges)
            {
                if (!context.TryStep() || edge < 0 || edge >= nodes.Count)
                    return false;
                if (index[edge] < 0)
                {
                    if (!VisitReconcileNode(edge, nodes, index, low, onStack,
                            stack, ref nextIndex, components, context))
                        return false;
                    low[node] = Math.Min(low[node], low[edge]);
                }
                else if (onStack[edge])
                {
                    low[node] = Math.Min(low[node], index[edge]);
                }
            }
            if (low[node] != index[node])
                return true;
            var component = new List<int>();
            while (stack.Count > 0)
            {
                if (!context.TryStep())
                    return false;
                int current = stack.Pop();
                onStack[current] = false;
                component.Add(current);
                if (current == node)
                    break;
            }
            components.Add(component);
            return !context.WorkExhausted;
        }

        private static bool AppendReconcileComponent(List<PdfReviewWordOp> operations,
            IList<ReconcileNode> nodes, IList<int> component, DiffContext context)
        {
            if (operations == null || nodes == null || component == null ||
                component.Count == 0)
                return false;
            bool reconciled = false;
            bool exact = false;
            int leftPage = -1;
            int rightPage = -1;
            foreach (int nodeIndex in component)
            {
                if (!context.TryStep() || nodeIndex < 0 || nodeIndex >= nodes.Count)
                    return false;
                ReconcileNode node = nodes[nodeIndex];
                if (node == null || (node.MatchKind != PdfReviewMatchKind.None &&
                    node.MatchKind != PdfReviewMatchKind.Exact &&
                    node.MatchKind != PdfReviewMatchKind.ReconciledOrder))
                    return false;
                reconciled |= node.MatchKind == PdfReviewMatchKind.ReconciledOrder;
                exact |= node.MatchKind == PdfReviewMatchKind.Exact;
                if (node.MatchKind != PdfReviewMatchKind.ReconciledOrder)
                    continue;
                if (node.Left == null || node.Right == null)
                    return false;
                if (leftPage < 0)
                {
                    leftPage = node.Left.PageIndex;
                    rightPage = node.Right.PageIndex;
                }
                else if (leftPage != node.Left.PageIndex ||
                    rightPage != node.Right.PageIndex)
                {
                    return false;
                }
            }

            if (!reconciled && component.Count == 1)
            {
                if (!context.TryStep())
                    return false;
                ReconcileNode node = nodes[component[0]];
                if (node.Left != null && node.Right != null &&
                    node.MatchKind == PdfReviewMatchKind.Exact)
                {
                    if (!context.TryReserve(3))
                        return false;
                    AppendEqual(operations, node.Left, node.Right,
                        PdfReviewMatchKind.Exact);
                }
                else if (node.Left != null && node.Right == null)
                {
                    if (!context.TryReserve(1))
                        return false;
                    Append(operations, PdfReviewDiffKind.Delete, node.Left);
                }
                else if (node.Left == null && node.Right != null)
                {
                    if (!context.TryReserve(1))
                        return false;
                    Append(operations, PdfReviewDiffKind.Insert, node.Right);
                }
                else
                {
                    return false;
                }
                return !context.WorkExhausted;
            }
            if (!reconciled)
                return false;
            foreach (int nodeIndex in component)
            {
                if (!context.TryStep())
                    return false;
                ReconcileNode node = nodes[nodeIndex];
                if (node.Left == null || node.Right == null ||
                    (node.MatchKind != PdfReviewMatchKind.Exact &&
                     node.MatchKind != PdfReviewMatchKind.ReconciledOrder) ||
                    node.Left.PageIndex != leftPage ||
                    node.Right.PageIndex != rightPage)
                    return false;
            }

            if (!context.TryReserve(2L * component.Count) ||
                !context.TryReserveSort(component.Count) ||
                !context.TryReserveSort(component.Count))
                return false;
            var leftOrder = new List<int>(component);
            var rightOrder = new List<int>(component);
            leftOrder.Sort(delegate(int a, int b)
            {
                int comparison = nodes[a].LeftOrder.CompareTo(nodes[b].LeftOrder);
                return comparison != 0 ? comparison : a.CompareTo(b);
            });
            rightOrder.Sort(delegate(int a, int b)
            {
                int comparison = nodes[a].RightOrder.CompareTo(nodes[b].RightOrder);
                return comparison != 0 ? comparison : a.CompareTo(b);
            });
            var equal = new PdfReviewWordOp
            {
                Kind = PdfReviewDiffKind.Equal,
                MatchKind = exact ? PdfReviewMatchKind.MixedOrder :
                    PdfReviewMatchKind.ReconciledOrder
            };
            foreach (int nodeIndex in leftOrder)
            {
                if (!context.TryStep())
                    return false;
                equal.LeftWords.Add(nodes[nodeIndex].Left);
            }
            foreach (int nodeIndex in rightOrder)
            {
                if (!context.TryStep())
                    return false;
                equal.RightWords.Add(nodes[nodeIndex].Right);
            }
            foreach (int nodeIndex in leftOrder)
            {
                if (!context.TryStep())
                    return false;
                equal.Matches.Add(new PdfReviewWordMatch
                {
                    Left = nodes[nodeIndex].Left,
                    Right = nodes[nodeIndex].Right,
                    Kind = nodes[nodeIndex].MatchKind
                });
            }
            if (!TryReserveOperationCopy(equal, context))
                return false;
            AppendOperation(operations, equal);
            return !context.WorkExhausted;
        }

        private static bool StableReconcileAnchor(IList<ReconcileItem> items, int index,
            DiffContext context)
        {
            if (items == null || index < 0 || index >= items.Count)
                return false;
            ReconcileItem anchor = items[index];
            if (!anchor.IsExact || WordKey(anchor.Left).Length == 0 ||
                WordKey(anchor.Left) != WordKey(anchor.Right) ||
                !ValidReviewBox(anchor.Left.Box) || !ValidReviewBox(anchor.Right.Box))
                return false;
            string key = WordKey(anchor.Left);
            int first = Math.Max(0, index - ReconcileAnchorNeighborhood);
            int last = Math.Min(items.Count - 1, index + ReconcileAnchorNeighborhood);
            int leftCount = 0, rightCount = 0;
            for (int i = first; i <= last; i++)
            {
                if (!context.TryStep())
                    return false;
                if (items[i].Left != null && WordKey(items[i].Left) == key)
                    leftCount++;
                if (items[i].Right != null && WordKey(items[i].Right) == key)
                    rightCount++;
            }
            return !context.WorkExhausted && leftCount == 1 && rightCount == 1;
        }

        private static bool SameKeyMultiset(IList<PdfReviewWord> left,
            IList<PdfReviewWord> right, DiffContext context)
        {
            if (left == null || right == null || left.Count != right.Count)
                return false;
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (PdfReviewWord word in left)
            {
                if (!context.TryStep())
                    return false;
                string key = WordKey(word);
                int count;
                counts.TryGetValue(key, out count);
                counts[key] = count + 1;
            }
            foreach (PdfReviewWord word in right)
            {
                if (!context.TryStep())
                    return false;
                string key = WordKey(word);
                int count;
                if (!counts.TryGetValue(key, out count) || count == 0)
                    return false;
                if (count == 1)
                    counts.Remove(key);
                else
                    counts[key] = count - 1;
            }
            return !context.WorkExhausted && counts.Count == 0;
        }

        private static bool SinglePage(IList<PdfReviewWord> words,
            DiffContext context, out int pageIndex)
        {
            pageIndex = -1;
            if (words == null)
                return false;
            foreach (PdfReviewWord word in words)
            {
                if (!context.TryStep())
                    return false;
                if (word == null || word.PageIndex < 0)
                    return false;
                if (pageIndex < 0)
                    pageIndex = word.PageIndex;
                else if (pageIndex != word.PageIndex)
                    return false;
            }
            return !context.WorkExhausted && pageIndex >= 0;
        }

        private static bool WindowGeometryBounded(IList<PdfReviewWord> words,
            PdfReviewPage page, DiffContext context)
        {
            double pageWidth = PageViewWidth(page);
            double pageHeight = PageViewHeight(page);
            double pageArea = pageWidth * pageHeight;
            if (words == null || words.Count == 0 ||
                !FinitePositive(pageWidth) || !FinitePositive(pageHeight) ||
                !FinitePositive(pageArea))
                return false;
            double area = 0;
            double left = double.MaxValue, bottom = double.MaxValue;
            double right = double.MinValue, top = double.MinValue;
            foreach (PdfReviewWord word in words)
            {
                if (!context.TryStep())
                    return false;
                if (word == null || !ValidReviewBox(word.Box))
                    return false;
                double width = word.Box.Right - word.Box.Left;
                double height = word.Box.Top - word.Box.Bottom;
                area += width * height;
                left = Math.Min(left, word.Box.Left);
                bottom = Math.Min(bottom, word.Box.Bottom);
                right = Math.Max(right, word.Box.Right);
                top = Math.Max(top, word.Box.Top);
            }
            double boundsArea = (right - left) * (top - bottom);
            return !context.WorkExhausted && Finite(area) && Finite(boundsArea) &&
                area <= ReconcileMaxWordAreaShare * pageArea &&
                boundsArea <= ReconcileMaxBoundsAreaShare * pageArea;
        }

        private static bool TryLocalRegistration(ReconcileItem before,
            ReconcileItem after, PdfReviewPage leftPage, PdfReviewPage rightPage,
            out double scaleX, out double scaleY, out double dx, out double dy)
        {
            scaleX = scaleY = 1;
            dx = dy = 0;
            double leftWidth = PageViewWidth(leftPage);
            double leftHeight = PageViewHeight(leftPage);
            double rightWidth = PageViewWidth(rightPage);
            double rightHeight = PageViewHeight(rightPage);
            if (!FinitePositive(leftWidth) || !FinitePositive(leftHeight) ||
                !FinitePositive(rightWidth) || !FinitePositive(rightHeight))
                return false;
            scaleX = rightWidth / leftWidth;
            scaleY = rightHeight / leftHeight;
            if (scaleX < ReconcileMinPageScale || scaleX > ReconcileMaxPageScale ||
                scaleY < ReconcileMinPageScale || scaleY > ReconcileMaxPageScale)
                return false;

            double beforeDx = CenterX(before.Right.Box) -
                CenterX(before.Left.Box) * scaleX;
            double beforeDy = CenterY(before.Right.Box) -
                CenterY(before.Left.Box) * scaleY;
            double afterDx = CenterX(after.Right.Box) -
                CenterX(after.Left.Box) * scaleX;
            double afterDy = CenterY(after.Right.Box) -
                CenterY(after.Left.Box) * scaleY;
            double height = Math.Max(Math.Max(BoxHeight(before.Left.Box),
                BoxHeight(before.Right.Box)), Math.Max(BoxHeight(after.Left.Box),
                BoxHeight(after.Right.Box)));
            double tolerance = Math.Max(2.5, 0.35 * height);
            if (Math.Abs(beforeDx - afterDx) > tolerance ||
                Math.Abs(beforeDy - afterDy) > tolerance)
                return false;
            dx = (beforeDx + afterDx) / 2.0;
            dy = (beforeDy + afterDy) / 2.0;
            return Finite(dx) && Finite(dy);
        }

        /// <summary>
        /// Marked-content identifiers are local to one PDF page, so their numeric values
        /// cannot be compared across revisions. Trusted IDs still corroborate structure:
        /// the geometry-selected pairs, including both anchors, must induce one bijection
        /// between the known left and right block partitions.
        /// </summary>
        private static bool ConsistentBlockCorrespondence(ReconcileItem before,
            ReconcileItem after, IList<PdfReviewWord> left,
            IList<PdfReviewWord> right, int[] rightByLeft, DiffContext context)
        {
            if (left == null || right == null || rightByLeft == null ||
                rightByLeft.Length != left.Count ||
                (before == null && after == null))
                return false;
            var rightByLeftBlock = new Dictionary<int, int>();
            var leftByRightBlock = new Dictionary<int, int>();
            if ((before != null && (!context.TryStep() ||
                    !AddBlockCorrespondence(before.Left, before.Right,
                        rightByLeftBlock, leftByRightBlock))) ||
                (after != null && (!context.TryStep() ||
                    !AddBlockCorrespondence(after.Left, after.Right,
                        rightByLeftBlock, leftByRightBlock))))
                return false;
            for (int i = 0; i < left.Count; i++)
            {
                if (!context.TryStep())
                    return false;
                int rightIndex = rightByLeft[i];
                if (rightIndex < 0 || rightIndex >= right.Count ||
                    !AddBlockCorrespondence(left[i], right[rightIndex], rightByLeftBlock,
                        leftByRightBlock))
                    return false;
            }
            return !context.WorkExhausted;
        }

        private static bool AddBlockCorrespondence(PdfReviewWord left,
            PdfReviewWord right, IDictionary<int, int> rightByLeftBlock,
            IDictionary<int, int> leftByRightBlock)
        {
            if (left == null || right == null)
                return false;
            if (left.BlockId < 0 || right.BlockId < 0)
                return true;
            int mapped;
            if (rightByLeftBlock.TryGetValue(left.BlockId, out mapped) &&
                mapped != right.BlockId)
                return false;
            if (leftByRightBlock.TryGetValue(right.BlockId, out mapped) &&
                mapped != left.BlockId)
                return false;
            rightByLeftBlock[left.BlockId] = right.BlockId;
            leftByRightBlock[right.BlockId] = left.BlockId;
            return true;
        }

        private static void CountGeometryMatchings(int depth, int[] order,
            IList<List<int>> edges, bool[] used, int[] current, DiffContext context,
            ref int solutions, ref int[] unique)
        {
            if (solutions > 1 || (context != null && !context.TryStep()))
                return;
            if (depth == order.Length)
            {
                if (context != null && !context.TryReserve(current.Length))
                    return;
                solutions++;
                if (solutions == 1)
                    unique = (int[])current.Clone();
                return;
            }
            int leftIndex = order[depth];
            foreach (int rightIndex in edges[leftIndex])
            {
                if (context != null && !context.TryStep())
                    return;
                if (used[rightIndex])
                    continue;
                used[rightIndex] = true;
                current[leftIndex] = rightIndex;
                CountGeometryMatchings(depth + 1, order, edges, used, current,
                    context, ref solutions, ref unique);
                current[leftIndex] = -1;
                used[rightIndex] = false;
                if (solutions > 1 || (context != null && context.WorkExhausted))
                    return;
            }
        }

        private static bool GeometryColocated(PdfReviewBox left, PdfReviewBox right,
            double scaleX, double scaleY, double dx, double dy)
        {
            double mappedLeft = left.Left * scaleX + dx;
            double mappedRight = left.Right * scaleX + dx;
            double mappedBottom = left.Bottom * scaleY + dy;
            double mappedTop = left.Top * scaleY + dy;
            double mappedWidth = mappedRight - mappedLeft;
            double mappedHeight = mappedTop - mappedBottom;
            double rightWidth = right.Right - right.Left;
            double rightHeight = right.Top - right.Bottom;
            if (!FinitePositive(mappedWidth) || !FinitePositive(mappedHeight) ||
                !FinitePositive(rightWidth) || !FinitePositive(rightHeight) ||
                Math.Min(mappedWidth, rightWidth) / Math.Max(mappedWidth, rightWidth) < 0.55 ||
                Math.Min(mappedHeight, rightHeight) / Math.Max(mappedHeight, rightHeight) < 0.65)
                return false;

            double xOverlap = Math.Min(mappedRight, right.Right) -
                Math.Max(mappedLeft, right.Left);
            double yOverlap = Math.Min(mappedTop, right.Top) -
                Math.Max(mappedBottom, right.Bottom);
            if (xOverlap < 0.45 * Math.Min(mappedWidth, rightWidth) ||
                yOverlap < 0.55 * Math.Min(mappedHeight, rightHeight))
                return false;
            double xDistance = Math.Abs((mappedLeft + mappedRight) / 2.0 -
                CenterX(right));
            double yDistance = Math.Abs((mappedBottom + mappedTop) / 2.0 -
                CenterY(right));
            return xDistance <= Math.Max(4.0, 0.25 * Math.Max(mappedWidth, rightWidth)) &&
                yDistance <= Math.Max(3.0, 0.35 * Math.Max(mappedHeight, rightHeight));
        }

        private static bool CompatibleBlocks(PdfReviewWord left, PdfReviewWord right)
        {
            return left != null && right != null &&
                (left.BlockId < 0 || right.BlockId < 0 || left.BlockId == right.BlockId);
        }

        private static bool ValidReviewBox(PdfReviewBox box)
        {
            return Finite(box.Left) && Finite(box.Right) && Finite(box.Bottom) &&
                Finite(box.Top) && box.Right > box.Left && box.Top > box.Bottom;
        }

        private static bool FinitePositive(double value)
        {
            return Finite(value) && value > 0;
        }

        private static bool Finite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static double CenterX(PdfReviewBox box)
        {
            return box.Left + (box.Right - box.Left) / 2.0;
        }

        private static double CenterY(PdfReviewBox box)
        {
            return box.Bottom + (box.Top - box.Bottom) / 2.0;
        }

        private static double BoxHeight(PdfReviewBox box)
        {
            return box.Top - box.Bottom;
        }

        private static bool TryReserveOperationCopy(PdfReviewWordOp operation,
            DiffContext context)
        {
            if (operation == null)
                return false;
            long units = 1L + operation.LeftWords.Count + operation.RightWords.Count +
                operation.Matches.Count;
            return context.TryReserve(units);
        }

        private static bool TryAppendReconcileItem(List<PdfReviewWordOp> operations,
            ReconcileItem item, DiffContext context)
        {
            if (!context.TryStep() || operations == null || item == null)
                return false;
            if (item.Barrier != null)
            {
                if (!TryReserveOperationCopy(item.Barrier, context))
                    return false;
                AppendOperation(operations, item.Barrier);
                return !context.WorkExhausted;
            }
            if (item.Kind == PdfReviewDiffKind.Delete && item.Left != null)
            {
                if (!context.TryReserve(1))
                    return false;
                Append(operations, PdfReviewDiffKind.Delete, item.Left);
            }
            else if (item.Kind == PdfReviewDiffKind.Insert && item.Right != null)
            {
                if (!context.TryReserve(1))
                    return false;
                Append(operations, PdfReviewDiffKind.Insert, item.Right);
            }
            else if (item.Kind == PdfReviewDiffKind.Equal && item.Left != null &&
                item.Right != null && item.MatchKind != PdfReviewMatchKind.None)
            {
                if (!context.TryReserve(3))
                    return false;
                AppendEqual(operations, item.Left, item.Right, item.MatchKind);
            }
            else
            {
                return false;
            }
            return !context.WorkExhausted;
        }

        private sealed class SplitJoinWindow
        {
            public int Start;
            public int End;
            public PdfReviewWordOp Equal;
        }

        /// <summary>
        /// Превращает только доказанный 1↔2/2↔1 hunk в Equal(SplitJoin). Совпадения
        /// текста недостаточно: нужны два устойчивых Exact-anchor, одна page-pair,
        /// точный source-backed separator на split-стороне и посимвольный trusted
        /// source-span без разделителя внутри joined-слова. При сомнении hunk остаётся
        /// обычным Delete/Insert.
        /// </summary>
        internal static void ClassifySplitJoin(PdfReviewResult result)
        {
            ClassifySplitJoin(result, PdfReviewLimits.Default(), null);
        }

        internal static void ClassifySplitJoin(PdfReviewResult result,
            PdfReviewLimits limits, Func<bool> cancelled)
        {
            if (limits == null)
                limits = PdfReviewLimits.Default();
            ClassifySplitJoin(result, new ReviewPageLookup(result),
                new WhitespaceEvidenceCache(), new DiffContext(limits, cancelled));
        }

        private static void ClassifySplitJoin(PdfReviewResult result,
            ReviewPageLookup pages, WhitespaceEvidenceCache evidence,
            DiffContext context)
        {
            context.ThrowIfCancellation();
            if (result == null || result.Left == null || result.Right == null ||
                result.Operations.Count == 0)
                return;
            List<ReconcileItem> items;
            if (!TryReconcileItems(result.Operations, context, out items) ||
                items.Count < 5 || context.WorkExhausted)
                return;

            WhitespaceEvidenceIndex leftEvidence;
            WhitespaceEvidenceIndex rightEvidence;
            if (!evidence.TryGet(result, context, out leftEvidence, out rightEvidence))
                return;
            var accepted = new List<SplitJoinWindow>();
            int anchorIndex = 0;
            while (anchorIndex < items.Count)
            {
                if (!context.TryStep())
                    return;
                if (!items[anchorIndex].IsExact)
                {
                    anchorIndex++;
                    continue;
                }
                int start = anchorIndex + 1;
                int after = start;
                while (after < items.Count && !items[after].IsExact &&
                    items[after].Barrier == null)
                {
                    if (!context.TryStep())
                        return;
                    after++;
                }
                if (after > start && after < items.Count && items[after].IsExact)
                {
                    SplitJoinWindow window;
                    if (TrySplitJoinWindow(pages, items, anchorIndex, after,
                        leftEvidence, rightEvidence, context, out window))
                        accepted.Add(window);
                    if (context.WorkExhausted)
                        return;
                }
                anchorIndex = after > anchorIndex ? after : anchorIndex + 1;
            }
            if (accepted.Count == 0)
                return;

            var rebuilt = new List<PdfReviewWordOp>();
            int windowIndex = 0;
            for (int i = 0; i < items.Count;)
            {
                if (!context.TryStep())
                    return;
                SplitJoinWindow window = windowIndex < accepted.Count
                    ? accepted[windowIndex] : null;
                if (window != null && window.Start < i)
                    return;
                if (window != null && i == window.Start)
                {
                    if (!TryReserveOperationCopy(window.Equal, context))
                        return;
                    AppendOperation(rebuilt, window.Equal);
                    i = window.End + 1;
                    windowIndex++;
                }
                else
                {
                    if (!TryAppendReconcileItem(rebuilt, items[i], context))
                        return;
                    i++;
                }
            }
            if (windowIndex != accepted.Count || context.WorkExhausted ||
                !context.TryReserve(rebuilt.Count))
                return;
            context.ThrowIfCancellation();
            result.ReplaceOperations(rebuilt);
        }

        private static bool TrySplitJoinWindow(ReviewPageLookup pages,
            IList<ReconcileItem> items, int beforeIndex, int afterIndex,
            WhitespaceEvidenceIndex leftEvidence,
            WhitespaceEvidenceIndex rightEvidence, DiffContext context,
            out SplitJoinWindow window)
        {
            window = null;
            if (!context.TryStep() || beforeIndex < 0 ||
                afterIndex <= beforeIndex + 1 || afterIndex >= items.Count)
                return false;

            // Окно собирается до проверки якорей: третье слово на любой стороне уже
            // исключает 1↔2/2↔1, и широкое окно (вставленная страница) отвергается за
            // несколько шагов, а не обходом целиком плюс двумя окрестностями якорей.
            var left = new List<PdfReviewWord>();
            var right = new List<PdfReviewWord>();
            for (int i = beforeIndex + 1; i < afterIndex; i++)
            {
                if (!context.TryStep())
                    return false;
                ReconcileItem item = items[i];
                if (item.Barrier != null || (item.Kind != PdfReviewDiffKind.Delete &&
                    item.Kind != PdfReviewDiffKind.Insert))
                    return false;
                if (item.Kind == PdfReviewDiffKind.Delete && item.Left != null)
                    left.Add(item.Left);
                else if (item.Kind == PdfReviewDiffKind.Insert && item.Right != null)
                    right.Add(item.Right);
                else
                    return false;
                if (left.Count > 2 || right.Count > 2)
                    return false;
            }
            if (!((left.Count == 1 && right.Count == 2) ||
                  (left.Count == 2 && right.Count == 1)) ||
                !StableReconcileAnchor(items, beforeIndex, context) ||
                !StableReconcileAnchor(items, afterIndex, context))
                return false;

            ReconcileItem before = items[beforeIndex];
            ReconcileItem after = items[afterIndex];
            int leftPage, rightPage;
            if (!SinglePage(left, context, out leftPage) ||
                !SinglePage(right, context, out rightPage) ||
                before.Left.PageIndex != leftPage || before.Right.PageIndex != rightPage ||
                after.Left.PageIndex != leftPage || after.Right.PageIndex != rightPage ||
                !pages.IsExplicitPair(leftPage, rightPage, context))
                return false;
            PdfReviewPage leftOwner = pages.Page(true, leftPage);
            PdfReviewPage rightOwner = pages.Page(false, rightPage);
            if (leftOwner == null || rightOwner == null)
                return false;

            bool leftSplit = left.Count == 2;
            IList<PdfReviewWord> split = leftSplit
                ? (IList<PdfReviewWord>)left : right;
            PdfReviewWord joined = leftSplit ? right[0] : left[0];
            int splitOffset;
            if (!TrySplitJoinText(joined, split[0], split[1], context,
                    out splitOffset))
                return false;

            PdfReviewWhitespaceEvidence splitBoundary;
            List<PdfReviewWhitespaceAtom> splitAtoms;
            WhitespaceEvidenceIndex splitIndex = leftSplit
                ? leftEvidence : rightEvidence;
            bool lineBreak;
            if (!splitIndex.TryGet(split[0], split[1], out splitBoundary,
                    out splitAtoms) || splitAtoms.Count == 0 ||
                !SplitFlowCorroborated(split[0], split[1], splitAtoms, context,
                    out lineBreak))
                return false;

            double scaleX, scaleY, dx, dy;
            if (!TryLocalRegistration(before, after, leftOwner, rightOwner,
                    out scaleX, out scaleY, out dx, out dy) ||
                !SplitJoinGeometryColocated(left, right, scaleX, scaleY, dx, dy,
                    lineBreak))
                return false;

            PdfReviewWhitespaceEvidence joinedBoundary =
                JoinedEmptyBoundary(leftSplit ? rightOwner : leftOwner, joined,
                    splitOffset, context);
            if (joinedBoundary == null)
                return false;
            var equal = new PdfReviewWordOp
            {
                Kind = PdfReviewDiffKind.Equal,
                MatchKind = PdfReviewMatchKind.SplitJoin
            };
            equal.LeftWords.AddRange(left);
            equal.RightWords.AddRange(right);
            equal.SplitJoinLeftBoundary = leftSplit
                ? splitBoundary : joinedBoundary;
            equal.SplitJoinRightBoundary = leftSplit
                ? joinedBoundary : splitBoundary;
            window = new SplitJoinWindow
            {
                Start = beforeIndex + 1,
                End = afterIndex - 1,
                Equal = equal
            };
            return true;
        }

        private static bool TrySplitJoinText(PdfReviewWord joined,
            PdfReviewWord first, PdfReviewWord second, DiffContext context,
            out int splitOffset)
        {
            splitOffset = -1;
            if (!context.TryStep() || !SourceUnitAligned(joined, context) ||
                !SourceUnitAligned(first, context) ||
                !SourceUnitAligned(second, context) ||
                !CompatibleBlocks(joined, first) || !CompatibleBlocks(joined, second) ||
                !SafeSplitJoinText(joined.Text, context) ||
                !SafeSplitJoinText(first.Text, context) ||
                !SafeSplitJoinText(second.Text, context))
                return false;

            string joinedText = joined.Text;
            string joinedKey = WordKey(joined);
            string firstKey = WordKey(first);
            string secondKey = WordKey(second);
            long stringWork = 1;
            stringWork = SaturatingAdd(stringWork, joinedText.Length);
            stringWork = SaturatingAdd(stringWork, first.Text.Length);
            stringWork = SaturatingAdd(stringWork, second.Text.Length);
            stringWork = SaturatingAdd(stringWork, joinedKey.Length);
            stringWork = SaturatingAdd(stringWork, firstKey.Length);
            stringWork = SaturatingAdd(stringWork, secondKey.Length);
            stringWork = SaturatingAdd(stringWork, stringWork);
            if (!context.TryReserve(stringWork))
                return false;

            string splitText = first.Text + second.Text;
            string splitKey = firstKey + secondKey;
            if (!string.Equals(joinedText, splitText, StringComparison.Ordinal) ||
                !string.Equals(joinedKey, splitKey, StringComparison.Ordinal))
                return false;
            splitOffset = first.Text.Length;
            return splitOffset > 0 && splitOffset < joinedText.Length;
        }

        private static bool SourceUnitAligned(PdfReviewWord word,
            DiffContext context)
        {
            return TrustedSourceAnchor(word, context) &&
                (long)word.SourceEnd - word.SourceStart + 1 == word.SourceText.Length;
        }

        private static bool SafeSplitJoinText(string text, DiffContext context)
        {
            if (!context.TryStep() || string.IsNullOrEmpty(text))
                return false;
            for (int i = 0; i < text.Length; i++)
            {
                if (!context.TryStep())
                    return false;
                char value = text[i];
                UnicodeCategory category = char.GetUnicodeCategory(value);
                if (char.IsSurrogate(value) || char.IsWhiteSpace(value) || value == '�' ||
                    category == UnicodeCategory.NonSpacingMark ||
                    category == UnicodeCategory.SpacingCombiningMark ||
                    category == UnicodeCategory.EnclosingMark ||
                    category == UnicodeCategory.Format ||
                    category == UnicodeCategory.PrivateUse ||
                    category == UnicodeCategory.OtherNotAssigned)
                    return false;
            }
            long normalizationWork = 0;
            normalizationWork = SaturatingAdd(normalizationWork, text.Length);
            normalizationWork = SaturatingAdd(normalizationWork, text.Length);
            normalizationWork = SaturatingAdd(normalizationWork, text.Length);
            if (!context.TryReserve(normalizationWork))
                return false;
            return text.IsNormalized(NormalizationForm.FormC) &&
                string.Equals(text, text.Normalize(NormalizationForm.FormKC),
                    StringComparison.Ordinal);
        }

        private static bool SplitFlowCorroborated(PdfReviewWord first,
            PdfReviewWord second, IList<PdfReviewWhitespaceAtom> atoms,
            DiffContext context, out bool lineBreak)
        {
            lineBreak = false;
            if (!context.TryStep() || first == null || second == null ||
                !ValidReviewBox(first.Box) || !ValidReviewBox(second.Box) ||
                !TryContainsLineBreak(atoms, context, out lineBreak))
                return false;
            if (!lineBreak)
                return SameTextBand(first.Box, second.Box);
            if (SameTextBand(first.Box, second.Box))
                return true;
            double firstHeight = BoxHeight(first.Box);
            double secondHeight = BoxHeight(second.Box);
            double maxHeight = Math.Max(firstHeight, secondHeight);
            double step = CenterY(first.Box) - CenterY(second.Box);
            return maxHeight > 0 && step >= 0.35 * Math.Min(firstHeight, secondHeight) &&
                step <= 3.5 * maxHeight &&
                Math.Abs(first.Box.Left - second.Box.Left) <= 4.0 * maxHeight;
        }

        private static bool SplitJoinGeometryColocated(IList<PdfReviewWord> left,
            IList<PdfReviewWord> right, double scaleX, double scaleY,
            double dx, double dy, bool lineBreak)
        {
            PdfReviewBox leftBounds, rightBounds;
            if (!TryWordBounds(left, out leftBounds) ||
                !TryWordBounds(right, out rightBounds))
                return false;
            double mappedLeft = leftBounds.Left * scaleX + dx;
            double mappedRight = leftBounds.Right * scaleX + dx;
            double mappedBottom = leftBounds.Bottom * scaleY + dy;
            double mappedTop = leftBounds.Top * scaleY + dy;
            double mappedWidth = mappedRight - mappedLeft;
            double mappedHeight = mappedTop - mappedBottom;
            double rightWidth = rightBounds.Right - rightBounds.Left;
            double rightHeight = rightBounds.Top - rightBounds.Bottom;
            if (!FinitePositive(mappedWidth) || !FinitePositive(mappedHeight) ||
                !FinitePositive(rightWidth) || !FinitePositive(rightHeight))
                return false;
            double widthRatio = Math.Min(mappedWidth, rightWidth) /
                Math.Max(mappedWidth, rightWidth);
            double xOverlap = Math.Min(mappedRight, rightBounds.Right) -
                Math.Max(mappedLeft, rightBounds.Left);
            double yOverlap = Math.Min(mappedTop, rightBounds.Top) -
                Math.Max(mappedBottom, rightBounds.Bottom);
            if (widthRatio < (lineBreak ? 0.35 : 0.55) ||
                xOverlap < (lineBreak ? 0.20 : 0.40) *
                    Math.Min(mappedWidth, rightWidth))
                return false;
            if (!lineBreak && (Math.Min(mappedHeight, rightHeight) /
                    Math.Max(mappedHeight, rightHeight) < 0.60 ||
                    yOverlap < 0.50 * Math.Min(mappedHeight, rightHeight)))
                return false;
            double xDistance = Math.Abs((mappedLeft + mappedRight) / 2.0 -
                CenterX(rightBounds));
            double yDistance = Math.Abs((mappedBottom + mappedTop) / 2.0 -
                CenterY(rightBounds));
            return xDistance <= Math.Max(5.0,
                       0.35 * Math.Max(mappedWidth, rightWidth)) &&
                yDistance <= Math.Max(4.0, (lineBreak ? 2.0 : 0.45) *
                    Math.Max(mappedHeight, rightHeight));
        }

        private static bool TryWordBounds(IList<PdfReviewWord> words,
            out PdfReviewBox bounds)
        {
            bounds = new PdfReviewBox();
            if (words == null || words.Count == 0)
                return false;
            double left = double.MaxValue, bottom = double.MaxValue;
            double right = double.MinValue, top = double.MinValue;
            foreach (PdfReviewWord word in words)
            {
                if (word == null || !ValidReviewBox(word.Box))
                    return false;
                left = Math.Min(left, word.Box.Left);
                bottom = Math.Min(bottom, word.Box.Bottom);
                right = Math.Max(right, word.Box.Right);
                top = Math.Max(top, word.Box.Top);
            }
            bounds = new PdfReviewBox
            {
                Left = left,
                Bottom = bottom,
                Right = right,
                Top = top
            };
            return true;
        }

        private static PdfReviewWhitespaceEvidence JoinedEmptyBoundary(
            PdfReviewPage page, PdfReviewWord joined, int splitOffset,
            DiffContext context)
        {
            if (!context.TryStep() || page == null || joined == null || splitOffset <= 0 ||
                splitOffset >= joined.Text.Length || !SourceUnitAligned(joined, context) ||
                joined.SourceStart + splitOffset > joined.SourceEnd)
                return null;
            double ratio = (double)splitOffset / joined.Text.Length;
            double x = joined.Box.Left + (joined.Box.Right - joined.Box.Left) * ratio;
            double pageWidth = PageViewWidth(page);
            double pageHeight = PageViewHeight(page);
            if (!FinitePositive(pageWidth) || !FinitePositive(pageHeight))
                return null;
            x = Math.Max(0.5, Math.Min(pageWidth - 0.5, x));
            double bottom = joined.Box.Top + 0.5;
            if (bottom + 1.0 > pageHeight)
                bottom = Math.Max(0, joined.Box.Bottom - 1.5);
            return new PdfReviewWhitespaceEvidence
            {
                PageIndex = joined.PageIndex,
                Within = joined,
                TextOffset = splitOffset,
                RawText = "",
                LogicalText = "",
                MarkerBox = new PdfReviewBox
                {
                    Left = x - 0.5,
                    Right = x + 0.5,
                    Bottom = bottom,
                    Top = Math.Min(pageHeight, bottom + 1.0)
                }
            };
        }

        private sealed class ExactWordMatches
        {
            private readonly Dictionary<PdfReviewWord, PdfReviewWord> _rightByLeft =
                new Dictionary<PdfReviewWord, PdfReviewWord>();

            public bool TryRight(PdfReviewWord left, out PdfReviewWord right)
            {
                right = null;
                return left != null && _rightByLeft.TryGetValue(left, out right);
            }

            public static bool TryFrom(IList<PdfReviewWordOp> operations,
                DiffContext context, out ExactWordMatches result)
            {
                result = new ExactWordMatches();
                var matches = new List<PdfReviewWordMatch>();
                var leftCounts = new Dictionary<PdfReviewWord, int>();
                var rightCounts = new Dictionary<PdfReviewWord, int>();
                if (!context.TryStep())
                    return false;
                if (operations != null)
                {
                    foreach (PdfReviewWordOp operation in operations)
                    {
                        if (!context.TryStep())
                            return false;
                        if (operation == null || operation.Kind != PdfReviewDiffKind.Equal)
                            continue;
                        foreach (PdfReviewWordMatch match in operation.Matches)
                        {
                            if (!context.TryStep())
                                return false;
                            if (match == null || match.Kind != PdfReviewMatchKind.Exact ||
                                match.Left == null || match.Right == null)
                                continue;
                            string leftKey = WordKey(match.Left);
                            string rightKey = WordKey(match.Right);
                            if (!context.TryReserve(Math.Max(leftKey.Length,
                                    rightKey.Length)))
                                return false;
                            if (!string.Equals(leftKey, rightKey,
                                    StringComparison.Ordinal))
                                continue;
                            if (!context.TryReserve(3))
                                return false;
                            matches.Add(match);
                            Increment(leftCounts, match.Left);
                            Increment(rightCounts, match.Right);
                        }
                    }
                }
                foreach (PdfReviewWordMatch match in matches)
                {
                    if (!context.TryStep())
                        return false;
                    if (leftCounts[match.Left] == 1 && rightCounts[match.Right] == 1)
                    {
                        if (!context.TryStep())
                            return false;
                        result._rightByLeft.Add(match.Left, match.Right);
                    }
                }
                return !context.WorkExhausted;
            }

            private static void Increment(Dictionary<PdfReviewWord, int> counts,
                PdfReviewWord word)
            {
                int value;
                counts.TryGetValue(word, out value);
                counts[word] = value + 1;
            }
        }

        private sealed class WhitespaceBoundaryKey : IEquatable<WhitespaceBoundaryKey>
        {
            private readonly PdfReviewWord _before;
            private readonly PdfReviewWord _after;

            public WhitespaceBoundaryKey(PdfReviewWord before, PdfReviewWord after)
            {
                _before = before;
                _after = after;
            }

            public bool Equals(WhitespaceBoundaryKey other)
            {
                return other != null && ReferenceEquals(_before, other._before) &&
                    ReferenceEquals(_after, other._after);
            }

            public override bool Equals(object obj)
            {
                return Equals(obj as WhitespaceBoundaryKey);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int before = _before == null ? 0 :
                        System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_before);
                    int after = _after == null ? 0 :
                        System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_after);
                    return before * 397 ^ after;
                }
            }
        }

        private sealed class WhitespaceEvidenceIndex
        {
            private readonly Dictionary<WhitespaceBoundaryKey, PdfReviewWhitespaceEvidence>
                _byBoundary =
                    new Dictionary<WhitespaceBoundaryKey, PdfReviewWhitespaceEvidence>();
            private readonly HashSet<WhitespaceBoundaryKey> _ambiguous =
                new HashSet<WhitespaceBoundaryKey>();
            private readonly Dictionary<PdfReviewWhitespaceEvidence,
                List<PdfReviewWhitespaceAtom>> _atoms =
                    new Dictionary<PdfReviewWhitespaceEvidence,
                        List<PdfReviewWhitespaceAtom>>();
            private readonly List<PdfReviewWhitespaceEvidence> _ordered =
                new List<PdfReviewWhitespaceEvidence>();

            public IList<PdfReviewWhitespaceEvidence> Ordered
            {
                get { return _ordered; }
            }

            public void Add(PdfReviewWhitespaceEvidence evidence,
                List<PdfReviewWhitespaceAtom> atoms)
            {
                var key = new WhitespaceBoundaryKey(evidence.Before, evidence.After);
                if (_ambiguous.Contains(key))
                    return;
                PdfReviewWhitespaceEvidence existing;
                if (_byBoundary.TryGetValue(key, out existing))
                {
                    _byBoundary.Remove(key);
                    _ambiguous.Add(key);
                    return;
                }
                _byBoundary.Add(key, evidence);
                _atoms.Add(evidence, atoms);
                _ordered.Add(evidence);
            }

            public bool TryGet(PdfReviewWord before, PdfReviewWord after,
                out PdfReviewWhitespaceEvidence evidence,
                out List<PdfReviewWhitespaceAtom> atoms)
            {
                atoms = null;
                var key = new WhitespaceBoundaryKey(before, after);
                if (!_byBoundary.TryGetValue(key, out evidence) || evidence == null ||
                    !_atoms.TryGetValue(evidence, out atoms))
                {
                    evidence = null;
                    atoms = null;
                    return false;
                }
                return true;
            }
        }

        /// <summary>
        /// Индексы пробельных свидетельств зависят только от документов, а не от операций,
        /// поэтому split/join и сравнение пробелов одного Compare строят и оплачивают их один
        /// раз. Запоминается только полностью построенная пара: индекс, оборванный
        /// исчерпанием, следующей стадии не достаётся.
        /// </summary>
        private sealed class WhitespaceEvidenceCache
        {
            private WhitespaceEvidenceIndex _left;
            private WhitespaceEvidenceIndex _right;

            public bool TryGet(PdfReviewResult result, DiffContext context,
                out WhitespaceEvidenceIndex left, out WhitespaceEvidenceIndex right)
            {
                if (_left == null)
                {
                    WhitespaceEvidenceIndex builtLeft =
                        BuildWhitespaceEvidenceIndex(result.Left, context);
                    WhitespaceEvidenceIndex builtRight = context.WorkExhausted ? null :
                        BuildWhitespaceEvidenceIndex(result.Right, context);
                    if (context.WorkExhausted)
                    {
                        left = null;
                        right = null;
                        return false;
                    }
                    _left = builtLeft;
                    _right = builtRight;
                }
                left = _left;
                right = _right;
                return true;
            }
        }

        /// <summary>
        /// Сравнивает только буквально декодированные пробельные границы, привязанные к
        /// однозначным Exact-парам слов. Геометрия может лишь отклонить сомнительный перенос;
        /// отсутствие пробела должно быть так же положительно доказано пустым source-span.
        /// </summary>
        internal static void CompareWhitespace(PdfReviewResult result, PdfReviewLimits limits)
        {
            CompareWhitespace(result, limits, (Func<bool>)null);
        }

        internal static void CompareWhitespace(PdfReviewResult result,
            PdfReviewLimits limits, Func<bool> cancelled)
        {
            if (limits == null)
                limits = PdfReviewLimits.Default();
            CompareWhitespace(result, limits, new WhitespaceEvidenceCache(),
                new DiffContext(limits, cancelled));
        }

        private static void CompareWhitespace(PdfReviewResult result,
            PdfReviewLimits limits, WhitespaceEvidenceCache evidence, DiffContext context)
        {
            context.ThrowIfCancellation();
            if (result == null || result.Left == null || result.Right == null)
                return;
            if (limits == null)
                limits = PdfReviewLimits.Default();

            ExactWordMatches exact;
            if (!ExactWordMatches.TryFrom(result.Operations, context, out exact))
                return;
            WhitespaceEvidenceIndex leftIndex;
            WhitespaceEvidenceIndex rightIndex;
            if (!evidence.TryGet(result, context, out leftIndex, out rightIndex))
                return;
            var changes = new List<PdfReviewWhitespaceChange>();

            foreach (PdfReviewWhitespaceEvidence leftEvidence in leftIndex.Ordered)
            {
                if (!context.TryStep())
                    return;
                PdfReviewWhitespaceEvidence uniqueLeft;
                List<PdfReviewWhitespaceAtom> leftAtoms;
                if (leftEvidence == null || !leftIndex.TryGet(leftEvidence.Before,
                    leftEvidence.After, out uniqueLeft, out leftAtoms) ||
                    !ReferenceEquals(leftEvidence, uniqueLeft))
                    continue;

                PdfReviewWord rightBefore = null;
                PdfReviewWord rightAfter = null;
                if (leftEvidence.Before != null &&
                    !exact.TryRight(leftEvidence.Before, out rightBefore))
                    continue;
                if (leftEvidence.After != null &&
                    !exact.TryRight(leftEvidence.After, out rightAfter))
                    continue;

                PdfReviewWhitespaceEvidence rightEvidence;
                List<PdfReviewWhitespaceAtom> rightAtoms;
                if (!rightIndex.TryGet(rightBefore, rightAfter, out rightEvidence,
                    out rightAtoms))
                    continue;
                bool atomsEqual;
                if (!TryAtomSequencesEqual(leftAtoms, rightAtoms, context,
                        out atomsEqual))
                    return;
                if (atomsEqual)
                    continue;

                var change = new PdfReviewWhitespaceChange
                {
                    Left = leftEvidence,
                    Right = rightEvidence
                };
                if (!DiffWhitespaceAtoms(leftAtoms, rightAtoms, limits,
                        change.DeletedAtoms, change.InsertedAtoms, context))
                    return;
                if (change.DeletedAtoms.Count > 0 || change.InsertedAtoms.Count > 0)
                {
                    if (!context.TryStep())
                        return;
                    changes.Add(change);
                }
            }
            if (!AppendSplitJoinWhitespace(result, leftIndex, rightIndex, limits,
                    changes, context) || context.WorkExhausted)
                return;

            long publicationWork = 1;
            publicationWork = SaturatingAdd(publicationWork,
                result.WhitespaceChanges.Count);
            publicationWork = SaturatingAdd(publicationWork, changes.Count);
            if (!context.TryReserve(publicationWork))
                return;
            context.ThrowIfCancellation();
            result.ReplaceWhitespaceChanges(changes);
        }

        private static bool AppendSplitJoinWhitespace(PdfReviewResult result,
            WhitespaceEvidenceIndex leftIndex, WhitespaceEvidenceIndex rightIndex,
            PdfReviewLimits limits, ICollection<PdfReviewWhitespaceChange> changes,
            DiffContext context)
        {
            foreach (PdfReviewWordOp op in result.Operations)
            {
                if (!context.TryStep())
                    return false;
                if (op == null || op.Kind != PdfReviewDiffKind.Equal ||
                    op.MatchKind != PdfReviewMatchKind.SplitJoin)
                    continue;
                bool leftSplit = op.LeftWords.Count == 2 && op.RightWords.Count == 1;
                bool rightSplit = op.LeftWords.Count == 1 && op.RightWords.Count == 2;
                if (!leftSplit && !rightSplit)
                    continue;

                PdfReviewWhitespaceEvidence leftEvidence = op.SplitJoinLeftBoundary;
                PdfReviewWhitespaceEvidence rightEvidence = op.SplitJoinRightBoundary;
                PdfReviewWhitespaceEvidence splitEvidence = leftSplit
                    ? leftEvidence : rightEvidence;
                PdfReviewWhitespaceEvidence joinedEvidence = leftSplit
                    ? rightEvidence : leftEvidence;
                IList<PdfReviewWord> splitWords = leftSplit
                    ? (IList<PdfReviewWord>)op.LeftWords : op.RightWords;
                PdfReviewWord joinedWord = leftSplit
                    ? op.RightWords[0] : op.LeftWords[0];

                PdfReviewWhitespaceEvidence indexedEvidence;
                List<PdfReviewWhitespaceAtom> splitAtoms;
                WhitespaceEvidenceIndex splitIndex = leftSplit
                    ? leftIndex : rightIndex;
                int splitOffset;
                if (!TrySplitJoinText(joinedWord, splitWords[0], splitWords[1],
                        context, out splitOffset) || splitEvidence == null ||
                    !splitIndex.TryGet(splitWords[0], splitWords[1],
                        out indexedEvidence, out splitAtoms) ||
                    !ReferenceEquals(splitEvidence, indexedEvidence) ||
                    splitAtoms.Count == 0 || joinedEvidence == null ||
                    !ReferenceEquals(joinedEvidence.Within, joinedWord) ||
                    joinedEvidence.Before != null || joinedEvidence.After != null ||
                    joinedEvidence.AtPageStart || joinedEvidence.AtPageEnd ||
                    joinedEvidence.PageIndex != joinedWord.PageIndex ||
                    joinedEvidence.TextOffset != splitOffset ||
                    joinedEvidence.RawText == null || joinedEvidence.RawText.Length != 0 ||
                    joinedEvidence.LogicalText == null ||
                    joinedEvidence.LogicalText.Length != 0)
                    continue;

                var joinedAtoms = new List<PdfReviewWhitespaceAtom>();
                IList<PdfReviewWhitespaceAtom> leftAtoms = leftSplit
                    ? (IList<PdfReviewWhitespaceAtom>)splitAtoms : joinedAtoms;
                IList<PdfReviewWhitespaceAtom> rightAtoms = leftSplit
                    ? (IList<PdfReviewWhitespaceAtom>)joinedAtoms : splitAtoms;
                var change = new PdfReviewWhitespaceChange
                {
                    Left = leftEvidence,
                    Right = rightEvidence
                };
                if (!DiffWhitespaceAtoms(leftAtoms, rightAtoms, limits,
                        change.DeletedAtoms, change.InsertedAtoms, context))
                    return false;
                if (change.DeletedAtoms.Count > 0 || change.InsertedAtoms.Count > 0)
                {
                    if (!context.TryStep())
                        return false;
                    changes.Add(change);
                }
            }
            return true;
        }

        private static WhitespaceEvidenceIndex BuildWhitespaceEvidenceIndex(
            PdfReviewDocument document, DiffContext context)
        {
            var result = new WhitespaceEvidenceIndex();
            if (document == null)
                return result;

            PdfReviewWord first = null;
            PdfReviewWord last = null;
            foreach (PdfReviewPage page in document.Pages)
            {
                if (!context.TryStep())
                    return result;
                if (page == null)
                    continue;
                foreach (PdfReviewWord word in page.Words)
                {
                    if (!context.TryStep())
                        return result;
                    if (word == null)
                        continue;
                    if (first == null)
                        first = word;
                    last = word;
                }
            }

            foreach (PdfReviewPage page in document.Pages)
            {
                if (!context.TryStep())
                    return result;
                if (page == null || page.Words.Count == 0)
                    continue;
                var positions = new Dictionary<PdfReviewWord, int>();
                for (int i = 0; i < page.Words.Count; i++)
                {
                    if (!context.TryStep())
                        return result;
                    PdfReviewWord word = page.Words[i];
                    if (word == null)
                        continue;
                    int existing;
                    if (positions.TryGetValue(word, out existing))
                        positions[word] = -1;
                    else
                        positions.Add(word, i);
                }

                foreach (PdfReviewWhitespaceEvidence evidence in page.WhitespaceBoundaries)
                {
                    if (!context.TryStep())
                        return result;
                    List<PdfReviewWhitespaceAtom> atoms;
                    if (TrustedWhitespaceEvidence(page, positions, first, last, evidence,
                        context, out atoms))
                        result.Add(evidence, atoms);
                    if (context.WorkExhausted)
                        return result;
                }
            }
            return result;
        }

        private static bool TrustedWhitespaceEvidence(PdfReviewPage page,
            IDictionary<PdfReviewWord, int> positions, PdfReviewWord first,
            PdfReviewWord last, PdfReviewWhitespaceEvidence evidence,
            DiffContext context, out List<PdfReviewWhitespaceAtom> atoms)
        {
            atoms = null;
            if (page == null || evidence == null || evidence.PageIndex != page.PageIndex ||
                evidence.RawText == null || evidence.LogicalText == null ||
                !TryWhitespaceAtoms(evidence.RawText, evidence.LogicalText, context,
                    out atoms))
                return false;

            int beforePosition = -1;
            int afterPosition = -1;
            bool hasBefore = evidence.Before != null &&
                positions.TryGetValue(evidence.Before, out beforePosition) && beforePosition >= 0;
            bool hasAfter = evidence.After != null &&
                positions.TryGetValue(evidence.After, out afterPosition) && afterPosition >= 0;

            if (!hasBefore)
                beforePosition = -1;
            if (!hasAfter)
                afterPosition = -1;

            if (evidence.AtPageStart)
            {
                return !evidence.AtPageEnd && evidence.Before == null && hasAfter &&
                    afterPosition == 0 && ReferenceEquals(evidence.After, first) &&
                    TrustedSourceAnchor(evidence.After, context);
            }
            if (evidence.AtPageEnd)
            {
                return evidence.Before != null && hasBefore && evidence.After == null &&
                    beforePosition == page.Words.Count - 1 &&
                    ReferenceEquals(evidence.Before, last) &&
                    TrustedSourceAnchor(evidence.Before, context);
            }
            if (!hasBefore || !hasAfter || afterPosition != beforePosition + 1 ||
                evidence.Before.PageIndex != page.PageIndex ||
                evidence.After.PageIndex != page.PageIndex ||
                !TrustedSourceAnchor(evidence.Before, context) ||
                !TrustedSourceAnchor(evidence.After, context) ||
                evidence.Before.SourceEnd >= evidence.After.SourceStart ||
                (evidence.Before.BlockId >= 0 && evidence.After.BlockId >= 0 &&
                 evidence.Before.BlockId != evidence.After.BlockId))
                return false;

            // Непустой explicit CR/LF сам доказывает перевод строки. Для остальных
            // границ требуем одну визуальную строку, чтобы reflow/wrap не стал «пробелом».
            bool lineBreak;
            return TryContainsLineBreak(atoms, context, out lineBreak) &&
                (lineBreak || SameTextBand(evidence.Before.Box, evidence.After.Box));
        }

        private static bool TrustedSourceAnchor(PdfReviewWord word,
            DiffContext context)
        {
            if (!context.TryStep() || word == null || !word.SourceTrusted ||
                word.SourceStart < 0 || word.SourceEnd < word.SourceStart ||
                string.IsNullOrEmpty(word.SourceText) || word.Text == null)
                return false;
            if (!context.TryReserve(Math.Max(word.SourceText.Length,
                    word.Text.Length)))
                return false;
            return string.Equals(word.SourceText, word.Text,
                StringComparison.Ordinal);
        }

        private static bool SameTextBand(PdfReviewBox left, PdfReviewBox right)
        {
            if (!ValidReviewBox(left) || !ValidReviewBox(right))
                return false;
            double overlap = Math.Min(left.Top, right.Top) -
                Math.Max(left.Bottom, right.Bottom);
            double height = Math.Min(left.Top - left.Bottom, right.Top - right.Bottom);
            return height > 0 && overlap >= 0.35 * height;
        }

        private static bool TryContainsLineBreak(
            IList<PdfReviewWhitespaceAtom> atoms, DiffContext context,
            out bool lineBreak)
        {
            lineBreak = false;
            if (!context.TryStep())
                return false;
            if (atoms == null)
                return true;
            foreach (PdfReviewWhitespaceAtom atom in atoms)
            {
                if (!context.TryStep())
                    return false;
                if (atom != null &&
                    atom.Kind == PdfReviewWhitespaceAtomKind.LineBreak)
                {
                    lineBreak = true;
                    return true;
                }
            }
            return true;
        }

        private static bool TryWhitespaceAtoms(string raw, string logical,
            DiffContext context, out List<PdfReviewWhitespaceAtom> atoms)
        {
            atoms = new List<PdfReviewWhitespaceAtom>();
            if (!context.TryStep() || raw == null || logical == null)
                return false;
            var rebuiltLogical = new StringBuilder(raw.Length);
            for (int i = 0; i < raw.Length; i++)
            {
                if (!context.TryStep())
                {
                    atoms = null;
                    return false;
                }
                char value = raw[i];
                PdfReviewWhitespaceAtomKind kind;
                string atomRaw;
                if (value == '\r')
                {
                    if (i + 1 < raw.Length && raw[i + 1] == '\n')
                    {
                        if (!context.TryStep())
                        {
                            atoms = null;
                            return false;
                        }
                        atomRaw = "\r\n";
                        i++;
                    }
                    else
                    {
                        atomRaw = "\r";
                    }
                    kind = PdfReviewWhitespaceAtomKind.LineBreak;
                    rebuiltLogical.Append('\n');
                }
                else if (value == '\n')
                {
                    atomRaw = "\n";
                    kind = PdfReviewWhitespaceAtomKind.LineBreak;
                    rebuiltLogical.Append('\n');
                }
                else
                {
                    if (!char.IsWhiteSpace(value))
                    {
                        atoms = null;
                        return false;
                    }
                    atomRaw = value.ToString();
                    if (value == ' ')
                        kind = PdfReviewWhitespaceAtomKind.Space;
                    else if (value == ' ')
                        kind = PdfReviewWhitespaceAtomKind.NoBreakSpace;
                    else if (value == '\t')
                        kind = PdfReviewWhitespaceAtomKind.Tab;
                    else
                        kind = PdfReviewWhitespaceAtomKind.Other;
                    rebuiltLogical.Append(value);
                }
                atoms.Add(new PdfReviewWhitespaceAtom { Kind = kind, RawText = atomRaw });
            }
            long comparisonWork = 1;
            comparisonWork = SaturatingAdd(comparisonWork, rebuiltLogical.Length);
            comparisonWork = SaturatingAdd(comparisonWork, rebuiltLogical.Length);
            comparisonWork = SaturatingAdd(comparisonWork, logical.Length);
            if (!context.TryReserve(comparisonWork))
            {
                atoms = null;
                return false;
            }
            if (!string.Equals(rebuiltLogical.ToString(), logical,
                StringComparison.Ordinal))
            {
                atoms = null;
                return false;
            }
            return true;
        }

        private static bool TryAtomSequencesEqual(
            IList<PdfReviewWhitespaceAtom> left,
            IList<PdfReviewWhitespaceAtom> right, DiffContext context,
            out bool equal)
        {
            equal = false;
            if (!context.TryStep())
                return false;
            if (left == null || right == null || left.Count != right.Count)
                return true;
            for (int i = 0; i < left.Count; i++)
            {
                if (!context.TryStep())
                    return false;
                if (!WhitespaceAtomsEqual(left[i], right[i]))
                    return true;
            }
            equal = true;
            return true;
        }

        private static bool WhitespaceAtomsEqual(PdfReviewWhitespaceAtom left,
            PdfReviewWhitespaceAtom right)
        {
            if (left == null || right == null || left.Kind != right.Kind)
                return false;
            return left.Kind == PdfReviewWhitespaceAtomKind.LineBreak ||
                string.Equals(left.RawText, right.RawText, StringComparison.Ordinal);
        }

        private enum WhitespaceEditKind
        {
            Delete,
            Insert
        }

        private sealed class WhitespaceEdit
        {
            public WhitespaceEditKind Kind;
            public PdfReviewWhitespaceAtom Atom;
        }

        private static bool DiffWhitespaceAtoms(IList<PdfReviewWhitespaceAtom> left,
            IList<PdfReviewWhitespaceAtom> right, PdfReviewLimits limits,
            ICollection<PdfReviewWhitespaceAtom> deleted,
            ICollection<PdfReviewWhitespaceAtom> inserted, DiffContext context)
        {
            int sequenceOrder;
            if (left == null || right == null || deleted == null || inserted == null ||
                !TryCompareAtomSequences(left, right, context, out sequenceOrder))
                return false;
            bool swapped = sequenceOrder > 0;
            IList<PdfReviewWhitespaceAtom> a = swapped ? right : left;
            IList<PdfReviewWhitespaceAtom> b = swapped ? left : right;
            var edits = new List<WhitespaceEdit>();
            if (!DiffWhitespaceRange(a, 0, a.Count, b, 0, b.Count, limits,
                    edits, context, 0))
                return false;
            foreach (WhitespaceEdit edit in edits)
            {
                if (!context.TryStep())
                    return false;
                if ((!swapped && edit.Kind == WhitespaceEditKind.Delete) ||
                    (swapped && edit.Kind == WhitespaceEditKind.Insert))
                    deleted.Add(edit.Atom);
                else
                    inserted.Add(edit.Atom);
            }
            return !context.WorkExhausted;
        }

        private static bool TryCompareAtomSequences(
            IList<PdfReviewWhitespaceAtom> left,
            IList<PdfReviewWhitespaceAtom> right, DiffContext context,
            out int comparison)
        {
            comparison = 0;
            if (!context.TryStep())
                return false;
            int common = Math.Min(left == null ? 0 : left.Count,
                right == null ? 0 : right.Count);
            for (int i = 0; i < common; i++)
            {
                if (!context.TryStep())
                    return false;
                comparison = CompareAtoms(left[i], right[i]);
                if (comparison != 0)
                    return true;
            }
            int leftCount = left == null ? 0 : left.Count;
            int rightCount = right == null ? 0 : right.Count;
            comparison = leftCount.CompareTo(rightCount);
            return true;
        }

        private static int CompareAtoms(PdfReviewWhitespaceAtom left,
            PdfReviewWhitespaceAtom right)
        {
            if (ReferenceEquals(left, right))
                return 0;
            if (left == null)
                return -1;
            if (right == null)
                return 1;
            int kind = left.Kind.CompareTo(right.Kind);
            if (kind != 0)
                return kind;
            if (left.Kind == PdfReviewWhitespaceAtomKind.LineBreak)
                return 0;
            return string.Compare(left.RawText, right.RawText, StringComparison.Ordinal);
        }

        private static bool DiffWhitespaceRange(IList<PdfReviewWhitespaceAtom> a,
            int aStart, int aEnd, IList<PdfReviewWhitespaceAtom> b, int bStart,
            int bEnd, PdfReviewLimits limits, List<WhitespaceEdit> edits,
            DiffContext context, int recursionDepth)
        {
            if (!context.TryStep())
                return false;
            while (aStart < aEnd && bStart < bEnd)
            {
                if (!context.TryStep())
                    return false;
                if (!WhitespaceAtomsEqual(a[aStart], b[bStart]))
                    break;
                aStart++;
                bStart++;
            }
            while (aStart < aEnd && bStart < bEnd)
            {
                if (!context.TryStep())
                    return false;
                if (!WhitespaceAtomsEqual(a[aEnd - 1], b[bEnd - 1]))
                    break;
                aEnd--;
                bEnd--;
            }

            int n = aEnd - aStart;
            int m = bEnd - bStart;
            int matrixRows;
            int matrixColumns;
            long matrixCells;
            if (n == 0)
            {
                for (int j = bStart; j < bEnd; j++)
                {
                    if (!context.TryStep())
                        return false;
                    AppendWhitespaceEdit(edits, WhitespaceEditKind.Insert, b[j]);
                }
            }
            else if (m == 0)
            {
                for (int i = aStart; i < aEnd; i++)
                {
                    if (!context.TryStep())
                        return false;
                    AppendWhitespaceEdit(edits, WhitespaceEditKind.Delete, a[i]);
                }
            }
            else if (TryGetMatrixSize(n, m, limits.MaxDiffCells,
                out matrixRows, out matrixColumns, out matrixCells))
            {
                if (!WhitespaceLcsCore(a, aStart, aEnd, b, bStart, bEnd,
                        edits, context, matrixRows, matrixColumns, matrixCells))
                    return false;
            }
            else if (recursionDepth >= MaxDiffRecursionDepth)
            {
                context.ExhaustWork();
                return false;
            }
            else
            {
                int aMid;
                int bMid;
                bool bisected = TryBisectCore(a, aStart, aEnd, b, bStart, bEnd,
                    WhitespaceAtomsEqual, context, out aMid, out bMid);
                if (context.WorkExhausted)
                    return false;
                if (bisected && !(aMid == aStart && bMid == bStart) &&
                    !(aMid == aEnd && bMid == bEnd))
                {
                    if (!DiffWhitespaceRange(a, aStart, aMid, b, bStart, bMid,
                            limits, edits, context, recursionDepth + 1) ||
                        !DiffWhitespaceRange(a, aMid, aEnd, b, bMid, bEnd,
                            limits, edits, context, recursionDepth + 1))
                        return false;
                }
                else
                {
                    for (int i = aStart; i < aEnd; i++)
                    {
                        if (!context.TryStep())
                            return false;
                        AppendWhitespaceEdit(edits, WhitespaceEditKind.Delete, a[i]);
                    }
                    for (int j = bStart; j < bEnd; j++)
                    {
                        if (!context.TryStep())
                            return false;
                        AppendWhitespaceEdit(edits, WhitespaceEditKind.Insert, b[j]);
                    }
                }
            }
            return !context.WorkExhausted;
        }

        private static bool WhitespaceLcsCore(IList<PdfReviewWhitespaceAtom> a,
            int aStart, int aEnd, IList<PdfReviewWhitespaceAtom> b, int bStart,
            int bEnd, List<WhitespaceEdit> edits, DiffContext context,
            int rows, int columns, long cells)
        {
            int n = aEnd - aStart;
            int m = bEnd - bStart;
            long work = SaturatingAdd(cells, (long)n + m);
            if (!context.TryReserve(work))
                return false;
            var dp = new int[rows, columns];
            for (int i = n - 1; i >= 0; i--)
            {
                for (int j = m - 1; j >= 0; j--)
                {
                    context.PollCancellation();
                    if (WhitespaceAtomsEqual(a[aStart + i], b[bStart + j]))
                        dp[i, j] = dp[i + 1, j + 1] + 1;
                    else
                        dp[i, j] = Math.Max(dp[i + 1, j], dp[i, j + 1]);
                }
            }

            int ai = 0;
            int bi = 0;
            while (ai < n || bi < m)
            {
                context.PollCancellation();
                if (ai < n && bi < m &&
                    WhitespaceAtomsEqual(a[aStart + ai], b[bStart + bi]) &&
                    dp[ai, bi] == dp[ai + 1, bi + 1] + 1)
                {
                    ai++;
                    bi++;
                }
                else if (bi >= m || (ai < n && dp[ai, bi] == dp[ai + 1, bi]))
                {
                    AppendWhitespaceEdit(edits, WhitespaceEditKind.Delete,
                        a[aStart + ai]);
                    ai++;
                }
                else
                {
                    AppendWhitespaceEdit(edits, WhitespaceEditKind.Insert,
                        b[bStart + bi]);
                    bi++;
                }
            }
            return true;
        }

        private static void AppendWhitespaceEdit(ICollection<WhitespaceEdit> edits,
            WhitespaceEditKind kind, PdfReviewWhitespaceAtom atom)
        {
            edits.Add(new WhitespaceEdit { Kind = kind, Atom = atom });
        }

        private sealed class PairProjection
        {
            public PdfReviewPagePair Pair;
            public int LeftPageIndex = -1;
            public int RightPageIndex = -1;
            public PdfReviewPairStatus Status;
        }

        /// <summary>
        /// Пересобрать все производные данные из единственных источников семантики —
        /// result.Operations и result.WhitespaceChanges. Вызывать после первичного diff
        /// и изменения строк viewer; растровый слой публикует свой candidate через
        /// PublishProjection, чтобы семантика и её проекция сменились вместе.
        /// </summary>
        public static void Project(PdfReviewResult result)
        {
            Project(result, new DiffContext(PdfReviewLimits.Default(), null));
        }

        internal static void Project(PdfReviewResult result, Func<bool> cancelled)
        {
            Project(result, new DiffContext(PdfReviewLimits.Default(), cancelled));
        }

        private static void Project(PdfReviewResult result, DiffContext context)
        {
            context.ThrowIfCancellation();
            if (result == null)
                return;
            BuildAndPublishProjection(result, result.Operations,
                result.WhitespaceChanges, context);
        }

        /// <summary>
        /// Атомарно публикует новый authoritative semantic snapshot вместе со всеми
        /// производными индексами, статусами и статистикой. До commit result не меняется.
        /// </summary>
        internal static void PublishProjection(PdfReviewResult result,
            List<PdfReviewWordOp> operations,
            List<PdfReviewWhitespaceChange> whitespaceChanges,
            Func<bool> cancelled)
        {
            var context = new DiffContext(PdfReviewLimits.Default(), cancelled);
            context.ThrowIfCancellation();
            if (result == null)
                return;
            BuildAndPublishProjection(result, operations, whitespaceChanges, context);
        }

        private static void BuildAndPublishProjection(PdfReviewResult result,
            List<PdfReviewWordOp> operations,
            List<PdfReviewWhitespaceChange> whitespaceChanges, DiffContext context)
        {
            if (operations == null) throw new ArgumentNullException("operations");
            if (whitespaceChanges == null)
                throw new ArgumentNullException("whitespaceChanges");

            PdfReviewDocument left = result.Left;
            PdfReviewDocument right = result.Right;
            List<PdfReviewWordOp> previouslyPublishedOperations = result.Operations;
            List<PdfReviewWhitespaceChange> previouslyPublishedWhitespace =
                result.WhitespaceChanges;
            List<PairProjection> pairProjections = CapturePairProjections(result, context);
            var deletedWordsByPage = new Dictionary<int, List<PdfReviewWord>>();
            var insertedWordsByPage = new Dictionary<int, List<PdfReviewWord>>();
            var deletedWhitespaceByPage =
                new Dictionary<int, List<PdfReviewWhitespaceMarker>>();
            var insertedWhitespaceByPage =
                new Dictionary<int, List<PdfReviewWhitespaceMarker>>();

            foreach (PdfReviewWordOp op in operations)
            {
                context.PollCancellation();
                if (op == null || (op.Kind != PdfReviewDiffKind.Delete &&
                    op.Kind != PdfReviewDiffKind.Insert))
                    continue;
                Dictionary<int, List<PdfReviewWord>> index =
                    op.Kind == PdfReviewDiffKind.Delete
                    ? deletedWordsByPage : insertedWordsByPage;
                IList<PdfReviewWord> words = op.Kind == PdfReviewDiffKind.Delete
                    ? (IList<PdfReviewWord>)op.LeftWords : op.RightWords;
                foreach (PdfReviewWord word in words)
                {
                    context.PollCancellation();
                    if (word == null || word.PageIndex < 0)
                        continue;
                    List<PdfReviewWord> pageWords;
                    if (!index.TryGetValue(word.PageIndex, out pageWords))
                    {
                        pageWords = new List<PdfReviewWord>();
                        index.Add(word.PageIndex, pageWords);
                    }
                    pageWords.Add(word);
                }
            }

            var pages = new ReviewPageLookup(result);
            foreach (PdfReviewWhitespaceChange change in whitespaceChanges)
            {
                context.PollCancellation();
                if (change == null)
                    continue;
                AddWhitespaceProjection(deletedWhitespaceByPage, pages,
                    change.Left, change.DeletedAtoms, true, context);
                AddWhitespaceProjection(insertedWhitespaceByPage, pages,
                    change.Right, change.InsertedAtoms, false, context);
            }

            CalculatePairStatuses(pairProjections, pages,
                deletedWordsByPage, insertedWordsByPage,
                deletedWhitespaceByPage, insertedWhitespaceByPage, context);
            PdfReviewStats stats = Statistics(left, right, operations,
                whitespaceChanges, pairProjections, context);

            // Последний callback и вся проверка согласованности предшествуют первой записи.
            // После Validate — только невыделяющие память присваивания, которые не отменяются.
            context.ThrowIfCancellation();
            ValidateProjectionSnapshot(result, left, right,
                previouslyPublishedOperations, previouslyPublishedWhitespace,
                pairProjections);
            result.PublishState(operations, whitespaceChanges,
                deletedWordsByPage, insertedWordsByPage,
                deletedWhitespaceByPage, insertedWhitespaceByPage, stats);
            for (int i = 0; i < pairProjections.Count; i++)
            {
                PairProjection projection = pairProjections[i];
                if (projection.Pair != null)
                    projection.Pair.Status = projection.Status;
            }
        }

        private static List<PairProjection> CapturePairProjections(
            PdfReviewResult result, DiffContext context)
        {
            int count = result.Pairs.Count;
            var projections = new List<PairProjection>(count);
            for (int i = 0; i < count; i++)
            {
                context.PollCancellation();
                PdfReviewPagePair pair = result.Pairs[i];
                projections.Add(new PairProjection
                {
                    Pair = pair,
                    LeftPageIndex = pair == null ? -1 : pair.LeftPageIndex,
                    RightPageIndex = pair == null ? -1 : pair.RightPageIndex,
                    Status = pair == null ? PdfReviewPairStatus.Unchanged : pair.Status
                });
            }
            return projections;
        }

        private static void CalculatePairStatuses(
            IList<PairProjection> pairProjections, ReviewPageLookup pages,
            Dictionary<int, List<PdfReviewWord>> deletedWordsByPage,
            Dictionary<int, List<PdfReviewWord>> insertedWordsByPage,
            Dictionary<int, List<PdfReviewWhitespaceMarker>> deletedWhitespaceByPage,
            Dictionary<int, List<PdfReviewWhitespaceMarker>> insertedWhitespaceByPage,
            DiffContext context)
        {
            foreach (PairProjection projection in pairProjections)
            {
                context.PollCancellation();
                if (projection.Pair == null)
                    continue;
                bool hasLeft = projection.LeftPageIndex >= 0;
                bool hasRight = projection.RightPageIndex >= 0;
                bool deleted = hasLeft &&
                    deletedWordsByPage.ContainsKey(projection.LeftPageIndex);
                bool inserted = hasRight &&
                    insertedWordsByPage.ContainsKey(projection.RightPageIndex);
                bool deletedWhitespace = hasLeft &&
                    deletedWhitespaceByPage.ContainsKey(projection.LeftPageIndex);
                bool insertedWhitespace = hasRight &&
                    insertedWhitespaceByPage.ContainsKey(projection.RightPageIndex);
                if (hasLeft && hasRight)
                    projection.Status = deleted || inserted || deletedWhitespace ||
                        insertedWhitespace
                        ? PdfReviewPairStatus.Changed : PdfReviewPairStatus.Unchanged;
                else if (hasLeft)
                    projection.Status = deleted || deletedWhitespace ||
                        IsEmptyPage(pages.Page(true, projection.LeftPageIndex))
                        ? PdfReviewPairStatus.LeftOnly : PdfReviewPairStatus.Unchanged;
                else if (hasRight)
                    projection.Status = inserted || insertedWhitespace ||
                        IsEmptyPage(pages.Page(false, projection.RightPageIndex))
                        ? PdfReviewPairStatus.RightOnly : PdfReviewPairStatus.Unchanged;
                else
                    projection.Status = PdfReviewPairStatus.Unchanged;
            }
        }

        private static void ValidateProjectionSnapshot(PdfReviewResult result,
            PdfReviewDocument left, PdfReviewDocument right,
            List<PdfReviewWordOp> previouslyPublishedOperations,
            List<PdfReviewWhitespaceChange> previouslyPublishedWhitespace,
            IList<PairProjection> pairProjections)
        {
            if (!ReferenceEquals(result.Left, left) || !ReferenceEquals(result.Right, right) ||
                !ReferenceEquals(result.Operations, previouslyPublishedOperations) ||
                !ReferenceEquals(result.WhitespaceChanges, previouslyPublishedWhitespace) ||
                result.Pairs.Count != pairProjections.Count)
                throw new InvalidOperationException(
                    "PdfReviewResult changed during projection.");
            for (int i = 0; i < pairProjections.Count; i++)
            {
                PairProjection projection = pairProjections[i];
                PdfReviewPagePair pair = result.Pairs[i];
                if (!ReferenceEquals(pair, projection.Pair) || (pair != null &&
                    (pair.LeftPageIndex != projection.LeftPageIndex ||
                     pair.RightPageIndex != projection.RightPageIndex)))
                    throw new InvalidOperationException(
                        "PdfReviewResult page pairs changed during projection.");
            }
        }

        private static void AddWhitespaceProjection(
            Dictionary<int, List<PdfReviewWhitespaceMarker>> index,
            ReviewPageLookup pages, PdfReviewWhitespaceEvidence evidence,
            IList<PdfReviewWhitespaceAtom> atoms, bool deleted, DiffContext context)
        {
            // Удалённое принадлежит ранней (левой) стороне, добавленное — поздней.
            if (index == null || evidence == null || evidence.PageIndex < 0 ||
                pages.Page(deleted, evidence.PageIndex) == null || atoms == null ||
                atoms.Count == 0)
                return;
            string summary = WhitespaceTokenSummary(atoms, context);
            if (summary.Length == 0)
                return;
            var marker = new PdfReviewWhitespaceMarker
            {
                Box = evidence.MarkerBox,
                Text = (deleted ? "− " : "+ ") + summary,
                AccessibleDescription = string.Format(CultureInfo.CurrentCulture,
                    Loc.T(deleted ? "review.whitespace.removed" :
                        "review.whitespace.added"), summary),
                Style = deleted ? PdfReviewHighlightStyle.Removed :
                    PdfReviewHighlightStyle.Added
            };
            List<PdfReviewWhitespaceMarker> markers;
            if (!index.TryGetValue(evidence.PageIndex, out markers))
            {
                markers = new List<PdfReviewWhitespaceMarker>();
                index.Add(evidence.PageIndex, markers);
            }
            markers.Add(marker);
        }

        private static string WhitespaceTokenSummary(
            IList<PdfReviewWhitespaceAtom> atoms, DiffContext context)
        {
            const int MaxGroups = 12;
            var result = new StringBuilder();
            int groups = 0;
            int i = 0;
            while (i < atoms.Count && groups < MaxGroups)
            {
                context.PollCancellation();
                PdfReviewWhitespaceAtom atom = atoms[i];
                if (atom == null)
                {
                    i++;
                    continue;
                }
                int count = 1;
                while (i + count < atoms.Count &&
                    WhitespaceAtomsEqual(atom, atoms[i + count]))
                {
                    context.PollCancellation();
                    count++;
                }
                if (result.Length > 0)
                    result.Append(' ');
                result.Append(WhitespaceToken(atom));
                if (count > 1)
                    result.Append('×').Append(count.ToString(CultureInfo.InvariantCulture));
                i += count;
                groups++;
            }
            if (i < atoms.Count)
            {
                if (result.Length > 0)
                    result.Append(' ');
                result.Append("…+").Append((atoms.Count - i).ToString(
                    CultureInfo.InvariantCulture));
            }
            return result.ToString();
        }

        private static string WhitespaceToken(PdfReviewWhitespaceAtom atom)
        {
            if (atom == null)
                return "";
            if (atom.Kind == PdfReviewWhitespaceAtomKind.Space)
                return "␠";
            if (atom.Kind == PdfReviewWhitespaceAtomKind.NoBreakSpace)
                return "NBSP";
            if (atom.Kind == PdfReviewWhitespaceAtomKind.Tab)
                return "⇥";
            if (atom.Kind == PdfReviewWhitespaceAtomKind.LineBreak)
                return "↵";
            if (string.IsNullOrEmpty(atom.RawText))
                return "U+????";
            int codePoint = char.ConvertToUtf32(atom.RawText, 0);
            return "U+" + codePoint.ToString(codePoint <= 0xFFFF ? "X4" : "X6",
                CultureInfo.InvariantCulture);
        }

        /// <summary>Счётчики строятся из глобальных операций и спроецированных строк.</summary>
        public static PdfReviewStats Statistics(PdfReviewResult result)
        {
            return Statistics(result,
                new DiffContext(PdfReviewLimits.Default(), null));
        }

        private static PdfReviewStats Statistics(PdfReviewResult result,
            DiffContext context)
        {
            context.ThrowIfCancellation();
            if (result == null)
                return new PdfReviewStats();
            List<PairProjection> pairProjections =
                CapturePairProjections(result, context);
            return Statistics(result.Left, result.Right, result.Operations,
                result.WhitespaceChanges, pairProjections, context);
        }

        private static PdfReviewStats Statistics(PdfReviewDocument left,
            PdfReviewDocument right, IList<PdfReviewWordOp> operations,
            IList<PdfReviewWhitespaceChange> whitespaceChanges,
            IList<PairProjection> pairProjections, DiffContext context)
        {
            var stats = new PdfReviewStats { PagePairs = pairProjections.Count };
            foreach (PairProjection projection in pairProjections)
            {
                context.PollCancellation();
                if (projection.Pair == null)
                    continue;
                if (projection.Status == PdfReviewPairStatus.Changed)
                    stats.ChangedPages++;
                else if (projection.Status == PdfReviewPairStatus.LeftOnly)
                    stats.LeftOnlyPages++;
                else if (projection.Status == PdfReviewPairStatus.RightOnly)
                    stats.RightOnlyPages++;
            }

            int hunkDeletes = 0;
            int hunkInserts = 0;
            foreach (PdfReviewWordOp op in operations)
            {
                context.PollCancellation();
                if (op == null)
                    continue;
                if (op.Kind == PdfReviewDiffKind.Equal)
                {
                    stats.Replacements += Math.Min(hunkDeletes, hunkInserts);
                    hunkDeletes = hunkInserts = 0;
                }
                else if (op.Kind == PdfReviewDiffKind.Delete)
                {
                    stats.DeletedWords += op.LeftWords.Count;
                    hunkDeletes += op.LeftWords.Count;
                }
                else
                {
                    stats.InsertedWords += op.RightWords.Count;
                    hunkInserts += op.RightWords.Count;
                }
            }
            stats.Replacements += Math.Min(hunkDeletes, hunkInserts);

            foreach (PdfReviewWhitespaceChange change in whitespaceChanges)
            {
                context.PollCancellation();
                if (change == null)
                    continue;
                int deletedWhitespace = CountWhitespaceAtoms(change.DeletedAtoms, context);
                int insertedWhitespace = CountWhitespaceAtoms(change.InsertedAtoms, context);
                if (deletedWhitespace == 0 && insertedWhitespace == 0)
                    continue;
                stats.WhitespaceChanges++;
                stats.DeletedWhitespaceAtoms += deletedWhitespace;
                stats.InsertedWhitespaceAtoms += insertedWhitespace;
            }

            // Процент остаётся метрикой слов: пробельная правка видна отдельными
            // счётчиками и статусом страницы, но не притворяется изменённым словом.
            long total = (long)DocumentWordCount(left, context) +
                DocumentWordCount(right, context);
            long changed = (long)stats.DeletedWords + stats.InsertedWords;
            stats.ChangedPercent = total <= 0 ? 0 :
                Math.Min(100, Math.Max(changed > 0 ? 1 : 0,
                    (int)Math.Round(100.0 * changed / total)));
            return stats;
        }

        private static int CountWhitespaceAtoms(IList<PdfReviewWhitespaceAtom> atoms,
            DiffContext context)
        {
            if (atoms == null)
                return 0;
            int count = 0;
            foreach (PdfReviewWhitespaceAtom atom in atoms)
            {
                context.PollCancellation();
                if (atom != null)
                    count++;
            }
            return count;
        }

        private static int DocumentWordCount(PdfReviewDocument document,
            DiffContext context)
        {
            if (document == null)
                return 0;
            int count = 0;
            foreach (PdfReviewPage page in document.Pages)
            {
                context.PollCancellation();
                if (page != null)
                    count += page.Words.Count;
            }
            return count;
        }

        private static bool IsEmptyPage(PdfReviewPage page)
        {
            return page != null && page.Words.Count == 0;
        }

        /// <summary>
        /// Похожесть двух страниц: 0.9 — пересечение слов, 0.1 — совпадение размеров.
        /// Слова берутся из готового списка страницы; пустой список (страницы, собранные
        /// вне сервиса) токенизируется из нормализованного текста.
        /// </summary>
        private static double Similarity(PdfReviewPage a, PdfReviewPage b)
        {
            if (a == null || b == null) return 0;
            return Similarity(a, b, WordSet(a), WordSet(b), WordsEqual(a.Words, b.Words));
        }

        private static double Similarity(PdfReviewPage a, PdfReviewPage b,
            HashSet<string> aw, HashSet<string> bw, bool exactWords)
        {
            if (a == null || b == null) return 0;
            if (exactWords) return 1;
            int common = 0;
            HashSet<string> smaller = aw.Count <= bw.Count ? aw : bw;
            HashSet<string> larger = aw.Count <= bw.Count ? bw : aw;
            foreach (string word in smaller)
                if (larger.Contains(word)) common++;
            return Similarity(a, b, aw.Count, bw.Count, common);
        }

        /// <summary>
        /// Сходство по числу общих слов. При неизменных размерах множеств монотонно
        /// неубывает по common — на этом держится отсев пар в TryAlignmentSimilarities.
        /// </summary>
        private static double Similarity(PdfReviewPage a, PdfReviewPage b,
            int aCount, int bCount, int common)
        {
            int union = aCount + bCount - common;
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
