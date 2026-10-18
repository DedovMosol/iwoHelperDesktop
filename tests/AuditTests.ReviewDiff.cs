using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Аудит F1: ядро сравнения PDF (PdfReviewDiff). Каждый тест воспроизводит находку
    /// аудита на синтетических моделях и проверяет наблюдаемый результат Compare/стадий.
    /// </summary>
    internal static partial class UnitTests
    {
        private static List<string> AuditF1Tokens(string prefix, int count)
        {
            var tokens = new List<string>(count);
            for (int i = 0; i < count; i++)
                tokens.Add(prefix + i.ToString("D6", CultureInfo.InvariantCulture));
            return tokens;
        }

        /// <summary>Документ, свёрстанный по pageWords слов на страницу.</summary>
        private static PdfReviewDocument AuditF1Document(string path, IList<string> tokens,
            int pageWords)
        {
            var pages = new List<string>();
            for (int start = 0; start < tokens.Count; start += pageWords)
            {
                int count = Math.Min(pageWords, tokens.Count - start);
                var page = new StringBuilder();
                for (int i = 0; i < count; i++)
                {
                    if (i > 0) page.Append(' ');
                    page.Append(tokens[start + i]);
                }
                pages.Add(page.ToString());
            }
            return ReviewDocument(path, pages.ToArray());
        }

        private static void AuditF1AssertExactOperations(PdfReviewResult result,
            int deleted, int inserted, string context)
        {
            AssertEq(deleted, result.Stats.DeletedWords, context + ": удалённых слов");
            AssertEq(inserted, result.Stats.InsertedWords, context + ": добавленных слов");
            AssertEq(ReviewDocumentKeys(result.Left),
                RebuildReviewSide(result.Operations, true),
                context + ": левая сторона восстанавливается из операций");
            AssertEq(ReviewDocumentKeys(result.Right),
                RebuildReviewSide(result.Operations, false),
                context + ": правая сторона восстанавливается из операций");
        }

        /// <summary>
        /// Находка 1: CanTrim не давал срезать совпадения через границу пары страниц, а Myers
        /// для вырожденного D=0/D=1 возвращал конец диапазона, и весь хвост уходил в сырые
        /// Delete/Insert (≈91% документа «изменено»). Правок должно быть ровно столько,
        /// сколько вставлено, при любом MaxDiffCells и в обе стороны.
        /// </summary>
        [Test("Аудит F1: вставка на второй странице при перевёрстке даёт ровно вставленные слова в обе стороны")]
        private static void AuditF1RepaginatedInsertionIsExact()
        {
            const int PageWords = 350;
            const int Inserted = 5;
            List<string> baseline = AuditF1Tokens("w", 12 * PageWords);
            var revised = new List<string>(baseline);
            var added = new List<string>();
            for (int i = 0; i < Inserted; i++)
            {
                string token = "added" + i.ToString(CultureInfo.InvariantCulture);
                revised.Insert(PageWords + 100 + i, token);
                added.Add(token);
            }
            PdfReviewDocument left = AuditF1Document("baseline.pdf", baseline, PageWords);
            PdfReviewDocument right = AuditF1Document("revised.pdf", revised, PageWords);
            AssertTrue(right.Pages.Count > left.Pages.Count,
                "fixture действительно сдвигает границы всех последующих страниц");
            string expected = string.Join("|", added.ToArray());

            foreach (int maxCells in new[] { PdfReviewLimits.AbsoluteMaxDiffCells, 4096 })
            {
                var limits = PdfReviewLimits.Default();
                limits.MaxDiffCells = maxCells;
                string context = "MaxDiffCells=" + maxCells.ToString(CultureInfo.InvariantCulture);

                PdfReviewResult forward = PdfReviewDiff.Compare(left, right, limits);
                AuditF1AssertExactOperations(forward, 0, Inserted, context + ", прямо");
                AssertEq(expected, ReviewChangedWordsAcrossPairs(forward,
                    PdfReviewDiffKind.Insert), context + ", прямо: вставлены именно новые слова");
                AssertEq(false, forward.WorkExhausted, context + ", прямо: бюджет не исчерпан");

                PdfReviewResult reverse = PdfReviewDiff.Compare(right, left, limits);
                AuditF1AssertExactOperations(reverse, Inserted, 0, context + ", обратно");
                AssertEq(expected, ReviewChangedWordsAcrossPairs(reverse,
                    PdfReviewDiffKind.Delete), context + ", обратно: удалены именно новые слова");
                AssertEq(false, reverse.WorkExhausted, context + ", обратно: бюджет не исчерпан");
            }
        }

        /// <summary>
        /// Находка 2: выравнивание страниц списывало n·m·(слов на страницу) с общего бюджета;
        /// на 340 страницах по 200 разных слов бюджет кончался до сравнения слов, и две копии
        /// одного документа выходили «изменены на 100%». С общими служебными словами на каждой
        /// странице (вторая итерация) кандидатов тоже не должно становиться n·m.
        /// </summary>
        [Test("Аудит F1: две копии сотен страниц с разными словами — без изменений и без исчерпания")]
        private static void AuditF1LargeIdenticalDocumentsStayUnchanged()
        {
            const int Pages = 340;
            const int PageWords = 200;
            foreach (int shared in new[] { 0, 40 })
            {
                var tokens = new List<string>(Pages * PageWords);
                for (int page = 0; page < Pages; page++)
                    for (int word = 0; word < PageWords; word++)
                        tokens.Add(word < shared
                            ? "common" + word.ToString(CultureInfo.InvariantCulture)
                            : "p" + page.ToString(CultureInfo.InvariantCulture) + "w" +
                                word.ToString(CultureInfo.InvariantCulture));
                PdfReviewDocument left = AuditF1Document("copy-left.pdf", tokens, PageWords);
                PdfReviewDocument right = AuditF1Document("copy-right.pdf", tokens, PageWords);
                string context = "общих слов на странице " +
                    shared.ToString(CultureInfo.InvariantCulture);

                Stopwatch watch = Stopwatch.StartNew();
                PdfReviewResult result = PdfReviewDiff.Compare(left, right,
                    PdfReviewLimits.Default());
                watch.Stop();

                AuditF1AssertExactOperations(result, 0, 0, context);
                AssertEq(0, result.Stats.ChangedPercent, context + ": процент изменений");
                AssertEq(false, result.WorkExhausted, context + ": бюджет не исчерпан");
                AssertEq(Pages, result.Pairs.Count, context + ": строк viewer");
                for (int i = 0; i < result.Pairs.Count; i++)
                {
                    PdfReviewPagePair pair = result.Pairs[i];
                    AssertTrue(pair.LeftPageIndex == i && pair.RightPageIndex == i &&
                               pair.Status == PdfReviewPairStatus.Unchanged,
                        context + ": строка " + i + " = " + pair.LeftPageIndex + ":" +
                        pair.RightPageIndex + ":" + pair.Status + ", ожидалась неизменная пара");
                }
                AssertTrue(watch.Elapsed < TimeSpan.FromSeconds(30),
                    context + ": сравнение заняло " + watch.Elapsed + " (порог 30 с)");
            }
        }

        /// <summary>
        /// Находка 3: листья матрицы оплачивались n·m независимо от числа правок, и после
        /// исчерпания даже линейное точное срезание краёв пропускалось — 500 реальных замен
        /// в длинном документе давали 62-84% «изменено». Ожидается ровно 500/500.
        /// </summary>
        [Test("Аудит F1: 500 замен в документах по 40 тыс. слов находятся точно и в пределах бюджета")]
        private static void AuditF1ManyReplacementsInLargeDocumentAreExact()
        {
            const int Words = 40000;
            const int PageWords = 400;
            const int Replacements = 500;
            List<string> baseline = AuditF1Tokens("w", Words);
            var revised = new List<string>(baseline);
            var random = new Random(20261016);
            var positions = new HashSet<int>();
            while (positions.Count < Replacements)
                positions.Add(random.Next(Words));
            foreach (int position in positions)
                revised[position] = "r" + position.ToString("D6", CultureInfo.InvariantCulture);
            PdfReviewDocument left = AuditF1Document("many-left.pdf", baseline, PageWords);
            PdfReviewDocument right = AuditF1Document("many-right.pdf", revised, PageWords);

            Stopwatch watch = Stopwatch.StartNew();
            PdfReviewResult result = PdfReviewDiff.Compare(left, right,
                PdfReviewLimits.Default());
            watch.Stop();

            AuditF1AssertExactOperations(result, Replacements, Replacements, "500 замен");
            AssertEq(Replacements, result.Stats.Replacements, "500 замен: число замен");
            AssertEq(false, result.WorkExhausted, "500 замен: бюджет не исчерпан");
            AssertTrue(watch.Elapsed < TimeSpan.FromSeconds(60),
                "500 замен: сравнение заняло " + watch.Elapsed + " (порог 60 с)");
        }

        /// <summary>
        /// Находка 3 (вторая часть): когда бюджета не хватает на доказательство середины,
        /// общие края всё равно срезаются линейным точным сравнением ключей — сырыми
        /// остаются только непроверенные слова, а не весь диапазон.
        /// </summary>
        [Test("Аудит F1: после исчерпания бюджета точные общие края остаются Equal")]
        private static void AuditF1ExactEndsSurviveExhaustion()
        {
            List<string> first = AuditF1Tokens("a", 100);
            List<string> second = AuditF1Tokens("b", 100);
            var baseline = new List<string>(first);
            baseline.AddRange(second);
            var revised = new List<string> { "inserted" };
            revised.AddRange(first);
            revised.AddRange(second);
            // Слева страница 0 = a*, справа a100 перешло на страницу 1: восстановленный
            // по страницам префикс обрывается на межстраничной паре.
            PdfReviewDocument left = AuditF1Document("ends-left.pdf", baseline, 100);
            PdfReviewDocument right = AuditF1Document("ends-right.pdf", revised, 100);
            var limits = PdfReviewLimits.Default();
            limits.MaxDiffWork = 8000;

            PdfReviewResult forward = PdfReviewDiff.Compare(left, right, limits);
            AssertEq(true, forward.WorkExhausted,
                "fixture действительно исчерпывает бюджет на середине");
            AuditF1AssertExactOperations(forward, 0, 1, "исчерпание, прямо");

            PdfReviewResult reverse = PdfReviewDiff.Compare(right, left, limits);
            AssertEq(true, reverse.WorkExhausted, "обратно: бюджет исчерпан");
            AuditF1AssertExactOperations(reverse, 1, 0, "исчерпание, обратно");
        }

        [Test("Аудит F1: бюджет работы растёт с длиной документа до потолка без переполнения")]
        private static void AuditF1DiffWorkScalesWithDocumentLength()
        {
            var limits = PdfReviewLimits.Default();
            long baseWork = limits.MaxDiffWork;
            long reference = PdfReviewLimits.DiffWorkReferenceWords;
            AssertEq(baseWork, limits.DiffWorkFor(0), "пустой документ получает базовый бюджет");
            AssertEq(baseWork, limits.DiffWorkFor(-5), "отрицательная длина не уменьшает бюджет");
            AssertEq(baseWork, limits.DiffWorkFor(reference),
                "документ эталонной длины получает базовый бюджет");
            AssertEq(baseWork * 2, limits.DiffWorkFor(reference * 2),
                "вдвое длиннее эталона — вдвое больше бюджета");
            AssertEq(baseWork * 3 / 2, limits.DiffWorkFor(reference * 3 / 2),
                "масштаб линеен и между кратными эталона");
            AssertEq(baseWork * PdfReviewLimits.MaxDiffWorkScale,
                limits.DiffWorkFor(long.MaxValue), "масштаб ограничен потолком");

            limits.MaxDiffWork = 0;
            AssertEq(0L, limits.DiffWorkFor(reference * 4), "нулевой бюджет не масштабируется");
            limits.MaxDiffWork = -7;
            AssertEq(0L, limits.DiffWorkFor(reference * 4), "отрицательный бюджет равен нулю");
            limits.MaxDiffWork = long.MaxValue;
            AssertEq(long.MaxValue, limits.DiffWorkFor(reference * 4),
                "огромный бюджет насыщается, а не переполняется");
        }

        /// <summary>
        /// Находка 4: при исчерпании стадии молча пропускались, и результат выглядел полным.
        /// Теперь результат помечен и называет недоделанные стадии ещё до проекции.
        /// </summary>
        [Test("Аудит F1: исчерпание бюджета помечает результат и называет недоделанные стадии")]
        private static void AuditF1ExhaustionIsReported()
        {
            PdfReviewResult full = PdfReviewDiff.Compare(
                ReviewDocument("flag-left.pdf", "a b c"),
                ReviewDocument("flag-right.pdf", "x b y"), PdfReviewLimits.Default());
            AssertEq(false, full.WorkExhausted, "полное сравнение не помечено");
            AssertEq(PdfReviewStage.None, full.IncompleteStages,
                "полное сравнение не называет недоделанных стадий");

            var limits = PdfReviewLimits.Default();
            limits.MaxDiffWork = 0;
            PdfReviewResult exhausted = PdfReviewDiff.Compare(
                ReviewDocument("flag-left.pdf", "a b c"),
                ReviewDocument("flag-right.pdf", "x b y"), limits);
            AssertEq(true, exhausted.WorkExhausted, "исчерпанное сравнение помечено");
            PdfReviewStage all = PdfReviewStage.PageAlignment | PdfReviewStage.WordDiff |
                PdfReviewStage.OrderReconciliation | PdfReviewStage.SplitJoin |
                PdfReviewStage.Whitespace;
            AssertEq(all, exhausted.IncompleteStages,
                "при нулевом бюджете ни одна стадия не доказана");
            AssertTrue(exhausted.Stats != null && exhausted.Stats.DeletedWords == 3 &&
                       exhausted.Stats.InsertedWords == 3,
                "непроверенный текст остаётся консервативными Delete/Insert");
        }

        /// <summary>
        /// Страница с доверенными source-пробелами между всеми словами, построенная напрямую
        /// (без PdfPig): так тесты бюджета управляют числом границ и длиной ключей.
        /// </summary>
        private static PdfReviewPage AuditF1WhitespacePage(int pageIndex, IList<string> words,
            IList<string> boundaries)
        {
            var page = new PdfReviewPage
            {
                PageIndex = pageIndex,
                WidthPt = 600,
                HeightPt = 800,
                ViewWidthPt = 600,
                ViewHeightPt = 800
            };
            var text = new StringBuilder();
            int source = 0;
            for (int i = 0; i < words.Count; i++)
            {
                if (i > 0)
                {
                    text.Append(' ');
                    source += 2;
                }
                page.Words.Add(new PdfReviewWord
                {
                    Text = words[i],
                    Key = words[i],
                    PageIndex = pageIndex,
                    SourceStart = source,
                    SourceEnd = source + words[i].Length - 1,
                    SourceText = words[i],
                    SourceTrusted = true,
                    Box = new PdfReviewBox
                    {
                        Left = 1 + i * 0.5,
                        Right = 1.4 + i * 0.5,
                        Bottom = 400,
                        Top = 410
                    }
                });
                text.Append(words[i]);
                source += words[i].Length;
            }
            if (boundaries != null)
                for (int i = 0; i + 1 < page.Words.Count; i++)
                    page.WhitespaceBoundaries.Add(new PdfReviewWhitespaceEvidence
                    {
                        PageIndex = pageIndex,
                        Before = page.Words[i],
                        After = page.Words[i + 1],
                        RawText = boundaries[i],
                        LogicalText = boundaries[i]
                    });
            page.Text = text.ToString();
            page.NormalizedText = PdfReviewDiff.Normalize(page.Text);
            page.Fingerprint = PdfReviewDiff.Fingerprint(page.NormalizedText,
                page.WidthPt, page.HeightPt);
            return page;
        }

        private static PdfReviewDocument AuditF1PagesDocument(string path,
            params PdfReviewPage[] pages)
        {
            var document = new PdfReviewDocument { Path = path };
            foreach (PdfReviewPage page in pages)
            {
                document.Pages.Add(page);
                document.WordCount += page.Words.Count;
                document.CharacterCount += page.NormalizedText.Length;
            }
            return document;
        }

        private static List<string> AuditF1Spaces(int count)
        {
            var spaces = new List<string>(count);
            for (int i = 0; i < count; i++)
                spaces.Add(" ");
            return spaces;
        }

        /// <summary>
        /// Находка 4 (вторая часть): сравнение пробелов получает гарантированную долю бюджета.
        /// Здесь словесный diff исчерпывает свою долю на чужой странице, но пробельная правка
        /// между доказанными Exact-словами первой страницы всё равно найдена.
        /// </summary>
        [Test("Аудит F1: сравнение пробелов сохраняет свою долю бюджета после исчерпания diff")]
        private static void AuditF1WhitespaceKeepsGuaranteedShare()
        {
            List<string> words = AuditF1Tokens("s", 20);
            List<string> leftSpaces = AuditF1Spaces(words.Count - 1);
            List<string> rightSpaces = AuditF1Spaces(words.Count - 1);
            rightSpaces[7] = "  ";
            PdfReviewDocument left = AuditF1PagesDocument("share-left.pdf",
                AuditF1WhitespacePage(0, words, leftSpaces),
                AuditF1WhitespacePage(1, AuditF1Tokens("x", 250), null));
            PdfReviewDocument right = AuditF1PagesDocument("share-right.pdf",
                AuditF1WhitespacePage(0, words, rightSpaces),
                AuditF1WhitespacePage(1, AuditF1Tokens("y", 250), null));
            var limits = PdfReviewLimits.Default();
            limits.MaxDiffWork = 40000;

            PdfReviewResult result = PdfReviewDiff.Compare(left, right, limits);
            AssertEq(true, result.WorkExhausted, "словесная доля бюджета исчерпана");
            AssertTrue((result.IncompleteStages & PdfReviewStage.WordDiff) != 0,
                "недоделанный word diff назван: " + result.IncompleteStages);
            AssertEq(PdfReviewStage.None, result.IncompleteStages & PdfReviewStage.Whitespace,
                "сравнение пробелов завершилось своей долей");
            AssertEq(1, result.WhitespaceChanges.Count,
                "двойной пробел между Exact-словами найден");
            AssertEq(1, result.WhitespaceChanges[0].InsertedAtoms.Count,
                "добавлен ровно один пробел");
            AssertEq(1, result.Stats.WhitespaceChanges, "статистика видит пробельную правку");
        }

        /// <summary>
        /// Находка 5: ExplicitPagePair перебирал все строки viewer на каждый запрос во
        /// внутренних циклах. Посторонние строки не должны менять стоимость стадии.
        /// </summary>
        [Test("Аудит F1: стоимость стадий сверки не зависит от числа посторонних строк viewer")]
        private static void AuditF1StageWorkIndependentOfUnrelatedPairs()
        {
            const int Extra = 2000;
            long reconcile = FindMinimumReviewWork(delegate(long work)
            {
                return AuditF1ReconcileSucceeds(0, work);
            }, "reconcile");
            long reconcilePadded = FindMinimumReviewWork(delegate(long work)
            {
                return AuditF1ReconcileSucceeds(Extra, work);
            }, "reconcile + строки");
            AssertEq(reconcile, reconcilePadded,
                "минимальный бюджет reconciliation с " + Extra + " посторонними строками");

            long splitJoin = FindMinimumReviewWork(delegate(long work)
            {
                return AuditF1SplitJoinSucceeds(AuditF1PadPairs(ReviewSplitJoinStageFixture(), 0),
                    work);
            }, "split/join");
            long splitJoinPadded = FindMinimumReviewWork(delegate(long work)
            {
                return AuditF1SplitJoinSucceeds(AuditF1PadPairs(ReviewSplitJoinStageFixture(),
                    Extra), work);
            }, "split/join + строки");
            AssertEq(splitJoin, splitJoinPadded,
                "минимальный бюджет split/join с " + Extra + " посторонними строками");
        }

        private static PdfReviewResult AuditF1PadPairs(PdfReviewResult result, int extra)
        {
            for (int i = 0; i < extra; i++)
                result.Pairs.Add(new PdfReviewPagePair
                {
                    LeftPageIndex = 1000 + i,
                    RightPageIndex = 1000 + i
                });
            return result;
        }

        private static bool AuditF1ReconcileSucceeds(int extra, long work)
        {
            PdfReviewResult result = AuditF1PadPairs(ReviewOrderPermutationStageFixture(), extra);
            var limits = PdfReviewLimits.Default();
            limits.MaxDiffWork = work;
            PdfReviewDiff.ReconcileOrderArtifacts(result, limits, null);
            return ReviewMatchCount(result, PdfReviewMatchKind.ReconciledOrder) == 1;
        }

        private static bool AuditF1SplitJoinSucceeds(PdfReviewResult result, long work)
        {
            var limits = PdfReviewLimits.Default();
            limits.MaxDiffWork = work;
            PdfReviewDiff.ClassifySplitJoin(result, limits, null);
            return ReviewSplitJoinOperation(result) != null;
        }

        private static PdfReviewWord AuditF1AnchorWord(string key)
        {
            return new PdfReviewWord
            {
                Text = key,
                Key = key,
                PageIndex = 0,
                Box = new PdfReviewBox { Left = 1, Right = 20, Bottom = 190, Top = 198 }
            };
        }

        /// <summary>
        /// Окно split/join между двумя Exact-якорями: [якорь][K удалений][K вставок] перед
        /// настоящим 2↔1 окном fixture или, для сравнения, те же слова до первого якоря.
        /// </summary>
        private static PdfReviewResult AuditF1SplitJoinWithWideWindow(bool between, int size)
        {
            PdfReviewResult result = ReviewSplitJoinStageFixture();
            var original = new List<PdfReviewWordOp>(result.Operations);
            var prefix = new List<PdfReviewWordOp>();
            var anchor = new List<PdfReviewWordOp>();
            PdfReviewDiff.AppendEqual(anchor, AuditF1AnchorWord("outer-anchor"),
                AuditF1AnchorWord("outer-anchor"), PdfReviewMatchKind.Exact);
            if (between)
                prefix.AddRange(anchor);
            for (int i = 0; i < size; i++)
                PdfReviewDiff.Append(prefix, PdfReviewDiffKind.Delete,
                    AuditF1AnchorWord("gone-" + i.ToString(CultureInfo.InvariantCulture)));
            for (int i = 0; i < size; i++)
                PdfReviewDiff.Append(prefix, PdfReviewDiffKind.Insert,
                    AuditF1AnchorWord("new-" + i.ToString(CultureInfo.InvariantCulture)));
            if (!between)
                prefix.AddRange(anchor);
            prefix.AddRange(original);
            result.ReplaceOperations(prefix);
            return result;
        }

        /// <summary>
        /// Находка 7 (вторая часть): окно, где с любой стороны больше двух слов, не может
        /// быть 1↔2/2↔1, и его проверка не должна стоить пропорционально его ширине.
        /// </summary>
        [Test("Аудит F1: широкое окно между якорями отвергается split/join без обхода")]
        private static void AuditF1WideSplitJoinWindowExitsEarly()
        {
            const int Size = 200;
            long outside = FindMinimumReviewWork(delegate(long work)
            {
                return AuditF1SplitJoinSucceeds(AuditF1SplitJoinWithWideWindow(false, Size), work);
            }, "окно вне якорей");
            long between = FindMinimumReviewWork(delegate(long work)
            {
                return AuditF1SplitJoinSucceeds(AuditF1SplitJoinWithWideWindow(true, Size), work);
            }, "окно между якорями");
            AssertTrue(between - outside <= 16,
                "широкое окно между якорями стоит " + (between - outside) +
                " единиц сверх того же окна вне якорей (ожидалось не больше 16)");
        }

        private static PdfReviewDocument AuditF1LongWordDocument(string path, bool boundaries)
        {
            var words = new List<string>();
            for (int i = 0; i < 600; i++)
                words.Add("long" + i.ToString("D4", CultureInfo.InvariantCulture) +
                    new string('x', 196));
            return AuditF1PagesDocument(path, AuditF1WhitespacePage(0, words,
                boundaries ? AuditF1Spaces(words.Count - 1) : null));
        }

        private static bool AuditF1CompareCompletes(bool boundaries, long work)
        {
            var limits = PdfReviewLimits.Default();
            limits.MaxDiffWork = work;
            PdfReviewResult result = PdfReviewDiff.Compare(
                AuditF1LongWordDocument("index-left.pdf", boundaries),
                AuditF1LongWordDocument("index-right.pdf", boundaries), limits);
            return !result.WorkExhausted;
        }

        private static bool AuditF1WhitespaceCompletes(bool boundaries, long work)
        {
            PdfReviewResult result = PdfReviewDiff.Compare(
                AuditF1LongWordDocument("index-left.pdf", boundaries),
                AuditF1LongWordDocument("index-right.pdf", boundaries),
                PdfReviewLimits.Default());
            // Прежняя публикация заменяется пустым списком только полностью доказанной стадией.
            result.ReplaceWhitespaceChanges(new List<PdfReviewWhitespaceChange>
            {
                new PdfReviewWhitespaceChange()
            });
            var limits = PdfReviewLimits.Default();
            limits.MaxDiffWork = work;
            PdfReviewDiff.CompareWhitespace(result, limits);
            return result.WhitespaceChanges.Count == 0;
        }

        /// <summary>
        /// Находка 7: индекс пробельных свидетельств строился и оплачивался дважды — в
        /// split/join и в сравнении пробелов. Добавочная стоимость границ в полном Compare
        /// должна совпадать с добавочной стоимостью одной стадии сравнения пробелов.
        /// </summary>
        [Test("Аудит F1: Compare оплачивает индекс пробельных свидетельств один раз")]
        private static void AuditF1WhitespaceIndexBuiltOnce()
        {
            long compareWith = FindMinimumReviewWork(delegate(long work)
            {
                return AuditF1CompareCompletes(true, work);
            }, "Compare с границами");
            long compareWithout = FindMinimumReviewWork(delegate(long work)
            {
                return AuditF1CompareCompletes(false, work);
            }, "Compare без границ");
            long stageWith = FindMinimumReviewWork(delegate(long work)
            {
                return AuditF1WhitespaceCompletes(true, work);
            }, "пробелы с границами");
            long stageWithout = FindMinimumReviewWork(delegate(long work)
            {
                return AuditF1WhitespaceCompletes(false, work);
            }, "пробелы без границ");
            AssertTrue(stageWith - stageWithout > 100000,
                "fixture делает индекс границ заметной частью работы: " +
                (stageWith - stageWithout));
            AssertEq(stageWith - stageWithout, compareWith - compareWithout,
                "добавочная стоимость границ в Compare (индекс строится один раз)");
        }

        /// <summary>
        /// Находка 8: Compare переписывал PageIndex слов разделяемых кэшированных документов,
        /// даже когда значение не менялось, а откат затирал их снова. Слово с уже верным
        /// владельцем Compare не трогает ни при применении, ни при откате.
        /// </summary>
        [Test("Аудит F1: Compare не пишет в слова с уже верным владельцем и при откате")]
        private static void AuditF1OwnershipSkipsUnchangedWords()
        {
            PdfReviewDocument left = ReviewDocument("owner-left.pdf", "alpha beta gamma");
            PdfReviewDocument right = ReviewDocument("owner-right.pdf", "alpha delta gamma");
            PdfReviewWord stale = left.Pages[0].Words[0];
            PdfReviewWord settled = right.Pages[0].Words[2];
            stale.PageIndex = 42;
            const int Sentinel = -777;
            bool thrown = false;
            try
            {
                PdfReviewDiff.Compare(left, right, PdfReviewLimits.Default(), delegate
                {
                    // Владелец применён: другой читатель того же кэша меняет слово, которое
                    // Compare переписывать не должен.
                    if (stale.PageIndex == 0 && settled.PageIndex != Sentinel)
                        settled.PageIndex = Sentinel;
                    return settled.PageIndex == Sentinel;
                });
            }
            catch (OperationCanceledException)
            {
                thrown = true;
            }
            AssertTrue(thrown, "отмена после применения владельцев сработала");
            AssertEq(42, stale.PageIndex, "изменённый владелец откатан к исходному значению");
            AssertEq(Sentinel, settled.PageIndex,
                "слово с верным владельцем не затёрто откатом");
        }
    }
}

namespace ExcelMerger.Tests
{
    /// <summary>Аудит F1: попытки сломать исправления — свойства на случайных входах.</summary>
    internal static partial class UnitTests
    {
        private static int AuditF1LcsLength(IList<string> left, IList<string> right)
        {
            var lengths = new int[left.Count + 1, right.Count + 1];
            for (int i = left.Count - 1; i >= 0; i--)
                for (int j = right.Count - 1; j >= 0; j--)
                    lengths[i, j] = left[i] == right[j]
                        ? lengths[i + 1, j + 1] + 1
                        : Math.Max(lengths[i + 1, j], lengths[i, j + 1]);
            return lengths[0, 0];
        }

        /// <summary>Случайная перевёрстка: те же токены, другие границы страниц и пустые страницы.</summary>
        private static PdfReviewDocument AuditF1RandomPages(Random random, string path,
            IList<string> tokens)
        {
            var pages = new List<string>();
            int index = 0;
            while (index < tokens.Count || pages.Count == 0)
            {
                int size = random.Next(5) == 0 ? 0 : random.Next(1, 9);
                var page = new StringBuilder();
                for (int i = 0; i < size && index < tokens.Count; i++, index++)
                {
                    if (page.Length > 0) page.Append(' ');
                    page.Append(tokens[index]);
                }
                pages.Add(page.ToString());
                if (pages.Count >= 5 && index < tokens.Count)
                {
                    // Не больше пяти страниц: остаток уходит на последнюю.
                    for (; index < tokens.Count; index++)
                        pages[pages.Count - 1] += (pages[pages.Count - 1].Length > 0
                            ? " " : "") + tokens[index];
                }
            }
            return ReviewDocument(path, pages.ToArray());
        }

        private static List<string> AuditF1Keys(PdfReviewDocument document)
        {
            var keys = new List<string>();
            foreach (PdfReviewPage page in document.Pages)
                foreach (PdfReviewWord word in page.Words)
                    keys.Add(word.Key);
            return keys;
        }

        /// <summary>
        /// Свойство исправления 1: предпочтение страниц разрешает лишь равные по длине LCS.
        /// На случайной перевёрстке с повторными ключами Compare при любом пределе матрицы и
        /// в обе стороны находит ровно столько Equal, сколько слов в оптимальном LCS; при
        /// малом бюджете обе стороны всё равно восстанавливаются без потерь.
        /// </summary>
        [Test("Аудит F1: случайная перевёрстка с повторами сохраняет оптимальный LCS при любом пределе матрицы")]
        private static void AuditF1RandomRepaginationKeepsOptimalLcs()
        {
            var random = new Random(160);
            for (int sample = 0; sample < 300; sample++)
            {
                var baseline = new List<string>();
                int length = random.Next(0, 26);
                for (int i = 0; i < length; i++)
                    baseline.Add("k" + random.Next(0, 6).ToString(CultureInfo.InvariantCulture));
                var revised = new List<string>();
                foreach (string token in baseline)
                {
                    int edit = random.Next(10);
                    if (edit == 0) continue;
                    if (edit == 1) revised.Add("k" + random.Next(0, 6).ToString(CultureInfo.InvariantCulture));
                    revised.Add(token);
                }
                PdfReviewDocument left = AuditF1RandomPages(random, "random-left.pdf", baseline);
                PdfReviewDocument right = AuditF1RandomPages(random, "random-right.pdf", revised);
                List<string> leftKeys = AuditF1Keys(left);
                List<string> rightKeys = AuditF1Keys(right);
                int lcs = AuditF1LcsLength(leftKeys, rightKeys);
                string name = "выборка " + sample.ToString(CultureInfo.InvariantCulture);

                foreach (int maxCells in new[] { 36, 64, 256, PdfReviewLimits.AbsoluteMaxDiffCells })
                {
                    var limits = PdfReviewLimits.Default();
                    limits.MaxDiffCells = maxCells;
                    string context = name + ", MaxDiffCells=" +
                        maxCells.ToString(CultureInfo.InvariantCulture);
                    PdfReviewResult forward = PdfReviewDiff.Compare(left, right, limits);
                    PdfReviewResult reverse = PdfReviewDiff.Compare(right, left, limits);
                    AssertEq(lcs, ReviewOperationWordCount(forward.Operations,
                        PdfReviewDiffKind.Equal, -1), context + ": Equal прямо = LCS");
                    AssertEq(lcs, ReviewOperationWordCount(reverse.Operations,
                        PdfReviewDiffKind.Equal, -1), context + ": Equal обратно = LCS");
                    AssertEq(false, forward.WorkExhausted, context + ": без исчерпания");
                    AuditF1AssertExactOperations(forward, leftKeys.Count - lcs,
                        rightKeys.Count - lcs, context + ", прямо");
                    AuditF1AssertExactOperations(reverse, rightKeys.Count - lcs,
                        leftKeys.Count - lcs, context + ", обратно");
                }

                foreach (long work in new long[] { 0, 30, 300 })
                {
                    var limits = PdfReviewLimits.Default();
                    limits.MaxDiffCells = 36;
                    limits.MaxDiffWork = work;
                    string context = name + ", MaxDiffWork=" +
                        work.ToString(CultureInfo.InvariantCulture);
                    PdfReviewResult result = PdfReviewDiff.Compare(left, right, limits);
                    int equal = ReviewOperationWordCount(result.Operations,
                        PdfReviewDiffKind.Equal, -1);
                    AssertTrue(equal <= lcs, context + ": Equal " + equal + " не больше LCS " + lcs);
                    AuditF1AssertExactOperations(result, leftKeys.Count - equal,
                        rightKeys.Count - equal, context);
                    AssertTrue(result.WorkExhausted || equal == lcs,
                        context + ": неполный результат обязан быть помечен");
                }
            }
        }

        private static string AuditF1PageOrderKey(PdfReviewPage page)
        {
            if (page.Words.Count == 0)
                return PdfReviewDiff.Normalize(page.NormalizedText ?? page.Text ?? "");
            var key = new StringBuilder();
            foreach (PdfReviewWord word in page.Words)
            {
                string value = PdfReviewDiff.WordKey(word);
                key.Append(value.Length.ToString(CultureInfo.InvariantCulture))
                    .Append(':').Append(value);
            }
            return key.ToString();
        }

        private static bool AuditF1SameCost(double left, double right)
        {
            return left == right || Math.Abs(left - right) <= 1e-12;
        }

        private static bool AuditF1Better(double cost, int words, double bestCost, int bestWords)
        {
            if (!AuditF1SameCost(cost, bestCost))
                return cost < bestCost;
            return words > bestWords;
        }

        /// <summary>
        /// Эталон выравнивания до исправления: сходство каждой из n·m пар считается полным
        /// перебором множеств (через публичный Pair), DP и tie-break — как в Align.
        /// </summary>
        private static string AuditF1ReferenceAlignment(IList<PdfReviewPage> left,
            IList<PdfReviewPage> right)
        {
            const double Threshold = 0.24;
            const double Gap = 0.58;
            int n = left.Count, m = right.Count;
            var cost = new double[n + 1, m + 1];
            var words = new int[n + 1, m + 1];
            var move = new byte[n + 1, m + 1];
            for (int i = n - 1; i >= 0; i--) { cost[i, m] = cost[i + 1, m] + Gap; move[i, m] = 2; }
            for (int j = m - 1; j >= 0; j--) { cost[n, j] = cost[n, j + 1] + Gap; move[n, j] = 3; }
            for (int i = n - 1; i >= 0; i--)
            {
                for (int j = m - 1; j >= 0; j--)
                {
                    string leftKey = AuditF1PageOrderKey(left[i]);
                    string rightKey = AuditF1PageOrderKey(right[j]);
                    double sim = leftKey == rightKey ? 1 :
                        PdfReviewDiff.Pair(left[i], right[j], PdfReviewLimits.Default()).Similarity;
                    double bestCost = sim >= Threshold ? cost[i + 1, j + 1] + (1.0 - sim) : double.MaxValue;
                    int bestWords = sim >= Threshold ? words[i + 1, j + 1] +
                        Math.Min(left[i].Words.Count, right[j].Words.Count) : int.MinValue;
                    byte bestMove = 1;
                    if (AuditF1Better(cost[i + 1, j] + Gap, words[i + 1, j], bestCost, bestWords))
                    {
                        bestCost = cost[i + 1, j] + Gap;
                        bestWords = words[i + 1, j];
                        bestMove = 2;
                    }
                    double dropRight = cost[i, j + 1] + Gap;
                    int dropRightWords = words[i, j + 1];
                    int identity = string.CompareOrdinal(leftKey, rightKey);
                    if (identity == 0) identity = left[i].ViewWidthPt.CompareTo(right[j].ViewWidthPt);
                    if (identity == 0) identity = left[i].ViewHeightPt.CompareTo(right[j].ViewHeightPt);
                    if (AuditF1Better(dropRight, dropRightWords, bestCost, bestWords) ||
                        (AuditF1SameCost(dropRight, bestCost) && dropRightWords == bestWords &&
                         bestMove == 2 && identity > 0))
                    {
                        bestCost = dropRight;
                        bestWords = dropRightWords;
                        bestMove = 3;
                    }
                    cost[i, j] = bestCost;
                    words[i, j] = bestWords;
                    move[i, j] = bestMove;
                }
            }
            var parts = new List<string>();
            int li = 0, ri = 0;
            while (li < n || ri < m)
            {
                byte step = move[li, ri];
                if (step == 1)
                {
                    parts.Add(li + ":" + ri + ":" + (PdfReviewDiff.WordsEqual(left[li].Words,
                        right[ri].Words) ? PdfReviewPairStatus.Unchanged : PdfReviewPairStatus.Changed));
                    li++; ri++;
                }
                else if (step == 2 || ri >= m)
                {
                    parts.Add(li + ":-1:" + PdfReviewPairStatus.LeftOnly);
                    li++;
                }
                else
                {
                    parts.Add("-1:" + ri + ":" + PdfReviewPairStatus.RightOnly);
                    ri++;
                }
            }
            return string.Join("|", parts.ToArray());
        }

        private static List<PdfReviewPage> AuditF1RandomAlignmentPages(Random random,
            IList<PdfReviewPage> source, int count)
        {
            var pages = new List<PdfReviewPage>();
            for (int p = 0; p < count; p++)
            {
                var words = new List<string>();
                if (source != null && source.Count > 0 && random.Next(4) != 0)
                {
                    // Возмущённая копия страницы другой стороны: сходство около порога.
                    PdfReviewPage origin = source[random.Next(source.Count)];
                    foreach (PdfReviewWord word in origin.Words)
                        if (random.Next(10) >= 3)
                            words.Add(random.Next(6) == 0 ? word.Key.ToUpperInvariant() : word.Key);
                }
                int extra = random.Next(5) == 0 ? 0 : random.Next(1, 12);
                for (int i = 0; i < extra; i++)
                {
                    // Частые служебные ключи и редкие содержательные.
                    int rank = random.Next(40);
                    words.Add(rank < 20 ? "c" + (rank % 4).ToString(CultureInfo.InvariantCulture)
                        : "r" + random.Next(60).ToString(CultureInfo.InvariantCulture));
                }
                PdfReviewPage page = ReviewPage(p, string.Join(" ", words.ToArray()));
                if (random.Next(5) == 0)
                {
                    page.WidthPt = page.ViewWidthPt = 612;
                    page.HeightPt = page.ViewHeightPt = 792;
                }
                pages.Add(page);
            }
            return pages;
        }

        /// <summary>
        /// Свойство исправления 2: отсев пар по редким ключам и верхней оценке частых точен.
        /// На случайных страницах (частые служебные и редкие ключи, регистр, похожие копии
        /// около порога, пустые страницы, разные размеры) Align совпадает с полным перебором.
        /// </summary>
        [Test("Аудит F1: выравнивание с отсевом пар совпадает с полным перебором на случайных страницах")]
        private static void AuditF1AlignmentFilterMatchesBruteForce()
        {
            var random = new Random(2007);
            for (int sample = 0; sample < 400; sample++)
            {
                List<PdfReviewPage> left = AuditF1RandomAlignmentPages(random, null,
                    random.Next(0, 13));
                List<PdfReviewPage> right = AuditF1RandomAlignmentPages(random, left,
                    random.Next(0, 13));
                AssertEq(AuditF1ReferenceAlignment(left, right),
                    PairSig(PdfReviewDiff.Align(left, right, PdfReviewLimits.Default())),
                    "выборка " + sample.ToString(CultureInfo.InvariantCulture) + ": строки viewer");
            }
        }

        /// <summary>
        /// Снимок строк viewer сохраняет прежнюю семантику «ровно одна строка»: повторённая
        /// пара страниц не считается явной, и сверка воздерживается.
        /// </summary>
        [Test("Аудит F1: повторённая строка viewer не считается явной парой страниц")]
        private static void AuditF1DuplicatedPairIsNotExplicit()
        {
            PdfReviewResult single = ReviewOrderPermutationStageFixture();
            PdfReviewDiff.ReconcileOrderArtifacts(single, PdfReviewLimits.Default(), null);
            AssertEq(1, ReviewMatchCount(single, PdfReviewMatchKind.ReconciledOrder),
                "с одной строкой перестановка снимается");

            PdfReviewResult duplicated = ReviewOrderPermutationStageFixture();
            duplicated.Pairs.Add(new PdfReviewPagePair { LeftPageIndex = 0, RightPageIndex = 0 });
            PdfReviewDiff.ReconcileOrderArtifacts(duplicated, PdfReviewLimits.Default(), null);
            AssertEq(0, ReviewMatchCount(duplicated, PdfReviewMatchKind.ReconciledOrder),
                "две строки одной пары не являются явным сопоставлением");

            PdfReviewResult splitJoin = ReviewSplitJoinStageFixture();
            splitJoin.Pairs.Add(new PdfReviewPagePair { LeftPageIndex = 0, RightPageIndex = 0 });
            PdfReviewDiff.ClassifySplitJoin(splitJoin, PdfReviewLimits.Default(), null);
            AssertTrue(ReviewSplitJoinOperation(splitJoin) == null,
                "split/join тоже требует ровно одну строку пары");
        }

        /// <summary>
        /// Исправление 8 в другом потоке: документы сервиса общие для LRU, и параллельные
        /// Compare одной пары не пишут в слова и дают одинаковый результат.
        /// </summary>
        [Test("Аудит F1: параллельные Compare общих документов дают одинаковый результат")]
        private static void AuditF1ParallelComparesShareDocuments()
        {
            List<string> baseline = AuditF1Tokens("w", 3000);
            var revised = new List<string>(baseline);
            for (int i = 0; i < 30; i++)
                revised[i * 97 + 5] = "edit" + i.ToString(CultureInfo.InvariantCulture);
            PdfReviewDocument left = AuditF1Document("shared-left.pdf", baseline, 300);
            PdfReviewDocument right = AuditF1Document("shared-right.pdf", revised, 300);
            string expected = ReviewOpsSig(PdfReviewDiff.Compare(left, right,
                PdfReviewLimits.Default()).Operations);

            var failures = new List<string>();
            var threads = new List<System.Threading.Thread>();
            for (int t = 0; t < 4; t++)
            {
                var thread = new System.Threading.Thread(delegate()
                {
                    try
                    {
                        for (int i = 0; i < 5; i++)
                        {
                            string actual = ReviewOpsSig(PdfReviewDiff.Compare(left, right,
                                PdfReviewLimits.Default()).Operations);
                            if (actual != expected)
                                lock (failures) failures.Add("операции разошлись");
                        }
                    }
                    catch (Exception ex)
                    {
                        lock (failures) failures.Add(ex.GetType().Name + ": " + ex.Message);
                    }
                });
                thread.IsBackground = true;
                threads.Add(thread);
                thread.Start();
            }
            foreach (System.Threading.Thread thread in threads)
                AssertTrue(thread.Join(TimeSpan.FromSeconds(120)),
                    "поток параллельного Compare завершился");
            AssertEq(0, failures.Count, "сбои параллельного Compare: " +
                string.Join("; ", failures.ToArray()));
            for (int i = 0; i < left.Pages.Count; i++)
                foreach (PdfReviewWord word in left.Pages[i].Words)
                    AssertEq(i, word.PageIndex, "владелец слова общего документа не изменён");
        }
    }
}

namespace ExcelMerger.Tests
{
    /// <summary>Аудит F1, повторное ревью: зонды ревьюера, превращённые в тесты.</summary>
    internal static partial class UnitTests
    {
        private static string AuditF1Join(string first, IList<string> rest, string last)
        {
            var parts = new List<string>();
            if (first != null) parts.Add(first);
            parts.AddRange(rest);
            if (last != null) parts.Add(last);
            return string.Join(" ", parts.ToArray());
        }

        /// <summary>Односторонние слова операций в документном порядке: «ключ@pстраница».</summary>
        private static string AuditF1ChangedWords(PdfReviewResult result, PdfReviewDiffKind kind)
        {
            var parts = new List<string>();
            foreach (PdfReviewWordOp op in result.Operations)
                if (op.Kind == kind)
                    foreach (PdfReviewWord word in op.Words)
                        parts.Add(word.Key + "@p" +
                            word.PageIndex.ToString(CultureInfo.InvariantCulture));
            return string.Join(",", parts.ToArray());
        }

        /// <summary>
        /// H1 (зонд 3): служебные слова есть почти на каждой странице, поэтому оценка
        /// min(частых_i, частых_j) пропускала к поштучной проверке почти все n·m пар, и на
        /// 500 страницах выравнивание исчерпывало свой бюджет: у двух копий вместо 500 пар
        /// получалось 1000 односторонних строк.
        /// </summary>
        [Test("Аудит F1: две копии 500 страниц с частыми служебными словами сопоставляются постранично")]
        private static void AuditF1FrequentWordsStillAlignFiveHundredPages()
        {
            const int Pages = 500;
            // {словарь служебных слов, служебных на странице, своих слов на странице}
            foreach (int[] shape in new[] { new[] { 150, 80, 120 }, new[] { 100, 100, 100 } })
            {
                var random = new Random(shape[0]);
                var pages = new string[Pages];
                for (int page = 0; page < Pages; page++)
                {
                    var service = new List<string>();
                    for (int i = 0; i < shape[0]; i++)
                        service.Add("svc" + i.ToString(CultureInfo.InvariantCulture));
                    // Частичная тасовка Фишера-Йетса: shape[1] разных служебных слов.
                    for (int i = 0; i < shape[1]; i++)
                    {
                        int j = random.Next(i, service.Count);
                        string swap = service[i];
                        service[i] = service[j];
                        service[j] = swap;
                    }
                    List<string> words = service.GetRange(0, shape[1]);
                    words.AddRange(AuditF1Tokens(
                        "p" + page.ToString(CultureInfo.InvariantCulture) + "w", shape[2]));
                    pages[page] = string.Join(" ", words.ToArray());
                }
                PdfReviewDocument left = ReviewDocument("frequent-left.pdf", pages);
                PdfReviewDocument right = ReviewDocument("frequent-right.pdf", pages);
                string context = shape[1].ToString(CultureInfo.InvariantCulture) + " из " +
                    shape[0].ToString(CultureInfo.InvariantCulture) + " служебных + " +
                    shape[2].ToString(CultureInfo.InvariantCulture) + " своих";

                Stopwatch watch = Stopwatch.StartNew();
                PdfReviewResult result = PdfReviewDiff.Compare(left, right,
                    PdfReviewLimits.Default());
                watch.Stop();

                AssertEq(PdfReviewStage.None, result.IncompleteStages,
                    context + ": все стадии завершены");
                AssertEq(Pages, result.Pairs.Count, context + ": строк viewer");
                for (int i = 0; i < result.Pairs.Count; i++)
                {
                    PdfReviewPagePair pair = result.Pairs[i];
                    AssertTrue(pair.LeftPageIndex == i && pair.RightPageIndex == i &&
                               pair.Status == PdfReviewPairStatus.Unchanged,
                        context + ": строка " + i + " = " + pair.LeftPageIndex + ":" +
                        pair.RightPageIndex + ":" + pair.Status + ", ожидалась неизменная пара");
                }
                AssertTrue(watch.Elapsed < TimeSpan.FromSeconds(30),
                    context + ": сравнение заняло " + watch.Elapsed + " (порог 30 с)");
            }
        }

        /// <summary>
        /// M1 (зонд 1): перед документом вставлена обложка с тем же колонтитулом. Срезание
        /// краёв без предпочтений отдавало колонтитул первой страницы обложке, и неизменная
        /// страница выглядела изменённой. Совпадение должно остаться на сопоставленной паре.
        /// </summary>
        [Test("Аудит F1: колонтитул вставленной обложки не забирает совпадение неизменной страницы")]
        private static void AuditF1InsertedCoverKeepsHeaderOnPairedPage()
        {
            List<string> first = AuditF1Tokens("a", 300);
            List<string> second = AuditF1Tokens("b", 300);
            var changed = new List<string>(second);
            changed[150] = "CHANGED";
            PdfReviewDocument left = ReviewDocument("cover-left.pdf",
                AuditF1Join("HEADER", first, null), AuditF1Join("HEADER", second, null));
            PdfReviewDocument right = ReviewDocument("cover-right.pdf", "HEADER cover",
                AuditF1Join("HEADER", first, null), AuditF1Join("HEADER", changed, null));

            PdfReviewResult forward = PdfReviewDiff.Compare(left, right, PdfReviewLimits.Default());
            AssertEq("-1:0:RightOnly|0:1:Unchanged|1:2:Changed", PairSig(forward.Pairs),
                "прямо: строки viewer");
            AssertEq("HEADER@p0,cover@p0,CHANGED@p2",
                AuditF1ChangedWords(forward, PdfReviewDiffKind.Insert), "прямо: вставлено");
            AssertEq("b000150@p1", AuditF1ChangedWords(forward, PdfReviewDiffKind.Delete),
                "прямо: удалено");

            PdfReviewResult reverse = PdfReviewDiff.Compare(right, left, PdfReviewLimits.Default());
            AssertEq("0:-1:LeftOnly|1:0:Unchanged|2:1:Changed", PairSig(reverse.Pairs),
                "обратно: строки viewer");
            AssertEq("HEADER@p0,cover@p0,CHANGED@p2",
                AuditF1ChangedWords(reverse, PdfReviewDiffKind.Delete), "обратно: удалено");
            AssertEq("b000150@p1", AuditF1ChangedWords(reverse, PdfReviewDiffKind.Insert),
                "обратно: вставлено");
        }

        /// <summary>
        /// M1 (зонд 7): страница вставлена в середину 10-страничного документа с колонтитулом
        /// сверху и снизу, а дальше есть правка, которая обрывает срезание суффикса. Колонтитулы
        /// вставленной страницы должны быть вставкой, а соседние страницы — неизменными.
        /// </summary>
        [Test("Аудит F1: колонтитулы страницы, вставленной в середину, не ломают статусы соседних")]
        private static void AuditF1MiddleInsertionKeepsNeighbourPagesUnchanged()
        {
            const int Pages = 10;
            const int PageWords = 300;
            var leftPages = new List<string>();
            var rightPages = new List<string>();
            for (int page = 0; page < Pages; page++)
            {
                List<string> words = AuditF1Tokens(
                    "p" + page.ToString(CultureInfo.InvariantCulture) + "w", PageWords);
                leftPages.Add(AuditF1Join("HEADER", words, "RUNNING"));
                if (page == 6)
                    rightPages.Add(AuditF1Join("HEADER", AuditF1Tokens("new", PageWords),
                        "RUNNING"));
                if (page == 7)
                    words[150] = "CHANGED";
                rightPages.Add(AuditF1Join("HEADER", words, "RUNNING"));
            }
            PdfReviewDocument left = ReviewDocument("middle-left.pdf", leftPages.ToArray());
            PdfReviewDocument right = ReviewDocument("middle-right.pdf", rightPages.ToArray());

            var forwardRows = new List<string>();
            var reverseRows = new List<string>();
            for (int page = 0; page < Pages; page++)
            {
                int target = page < 6 ? page : page + 1;
                string status = page == 7 ? "Changed" : "Unchanged";
                if (page == 6)
                {
                    forwardRows.Add("-1:6:RightOnly");
                    reverseRows.Add("6:-1:LeftOnly");
                }
                forwardRows.Add(page + ":" + target + ":" + status);
                reverseRows.Add(target + ":" + page + ":" + status);
            }

            PdfReviewResult forward = PdfReviewDiff.Compare(left, right, PdfReviewLimits.Default());
            AssertEq(string.Join("|", forwardRows.ToArray()), PairSig(forward.Pairs),
                "прямо: строки viewer");
            AuditF1AssertExactOperations(forward, 1, PageWords + 3, "прямо");
            AssertEq(PageWords + 2, ReviewOperationWordCount(forward.Operations,
                PdfReviewDiffKind.Insert, 6), "прямо: вся вставленная страница — вставка");

            PdfReviewResult reverse = PdfReviewDiff.Compare(right, left, PdfReviewLimits.Default());
            AssertEq(string.Join("|", reverseRows.ToArray()), PairSig(reverse.Pairs),
                "обратно: строки viewer");
            AuditF1AssertExactOperations(reverse, PageWords + 3, 1, "обратно");
            AssertEq(PageWords + 2, ReviewOperationWordCount(reverse.Operations,
                PdfReviewDiffKind.Delete, 6), "обратно: вся удалённая страница — удаление");
        }

        /// <summary>
        /// M2 (зонды 4-5): при 3% равномерных замен в 250 тыс. слов верхняя бисекция тратила
        /// весь бюджет, не найдя середину, и весь документ становился сырыми Delete/Insert
        /// («изменено 100%» после ~20 с). Неудача одной бисекции не должна съедать всё.
        /// </summary>
        [Test("Аудит F1: 3% замен в 250 тыс. слов восстанавливаются, а не становятся «изменено 100%»")]
        private static void AuditF1DenseEditsInHugeDocumentKeepMostWords()
        {
            const int Words = 250000;
            const int PageWords = 500;
            List<string> baseline = AuditF1Tokens("w", Words);
            var revised = new List<string>(baseline);
            var random = new Random(4);
            int replaced = 0;
            for (int i = 0; i < Words; i++)
            {
                if (random.Next(100) >= 3)
                    continue;
                revised[i] = "r" + i.ToString("D6", CultureInfo.InvariantCulture);
                replaced++;
            }
            PdfReviewDocument left = AuditF1Document("dense-left.pdf", baseline, PageWords);
            PdfReviewDocument right = AuditF1Document("dense-right.pdf", revised, PageWords);

            Stopwatch watch = Stopwatch.StartNew();
            PdfReviewResult result = PdfReviewDiff.Compare(left, right,
                PdfReviewLimits.Default());
            watch.Stop();

            int equal = ReviewOperationWordCount(result.Operations, PdfReviewDiffKind.Equal, -1);
            AssertTrue(equal >= (Words - replaced) * 9 / 10,
                "совпало " + equal + " из " + (Words - replaced) + " неизменных слов (порог 90%)");
            AuditF1AssertExactOperations(result, Words - equal, Words - equal, "3% замен");
            AssertTrue(watch.Elapsed < TimeSpan.FromSeconds(60),
                "сравнение заняло " + watch.Elapsed + " (порог 60 с)");
        }

        /// <summary>
        /// L1: DiffWorkFor насыщался, как только baseWork·words переполнял long, хотя сам
        /// результат (baseWork·words/эталон) ещё помещается. Масштаб должен быть точным.
        /// </summary>
        [Test("Аудит F1: масштаб бюджета точен на больших значениях и насыщается только при переполнении")]
        private static void AuditF1DiffWorkScalesExactlyForLargeBudgets()
        {
            long reference = PdfReviewLimits.DiffWorkReferenceWords;
            long ceiling = reference * PdfReviewLimits.MaxDiffWorkScale;
            var limits = PdfReviewLimits.Default();
            foreach (long baseWork in new[] { long.MaxValue / 100000, long.MaxValue / 8,
                20000001L, long.MaxValue / 1000 - 1 })
            {
                foreach (long words in new[] { reference + 1, reference * 3 / 2 + 7, ceiling })
                {
                    limits.MaxDiffWork = baseWork;
                    // decimal держит 28 значащих цифр: произведение до ~1.8e24 и частное точны.
                    decimal exact = decimal.Floor((decimal)baseWork * words / reference);
                    AssertEq((long)exact, limits.DiffWorkFor(words),
                        "MaxDiffWork=" + baseWork.ToString(CultureInfo.InvariantCulture) +
                        ", слов " + words.ToString(CultureInfo.InvariantCulture));
                }
            }
            limits.MaxDiffWork = long.MaxValue / 100000;
            AssertEq(737869762948376L, limits.DiffWorkFor(200000),
                "значение из ревью: точный ×8 вместо насыщения");
            limits.MaxDiffWork = long.MaxValue / 8 + 1;
            AssertEq(long.MaxValue, limits.DiffWorkFor(ceiling),
                "результат больше long.MaxValue насыщается");
        }
    }
}

namespace ExcelMerger.Tests
{
    /// <summary>Аудит F1, повторное ревью: попытки сломать исправления.</summary>
    internal static partial class UnitTests
    {
        /// <summary>
        /// Исправление M1 на многословных колонтитулах с обеих сторон: обложка вставлена в
        /// начало (сдвиг вперёд, каскадом по словам шапки), последняя страница удалена
        /// (сдвиг назад, подвал), а правка в середине обрывает срезание краёв.
        /// </summary>
        [Test("Аудит F1: многословные шапка и подвал остаются на своих страницах при вставке и удалении")]
        private static void AuditF1MultiWordRunningHeadsStayOnPairedPages()
        {
            const string Head = "ACME QUARTERLY REPORT";
            const string Foot = "CONFIDENTIAL PAGE";
            var leftPages = new List<string>();
            var rightPages = new List<string> { Head + " cover " + Foot };
            for (int page = 0; page < 4; page++)
            {
                List<string> words = AuditF1Tokens(
                    "p" + page.ToString(CultureInfo.InvariantCulture) + "w", 280);
                leftPages.Add(AuditF1Join(Head, words, Foot));
                if (page == 1)
                    words[140] = "CHANGED";
                if (page < 3)
                    rightPages.Add(AuditF1Join(Head, words, Foot));
            }
            PdfReviewDocument left = ReviewDocument("heads-left.pdf", leftPages.ToArray());
            PdfReviewDocument right = ReviewDocument("heads-right.pdf", rightPages.ToArray());

            PdfReviewResult forward = PdfReviewDiff.Compare(left, right, PdfReviewLimits.Default());
            AssertEq("-1:0:RightOnly|0:1:Unchanged|1:2:Changed|2:3:Unchanged|3:-1:LeftOnly",
                PairSig(forward.Pairs), "прямо: строки viewer");
            AuditF1AssertExactOperations(forward, 1 + 285, 6 + 1, "прямо");

            PdfReviewResult reverse = PdfReviewDiff.Compare(right, left, PdfReviewLimits.Default());
            AssertEq("0:-1:LeftOnly|1:0:Unchanged|2:1:Changed|3:2:Unchanged|-1:3:RightOnly",
                PairSig(reverse.Pairs), "обратно: строки viewer");
            AuditF1AssertExactOperations(reverse, 6 + 1, 1 + 285, "обратно");
        }

        /// <summary>
        /// Исправление M2 на случайных входах: бюджета хватает на якоря, но не на точную
        /// бисекцию. Уникальные слова переставлены блоками (якорь может не входить в LCS),
        /// есть повторы. Обе стороны восстанавливаются без потерь и в своём порядке, Equal не
        /// больше LCS, а неоптимальный результат обязательно помечен.
        /// </summary>
        [Test("Аудит F1: якорное разбиение на случайных входах сохраняет обе стороны и честно помечено")]
        private static void AuditF1AnchoredSplitKeepsSidesAndIsFlagged()
        {
            var random = new Random(1999);
            int approximated = 0;
            for (int sample = 0; sample < 200; sample++)
            {
                var baseline = new List<string>();
                int length = random.Next(40, 161);
                for (int i = 0; i < length; i++)
                    baseline.Add(random.Next(2) == 0
                        ? "u" + i.ToString(CultureInfo.InvariantCulture)
                        : "k" + random.Next(5).ToString(CultureInfo.InvariantCulture));
                var revised = new List<string>();
                foreach (string token in baseline)
                {
                    int edit = random.Next(12);
                    if (edit == 0) continue;
                    if (edit == 1)
                        revised.Add("n" + random.Next(1000).ToString(CultureInfo.InvariantCulture));
                    if (edit != 2)
                        revised.Add(token);
                }
                // Перенос блока: уникальные слова блока оказываются якорями не по порядку.
                int blockLength = random.Next(1, 1 + revised.Count / 4);
                int blockStart = random.Next(revised.Count - blockLength + 1);
                List<string> block = revised.GetRange(blockStart, blockLength);
                revised.RemoveRange(blockStart, blockLength);
                revised.InsertRange(random.Next(revised.Count + 1), block);

                PdfReviewDocument left = AuditF1RandomPages(random, "anchor-left.pdf", baseline);
                PdfReviewDocument right = AuditF1RandomPages(random, "anchor-right.pdf", revised);
                List<string> leftKeys = AuditF1Keys(left);
                List<string> rightKeys = AuditF1Keys(right);
                int lcs = AuditF1LcsLength(leftKeys, rightKeys);
                foreach (int factor in new[] { 4, 8, 16 })
                {
                    var limits = PdfReviewLimits.Default();
                    limits.MaxDiffCells = 36;
                    limits.MaxDiffWork = factor * (leftKeys.Count + rightKeys.Count);
                    string context = "выборка " + sample.ToString(CultureInfo.InvariantCulture) +
                        ", бюджет ×" + factor.ToString(CultureInfo.InvariantCulture);
                    PdfReviewResult result = PdfReviewDiff.Compare(left, right, limits);
                    int equal = ReviewOperationWordCount(result.Operations,
                        PdfReviewDiffKind.Equal, -1);
                    AssertTrue(equal <= lcs, context + ": Equal " + equal + " не больше LCS " + lcs);
                    AuditF1AssertExactOperations(result, leftKeys.Count - equal,
                        rightKeys.Count - equal, context);
                    AssertTrue(result.WorkExhausted || equal == lcs,
                        context + ": неоптимальный результат обязан быть помечен");
                    if (result.WorkExhausted && equal > 0)
                        approximated++;
                }
            }
            AssertTrue(approximated > 0,
                "fixture действительно доходит до помеченных результатов с совпадениями");
        }

        /// <summary>
        /// Исправление H1: частые ключи считаются масками из нескольких машинных слов.
        /// Больше 64 частых ключей (старшие биты, перенос между словами маски, регистр)
        /// на случайных страницах — Align по-прежнему совпадает с полным перебором.
        /// </summary>
        [Test("Аудит F1: выравнивание с масками более 64 частых ключей совпадает с полным перебором")]
        private static void AuditF1WideFrequentMasksMatchBruteForce()
        {
            var random = new Random(64);
            for (int sample = 0; sample < 150; sample++)
            {
                var left = new List<PdfReviewPage>();
                var right = new List<PdfReviewPage>();
                int leftCount = random.Next(6, 13);
                int rightCount = random.Next(6, 13);
                for (int p = 0; p < leftCount + rightCount; p++)
                {
                    var words = new List<string>();
                    for (int k = 0; k < 140; k++)
                        if (random.Next(10) < 6)
                            words.Add((random.Next(8) == 0 ? "SVC" : "svc") +
                                k.ToString(CultureInfo.InvariantCulture));
                    int own = random.Next(0, 160);
                    for (int k = 0; k < own; k++)
                        words.Add("r" + random.Next(400).ToString(CultureInfo.InvariantCulture));
                    if (random.Next(6) == 0)
                        words.Clear();
                    PdfReviewPage page = ReviewPage(p < leftCount ? p : p - leftCount,
                        string.Join(" ", words.ToArray()));
                    (p < leftCount ? left : right).Add(page);
                }
                AssertEq(AuditF1ReferenceAlignment(left, right),
                    PairSig(PdfReviewDiff.Align(left, right, PdfReviewLimits.Default())),
                    "выборка " + sample.ToString(CultureInfo.InvariantCulture) + ": строки viewer");
            }
        }
    }
}
