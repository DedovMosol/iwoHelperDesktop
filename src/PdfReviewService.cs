using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Локальная оркестрация review двух явно выбранных PDF: проверки, извлечение,
    /// лёгкая проекция страниц и pure diff. Успешный текст кэшируется только в памяти.
    /// </summary>
    internal static class PdfReviewService
    {
        private const string AlgorithmVersion = "review-8";
        // Полные raw-слова одного текста схлопываются только при почти совпадающих рамках.
        // Ячейка больше максимального допуска центра, поэтому достаточно соседей 3×3.
        private const double OverlayCellSizePt = 4.0;
        private const double OverlayCenterFactor = 0.15;
        private const double OverlayCenterMinPt = 0.45;
        private const double OverlayCenterMaxPt = 2.0;
        private const double OverlayMinOverlapShare = 0.85;
        private const double OverlayMinSizeRatio = 0.85;
        private const int OverlayMaxCandidatesPerCell = 32;
        private static readonly object CacheGate = new object();
        private static readonly LruCache<PdfReviewDocument> Cache =
            new LruCache<PdfReviewDocument>(4, null);

        public static PdfReviewResult Compare(string leftPath, string rightPath,
            Action<int, int> progress = null, Func<bool> cancelled = null,
            PdfReviewLimits limits = null)
        {
            limits = limits ?? PdfReviewLimits.Default();
            string left = Canonical(leftPath), right = Canonical(rightPath);
            if (left == null || right == null)
                throw new PdfReviewException(PdfReviewFailure.Unreadable, null,
                    Loc.T("review.err.pickBoth"));
            if (OutputFile.IsSameFile(left, right))
                throw new PdfReviewException(PdfReviewFailure.Unreadable, left,
                    Loc.T("review.err.sameFile"));

            Action<int, int> leftProgress = progress == null ? null :
                (Action<int, int>)delegate(int done, int total)
                {
                    progress(total > 0 ? done : 0, Math.Max(1, total) * 2);
                };
            PdfReviewDocument leftDoc = Load(left, leftProgress, cancelled, limits);
            Cancellation.ThrowIf(cancelled);
            Action<int, int> rightProgress = progress == null ? null :
                (Action<int, int>)delegate(int done, int total)
                {
                    int t = Math.Max(1, total);
                    progress(t + done, t * 2);
                };
            PdfReviewDocument rightDoc = Load(right, rightProgress, cancelled, limits);
            Cancellation.ThrowIf(cancelled);
            PdfReviewResult result = PdfReviewDiff.Compare(leftDoc, rightDoc, limits,
                cancelled);
            // Текстовый слой PDF ненадёжен: одинаковые глифы могут иметь разные Unicode,
            // порядок и границы слов. Подтверждаем только найденные текстом кандидаты растром;
            // если рендер недоступен, fail-safe оставляет исходный diff, а не скрывает правку.
            PdfReviewVisualDiff.Refine(result, limits, cancelled);
            Cancellation.ThrowIf(cancelled);
            if (progress != null)
                progress(2, 2);
            return result;
        }

        public static PdfReviewDocument Load(string path, Action<int, int> progress,
            Func<bool> cancelled, PdfReviewLimits limits)
        {
            // PdfPig вшит в exe ресурсом и без перехвата разрешения не грузится: резолвер
            // обязан быть зарегистрирован ДО JIT-компиляции ядра, в теле которого есть типы
            // UglyToad (строка с «PdfDocument probe» одна роняла первое же сравнение
            // в однофайловой сборке: «Не удалось загрузить файл или сборку…»).
            EmbeddedAssemblies.Ensure();
            return LoadCore(path, progress, cancelled, limits);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static PdfReviewDocument LoadCore(string path, Action<int, int> progress,
            Func<bool> cancelled, PdfReviewLimits limits)
        {
            string full = Canonical(path);
            if (full == null || !File.Exists(full))
                throw Failure(PdfReviewFailure.Unreadable, path, Loc.T("review.err.unreadable"));
            FileStamp before = Stamp(full);
            if (before.Length > limits.MaxFileBytes)
                throw Failure(PdfReviewFailure.TooLarge, full, Loc.T("review.err.tooLarge"));
            string key = CacheKey(full, before);
            lock (CacheGate)
            {
                PdfReviewDocument cached;
                if (Cache.TryGet(key, out cached))
                    return cached;
            }

            // Открываем документ САМИ, а не через PdfPageProbe: тот глотает любое исключение
            // в «-1», и причина отказа терялась ещё до того, как её можно было понять. Здесь
            // исключение нужно живьём: по нему отличаем защищённый паролем файл (пользователю
            // будет предложен пароль) от битого/недоступного (сообщаем истинную причину).
            UglyToad.PdfPig.PdfDocument probe;
            try
            {
                probe = PdfPageProbe.OpenPig(full);
            }
            catch (Exception ex)
            {
                throw ReadFailure(full, ex);
            }
            int count;
            using (probe)
                count = probe.NumberOfPages;
            if (count <= 0)
                throw Failure(PdfReviewFailure.Unreadable, full, Loc.T("review.err.unreadable"));
            if (count > limits.MaxPages)
                throw Failure(PdfReviewFailure.TooLarge, full, Loc.T("review.err.tooLarge"));
            Cancellation.ThrowIf(cancelled);

            try
            {
                if (!PdfTextExtract.AnyPageHasText(full))
                    throw Failure(PdfReviewFailure.NoText, full, Loc.T("review.err.noText"));
                List<PdfPageText> extracted = PdfTextExtract.Extract(full, progress, null,
                    cancelled, PageLayoutMode.Review);
                var doc = new PdfReviewDocument { Path = full };
                foreach (PdfPageText page in extracted)
                {
                    Cancellation.ThrowIf(cancelled);
                    string text = PlainText.Page(page);
                    string normalized = PdfReviewDiff.Normalize(text);
                    doc.CharacterCount += PdfReviewDiff.CodePoints(normalized);
                    if (doc.CharacterCount > limits.MaxCharacters)
                        throw Failure(PdfReviewFailure.TooLarge, full, Loc.T("review.err.tooLarge"));
                    var reviewPage = new PdfReviewPage
                    {
                        PageIndex = page.PageIndex,
                        Text = text,
                        NormalizedText = normalized,
                        WidthPt = page.WidthPt,
                        HeightPt = page.HeightPt,
                        Fingerprint = PdfReviewDiff.Fingerprint(normalized, page.WidthPt, page.HeightPt)
                    };
                    double viewW, viewH;
                    PdfReviewGeometry.ViewSize(page.WidthPt, page.HeightPt, page.NativeRotation,
                        out viewW, out viewH);
                    reviewPage.ViewWidthPt = viewW;
                    reviewPage.ViewHeightPt = viewH;
                    BuildWords(page, reviewPage);
                    doc.WordCount += reviewPage.Words.Count;
                    doc.Pages.Add(reviewPage);
                }
                RetainDocumentEdgeWhitespace(doc);
                bool any = false;
                foreach (PdfReviewPage page in doc.Pages)
                    if (!string.IsNullOrWhiteSpace(page.NormalizedText)) { any = true; break; }
                if (!any)
                    throw Failure(PdfReviewFailure.NoText, full, Loc.T("review.err.noText"));
                FileStamp after = Stamp(full);
                if (before.Length != after.Length || before.WriteTicks != after.WriteTicks)
                    throw Failure(PdfReviewFailure.ChangedDuringRead, full,
                        Loc.T("review.err.changedDuringRead"));
                lock (CacheGate)
                    Cache.Add(key, doc);
                return doc;
            }
            catch (OperationCanceledException) { throw; }
            catch (PdfReviewException) { throw; }
            catch (Exception ex)
            {
                PdfReviewFailure reason = PdfPasswords.LooksPasswordProtected(ex)
                    ? PdfReviewFailure.PasswordRequired : PdfReviewFailure.Unreadable;
                string message = reason == PdfReviewFailure.PasswordRequired
                    ? Loc.T("review.err.password")
                    : string.Format(Loc.T("review.err.unreadableFile"), Path.GetFileName(full), ex.Message);
                throw Failure(reason, full, message);
            }
        }

        private struct FileStamp
        {
            public long Length;
            public long WriteTicks;
        }

        /// <summary>
        /// Видимые слова страницы в порядке чтения. PdfPig иногда отдаёт одно слово несколькими
        /// почти соприкасающимися фрагментами; внутри одной строки объединяем только такие
        /// фрагменты по общей геометрической политике OcrLayout. Настоящий пробел и пустой
        /// фрагмент остаются границей канонического слова. Единственное узкое исключение для
        /// границы строки — латинский перенос «inter-» + «national»: его восстанавливаем как
        /// одно слово. Дефисы кириллицы, цифр и дефисы внутри строки не меняем.
        /// </summary>
        internal static void BuildWords(PdfPageText page, PdfReviewPage reviewPage)
        {
            if (page.Words == null || page.Words.Count == 0)
                return;
            OcrLayout.Line previousLine = null;
            bool previousLineHasWords = false;
            foreach (OcrLayout.Line line in OcrLayout.ToLines(CollapseOverlayWords(page.Words)))
            {
                var lineWords = new List<PdfReviewWord>();
                var text = new StringBuilder();
                var fragments = new List<PdfWord>();
                PdfWord previous = null;
                double left = 0, bottom = 0, right = 0, top = 0;
                foreach (PdfWord word in line.Words)
                {
                    string fragment = word.Text == null ? "" : word.Text.Trim();
                    if (fragment.Length == 0)
                    {
                        AddWord(lineWords, reviewPage.PageIndex, page, text, fragments,
                            left, bottom, right, top);
                        text.Length = 0;
                        fragments.Clear();
                        previous = null;
                        continue;
                    }

                    if (text.Length > 0 && (OcrLayout.HasSpaceBetween(previous, word) ||
                        HasExplicitSourceWhitespace(page, previous, word)))
                    {
                        AddWord(lineWords, reviewPage.PageIndex, page, text, fragments,
                            left, bottom, right, top);
                        text.Length = 0;
                        fragments.Clear();
                    }
                    if (text.Length == 0)
                    {
                        left = word.Left;
                        bottom = word.Bottom;
                        right = word.Right;
                        top = word.Top;
                    }
                    else
                    {
                        left = Math.Min(left, word.Left);
                        bottom = Math.Min(bottom, word.Bottom);
                        right = Math.Max(right, word.Right);
                        top = Math.Max(top, word.Top);
                    }
                    text.Append(fragment);
                    fragments.Add(word);
                    previous = word;
                }
                AddWord(lineWords, reviewPage.PageIndex, page, text, fragments,
                    left, bottom, right, top);
                bool lineHasWords = lineWords.Count > 0;
                if (previousLineHasWords)
                    JoinLatinLineEnd(reviewPage.Words, lineWords, previousLine, line);
                reviewPage.Words.AddRange(lineWords);
                previousLine = line;
                previousLineHasWords = lineHasWords;
            }
            ResolveWhitespaceBoundaries(page, reviewPage);
        }

        private const double HyphenLineStepMinEm = 0.5;
        private const double HyphenLineStepMaxEm = 3.0;
        private const double HyphenLineLeftToleranceEm = 3.0;

        /// <summary>
        /// Соседство в списке извлечения недостаточно: там рядом могут оказаться разные колонки
        /// или блоки через большой вертикальный разрыв. Снимаем перенос только у двух строк одного
        /// потока: вторая физически ниже, близка по шагу и начинается у того же левого края.
        /// </summary>
        private static bool AreAdjacentContinuationLines(OcrLayout.Line previous,
            OcrLayout.Line next)
        {
            if (previous == null || next == null)
                return false;
            int previousBlock = previous.BlockId;
            int nextBlock = next.BlockId;
            if (previousBlock >= 0 && nextBlock >= 0 && previousBlock != nextBlock)
                return false;

            double em = Math.Max(previous.Height, next.Height);
            if (em <= 0)
                return false;
            double step = previous.MidY - next.MidY;
            if (step < HyphenLineStepMinEm * em || step > HyphenLineStepMaxEm * em)
                return false;
            return Math.Abs(previous.Left - next.Left) <= HyphenLineLeftToleranceEm * em;
        }

        /// <summary>
        /// Снять только доказуемый для Review латинский перенос на соседнюю физическую строку.
        /// ASCII '-' и soft hyphen допустимы, обе соседние буквы обязаны быть латинскими. Поэтому
        /// «информационно-» + «коммуникационный», «ISO-» + «9001» и обычный same-line дефис
        /// остаются видимыми частями текста, а не превращаются в неявную эквивалентность.
        /// </summary>
        private static void JoinLatinLineEnd(IList<PdfReviewWord> previousWords,
            IList<PdfReviewWord> lineWords, OcrLayout.Line previousLine,
            OcrLayout.Line line)
        {
            if (previousWords == null || previousWords.Count == 0 ||
                lineWords == null || lineWords.Count == 0 ||
                !AreAdjacentContinuationLines(previousLine, line))
                return;
            PdfReviewWord left = previousWords[previousWords.Count - 1];
            PdfReviewWord right = lineWords[0];
            string leftText = left == null ? null : left.Text;
            string rightText = right == null ? null : right.Text;
            if (string.IsNullOrEmpty(leftText) || string.IsNullOrEmpty(rightText) ||
                leftText.Length < 2)
                return;
            char hyphen = leftText[leftText.Length - 1];
            if ((hyphen != '-' && hyphen != '\x00AD') ||
                !IsLatinLetter(leftText[leftText.Length - 2]) ||
                !IsLatinLetter(rightText[0]))
                return;

            left.Text = leftText.Substring(0, leftText.Length - 1) + rightText;
            left.Key = left.Text.Normalize(NormalizationForm.FormC);
            left.Box = Union(left.Box, right.Box);
            left.BlockId = left.BlockId >= 0 && left.BlockId == right.BlockId
                ? left.BlockId : -1;
            InvalidateSource(left);
            lineWords.RemoveAt(0);
        }

        private static bool IsLatinLetter(char value)
        {
            if (!char.IsLetter(value))
                return false;
            return (value >= 'A' && value <= 'Z') ||
                   (value >= 'a' && value <= 'z') ||
                   (value >= 'À' && value <= 'ɏ') ||
                   (value >= 'Ḁ' && value <= 'ỿ');
        }

        private static PdfReviewBox Union(PdfReviewBox left, PdfReviewBox right)
        {
            return new PdfReviewBox
            {
                Left = Math.Min(left.Left, right.Left),
                Bottom = Math.Min(left.Bottom, right.Bottom),
                Right = Math.Max(left.Right, right.Right),
                Top = Math.Max(left.Top, right.Top)
            };
        }

        private sealed class OverlayInput
        {
            public PdfWord Word;
            public string Key;
        }

        private sealed class OverlayCandidate
        {
            // Anchor не расширяется при union: цепочка соседних, но уже разных экземпляров
            // не должна схлопываться транзитивно.
            public PdfWord Anchor;
            public PdfWord Word;
            public int ResultIndex;
        }

        private struct OverlayCell : IEquatable<OverlayCell>
        {
            public long X;
            public long Y;

            public bool Equals(OverlayCell other)
            {
                return X == other.X && Y == other.Y;
            }

            public override bool Equals(object value)
            {
                return value is OverlayCell && Equals((OverlayCell)value);
            }

            public override int GetHashCode()
            {
                unchecked { return (X.GetHashCode() * 397) ^ Y.GetHashCode(); }
            }
        }

        /// <summary>
        /// PdfPig может вернуть одну физическую надпись несколькими полными Word (например,
        /// два почти совпадающих слоя). До строк и LCS сворачиваем только NFC-равные слова с
        /// почти одинаковой геометрией. Поиск ограничен соседними spatial-bucket, поэтому частое
        /// слово на большой странице не превращает проход в O(n²). Исходные PdfWord не меняются.
        /// </summary>
        private static List<PdfWord> CollapseOverlayWords(IList<PdfWord> words)
        {
            var ordered = new List<OverlayInput>(words.Count);
            for (int i = 0; i < words.Count; i++)
            {
                PdfWord word = words[i];
                if (word == null)
                    continue;
                string text = word.Text == null ? "" : word.Text.Trim();
                ordered.Add(new OverlayInput
                {
                    Word = word,
                    Key = text.Length == 0 ? "" : text.Normalize(NormalizationForm.FormC)
                });
            }
            ordered.Sort(delegate(OverlayInput a, OverlayInput b)
            {
                int c = b.Word.MidY.CompareTo(a.Word.MidY);
                if (c != 0) return c;
                c = a.Word.Left.CompareTo(b.Word.Left);
                if (c != 0) return c;
                c = a.Word.Right.CompareTo(b.Word.Right);
                if (c != 0) return c;
                c = a.Word.Bottom.CompareTo(b.Word.Bottom);
                if (c != 0) return c;
                c = a.Word.Top.CompareTo(b.Word.Top);
                if (c != 0) return c;
                return string.CompareOrdinal(a.Word.Text ?? "", b.Word.Text ?? "");
            });

            var result = new List<PdfWord>(ordered.Count);
            var byText = new Dictionary<string,
                Dictionary<OverlayCell, List<OverlayCandidate>>>(StringComparer.Ordinal);
            for (int i = 0; i < ordered.Count; i++)
            {
                OverlayInput input = ordered[i];
                long cellX, cellY;
                if (input.Key.Length == 0 || !TryOverlayCell(input.Word, out cellX, out cellY))
                {
                    result.Add(input.Word);
                    continue;
                }

                Dictionary<OverlayCell, List<OverlayCandidate>> cells;
                if (!byText.TryGetValue(input.Key, out cells))
                {
                    cells = new Dictionary<OverlayCell, List<OverlayCandidate>>();
                    byText.Add(input.Key, cells);
                }

                OverlayCandidate best = null;
                double bestDistance = double.MaxValue;
                for (long dx = -1; dx <= 1; dx++)
                {
                    for (long dy = -1; dy <= 1; dy++)
                    {
                        List<OverlayCandidate> candidates;
                        var cell = new OverlayCell { X = cellX + dx, Y = cellY + dy };
                        if (!cells.TryGetValue(cell, out candidates))
                            continue;
                        for (int c = 0; c < candidates.Count; c++)
                        {
                            OverlayCandidate candidate = candidates[c];
                            if (!IsOverlayDuplicate(candidate.Anchor, input.Word))
                                continue;
                            double x = CenterX(candidate.Anchor) - CenterX(input.Word);
                            double y = candidate.Anchor.MidY - input.Word.MidY;
                            double distance = x * x + y * y;
                            if (best == null || distance < bestDistance ||
                                (distance == bestDistance && candidate.ResultIndex < best.ResultIndex))
                            {
                                best = candidate;
                                bestDistance = distance;
                            }
                        }
                    }
                }

                if (best != null)
                {
                    if (object.ReferenceEquals(best.Word, best.Anchor))
                    {
                        best.Word = CloneWord(best.Anchor);
                        result[best.ResultIndex] = best.Word;
                    }
                    best.Word.Left = Math.Min(best.Word.Left, input.Word.Left);
                    best.Word.Right = Math.Max(best.Word.Right, input.Word.Right);
                    best.Word.Bottom = Math.Min(best.Word.Bottom, input.Word.Bottom);
                    best.Word.Top = Math.Max(best.Word.Top, input.Word.Top);
                    InvalidateSource(best.Word);
                    continue;
                }

                var added = new OverlayCandidate
                {
                    Anchor = input.Word,
                    Word = input.Word,
                    ResultIndex = result.Count
                };
                result.Add(input.Word);
                var ownCell = new OverlayCell { X = cellX, Y = cellY };
                List<OverlayCandidate> bucket;
                if (!cells.TryGetValue(ownCell, out bucket))
                {
                    bucket = new List<OverlayCandidate>();
                    cells.Add(ownCell, bucket);
                }
                // В патологической ячейке не индексируем бесконечно много несовпадающих
                // экземпляров: пропуск лишь сохраняет лишний semantic candidate (fail-safe).
                if (bucket.Count < OverlayMaxCandidatesPerCell)
                    bucket.Add(added);
            }
            return result;
        }

        private static bool IsOverlayDuplicate(PdfWord a, PdfWord b)
        {
            double aw = a.Right - a.Left, bw = b.Right - b.Left;
            double ah = a.Top - a.Bottom, bh = b.Top - b.Bottom;
            if (aw <= 0 || bw <= 0 || ah <= 0 || bh <= 0)
                return false;
            if (Math.Min(aw, bw) / Math.Max(aw, bw) < OverlayMinSizeRatio ||
                Math.Min(ah, bh) / Math.Max(ah, bh) < OverlayMinSizeRatio)
                return false;

            double xOverlap = Math.Min(a.Right, b.Right) - Math.Max(a.Left, b.Left);
            double yOverlap = Math.Min(a.Top, b.Top) - Math.Max(a.Bottom, b.Bottom);
            if (xOverlap < OverlayMinOverlapShare * Math.Min(aw, bw) ||
                yOverlap < OverlayMinOverlapShare * Math.Min(ah, bh))
                return false;

            double tolerance = OverlayCenterFactor * Math.Min(ah, bh);
            tolerance = Math.Max(OverlayCenterMinPt, Math.Min(OverlayCenterMaxPt, tolerance));
            return Math.Abs(CenterX(a) - CenterX(b)) <= tolerance &&
                   Math.Abs(a.MidY - b.MidY) <= tolerance;
        }

        private static bool TryOverlayCell(PdfWord word, out long x, out long y)
        {
            x = y = 0;
            if (!Finite(word.Left) || !Finite(word.Right) ||
                !Finite(word.Bottom) || !Finite(word.Top) ||
                word.Right <= word.Left || word.Top <= word.Bottom)
                return false;
            double sx = Math.Floor(CenterX(word) / OverlayCellSizePt);
            double sy = Math.Floor(word.MidY / OverlayCellSizePt);
            // Соседние ячейки ниже прибавляются через ±1: края long оставляем недостижимыми.
            if (!Finite(sx) || !Finite(sy) ||
                sx <= long.MinValue + 1.0 || sx >= long.MaxValue - 1.0 ||
                sy <= long.MinValue + 1.0 || sy >= long.MaxValue - 1.0)
                return false;
            x = (long)sx;
            y = (long)sy;
            return true;
        }

        private static double CenterX(PdfWord word)
        {
            return word.Left + (word.Right - word.Left) / 2.0;
        }

        private static bool Finite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static PdfWord CloneWord(PdfWord word)
        {
            return new PdfWord
            {
                Text = word.Text,
                Left = word.Left,
                Right = word.Right,
                Bottom = word.Bottom,
                Top = word.Top,
                BaselineXPt = word.BaselineXPt,
                BaselineYPt = word.BaselineYPt,
                BlockId = word.BlockId,
                SourceStart = word.SourceStart,
                SourceEnd = word.SourceEnd,
                SourceText = word.SourceText,
                SourceTrusted = word.SourceTrusted,
                FontSizePt = word.FontSizePt,
                Bold = word.Bold,
                Italic = word.Italic,
                ColorArgb = word.ColorArgb,
                FontName = word.FontName,
                Super = word.Super,
                Sub = word.Sub,
                Underline = word.Underline,
                Uri = word.Uri
            };
        }

        private static void AddWord(ICollection<PdfReviewWord> words, int pageIndex,
            PdfPageText page, StringBuilder text, IList<PdfWord> fragments,
            double left, double bottom, double right, double top)
        {
            if (text.Length == 0)
                return;
            string display = text.ToString();
            int sourceStart, sourceEnd;
            string sourceText;
            bool sourceTrusted = TryAggregateSource(fragments, display,
                out sourceStart, out sourceEnd, out sourceText);
            words.Add(new PdfReviewWord
            {
                Text = display,
                Key = display.Normalize(NormalizationForm.FormC),
                PageIndex = pageIndex,
                Box = PdfReviewGeometry.RawToView(left, bottom, right, top,
                    page.NativeRotation, page.WidthPt, page.HeightPt),
                SourceStart = sourceTrusted ? sourceStart : -1,
                SourceEnd = sourceTrusted ? sourceEnd : -1,
                SourceText = sourceTrusted ? sourceText : null,
                SourceTrusted = sourceTrusted,
                BlockId = CommonBlock(fragments)
            });
        }

        private static bool TryAggregateSource(IList<PdfWord> fragments,
            string display, out int start, out int end, out string sourceText)
        {
            start = end = -1;
            sourceText = null;
            if (fragments == null || fragments.Count == 0)
                return false;
            var source = new StringBuilder();
            int previousEnd = -1;
            for (int i = 0; i < fragments.Count; i++)
            {
                PdfWord fragment = fragments[i];
                if (fragment == null || !fragment.SourceTrusted ||
                    fragment.SourceStart < 0 || fragment.SourceEnd < fragment.SourceStart ||
                    string.IsNullOrEmpty(fragment.SourceText) ||
                    (i > 0 && fragment.SourceStart != previousEnd + 1))
                    return false;
                if (i == 0)
                    start = fragment.SourceStart;
                previousEnd = fragment.SourceEnd;
                source.Append(fragment.SourceText);
            }
            end = previousEnd;
            sourceText = source.ToString();
            return string.Equals(sourceText, display, StringComparison.Ordinal);
        }

        private static int CommonBlock(IList<PdfWord> fragments)
        {
            if (fragments == null || fragments.Count == 0 || fragments[0] == null)
                return -1;
            int block = fragments[0].BlockId;
            for (int i = 1; i < fragments.Count; i++)
                if (fragments[i] == null || fragments[i].BlockId != block)
                    return -1;
            return block;
        }

        private static bool HasExplicitSourceWhitespace(PdfPageText page,
            PdfWord before, PdfWord after)
        {
            string raw;
            return TrySourceRun(page, before, after, out raw) && raw.Length > 0;
        }

        /// <summary>
        /// Положительное доказательство границы: оба retained-слова имеют соседние
        /// source-span в прямом порядке, а между ними лежат только явно декодированные
        /// whitespace-unit (включая ноль unit для доказуемо пустой границы).
        /// </summary>
        private static bool TrySourceRun(PdfPageText page, PdfWord before,
            PdfWord after, out string raw)
        {
            raw = null;
            if (before == null || after == null || !before.SourceTrusted ||
                !after.SourceTrusted || before.SourceEnd < before.SourceStart ||
                after.SourceEnd < after.SourceStart || before.SourceEnd >= after.SourceStart ||
                (before.BlockId >= 0 && after.BlockId >= 0 &&
                 before.BlockId != after.BlockId))
                return false;
            return TrySourceUnits(page, before.SourceEnd + 1, after.SourceStart, out raw);
        }

        private static bool TrySourceRun(PdfPageText page, PdfReviewWord before,
            PdfReviewWord after, out string raw)
        {
            raw = null;
            if (before == null || after == null || !before.SourceTrusted ||
                !after.SourceTrusted || before.SourceEnd < before.SourceStart ||
                after.SourceEnd < after.SourceStart || before.SourceEnd >= after.SourceStart ||
                (before.BlockId >= 0 && after.BlockId >= 0 &&
                 before.BlockId != after.BlockId))
                return false;
            return TrySourceUnits(page, before.SourceEnd + 1, after.SourceStart, out raw);
        }

        private static bool TrySourceUnits(PdfPageText page, int start, int end,
            out string raw)
        {
            raw = null;
            if (page == null || page.SourceUnits == null || start < 0 || end < start ||
                end > page.SourceUnits.Count)
                return false;
            var value = new StringBuilder();
            for (int i = start; i < end; i++)
            {
                PdfSourceTextUnit unit = page.SourceUnits[i];
                if (unit == null || !unit.Trusted || string.IsNullOrEmpty(unit.Text) ||
                    !IsWhitespace(unit.Text))
                    return false;
                value.Append(unit.Text);
            }
            raw = value.ToString();
            return true;
        }

        private static bool IsWhitespace(string value)
        {
            if (string.IsNullOrEmpty(value))
                return false;
            for (int i = 0; i < value.Length; i++)
                if (!char.IsWhiteSpace(value[i]))
                    return false;
            return true;
        }

        private static void ResolveWhitespaceBoundaries(PdfPageText page,
            PdfReviewPage reviewPage)
        {
            reviewPage.WhitespaceBoundaries.Clear();
            if (page == null || reviewPage.Words.Count == 0)
                return;

            PdfReviewWord first = reviewPage.Words[0];
            string raw;
            if (first.SourceTrusted && TrySourceUnits(page, 0, first.SourceStart, out raw))
                reviewPage.WhitespaceBoundaries.Add(WhitespaceEvidence(reviewPage,
                    null, first, raw, true, false));

            for (int i = 1; i < reviewPage.Words.Count; i++)
            {
                PdfReviewWord before = reviewPage.Words[i - 1];
                PdfReviewWord after = reviewPage.Words[i];
                if (TrySourceRun(page, before, after, out raw))
                    reviewPage.WhitespaceBoundaries.Add(WhitespaceEvidence(reviewPage,
                        before, after, raw, false, false));
            }

            PdfReviewWord last = reviewPage.Words[reviewPage.Words.Count - 1];
            if (last.SourceTrusted && TrySourceUnits(page, last.SourceEnd + 1,
                page.SourceUnits == null ? 0 : page.SourceUnits.Count, out raw))
                reviewPage.WhitespaceBoundaries.Add(WhitespaceEvidence(reviewPage,
                    last, null, raw, false, true));
        }

        private static PdfReviewWhitespaceEvidence WhitespaceEvidence(
            PdfReviewPage page, PdfReviewWord before, PdfReviewWord after,
            string raw, bool atStart, bool atEnd)
        {
            return new PdfReviewWhitespaceEvidence
            {
                PageIndex = page.PageIndex,
                Before = before,
                After = after,
                RawText = raw,
                LogicalText = LogicalWhitespace(raw),
                AtPageStart = atStart,
                AtPageEnd = atEnd,
                MarkerBox = BoundaryMarker(page, before, after, atStart, atEnd)
            };
        }

        private static string LogicalWhitespace(string raw)
        {
            if (string.IsNullOrEmpty(raw))
                return "";
            var result = new StringBuilder(raw.Length);
            for (int i = 0; i < raw.Length; i++)
            {
                char value = raw[i];
                if (value == '\r')
                {
                    if (i + 1 < raw.Length && raw[i + 1] == '\n')
                        i++;
                    result.Append('\n');
                }
                else
                {
                    result.Append(value);
                }
            }
            return result.ToString();
        }

        private static PdfReviewBox BoundaryMarker(PdfReviewPage page,
            PdfReviewWord before, PdfReviewWord after, bool atStart, bool atEnd)
        {
            PdfReviewBox anchor = after != null ? after.Box : before.Box;
            double x;
            if (atStart)
                x = anchor.Left - 1.0;
            else if (atEnd)
                x = anchor.Right + 1.0;
            else if (before.Box.Right <= after.Box.Left)
                x = before.Box.Right + (after.Box.Left - before.Box.Right) / 2.0;
            else
                x = after.Box.Left;
            double bottom = anchor.Bottom;
            double top = anchor.Top;
            if (before != null && after != null)
            {
                double overlapBottom = Math.Max(before.Box.Bottom, after.Box.Bottom);
                double overlapTop = Math.Min(before.Box.Top, after.Box.Top);
                if (overlapTop > overlapBottom)
                {
                    bottom = overlapBottom;
                    top = overlapTop;
                }
            }
            if (top <= bottom)
            {
                double center = (anchor.Bottom + anchor.Top) / 2.0;
                bottom = center - 1.0;
                top = center + 1.0;
            }
            double maxX = page.ViewWidthPt > 0 ? page.ViewWidthPt : page.WidthPt;
            x = Math.Max(0, Math.Min(maxX, x));
            return new PdfReviewBox
            {
                Left = Math.Max(0, x - 0.5),
                Right = Math.Min(maxX, x + 0.5),
                Bottom = bottom,
                Top = top
            };
        }

        private static void InvalidateSource(PdfReviewWord word)
        {
            if (word == null)
                return;
            word.SourceStart = -1;
            word.SourceEnd = -1;
            word.SourceText = null;
            word.SourceTrusted = false;
        }

        private static void InvalidateSource(PdfWord word)
        {
            if (word == null)
                return;
            word.SourceStart = -1;
            word.SourceEnd = -1;
            word.SourceText = null;
            word.SourceTrusted = false;
        }

        internal static void RetainDocumentEdgeWhitespace(PdfReviewDocument document)
        {
            if (document == null || document.Pages.Count == 0)
                return;
            PdfReviewPage first = null, last = null;
            foreach (PdfReviewPage page in document.Pages)
            {
                if (page == null || page.Words.Count == 0)
                    continue;
                if (first == null)
                    first = page;
                last = page;
            }
            foreach (PdfReviewPage page in document.Pages)
            {
                if (page == null)
                    continue;
                page.WhitespaceBoundaries.RemoveAll(delegate(PdfReviewWhitespaceEvidence evidence)
                {
                    return evidence == null ||
                        (evidence.AtPageStart && !object.ReferenceEquals(page, first)) ||
                        (evidence.AtPageEnd && !object.ReferenceEquals(page, last));
                });
            }
        }

        private static FileStamp Stamp(string path)
        {
            try
            {
                var file = new FileInfo(path);
                return new FileStamp { Length = file.Length, WriteTicks = file.LastWriteTimeUtc.Ticks };
            }
            catch
            {
                throw Failure(PdfReviewFailure.Unreadable, path, Loc.T("review.err.unreadable"));
            }
        }

        private static string CacheKey(string path, FileStamp stamp)
        {
            return path.ToLowerInvariant() + "|" + stamp.Length + "|" + stamp.WriteTicks + "|" + AlgorithmVersion;
        }

        private static string Canonical(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return null;
            try { return Path.GetFullPath(path); }
            catch { return null; } // недопустимые символы пути → вызывающий выдаст «не читается»
        }

        private static PdfReviewException Failure(PdfReviewFailure reason, string path, string message)
        {
            return new PdfReviewException(reason, path, message);
        }

        /// <summary>
        /// Разобрать исключение ОТКРЫТИЯ документа: защищённый паролем файл — отдельный исход
        /// (форма предложит ввести пароль), всё остальное — «не читается» с истинной причиной.
        /// Раньше сюда не доходило вовсе: подсчёт страниц глотал исключение в «-1», и человек
        /// получал безликое «не удалось прочитать» даже на паролем защищённом документе.
        /// </summary>
        private static PdfReviewException ReadFailure(string path, Exception ex)
        {
            PdfReviewFailure reason = PdfPasswords.LooksPasswordProtected(ex)
                ? PdfReviewFailure.PasswordRequired : PdfReviewFailure.Unreadable;
            string message = reason == PdfReviewFailure.PasswordRequired
                ? Loc.T("review.err.password")
                : string.Format(Loc.T("review.err.unreadableFile"), Path.GetFileName(path), ex.Message);
            return Failure(reason, path, message);
        }

    }
}
