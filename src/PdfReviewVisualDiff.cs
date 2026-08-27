using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace ExcelMerger
{
    /// <summary>
    /// Визуальная верификация текстового diff. PDF может показывать одинаковые глифы, но
    /// отдавать разные Unicode/порядок/границы слов из-за ToUnicode и устройства content stream.
    /// Текстовый LCS остаётся источником кандидатов, а этот слой снимает только те Delete/Insert,
    /// в чьей области два растра действительно эквивалентны. Сбой рендера ничего не скрывает:
    /// исходный текстовый diff остаётся без изменений.
    /// </summary>
    internal static class PdfReviewVisualDiff
    {
        private const int TargetWidth = 1200;
        private const long OwnPixelCap = 4000000;
        private const int MinimumUsefulWidth = 256;
        private const int MaximumSampleWidth = 640;
        private const int MaximumSampleHeight = 256;
        private const int MaximumRegistrationShift = 4;
        private const double BoxPaddingPt = 1.25;
        private const int InkFloor = 12;
        private const double EquivalentInkDifference = 0.035;
        private const int CancellationPollWork = 4096;
        private const int MaximumCandidatesPerHunk = 512;
        private const long MaximumRelationChecks = 1000000;
        private const long MaximumSamplePixelsPerHunk = 8000000;
        private const long MaximumSamplePixelsPerResult = 32000000;

        private enum RefinementOutcome
        {
            Unchanged,
            Refined,
            Abstain
        }

        /// <summary>
        /// Не вызывает пользовательский callback на коротком заведомо безопасном no-op пути,
        /// но ограничивает задержку отмены при обходе больших недоверенных коллекций и растра.
        /// </summary>
        private sealed class CancellationPoller
        {
            private readonly Func<bool> _cancelled;
            private long _pendingWork;

            public CancellationPoller(Func<bool> cancelled)
            {
                _cancelled = cancelled;
            }

            public void Check()
            {
                _pendingWork = 0;
                Cancellation.ThrowIf(_cancelled);
            }

            public void Worked(long amount)
            {
                if (_cancelled == null || amount <= 0)
                    return;
                if (amount >= CancellationPollWork - _pendingWork)
                {
                    _pendingWork = 0;
                    Cancellation.ThrowIf(_cancelled);
                }
                else
                {
                    _pendingWork += amount;
                }
            }
        }

        /// <summary>
        /// Детерминированный общий потолок необязательного raster-proof. Исчерпание не
        /// меняет текстовую семантику: текущая hunk целиком остаётся Delete/Insert.
        /// </summary>
        private sealed class VisualWorkBudget
        {
            private long _relationChecks;
            private long _samplePixels;

            public readonly CancellationPoller Poller;

            public VisualWorkBudget(CancellationPoller poller)
            {
                Poller = poller;
            }

            public bool HasProofCapacity
            {
                get
                {
                    return _relationChecks < MaximumRelationChecks &&
                        _samplePixels < MaximumSamplePixelsPerResult;
                }
            }

            public bool TryTakeRelations(long count)
            {
                if (count < 0 || count > MaximumRelationChecks - _relationChecks)
                    return false;
                _relationChecks += count;
                return true;
            }

            public bool TryTakeSamplePixels(long count, ref long hunkSamplePixels)
            {
                if (count < 0 || count > MaximumSamplePixelsPerHunk - hunkSamplePixels ||
                    count > MaximumSamplePixelsPerResult - _samplePixels)
                    return false;
                hunkSamplePixels += count;
                _samplePixels += count;
                return true;
            }
        }

        private sealed class ChangeHunk
        {
            public readonly List<PdfReviewWordOp> Operations = new List<PdfReviewWordOp>();
            public List<PdfReviewWordOp> Refined;
            public int LeftPageIndex = -1;
            public int RightPageIndex = -1;
        }

        private sealed class DiffSegment
        {
            public PdfReviewWordOp Equal;
            public ChangeHunk Change;
        }

        private sealed class HunkGroup
        {
            public int LeftPageIndex;
            public int RightPageIndex;
            public readonly List<ChangeHunk> Hunks = new List<ChangeHunk>();
        }

        /// <summary>
        /// Растрово проверяются только смешанные Delete/Insert hunks, целиком принадлежащие
        /// одной явно сопоставленной паре физических страниц. Pure и cross-page кандидаты
        /// остаются семантическими: растр не имеет права создавать или угадывать соответствие.
        /// </summary>
        public static void Refine(PdfReviewResult result, PdfReviewLimits limits,
            Func<bool> cancelled)
        {
            Refine(result, limits, cancelled, null);
        }

        /// <summary>
        /// Детерминированный seam для тестов orchestration: production передаёт null и владеет
        /// PdfThumbnailRenderer, тест может вернуть принадлежащие вызывающему bitmap без WinRT.
        /// Семантика, лимиты, отмена и атомарная публикация остаются теми же.
        /// </summary>
        internal static void Refine(PdfReviewResult result, PdfReviewLimits limits,
            Func<bool> cancelled, Func<string, int, int, int, Bitmap> renderPage)
        {
            if (result == null || result.Left == null || result.Right == null)
                return;
            List<PdfReviewWordOp> sourceOperations = result.Operations;
            List<PdfReviewWhitespaceChange> sourceWhitespace = result.WhitespaceChanges;
            if (sourceOperations.Count == 0)
                return;
            limits = limits ?? PdfReviewLimits.Default();
            long pixelCap = Math.Min(OwnPixelCap, limits.MaxRenderPixels);
            if (pixelCap <= 0)
                return;

            var poller = new CancellationPoller(cancelled);
            var work = new VisualWorkBudget(poller);
            List<DiffSegment> segments = Segments(sourceOperations, poller);
            List<HunkGroup> groups = SafeGroups(result, segments, poller);
            bool anyRefined = false;
            if (groups.Count > 0)
            {
                try
                {
                    poller.Check();
                    using (var renderer = renderPage == null ? new PdfThumbnailRenderer() : null)
                    {
                        Func<string, int, int, int, VisualPage> render =
                            delegate(string path, int pageIndex, int width, int maxHeight)
                            {
                                if (renderer != null)
                                    return new VisualPage(renderer.RenderOwned(path, pageIndex,
                                        width, maxHeight, pixelCap));
                                return new VisualPage(renderPage(path, pageIndex,
                                    width, maxHeight));
                            };
                        anyRefined = RefineGroups(result, groups, pixelCap, render, work);
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch
                {
                    // Не удалось создать renderer: весь текстовый diff остаётся без изменений.
                }
            }

            if (!anyRefined)
                return;

            var rebuilt = new List<PdfReviewWordOp>();
            foreach (DiffSegment segment in segments)
            {
                poller.Worked(1);
                IList<PdfReviewWordOp> source = segment.Change == null
                    ? (IList<PdfReviewWordOp>)new[] { segment.Equal }
                    : segment.Change.Refined ?? segment.Change.Operations;
                AppendOperations(rebuilt, source, poller);
            }
            if (rebuilt.Count == 0)
                rebuilt.Add(new PdfReviewWordOp { Kind = PdfReviewDiffKind.Equal });

            // Не затираем более свежую публикацию, если result изменили параллельно.
            if (!ReferenceEquals(result.Operations, sourceOperations) ||
                !ReferenceEquals(result.WhitespaceChanges, sourceWhitespace))
                return;
            PdfReviewDiff.PublishProjection(result, rebuilt, sourceWhitespace, cancelled);
        }

        private sealed class VisualPage : IDisposable
        {
            private BudgetedBitmap _owned;
            internal readonly Bitmap Bitmap;

            internal VisualPage(BudgetedBitmap owned)
            {
                _owned = owned;
                Bitmap = owned == null ? null : owned.Bitmap;
            }

            internal VisualPage(Bitmap bitmap)
            {
                Bitmap = bitmap;
            }

            public void Dispose()
            {
                if (_owned != null)
                    _owned.Dispose();
                else if (Bitmap != null)
                    Bitmap.Dispose();
                _owned = null;
            }
        }

        private static bool RefineGroups(PdfReviewResult result,
            IList<HunkGroup> groups, long pixelCap,
            Func<string, int, int, int, VisualPage> renderPage, VisualWorkBudget work)
        {
            bool anyRefined = false;
            foreach (HunkGroup group in groups)
            {
                if (!work.HasProofCapacity)
                    break;
                work.Poller.Check();
                PdfReviewPage leftPage = PdfReviewDiff.PageAt(result.Left,
                    group.LeftPageIndex);
                PdfReviewPage rightPage = PdfReviewDiff.PageAt(result.Right,
                    group.RightPageIndex);
                if (leftPage == null || rightPage == null)
                    continue;
                int width = RenderWidth(leftPage, rightPage, pixelCap);
                if (width < MinimumUsefulWidth)
                    continue;
                int maxHeight = (int)Math.Max(1,
                    Math.Min(20000, pixelCap / width));
                VisualPage leftPageRaster = null;
                VisualPage rightPageRaster = null;
                try
                {
                    leftPageRaster = renderPage(result.Left.Path, group.LeftPageIndex,
                        width, maxHeight);
                    rightPageRaster = renderPage(result.Right.Path, group.RightPageIndex,
                        width, maxHeight);
                    Bitmap left = leftPageRaster == null ? null : leftPageRaster.Bitmap;
                    Bitmap right = rightPageRaster == null ? null : rightPageRaster.Bitmap;
                    if (!WithinLimit(left, pixelCap) || !WithinLimit(right, pixelCap))
                        continue;
                    InkRaster leftInk = InkRaster.From(left, work.Poller);
                    InkRaster rightInk = InkRaster.From(right, work.Poller);
                    if (leftInk == null || rightInk == null)
                        continue;
                    foreach (ChangeHunk hunk in group.Hunks)
                    {
                        if (!work.HasProofCapacity)
                            break;
                        work.Poller.Check();
                        List<PdfReviewWordOp> refined;
                        if (RefineHunk(hunk.Operations, leftPage, rightPage,
                            leftInk, rightInk, work, out refined))
                        {
                            hunk.Refined = refined;
                            anyRefined = true;
                        }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch
                {
                    // Fail-safe: сбой одной пары сохраняет её исходные кандидаты и
                    // не мешает проверить независимые пары документа.
                }
                finally
                {
                    if (rightPageRaster != null && (leftPageRaster == null ||
                        !ReferenceEquals(rightPageRaster.Bitmap, leftPageRaster.Bitmap)))
                        rightPageRaster.Dispose();
                    if (leftPageRaster != null)
                        leftPageRaster.Dispose();
                }
            }
            return anyRefined;
        }

        private static List<DiffSegment> Segments(IList<PdfReviewWordOp> operations,
            CancellationPoller poller)
        {
            var result = new List<DiffSegment>();
            ChangeHunk current = null;
            foreach (PdfReviewWordOp op in operations)
            {
                poller.Worked(1);
                if (op == null)
                    continue;
                if (op.Kind == PdfReviewDiffKind.Equal)
                {
                    current = null;
                    result.Add(new DiffSegment { Equal = op });
                }
                else
                {
                    if (current == null)
                    {
                        current = new ChangeHunk();
                        result.Add(new DiffSegment { Change = current });
                    }
                    current.Operations.Add(op);
                }
            }
            return result;
        }

        private static List<HunkGroup> SafeGroups(PdfReviewResult result,
            IList<DiffSegment> segments, CancellationPoller poller)
        {
            var groups = new List<HunkGroup>();
            var byPair = new Dictionary<long, HunkGroup>();
            foreach (DiffSegment segment in segments)
            {
                poller.Worked(1);
                ChangeHunk hunk = segment.Change;
                if (hunk == null || !TryOwners(hunk.Operations,
                    out hunk.LeftPageIndex, out hunk.RightPageIndex, poller) ||
                    !IsDisplayPair(result.Pairs, hunk.LeftPageIndex,
                        hunk.RightPageIndex, poller))
                    continue;
                long key = ((long)hunk.LeftPageIndex << 32) |
                    (uint)hunk.RightPageIndex;
                HunkGroup group;
                if (!byPair.TryGetValue(key, out group))
                {
                    group = new HunkGroup
                    {
                        LeftPageIndex = hunk.LeftPageIndex,
                        RightPageIndex = hunk.RightPageIndex
                    };
                    byPair.Add(key, group);
                    groups.Add(group);
                }
                group.Hunks.Add(hunk);
            }
            return groups;
        }

        private static bool TryOwners(IList<PdfReviewWordOp> operations,
            out int leftPageIndex, out int rightPageIndex,
            CancellationPoller poller)
        {
            leftPageIndex = rightPageIndex = -1;
            bool deleted = false, inserted = false;
            int candidateCount = 0;
            foreach (PdfReviewWordOp op in operations)
            {
                poller.Worked(1);
                if (op == null || (op.Kind != PdfReviewDiffKind.Delete &&
                    op.Kind != PdfReviewDiffKind.Insert))
                    continue;
                IList<PdfReviewWord> words = op.Kind == PdfReviewDiffKind.Delete
                    ? (IList<PdfReviewWord>)op.LeftWords : op.RightWords;
                foreach (PdfReviewWord word in words)
                {
                    poller.Worked(1);
                    if (++candidateCount > MaximumCandidatesPerHunk ||
                        word == null || word.PageIndex < 0)
                        return false;
                    if (op.Kind == PdfReviewDiffKind.Delete)
                    {
                        if (deleted && leftPageIndex != word.PageIndex)
                            return false;
                        leftPageIndex = word.PageIndex;
                        deleted = true;
                    }
                    else
                    {
                        if (inserted && rightPageIndex != word.PageIndex)
                            return false;
                        rightPageIndex = word.PageIndex;
                        inserted = true;
                    }
                }
            }
            return deleted && inserted;
        }

        private static bool IsDisplayPair(IList<PdfReviewPagePair> pairs,
            int leftPageIndex, int rightPageIndex, CancellationPoller poller)
        {
            if (pairs == null)
                return false;
            foreach (PdfReviewPagePair pair in pairs)
            {
                poller.Worked(1);
                if (pair != null && pair.LeftPageIndex == leftPageIndex &&
                    pair.RightPageIndex == rightPageIndex)
                    return true;
            }
            return false;
        }

        private static void AppendOperations(List<PdfReviewWordOp> target,
            IList<PdfReviewWordOp> source, CancellationPoller poller)
        {
            foreach (PdfReviewWordOp op in source)
            {
                poller.Worked(1);
                PdfReviewDiff.AppendOperation(target, op);
            }
        }

        /// <summary>
        /// Чистое растровое ядро для одной безопасной локальной hunk. Возвращает новый список
        /// только если хотя бы один кандидат действительно снят.
        /// </summary>
        internal static bool RefineHunk(IList<PdfReviewWordOp> operations,
            PdfReviewPage leftPage, PdfReviewPage rightPage, Bitmap leftBitmap,
            Bitmap rightBitmap, out List<PdfReviewWordOp> refinedOperations)
        {
            return RefineHunk(operations, leftPage, rightPage, leftBitmap, rightBitmap,
                null, out refinedOperations);
        }

        internal static bool RefineHunk(IList<PdfReviewWordOp> operations,
            PdfReviewPage leftPage, PdfReviewPage rightPage, Bitmap leftBitmap,
            Bitmap rightBitmap, Func<bool> cancelled,
            out List<PdfReviewWordOp> refinedOperations)
        {
            refinedOperations = null;
            if (operations == null || leftPage == null || rightPage == null ||
                leftBitmap == null || rightBitmap == null)
                return false;

            var poller = new CancellationPoller(cancelled);
            var work = new VisualWorkBudget(poller);
            poller.Check();
            if (!WithinCandidateLimit(operations, poller))
                return false;
            InkRaster left = InkRaster.From(leftBitmap, poller);
            InkRaster right = InkRaster.From(rightBitmap, poller);
            if (left == null || right == null)
                return false;
            return RefineHunk(operations, leftPage, rightPage, left, right, work,
                out refinedOperations);
        }

        private static bool RefineHunk(IList<PdfReviewWordOp> operations,
            PdfReviewPage leftPage, PdfReviewPage rightPage, InkRaster left,
            InkRaster right, VisualWorkBudget work,
            out List<PdfReviewWordOp> refinedOperations)
        {
            refinedOperations = null;
            if (!WithinCandidateLimit(operations, work.Poller))
                return false;

            bool refined = false;
            long hunkSamplePixels = 0;
            var rebuilt = new List<PdfReviewWordOp>();
            var changeRun = new List<PdfReviewWordOp>();
            foreach (PdfReviewWordOp op in operations)
            {
                work.Poller.Worked(1);
                if (op != null && (op.Kind == PdfReviewDiffKind.Delete ||
                    op.Kind == PdfReviewDiffKind.Insert))
                {
                    changeRun.Add(op);
                    continue;
                }

                RefinementOutcome outcome = AppendRefinedRun(rebuilt, changeRun,
                    leftPage, rightPage, left, right, work, ref hunkSamplePixels);
                if (outcome == RefinementOutcome.Abstain)
                    return false;
                refined |= outcome == RefinementOutcome.Refined;
                changeRun.Clear();
                PdfReviewDiff.AppendOperation(rebuilt, op);
            }
            RefinementOutcome last = AppendRefinedRun(rebuilt, changeRun, leftPage,
                rightPage, left, right, work, ref hunkSamplePixels);
            if (last == RefinementOutcome.Abstain)
                return false;
            refined |= last == RefinementOutcome.Refined;

            if (!refined)
                return false;
            refinedOperations = rebuilt;
            return true;
        }

        private static bool WithinCandidateLimit(IList<PdfReviewWordOp> operations,
            CancellationPoller poller)
        {
            int count = 0;
            foreach (PdfReviewWordOp op in operations)
            {
                poller.Worked(1);
                if (op == null || (op.Kind != PdfReviewDiffKind.Delete &&
                    op.Kind != PdfReviewDiffKind.Insert))
                    continue;
                IList<PdfReviewWord> words = op.Kind == PdfReviewDiffKind.Delete
                    ? (IList<PdfReviewWord>)op.LeftWords : op.RightWords;
                if (words.Count > MaximumCandidatesPerHunk - count)
                    return false;
                count += words.Count;
            }
            return true;
        }

        /// <summary>
        /// Растровое решение атомарно для связной группы и хранит обе документные
        /// последовательности. Для many-to-many группы намеренно нет выдуманных пар слов.
        /// </summary>
        private static RefinementOutcome AppendRefinedRun(List<PdfReviewWordOp> target,
            IList<PdfReviewWordOp> operations, PdfReviewPage leftPage,
            PdfReviewPage rightPage, InkRaster left, InkRaster right,
            VisualWorkBudget work, ref long hunkSamplePixels)
        {
            if (operations == null || operations.Count == 0)
                return RefinementOutcome.Unchanged;

            var deletes = new List<Candidate>();
            var inserts = new List<Candidate>();
            foreach (PdfReviewWordOp op in operations)
            {
                work.Poller.Worked(1);
                if (op == null)
                    continue;
                bool leftSide = op.Kind == PdfReviewDiffKind.Delete;
                PdfReviewPage page = leftSide ? leftPage : rightPage;
                IList<PdfReviewWord> words = leftSide
                    ? (IList<PdfReviewWord>)op.LeftWords : op.RightWords;
                foreach (PdfReviewWord word in words)
                {
                    work.Poller.Worked(1);
                    NormalizedBox region;
                    bool valid = TryNormalize(word == null ? new PdfReviewBox() : word.Box,
                        page, out region);
                    var candidate = new Candidate
                    {
                        Word = word,
                        Region = region,
                        Order = leftSide ? deletes.Count : inserts.Count,
                        Valid = valid,
                        Component = -1
                    };
                    if (leftSide) deletes.Add(candidate); else inserts.Add(candidate);
                }
            }
            if (deletes.Count == 0 || inserts.Count == 0)
            {
                AppendOperations(target, operations, work.Poller);
                return RefinementOutcome.Unchanged;
            }

            if (!ConfirmEquivalentComponents(deletes, inserts, left, right, work,
                ref hunkSamplePixels))
                return RefinementOutcome.Abstain;
            List<EquivalentGroup> groups = EquivalentGroups(deletes, inserts,
                work.Poller);
            if (groups.Count == 0)
            {
                AppendOperations(target, operations, work.Poller);
                return RefinementOutcome.Unchanged;
            }

            int deleteIndex = 0, insertIndex = 0;
            foreach (EquivalentGroup group in groups)
            {
                work.Poller.Worked(1);
                AppendCandidates(target, deletes, deleteIndex, group.LeftStart,
                    PdfReviewDiffKind.Delete, work.Poller);
                AppendCandidates(target, inserts, insertIndex, group.RightStart,
                    PdfReviewDiffKind.Insert, work.Poller);

                var equal = new PdfReviewWordOp
                {
                    Kind = PdfReviewDiffKind.Equal,
                    MatchKind = PdfReviewMatchKind.RasterEquivalent
                };
                for (int i = group.LeftStart; i <= group.LeftEnd; i++)
                {
                    work.Poller.Worked(1);
                    equal.LeftWords.Add(deletes[i].Word);
                }
                for (int i = group.RightStart; i <= group.RightEnd; i++)
                {
                    work.Poller.Worked(1);
                    equal.RightWords.Add(inserts[i].Word);
                }
                if (group.LeftCount == 1 && group.RightCount == 1)
                    equal.Matches.Add(new PdfReviewWordMatch
                    {
                        Left = deletes[group.LeftStart].Word,
                        Right = inserts[group.RightStart].Word,
                        Kind = PdfReviewMatchKind.RasterEquivalent
                    });
                PdfReviewDiff.AppendOperation(target, equal);

                deleteIndex = group.LeftEnd + 1;
                insertIndex = group.RightEnd + 1;
            }
            AppendCandidates(target, deletes, deleteIndex, deletes.Count,
                PdfReviewDiffKind.Delete, work.Poller);
            AppendCandidates(target, inserts, insertIndex, inserts.Count,
                PdfReviewDiffKind.Insert, work.Poller);
            return RefinementOutcome.Refined;
        }

        private static void AppendCandidates(List<PdfReviewWordOp> target,
            IList<Candidate> candidates, int start, int end, PdfReviewDiffKind kind,
            CancellationPoller poller)
        {
            for (int i = start; i < end; i++)
            {
                poller.Worked(1);
                PdfReviewDiff.Append(target, kind, candidates[i].Word);
            }
        }

        private static bool ConfirmEquivalentComponents(List<Candidate> deletes,
            List<Candidate> inserts, InkRaster left, InkRaster right,
            VisualWorkBudget work, ref long hunkSamplePixels)
        {
            int deleteCount = deletes.Count;
            int count = deleteCount + inserts.Count;
            if (count == 0)
                return true;
            long relationChecks = (long)deleteCount * inserts.Count;
            if (!work.TryTakeRelations(relationChecks))
                return false;

            var parent = new int[count];
            var rank = new byte[count];
            for (int i = 0; i < count; i++)
            {
                work.Poller.Worked(1);
                parent[i] = i;
            }
            for (int d = 0; d < deleteCount; d++)
                for (int i = 0; i < inserts.Count; i++)
                {
                    work.Poller.Worked(1);
                    if (Related(deletes[d].Region, inserts[i].Region))
                        Union(parent, rank, d, deleteCount + i);
                }

            var used = new bool[count];
            var valid = new bool[count];
            var hasDelete = new bool[count];
            var hasInsert = new bool[count];
            var regions = new NormalizedBox[count];
            for (int i = 0; i < count; i++)
            {
                work.Poller.Worked(1);
                int root = Find(parent, i);
                Candidate candidate = CandidateAt(deletes, inserts, i);
                candidate.Component = root;
                if (i < deleteCount) hasDelete[root] = true; else hasInsert[root] = true;
                if (!used[root])
                {
                    used[root] = true;
                    valid[root] = candidate.Valid;
                    regions[root] = candidate.Region;
                }
                else
                {
                    valid[root] = valid[root] && candidate.Valid;
                    regions[root] = Union(regions[root], candidate.Region);
                }
            }

            int maxShift = RegistrationShift(left, right);
            long requiredSamplePixels = 0;
            for (int i = 0; i < count; i++)
            {
                work.Poller.Worked(1);
                if (!used[i] || !valid[i] || !hasDelete[i] || !hasInsert[i])
                    continue;
                long regionWork = RegionSampleWork(left, right, regions[i], maxShift);
                if (regionWork < 0 ||
                    regionWork > MaximumSamplePixelsPerHunk - requiredSamplePixels)
                    return false;
                requiredSamplePixels += regionWork;
            }
            if (!work.TryTakeSamplePixels(requiredSamplePixels, ref hunkSamplePixels))
                return false;

            var equivalent = new bool[count];
            for (int i = 0; i < count; i++)
            {
                work.Poller.Worked(1);
                if (used[i] && valid[i] && hasDelete[i] && hasInsert[i])
                    equivalent[i] = RegionEquivalent(left, right, regions[i], maxShift,
                        work.Poller);
            }

            for (int d = 0; d < deleteCount; d++)
            {
                work.Poller.Worked(1);
                deletes[d].Equivalent = equivalent[deletes[d].Component];
            }
            for (int i = 0; i < inserts.Count; i++)
            {
                work.Poller.Worked(1);
                inserts[i].Equivalent = equivalent[inserts[i].Component];
            }
            return true;
        }

        private static long RegionSampleWork(InkRaster left, InkRaster right,
            NormalizedBox region, int maxShift)
        {
            int width = Math.Max(PixelsWide(left, region), PixelsWide(right, region));
            int height = Math.Max(PixelsHigh(left, region), PixelsHigh(right, region));
            width = Math.Max(2, Math.Min(MaximumSampleWidth, width));
            height = Math.Max(2, Math.Min(MaximumSampleHeight, height));
            long positions = (long)(maxShift * 2 + 1) * (maxShift * 2 + 1);
            return (long)width * height * (positions + 1);
        }

        /// <summary>
        /// Equal обязан восстанавливать обе исходные стороны. Поэтому принимаются только
        /// непересекающиеся компоненты, занимающие непрерывный диапазон на каждой стороне.
        /// </summary>
        private static List<EquivalentGroup> EquivalentGroups(List<Candidate> deletes,
            List<Candidate> inserts, CancellationPoller poller)
        {
            var byComponent = new Dictionary<int, EquivalentGroup>();
            AddEquivalentCandidates(byComponent, deletes, true, poller);
            AddEquivalentCandidates(byComponent, inserts, false, poller);

            var result = new List<EquivalentGroup>();
            foreach (EquivalentGroup group in byComponent.Values)
            {
                poller.Worked(1);
                if (group.LeftCount == 0 || group.RightCount == 0 ||
                    group.LeftEnd - group.LeftStart + 1 != group.LeftCount ||
                    group.RightEnd - group.RightStart + 1 != group.RightCount ||
                    !RangeIsComponent(deletes, group.LeftStart, group.LeftEnd,
                        group.Component, poller) ||
                    !RangeIsComponent(inserts, group.RightStart, group.RightEnd,
                        group.Component, poller))
                    continue;
                result.Add(group);
            }
            result.Sort(delegate(EquivalentGroup a, EquivalentGroup b)
            {
                int byLeft = a.LeftStart.CompareTo(b.LeftStart);
                return byLeft != 0 ? byLeft : a.RightStart.CompareTo(b.RightStart);
            });

            int previousRight = -1;
            foreach (EquivalentGroup group in result)
            {
                poller.Worked(1);
                if (group.RightStart <= previousRight)
                    return new List<EquivalentGroup>();
                previousRight = group.RightEnd;
            }
            return result;
        }

        private static void AddEquivalentCandidates(
            IDictionary<int, EquivalentGroup> groups, IList<Candidate> candidates,
            bool leftSide, CancellationPoller poller)
        {
            foreach (Candidate candidate in candidates)
            {
                poller.Worked(1);
                if (!candidate.Equivalent || candidate.Component < 0)
                    continue;
                EquivalentGroup group;
                if (!groups.TryGetValue(candidate.Component, out group))
                {
                    group = new EquivalentGroup
                    {
                        Component = candidate.Component,
                        LeftStart = int.MaxValue,
                        RightStart = int.MaxValue,
                        LeftEnd = -1,
                        RightEnd = -1
                    };
                    groups.Add(candidate.Component, group);
                }
                if (leftSide)
                {
                    group.LeftStart = Math.Min(group.LeftStart, candidate.Order);
                    group.LeftEnd = Math.Max(group.LeftEnd, candidate.Order);
                    group.LeftCount++;
                }
                else
                {
                    group.RightStart = Math.Min(group.RightStart, candidate.Order);
                    group.RightEnd = Math.Max(group.RightEnd, candidate.Order);
                    group.RightCount++;
                }
            }
        }

        private static bool RangeIsComponent(IList<Candidate> candidates, int start,
            int end, int component, CancellationPoller poller)
        {
            if (start < 0 || end < start || end >= candidates.Count)
                return false;
            for (int i = start; i <= end; i++)
            {
                poller.Worked(1);
                if (!candidates[i].Equivalent || candidates[i].Component != component)
                    return false;
            }
            return true;
        }

        private static Candidate CandidateAt(List<Candidate> deletes,
            List<Candidate> inserts, int index)
        {
            return index < deletes.Count ? deletes[index] : inserts[index - deletes.Count];
        }

        private static NormalizedBox Union(NormalizedBox a, NormalizedBox b)
        {
            return new NormalizedBox
            {
                Left = Math.Min(a.Left, b.Left),
                Top = Math.Min(a.Top, b.Top),
                Right = Math.Max(a.Right, b.Right),
                Bottom = Math.Max(a.Bottom, b.Bottom)
            };
        }

        private static int RegistrationShift(InkRaster left, InkRaster right)
        {
            int width = Math.Min(left.Width, right.Width);
            int scaled = (int)Math.Ceiling(MaximumRegistrationShift *
                width / (double)TargetWidth);
            return Math.Max(1, Math.Min(MaximumRegistrationShift, scaled));
        }

        private static int Find(int[] parent, int item)
        {
            int root = item;
            while (parent[root] != root)
                root = parent[root];
            while (parent[item] != item)
            {
                int next = parent[item];
                parent[item] = root;
                item = next;
            }
            return root;
        }

        private static void Union(int[] parent, byte[] rank, int left, int right)
        {
            int a = Find(parent, left), b = Find(parent, right);
            if (a == b)
                return;
            if (rank[a] < rank[b])
                parent[a] = b;
            else
            {
                parent[b] = a;
                if (rank[a] == rank[b])
                    rank[a]++;
            }
        }

        private static int RenderWidth(PdfReviewPage left, PdfReviewPage right, long pixelCap)
        {
            double leftAspect = Aspect(left), rightAspect = Aspect(right);
            double aspect = Math.Max(leftAspect, rightAspect);
            if (aspect <= 0)
                return 0;
            int byPixels = (int)Math.Floor(Math.Sqrt(pixelCap / aspect));
            return Math.Min(TargetWidth, Math.Max(1, byPixels));
        }

        private static double Aspect(PdfReviewPage page)
        {
            if (page == null || page.ViewWidthPt <= 0 || page.ViewHeightPt <= 0)
                return 0;
            return page.ViewHeightPt / page.ViewWidthPt;
        }

        private static bool WithinLimit(Bitmap bitmap, long pixelCap)
        {
            return bitmap != null && bitmap.Width > 0 && bitmap.Height > 0 &&
                (long)bitmap.Width * bitmap.Height <= pixelCap + Math.Max(1, bitmap.Width);
        }

        private static bool TryNormalize(PdfReviewBox box, PdfReviewPage page,
            out NormalizedBox result)
        {
            result = new NormalizedBox();
            if (page == null || page.ViewWidthPt <= 0 || page.ViewHeightPt <= 0 ||
                box.Right <= box.Left || box.Top <= box.Bottom)
                return false;
            double padX = BoxPaddingPt / page.ViewWidthPt;
            double padY = BoxPaddingPt / page.ViewHeightPt;
            result.Left = Clamp(box.Left / page.ViewWidthPt - padX);
            result.Right = Clamp(box.Right / page.ViewWidthPt + padX);
            result.Top = Clamp((page.ViewHeightPt - box.Top) / page.ViewHeightPt - padY);
            result.Bottom = Clamp((page.ViewHeightPt - box.Bottom) / page.ViewHeightPt + padY);
            return result.Right > result.Left && result.Bottom > result.Top;
        }

        private static double Clamp(double value)
        {
            return value < 0 ? 0 : value > 1 ? 1 : value;
        }

        private static bool Related(NormalizedBox a, NormalizedBox b)
        {
            if (a.Right <= a.Left || a.Bottom <= a.Top ||
                b.Right <= b.Left || b.Bottom <= b.Top)
                return false;
            double vertical = Math.Min(a.Bottom, b.Bottom) - Math.Max(a.Top, b.Top);
            double minHeight = Math.Min(a.Bottom - a.Top, b.Bottom - b.Top);
            if (vertical < 0.3 * minHeight)
                return false;
            double horizontal = Math.Min(a.Right, b.Right) - Math.Max(a.Left, b.Left);
            if (horizontal > 0)
                return true;
            double gap = Math.Max(a.Left, b.Left) - Math.Min(a.Right, b.Right);
            return gap <= 0.25 * Math.Max(a.Right - a.Left, b.Right - b.Left);
        }

        private static bool RegionEquivalent(InkRaster left, InkRaster right,
            NormalizedBox region, int maxShift, CancellationPoller poller)
        {
            int width = Math.Max(PixelsWide(left, region), PixelsWide(right, region));
            int height = Math.Max(PixelsHigh(left, region), PixelsHigh(right, region));
            width = Math.Max(2, Math.Min(MaximumSampleWidth, width));
            height = Math.Max(2, Math.Min(MaximumSampleHeight, height));
            byte[] a = Sample(left, region, width, height, 0, 0, poller);
            byte[] da = Dilate(a, width, height, poller);

            // Ищем ближайшее допустимое совмещение первым. Радиус фиксирован и мал:
            // стоимость O((2r+1)^2), без неограниченной регистрации или полного page diff.
            for (int radius = 0; radius <= maxShift; radius++)
                for (int dy = -radius; dy <= radius; dy++)
                    for (int dx = -radius; dx <= radius; dx++)
                    {
                        if (Math.Max(Math.Abs(dx), Math.Abs(dy)) != radius)
                            continue;
                        poller.Worked(1);
                        byte[] b = Sample(right, region, width, height, dx, dy, poller);
                        if (EquivalentInk(a, da, b, width, height, poller))
                            return true;
                    }
            return false;
        }

        private static bool EquivalentInk(byte[] a, byte[] dilatedA, byte[] b,
            int width, int height, CancellationPoller poller)
        {
            byte[] dilatedB = Dilate(b, width, height, poller);
            long mass = 0, unmatched = 0;
            int pending = 0;
            for (int i = 0; i < a.Length; i++)
            {
                int av = a[i] < InkFloor ? 0 : a[i];
                int bv = b[i] < InkFloor ? 0 : b[i];
                int dav = dilatedA[i] < InkFloor ? 0 : dilatedA[i];
                int dbv = dilatedB[i] < InkFloor ? 0 : dilatedB[i];
                mass += av + bv;
                if (av > dbv) unmatched += av - dbv;
                if (bv > dav) unmatched += bv - dav;
                if (++pending == 1024)
                {
                    poller.Worked(pending);
                    pending = 0;
                }
            }
            poller.Worked(pending);
            if (mass == 0)
                return true; // одинаково невидимый/служебный текст не должен давать рамки
            return (double)unmatched / mass <= EquivalentInkDifference;
        }

        private static int PixelsWide(InkRaster raster, NormalizedBox box)
        {
            return Math.Max(1, (int)Math.Ceiling((box.Right - box.Left) * raster.Width));
        }

        private static int PixelsHigh(InkRaster raster, NormalizedBox box)
        {
            return Math.Max(1, (int)Math.Ceiling((box.Bottom - box.Top) * raster.Height));
        }

        private static byte[] Sample(InkRaster raster, NormalizedBox box, int width,
            int height, int offsetX, int offsetY, CancellationPoller poller)
        {
            var result = new byte[width * height];
            for (int y = 0; y < height; y++)
            {
                double ny = box.Top + (y + 0.5) * (box.Bottom - box.Top) / height;
                int sy = (int)Math.Floor(ny * raster.Height) + offsetY;
                if (sy >= 0 && sy < raster.Height)
                    for (int x = 0; x < width; x++)
                    {
                        double nx = box.Left + (x + 0.5) * (box.Right - box.Left) / width;
                        int sx = (int)Math.Floor(nx * raster.Width) + offsetX;
                        if (sx < 0 || sx >= raster.Width)
                            continue;
                        result[y * width + x] = raster.Darkness[sy * raster.Width + sx];
                    }
                poller.Worked(width);
            }
            return result;
        }

        /// <summary>Максимум в окрестности 3×3: допускает субпиксельный/однопиксельный AA-сдвиг.</summary>
        private static byte[] Dilate(byte[] source, int width, int height,
            CancellationPoller poller)
        {
            var result = new byte[source.Length];
            for (int y = 0; y < height; y++)
            {
                int y0 = Math.Max(0, y - 1), y1 = Math.Min(height - 1, y + 1);
                for (int x = 0; x < width; x++)
                {
                    int x0 = Math.Max(0, x - 1), x1 = Math.Min(width - 1, x + 1);
                    byte max = 0;
                    for (int yy = y0; yy <= y1; yy++)
                        for (int xx = x0; xx <= x1; xx++)
                            if (source[yy * width + xx] > max)
                                max = source[yy * width + xx];
                    result[y * width + x] = max;
                }
                poller.Worked(width);
            }
            return result;
        }

        private sealed class Candidate
        {
            public PdfReviewWord Word;
            public NormalizedBox Region;
            public int Order;
            public int Component;
            public bool Valid;
            public bool Equivalent;
        }

        private sealed class EquivalentGroup
        {
            public int Component;
            public int LeftStart;
            public int LeftEnd;
            public int LeftCount;
            public int RightStart;
            public int RightEnd;
            public int RightCount;
        }

        private struct NormalizedBox
        {
            public double Left;
            public double Top;
            public double Right;
            public double Bottom;
        }

        private sealed class InkRaster
        {
            public int Width;
            public int Height;
            public byte[] Darkness;

            public static InkRaster From(Bitmap source, CancellationPoller poller)
            {
                if (source == null || source.Width <= 0 || source.Height <= 0)
                    return null;
                try
                {
                    poller.Check();
                    using (var copy = new Bitmap(source.Width, source.Height,
                        PixelFormat.Format32bppArgb))
                    {
                        using (Graphics g = Graphics.FromImage(copy))
                        {
                            g.CompositingMode = CompositingMode.SourceCopy;
                            g.DrawImageUnscaled(source, 0, 0);
                        }
                        poller.Check();
                        Rectangle rect = new Rectangle(0, 0, copy.Width, copy.Height);
                        BitmapData data = copy.LockBits(rect, ImageLockMode.ReadOnly,
                            PixelFormat.Format32bppArgb);
                        try
                        {
                            int stride = Math.Abs(data.Stride);
                            var bytes = new byte[stride * copy.Height];
                            Marshal.Copy(data.Scan0, bytes, 0, bytes.Length);
                            var darkness = new byte[copy.Width * copy.Height];
                            for (int y = 0; y < copy.Height; y++)
                            {
                                int row = data.Stride >= 0 ? y * stride : (copy.Height - 1 - y) * stride;
                                for (int x = 0; x < copy.Width; x++)
                                {
                                    int p = row + x * 4;
                                    int blue = bytes[p], green = bytes[p + 1], red = bytes[p + 2];
                                    int alpha = bytes[p + 3];
                                    int luminance = (299 * red + 587 * green + 114 * blue + 500) / 1000;
                                    int value = (255 - luminance) * alpha / 255;
                                    darkness[y * copy.Width + x] = (byte)value;
                                }
                                poller.Worked(copy.Width);
                            }
                            return new InkRaster
                            {
                                Width = copy.Width,
                                Height = copy.Height,
                                Darkness = darkness
                            };
                        }
                        finally
                        {
                            copy.UnlockBits(data);
                        }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch
                {
                    return null;
                }
            }
        }
    }
}
