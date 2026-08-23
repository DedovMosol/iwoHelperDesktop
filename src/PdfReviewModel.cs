using System.Collections.Generic;

namespace ExcelMerger
{
    internal enum PdfReviewPairStatus
    {
        Unchanged,
        Changed,
        LeftOnly,
        RightOnly
    }

    internal enum PdfReviewDiffKind
    {
        Equal,
        Delete,
        Insert
    }

    internal enum PdfReviewFailure
    {
        None,
        Unreadable,
        PasswordRequired,
        NoText,
        TooLarge,
        ChangedDuringRead
    }

    internal sealed class PdfReviewLimits
    {
        public long MaxFileBytes = 200L * 1024 * 1024;
        public int MaxPages = 500;
        public int MaxCharacters = 2000000;
        public int MaxTokens = 250000;
        public long MaxRenderPixels = 25000000;
        public int MaxDiffCells = 1000000;

        public static PdfReviewLimits Default() { return new PdfReviewLimits(); }
    }

    internal sealed class PdfReviewPage
    {
        public int PageIndex;
        public string Text;
        public string NormalizedText;
        public double WidthPt;
        public double HeightPt;
        public string Fingerprint;
    }

    internal sealed class PdfReviewDocument
    {
        public string Path;
        public readonly List<PdfReviewPage> Pages = new List<PdfReviewPage>();
        public int CharacterCount;
    }

    internal sealed class PdfReviewDiffOp
    {
        public PdfReviewDiffKind Kind;
        public string Text;
        public bool Collapsed;
        public int HiddenCharacters;
    }

    internal sealed class PdfReviewPagePair
    {
        public int LeftPageIndex = -1;
        public int RightPageIndex = -1;
        public PdfReviewPairStatus Status;
        public double Similarity;
        public readonly List<PdfReviewDiffOp> Operations = new List<PdfReviewDiffOp>();
    }

    internal sealed class PdfReviewStats
    {
        public int PagePairs;
        public int ChangedPages;
        public int LeftOnlyPages;
        public int RightOnlyPages;
        public int DeletedCharacters;
        public int InsertedCharacters;
        public int DeletedWords;
        public int InsertedWords;
        public int Replacements;
        public int ChangedPercent;
    }

    internal sealed class PdfReviewResult
    {
        public PdfReviewDocument Left;
        public PdfReviewDocument Right;
        public readonly List<PdfReviewPagePair> Pairs = new List<PdfReviewPagePair>();
        public PdfReviewStats Stats;
    }

    internal sealed class PdfReviewException : MergeException
    {
        public readonly PdfReviewFailure Reason;
        public readonly string FilePath;

        public PdfReviewException(PdfReviewFailure reason, string filePath, string message)
            : base(message)
        {
            Reason = reason;
            FilePath = filePath;
        }
    }
}
