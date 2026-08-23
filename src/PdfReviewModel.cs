using System.Collections.Generic;
using System.Drawing;

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
        public long MaxRenderPixels = 25000000;
        public int MaxDiffCells = 1000000;

        public static PdfReviewLimits Default() { return new PdfReviewLimits(); }
    }

    /// <summary>Рамка слова в пространстве отображения страницы: пункты, X вправо, Y вверх.</summary>
    internal struct PdfReviewBox
    {
        public double Left;
        public double Bottom;
        public double Right;
        public double Top;
    }

    /// <summary>Слово страницы для сравнения: текст, ключ сравнения и рамка для подсветки.</summary>
    internal sealed class PdfReviewWord
    {
        public string Text;
        public string Key;
        public PdfReviewBox Box;
    }

    /// <summary>Операция ворд-диффа: серия слов одного вида в исходном порядке.</summary>
    internal sealed class PdfReviewWordOp
    {
        public PdfReviewDiffKind Kind;
        public readonly List<PdfReviewWord> Words = new List<PdfReviewWord>();
    }

    /// <summary>
    /// Подсветка поверх отрендеренной страницы: рамки слов в пространстве страницы
    /// (пункты, Y вверх) и цвет. Пустые рамки — страница без подсветки.
    /// </summary>
    internal sealed class PdfReviewHighlight
    {
        public readonly List<PdfReviewBox> Boxes = new List<PdfReviewBox>();
        public double ViewWidthPt;
        public double ViewHeightPt;
        public Color Color;
    }

    internal sealed class PdfReviewPage
    {
        public int PageIndex;
        public string Text;
        public string NormalizedText;
        public double WidthPt;
        public double HeightPt;
        public string Fingerprint;
        public readonly List<PdfReviewWord> Words = new List<PdfReviewWord>();
        // Размеры страницы в пространстве отображения (с учётом её собственного поворота) —
        // на них масштабируются рамки подсветки при наложении на отрендеренную страницу.
        public double ViewWidthPt;
        public double ViewHeightPt;
    }

    internal sealed class PdfReviewDocument
    {
        public string Path;
        public readonly List<PdfReviewPage> Pages = new List<PdfReviewPage>();
        public int CharacterCount;
        public int WordCount;
    }

    internal sealed class PdfReviewPagePair
    {
        public int LeftPageIndex = -1;
        public int RightPageIndex = -1;
        public PdfReviewPairStatus Status;
        public double Similarity;
        public readonly List<PdfReviewWordOp> Operations = new List<PdfReviewWordOp>();
    }

    internal sealed class PdfReviewStats
    {
        public int PagePairs;
        public int ChangedPages;
        public int LeftOnlyPages;
        public int RightOnlyPages;
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
