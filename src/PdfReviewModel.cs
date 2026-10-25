using System;
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
        // Жёсткий верхний предел защищает все DP-матрицы даже при ошибочно завышенном
        // экземпляре limits. Настраиваемый MaxDiffCells может только ужесточить его.
        internal const int AbsoluteMaxDiffCells = 1000000;

        public long MaxFileBytes = 200L * 1024 * 1024;
        public int MaxPages = 500;
        public int MaxCharacters = 2000000;
        public long MaxRenderPixels = 25000000;
        public int MaxDiffCells = AbsoluteMaxDiffCells;
        // Детерминированный потолок семантической работы. В отличие от timeout он даёт
        // одинаковый результат на быстрых и медленных машинах; при исчерпании diff
        // консервативно оставляет непроверенный остаток Delete/Insert. Отдельные стадии и
        // выравнивание страниц внутри Compare получают его как есть, словесные стадии
        // Compare — масштабированным по длине (DiffWorkFor).
        public long MaxDiffWork = 20000000;

        // MaxDiffWork рассчитан на пару документов до DiffWorkReferenceWords слов вместе.
        // Словесным стадиям достаётся 7/8 масштаба (1/8 гарантирована пробелам): 56 млн
        // при 40 тыс. слов на сторону и 112 млн при 80 тыс. Замер точного diff при 5%
        // случайных замен: ~24 и ~79 млн (бисекция стоит ~D² при равномерных правках).
        // Одна бисекция берёт не больше половины остатка, а её неудача делит диапазон
        // якорями, поэтому нехватка бюджета не превращает документ в «изменено 100%».
        // Потолок ×8 (160 млн) на 2 млн знаков — порядка 10-20 с на x64, а не единицы секунд.
        internal const long DiffWorkReferenceWords = 25000;
        internal const long MaxDiffWorkScale = 8;

        /// <summary>
        /// Бюджет семантической работы Compare для пары документов из words слов. Число
        /// реальных правок растёт с длиной документа, а Myers тратит O((N+M)·D), поэтому
        /// фиксированный бюджет кончался бы на длинных документах при обычной плотности
        /// правок. Масштаб линеен по N+M, но не больше MaxDiffWorkScale: работа ограничивает
        /// время, а память diff остаётся O(N+M) плюс ограниченный лист матрицы, так что
        /// потолок одинаково безопасен для x86 и x64. Целочисленная арифметика оставляет
        /// результат детерминированным на любой разрядности.
        /// </summary>
        internal long DiffWorkFor(long words)
        {
            long baseWork = Math.Max(0L, MaxDiffWork);
            if (words <= DiffWorkReferenceWords || baseWork == 0)
                return baseWork;
            long scaledWords = Math.Min(words, DiffWorkReferenceWords * MaxDiffWorkScale);
            // floor(base·s/R) = (base/R)·s + (base%R)·s/R: остаток меньше R, и его
            // произведение на s ≤ 8R всегда помещается; насыщается только сам результат.
            long whole = baseWork / DiffWorkReferenceWords;
            long fraction = baseWork % DiffWorkReferenceWords * scaledWords /
                DiffWorkReferenceWords;
            if (whole > (long.MaxValue - fraction) / scaledWords)
                return long.MaxValue;
            return whole * scaledWords + fraction;
        }

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

    /// <summary>Слово документа: текст, ключ, физический владелец и рамка для подсветки.</summary>
    internal sealed class PdfReviewWord
    {
        public string Text;
        public string Key;
        public int PageIndex = -1;
        public PdfReviewBox Box;

        // Компактная provenance текстового слоя. Диапазон относится к декодированным
        // source-unit текущей страницы; PdfPig-объекты за пределы extraction не выходят.
        internal int SourceStart = -1;
        internal int SourceEnd = -1;
        internal string SourceText;
        internal bool SourceTrusted;
        internal int BlockId = -1;
        // Слово, склеенное из переноса через строку: рамка каждой его строки отдельно. Box —
        // их объединение (оно нужно сравнению как одна точка привязки), но подсветка, рейка и
        // каретка по нему накрывали неизменённый текст обеих строк. null — слово в одной рамке.
        internal List<PdfReviewBox> Fragments;

        /// <summary>Рамки для показа: строки переноса — каждая своей, иначе — сама рамка слова.</summary>
        internal IList<PdfReviewBox> DisplayBoxes
        {
            get { return Fragments ?? new List<PdfReviewBox> { Box }; }
        }
    }

    internal enum PdfReviewMatchKind
    {
        None,
        Exact,
        ReconciledOrder,
        // Контейнер для цикла порядка, в котором соседствуют Exact и
        // ReconciledOrder. Фактическая provenance остаётся у каждого Matches[i].
        MixedOrder,
        SplitJoin,
        RasterEquivalent
    }

    /// <summary>Явная связь двух фактических объектов, а не предположение «Equal общий».</summary>
    internal sealed class PdfReviewWordMatch
    {
        public PdfReviewWord Left;
        public PdfReviewWord Right;
        public PdfReviewMatchKind Kind;
    }

    /// <summary>
    /// Операция ворд-диффа. Списки каждой стороны хранят документный порядок независимо:
    /// это позволяет представить локальную перестановку, не потеряв ни одну последовательность.
    /// </summary>
    internal sealed class PdfReviewWordOp
    {
        public PdfReviewDiffKind Kind;
        public PdfReviewMatchKind MatchKind;
        public readonly List<PdfReviewWord> LeftWords = new List<PdfReviewWord>();
        public readonly List<PdfReviewWord> RightWords = new List<PdfReviewWord>();
        public readonly List<PdfReviewWordMatch> Matches = new List<PdfReviewWordMatch>();
        // У SplitJoin одна доказанная граница на каждой стороне: обычная между
        // двумя словами и пустая внутри joined-слова. Такие операции не склеиваются.
        public PdfReviewWhitespaceEvidence SplitJoinLeftBoundary;
        public PdfReviewWhitespaceEvidence SplitJoinRightBoundary;

        // Узкая совместимость старых read-only обходов: Delete/Equal исторически читали
        // левую сторону, Insert — правую. Новая семантика всегда использует явные списки.
        public List<PdfReviewWord> Words
        {
            get { return Kind == PdfReviewDiffKind.Insert ? RightWords : LeftWords; }
        }
    }

    internal enum PdfReviewWhitespaceAtomKind
    {
        Space,
        NoBreakSpace,
        Tab,
        LineBreak,
        Other
    }

    internal sealed class PdfReviewWhitespaceAtom
    {
        public PdfReviewWhitespaceAtomKind Kind;
        public string RawText;
    }

    /// <summary>
    /// Положительное свидетельство текстового слоя о границе. Пустая LogicalText тоже
    /// содержательна: соседство source-unit доказывает отсутствие литерального пробела.
    /// </summary>
    internal sealed class PdfReviewWhitespaceEvidence
    {
        public int PageIndex = -1;
        public PdfReviewWord Before;
        public PdfReviewWord After;
        // Для доказанной границы внутри единственного joined-слова. TextOffset —
        // UTF-16 offset в точном SourceText; обычная межсловная граница его не задаёт.
        public PdfReviewWord Within;
        public int TextOffset = -1;
        public string RawText;
        public string LogicalText;
        public bool AtPageStart;
        public bool AtPageEnd;
        public PdfReviewBox MarkerBox;
    }

    internal sealed class PdfReviewWhitespaceChange
    {
        public PdfReviewWhitespaceEvidence Left;
        public PdfReviewWhitespaceEvidence Right;
        public readonly List<PdfReviewWhitespaceAtom> DeletedAtoms =
            new List<PdfReviewWhitespaceAtom>();
        public readonly List<PdfReviewWhitespaceAtom> InsertedAtoms =
            new List<PdfReviewWhitespaceAtom>();
    }

    internal sealed class PdfReviewWhitespaceMarker
    {
        public PdfReviewBox Box;
        public string Text;
        public string AccessibleDescription;
        public PdfReviewHighlightStyle Style;
    }

    internal enum PdfReviewChangeBarSide
    {
        None,
        Left,
        Right
    }

    internal enum PdfReviewHighlightStyle
    {
        Removed,
        Added
    }

    /// <summary>
    /// Визуальная проекция semantic word-box в пространстве страницы (пункты, Y вверх).
    /// Normal mode использует Word-подобный фон с сохранением тёмного PDF ink; high
    /// contrast — системные контуры/pattern. WhitespaceMarkers остаются отдельным слоем.
    /// </summary>
    internal sealed class PdfReviewHighlight
    {
        public readonly List<PdfReviewBox> Boxes = new List<PdfReviewBox>();
        // UI-only ownership list paired with Boxes: it lets a clicked change bar select
        // the trusted words that produced that bar without re-diffing the page.
        public readonly List<PdfReviewWord> Words = new List<PdfReviewWord>();
        public readonly List<PdfReviewWhitespaceMarker> WhitespaceMarkers =
            new List<PdfReviewWhitespaceMarker>();
        public double ViewWidthPt;
        public double ViewHeightPt;
        public Color Color;
        public Color EdgeColor;
        public PdfReviewChangeBarSide ChangeBarSide;
        public PdfReviewHighlightStyle Style;
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
        public readonly List<PdfReviewWhitespaceEvidence> WhitespaceBoundaries =
            new List<PdfReviewWhitespaceEvidence>();
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

    /// <summary>Строка сопоставления физических страниц в viewer; семантического diff не хранит.</summary>
    internal sealed class PdfReviewPagePair
    {
        public int LeftPageIndex = -1;
        public int RightPageIndex = -1;
        public PdfReviewPairStatus Status;
        public double Similarity;
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
        public int WhitespaceChanges;
        public int DeletedWhitespaceAtoms;
        public int InsertedWhitespaceAtoms;
        public int ChangedPercent;
    }

    /// <summary>Стадии Compare, которые могут не завершиться из-за бюджета работы.</summary>
    [Flags]
    internal enum PdfReviewStage
    {
        None = 0,
        PageAlignment = 1,
        WordDiff = 2,
        OrderReconciliation = 4,
        SplitJoin = 8,
        Whitespace = 16
    }

    internal sealed class PdfReviewResult
    {
        public PdfReviewDocument Left;
        public PdfReviewDocument Right;
        public readonly List<PdfReviewPagePair> Pairs = new List<PdfReviewPagePair>();

        /// <summary>
        /// Бюджет работы исчерпан, и часть документа сравнена грубо (сырые удаление/вставка,
        /// пропущенные уточняющие стадии). Окно обязано сказать об этом: иначе грубые
        /// кандидаты выглядели бы настоящими правками, а «пробелов не менялось» — проверкой.
        /// </summary>
        public bool WorkExhausted { get; internal set; }

        /// <summary>
        /// Какие стадии не завершились (пропущены или прерваны бюджетом). Непустое значение
        /// всегда сопровождается WorkExhausted; детали нужны окну для точного пояснения.
        /// </summary>
        public PdfReviewStage IncompleteStages { get; internal set; }

        // Семантические и производные коллекции публикуются готовыми snapshot-ссылками.
        // Это намеренно не readonly-поля: Clear()+AddRange()/повторное заполнение словаря
        // теряет прежний корректный результат, если подготовка отменена или падает по памяти.
        // Наружу ссылки доступны только для чтения; заменить их может лишь транзакционный код.
        private List<PdfReviewWordOp> _operations = new List<PdfReviewWordOp>();
        private List<PdfReviewWhitespaceChange> _whitespaceChanges =
            new List<PdfReviewWhitespaceChange>();
        private Dictionary<int, List<PdfReviewWord>> _deletedWordsByPage =
            new Dictionary<int, List<PdfReviewWord>>();
        private Dictionary<int, List<PdfReviewWord>> _insertedWordsByPage =
            new Dictionary<int, List<PdfReviewWord>>();
        private Dictionary<int, List<PdfReviewWhitespaceMarker>> _deletedWhitespaceByPage =
            new Dictionary<int, List<PdfReviewWhitespaceMarker>>();
        private Dictionary<int, List<PdfReviewWhitespaceMarker>> _insertedWhitespaceByPage =
            new Dictionary<int, List<PdfReviewWhitespaceMarker>>();

        public List<PdfReviewWordOp> Operations { get { return _operations; } }
        public List<PdfReviewWhitespaceChange> WhitespaceChanges
        {
            get { return _whitespaceChanges; }
        }
        public Dictionary<int, List<PdfReviewWord>> DeletedWordsByPage
        {
            get { return _deletedWordsByPage; }
        }
        public Dictionary<int, List<PdfReviewWord>> InsertedWordsByPage
        {
            get { return _insertedWordsByPage; }
        }
        public Dictionary<int, List<PdfReviewWhitespaceMarker>> DeletedWhitespaceByPage
        {
            get { return _deletedWhitespaceByPage; }
        }
        public Dictionary<int, List<PdfReviewWhitespaceMarker>> InsertedWhitespaceByPage
        {
            get { return _insertedWhitespaceByPage; }
        }
        public PdfReviewStats Stats;

        internal void ReplaceOperations(List<PdfReviewWordOp> operations)
        {
            if (operations == null) throw new ArgumentNullException("operations");
            _operations = operations;
        }

        internal void ReplaceWhitespaceChanges(
            List<PdfReviewWhitespaceChange> whitespaceChanges)
        {
            if (whitespaceChanges == null)
                throw new ArgumentNullException("whitespaceChanges");
            _whitespaceChanges = whitespaceChanges;
        }

        internal void PublishState(List<PdfReviewWordOp> operations,
            List<PdfReviewWhitespaceChange> whitespaceChanges,
            Dictionary<int, List<PdfReviewWord>> deletedWordsByPage,
            Dictionary<int, List<PdfReviewWord>> insertedWordsByPage,
            Dictionary<int, List<PdfReviewWhitespaceMarker>> deletedWhitespaceByPage,
            Dictionary<int, List<PdfReviewWhitespaceMarker>> insertedWhitespaceByPage,
            PdfReviewStats stats)
        {
            // Вся валидация идёт до первой записи. После неё только присваивания ссылок:
            // commit не выделяет память и не вызывает пользовательский cancellation callback.
            if (operations == null) throw new ArgumentNullException("operations");
            if (whitespaceChanges == null)
                throw new ArgumentNullException("whitespaceChanges");
            if (deletedWordsByPage == null)
                throw new ArgumentNullException("deletedWordsByPage");
            if (insertedWordsByPage == null)
                throw new ArgumentNullException("insertedWordsByPage");
            if (deletedWhitespaceByPage == null)
                throw new ArgumentNullException("deletedWhitespaceByPage");
            if (insertedWhitespaceByPage == null)
                throw new ArgumentNullException("insertedWhitespaceByPage");
            if (stats == null) throw new ArgumentNullException("stats");

            _operations = operations;
            _whitespaceChanges = whitespaceChanges;
            _deletedWordsByPage = deletedWordsByPage;
            _insertedWordsByPage = insertedWordsByPage;
            _deletedWhitespaceByPage = deletedWhitespaceByPage;
            _insertedWhitespaceByPage = insertedWhitespaceByPage;
            Stats = stats;
        }
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
