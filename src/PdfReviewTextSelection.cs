using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger
{
    internal enum PdfReviewClipboardResult
    {
        Success,
        Unavailable
    }

    internal interface IPdfReviewClipboardWriter
    {
        PdfReviewClipboardResult WriteUnicodeText(string text);
    }

    /// <summary>
    /// Единственная точка записи Review в системный clipboard. Clipboard требует STA;
    /// ограниченные повторы выполняет сама перегрузка WinForms SetDataObject.
    /// </summary>
    internal sealed class PdfReviewClipboardWriter : IPdfReviewClipboardWriter
    {
        internal const int RetryCount = 5;
        internal const int RetryDelayMilliseconds = 100;

        public PdfReviewClipboardResult WriteUnicodeText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return PdfReviewClipboardResult.Unavailable;
            try
            {
                Clipboard.SetDataObject(CreateDataObject(text), true, RetryCount,
                    RetryDelayMilliseconds);
                return PdfReviewClipboardResult.Success;
            }
            catch (ExternalException)
            {
                return PdfReviewClipboardResult.Unavailable;
            }
            catch (ThreadStateException)
            {
                return PdfReviewClipboardResult.Unavailable;
            }
        }

        internal static DataObject CreateDataObject(string text)
        {
            var data = new DataObject();
            data.SetData(DataFormats.UnicodeText, true, text ?? "");
            return data;
        }
    }

    internal sealed class PdfReviewSelectableWord
    {
        public readonly PdfReviewWord Word;
        public readonly int CanonicalIndex;

        public PdfReviewSelectableWord(PdfReviewWord word, int canonicalIndex)
        {
            Word = word;
            CanonicalIndex = canonicalIndex;
        }
    }

    internal sealed class PdfReviewCopyText
    {
        public readonly string Text;
        public readonly int WordCount;
        public readonly bool UsedFallbackSeparator;

        public PdfReviewCopyText(string text, int wordCount, bool usedFallbackSeparator)
        {
            Text = text ?? "";
            WordCount = wordCount;
            UsedFallbackSeparator = usedFallbackSeparator;
        }
    }

    /// <summary>
    /// Page-local выбор текста Review. Модель знает только опубликованные trusted-слова и
    /// буквально доказанные source-boundary; геометрия служит лишь hit-test и не создаёт текст.
    /// </summary>
    internal sealed class PdfReviewTextSelection
    {
        internal const float NearestDragDistance = 64f;

        private readonly PdfReviewPage _page;
        private readonly List<PdfReviewSelectableWord> _words =
            new List<PdfReviewSelectableWord>();
        private readonly ReadOnlyCollection<PdfReviewSelectableWord> _readOnlyWords;
        private readonly string[] _boundaryText;
        private readonly bool[] _boundaryProven;
        private int _anchor = -1;
        private int _focus = -1;

        public PdfReviewTextSelection(PdfReviewPage page)
        {
            _page = page;
            if (page != null)
            {
                for (int i = 0; i < page.Words.Count; i++)
                {
                    PdfReviewWord word = page.Words[i];
                    if (IsSelectable(word))
                        _words.Add(new PdfReviewSelectableWord(word, i));
                }
            }
            _readOnlyWords = _words.AsReadOnly();
            int boundaryCount = page == null ? 0 : Math.Max(0, page.Words.Count - 1);
            _boundaryText = new string[boundaryCount];
            _boundaryProven = new bool[boundaryCount];
            IndexBoundaries();
        }

        public ReadOnlyCollection<PdfReviewSelectableWord> Words
        {
            get { return _readOnlyWords; }
        }

        public int Count { get { return _words.Count; } }
        public bool HasSelection { get { return _anchor >= 0 && _focus >= 0; } }

        public int SelectedCount
        {
            get
            {
                if (!HasSelection) return 0;
                return Math.Abs(_focus - _anchor) + 1;
            }
        }

        public int SelectionStart
        {
            get { return HasSelection ? Math.Min(_anchor, _focus) : -1; }
        }

        public int SelectionEnd
        {
            get { return HasSelection ? Math.Max(_anchor, _focus) : -1; }
        }

        public int Anchor { get { return _anchor; } }
        public int Focus { get { return _focus; } }

        internal static bool IsSelectable(PdfReviewWord word)
        {
            return word != null && !string.IsNullOrEmpty(word.Text) &&
                word.SourceTrusted && word.SourceStart >= 0 &&
                word.SourceEnd >= word.SourceStart &&
                string.Equals(word.SourceText, word.Text, StringComparison.Ordinal) &&
                IsFinite(word.Box.Left) && IsFinite(word.Box.Bottom) &&
                IsFinite(word.Box.Right) && IsFinite(word.Box.Top) &&
                word.Box.Right > word.Box.Left && word.Box.Top > word.Box.Bottom;
        }

        private static bool IsFinite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private void IndexBoundaries()
        {
            if (_page == null || _boundaryText.Length == 0)
                return;
            var canonicalIndex = new Dictionary<PdfReviewWord, int>();
            for (int i = 0; i < _page.Words.Count; i++)
            {
                PdfReviewWord word = _page.Words[i];
                if (word != null && !canonicalIndex.ContainsKey(word))
                    canonicalIndex.Add(word, i);
            }

            var seen = new bool[_boundaryText.Length];
            foreach (PdfReviewWhitespaceEvidence evidence in _page.WhitespaceBoundaries)
            {
                if (evidence == null || evidence.Before == null || evidence.After == null ||
                    evidence.AtPageStart || evidence.AtPageEnd || evidence.Within != null ||
                    evidence.PageIndex != _page.PageIndex ||
                    !IsLiteralWhitespace(evidence.RawText))
                    continue;
                int before, after;
                if (!canonicalIndex.TryGetValue(evidence.Before, out before) ||
                    !canonicalIndex.TryGetValue(evidence.After, out after) || after != before + 1 ||
                    !IsSelectable(evidence.Before) || !IsSelectable(evidence.After) ||
                    evidence.Before.SourceEnd >= evidence.After.SourceStart)
                    continue;

                if (seen[before])
                {
                    // Две source-записи для одной границы неоднозначны, даже если сейчас
                    // содержат одинаковую строку: copy не выбирает одну из дубликатов.
                    _boundaryProven[before] = false;
                    _boundaryText[before] = null;
                    continue;
                }
                seen[before] = true;
                _boundaryProven[before] = true;
                _boundaryText[before] = evidence.RawText;
            }
        }

        private static bool IsLiteralWhitespace(string text)
        {
            if (text == null) return false;
            for (int i = 0; i < text.Length; i++)
                if (!char.IsWhiteSpace(text[i]))
                    return false;
            return true;
        }

        public bool Clear()
        {
            if (!HasSelection)
                return false;
            _anchor = _focus = -1;
            return true;
        }

        public bool Select(int selectableIndex, bool extend)
        {
            if (selectableIndex < 0 || selectableIndex >= _words.Count)
                return false;
            int oldAnchor = _anchor, oldFocus = _focus;
            if (!extend || !HasSelection)
                _anchor = selectableIndex;
            _focus = selectableIndex;
            return oldAnchor != _anchor || oldFocus != _focus;
        }

        public bool ExtendTo(int selectableIndex)
        {
            if (!HasSelection || selectableIndex < 0 || selectableIndex >= _words.Count)
                return false;
            if (_focus == selectableIndex)
                return false;
            _focus = selectableIndex;
            return true;
        }

        public bool SelectAll()
        {
            if (_words.Count == 0)
                return false;
            bool changed = _anchor != 0 || _focus != _words.Count - 1;
            _anchor = 0;
            _focus = _words.Count - 1;
            return changed;
        }

        internal bool SelectRange(int start, int end)
        {
            if (start < 0 || end < 0 || start >= _words.Count || end >= _words.Count)
                return false;
            bool changed = _anchor != start || _focus != end;
            _anchor = start;
            _focus = end;
            return changed;
        }

        public int HitTest(PointF point, Size surfaceSize)
        {
            if (_page == null || surfaceSize.Width <= 0 || surfaceSize.Height <= 0)
                return -1;
            for (int i = 0; i < _words.Count; i++)
            {
                RectangleF rect = WordRectangle(i, surfaceSize);
                if (!rect.IsEmpty && rect.Contains(point))
                    return i;
            }
            return -1;
        }

        public int HitTestNearest(PointF point, Size surfaceSize, float maxDistance)
        {
            int exact = HitTest(point, surfaceSize);
            if (exact >= 0)
                return exact;
            if (_page == null || surfaceSize.Width <= 0 || surfaceSize.Height <= 0 ||
                maxDistance < 0)
                return -1;

            double best = (double)maxDistance * maxDistance;
            int bestIndex = -1;
            for (int i = 0; i < _words.Count; i++)
            {
                RectangleF rect = WordRectangle(i, surfaceSize);
                if (rect.IsEmpty) continue;
                double dx = point.X < rect.Left ? rect.Left - point.X :
                    point.X > rect.Right ? point.X - rect.Right : 0;
                double dy = point.Y < rect.Top ? rect.Top - point.Y :
                    point.Y > rect.Bottom ? point.Y - rect.Bottom : 0;
                double distance = dx * dx + dy * dy;
                if (distance < best || (Math.Abs(distance - best) < 0.0001 &&
                    (bestIndex < 0 || i < bestIndex)))
                {
                    best = distance;
                    bestIndex = i;
                }
            }
            return bestIndex;
        }

        public RectangleF WordRectangle(int selectableIndex, Size surfaceSize)
        {
            if (_page == null || selectableIndex < 0 || selectableIndex >= _words.Count)
                return RectangleF.Empty;
            return PdfReviewGeometry.ToPixelRect(_words[selectableIndex].Word.Box,
                _page.ViewWidthPt, _page.ViewHeightPt,
                surfaceSize.Width, surfaceSize.Height);
        }

        public List<RectangleF> SelectedRectangles(Size surfaceSize)
        {
            var result = new List<RectangleF>();
            if (!HasSelection)
                return result;
            for (int i = SelectionStart; i <= SelectionEnd; i++)
            {
                RectangleF rect = WordRectangle(i, surfaceSize);
                if (!rect.IsEmpty)
                    result.Add(rect);
            }
            return result;
        }

        public PdfReviewCopyText BuildCopyText()
        {
            if (!HasSelection)
                return new PdfReviewCopyText("", 0, false);
            int start = SelectionStart, end = SelectionEnd;
            var text = new StringBuilder();
            bool fallback = false;
            for (int i = start; i <= end; i++)
            {
                PdfReviewSelectableWord current = _words[i];
                if (i > start)
                {
                    PdfReviewSelectableWord previous = _words[i - 1];
                    int boundary = previous.CanonicalIndex;
                    if (current.CanonicalIndex == boundary + 1 && boundary >= 0 &&
                        boundary < _boundaryProven.Length && _boundaryProven[boundary])
                    {
                        text.Append(_boundaryText[boundary]);
                    }
                    else
                    {
                        text.Append(' ');
                        fallback = true;
                    }
                }
                text.Append(current.Word.SourceText);
            }
            return new PdfReviewCopyText(text.ToString(), end - start + 1, fallback);
        }
    }
}
