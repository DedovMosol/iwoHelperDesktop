using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelMerger
{
    internal sealed class PdfReviewAutoScrollEventArgs : EventArgs
    {
        public readonly int DeltaX;
        public readonly int DeltaY;

        public PdfReviewAutoScrollEventArgs(int deltaX, int deltaY)
        {
            DeltaX = deltaX;
            DeltaY = deltaY;
        }
    }

    /// <summary>
    /// Изображение одной Review-страницы с отдельным read-only слоем выбора. Слой не
    /// распознаёт растр: диапазон и clipboard всегда строятся PdfReviewTextSelection.
    /// </summary>
    internal sealed class PdfReviewPageSurface : PictureBox
    {
        private const int AutoScrollMargin = 28;
        private const int AutoScrollStep = 22;
        private const int ChangeBarHitWidth = 24;
        private const int SelectionHitPadding = 6;
        private const int SelectionLinePadding = 10;

        private sealed class ChangeBarHit
        {
            public float Top;
            public float Bottom;
            public int Start;
            public int End;
        }

        private readonly ContextMenuStrip _selectionMenu;
        private readonly ToolStripMenuItem _copyItem;
        private readonly ToolStripMenuItem _selectAllItem;
        private readonly Timer _autoScrollTimer;
        private PdfReviewTextSelection _selection;
        private PdfReviewHighlight _highlight;
        private IPdfReviewClipboardWriter _clipboardWriter;
        private Panel _scrollViewport;
        private bool _dragging;
        private Size _rectangleSize;
        private readonly List<RectangleF> _wordRectangles = new List<RectangleF>();
        private readonly Dictionary<PdfReviewWord, int> _wordIndex =
            new Dictionary<PdfReviewWord, int>();
        private readonly List<ChangeBarHit> _changeBarHits = new List<ChangeBarHit>();
        private PdfReviewSurfaceFeedback _feedback;

        private enum PdfReviewSurfaceFeedback
        {
            None,
            Copied,
            CopiedWithFallback,
            ClipboardUnavailable
        }

        public event EventHandler SelectionStateChanged;
        public event EventHandler<PdfReviewAutoScrollEventArgs> AutoScrollRequested;

        public PdfReviewPageSurface()
        {
            SetStyle(ControlStyles.Selectable | ControlStyles.UserPaint |
                ControlStyles.OptimizedDoubleBuffer, true);
            TabStop = false;
            AccessibleRole = AccessibleRole.Document;
            Cursor = Cursors.Default;
            _clipboardWriter = new PdfReviewClipboardWriter();

            _selectionMenu = new ContextMenuStrip();
            _copyItem = new ToolStripMenuItem();
            _selectAllItem = new ToolStripMenuItem();
            _copyItem.Click += delegate { CopySelection(); };
            _selectAllItem.Click += delegate { SelectAllWords(); };
            _selectionMenu.Items.Add(_copyItem);
            _selectionMenu.Items.Add(_selectAllItem);
            _selectionMenu.Opening += delegate
            {
                _copyItem.Enabled = HasSelection;
                _selectAllItem.Enabled = HasSelectableText;
            };
            ContextMenuStrip = _selectionMenu;

            _autoScrollTimer = new Timer { Interval = 50 };
            _autoScrollTimer.Tick += delegate { AutoScrollTick(); };
            RefreshLocalization();
            Loc.Changed += RefreshLocalization;
        }

        internal Panel ScrollViewport
        {
            get { return _scrollViewport; }
            set { _scrollViewport = value; }
        }

        internal IPdfReviewClipboardWriter ClipboardWriter
        {
            get { return _clipboardWriter; }
            set { _clipboardWriter = value ?? new PdfReviewClipboardWriter(); }
        }

        internal PdfReviewTextSelection SelectionModel { get { return _selection; } }
        internal bool HasSelectableText { get { return _selection != null && _selection.Count > 0; } }
        internal bool HasSelection { get { return _selection != null && _selection.HasSelection; } }
        internal int SelectedWordCount { get { return _selection == null ? 0 : _selection.SelectedCount; } }

        internal string InteractionStatus
        {
            get
            {
                switch (_feedback)
                {
                    case PdfReviewSurfaceFeedback.Copied:
                        return Loc.T("review.selection.copied");
                    case PdfReviewSurfaceFeedback.CopiedWithFallback:
                        return Loc.T("review.selection.copiedFallback");
                    case PdfReviewSurfaceFeedback.ClipboardUnavailable:
                        return Loc.T("review.selection.clipboardUnavailable");
                }
                return HasSelection
                    ? string.Format(Loc.T("review.selection.count"), SelectedWordCount)
                    : "";
            }
        }

        internal void AttachPage(PdfReviewPage page)
        {
            AttachPage(page, null);
        }

        internal void AttachPage(PdfReviewPage page, PdfReviewHighlight highlight)
        {
            EndDrag();
            _selection = page == null ? null : new PdfReviewTextSelection(page);
            _highlight = highlight;
            _feedback = PdfReviewSurfaceFeedback.None;
            InvalidateRectangleCache();
            RebuildChangeBarHits();
            TabStop = HasSelectableText;
            Cursor = HasSelectableText ? Cursors.IBeam : Cursors.Default;
            NotifySelectionChanged();
        }

        internal void DetachPage()
        {
            EndDrag();
            _selection = null;
            _highlight = null;
            _changeBarHits.Clear();
            _feedback = PdfReviewSurfaceFeedback.None;
            InvalidateRectangleCache();
            TabStop = false;
            Cursor = Cursors.Default;
            NotifySelectionChanged();
        }

        internal bool ClearSelection()
        {
            EndDrag();
            if (_selection == null || !_selection.Clear())
                return false;
            _feedback = PdfReviewSurfaceFeedback.None;
            NotifySelectionChanged();
            return true;
        }

        internal bool SelectAllWords()
        {
            if (_selection == null || !_selection.SelectAll())
                return false;
            _feedback = PdfReviewSurfaceFeedback.None;
            NotifySelectionChanged();
            return true;
        }

        internal bool CopySelection()
        {
            if (_selection == null || !_selection.HasSelection)
                return false;
            PdfReviewCopyText copy = _selection.BuildCopyText();
            if (copy.WordCount <= 0 || string.IsNullOrEmpty(copy.Text))
                return false;
            PdfReviewClipboardResult result = _clipboardWriter.WriteUnicodeText(copy.Text);
            _feedback = result == PdfReviewClipboardResult.Success
                ? (copy.UsedFallbackSeparator
                    ? PdfReviewSurfaceFeedback.CopiedWithFallback
                    : PdfReviewSurfaceFeedback.Copied)
                : PdfReviewSurfaceFeedback.ClipboardUnavailable;
            NotifySelectionChanged();
            return result == PdfReviewClipboardResult.Success;
        }

        private void RefreshLocalization()
        {
            if (IsDisposed)
                return;
            _copyItem.Text = Loc.T("common.copy");
            _selectAllItem.Text = Loc.T("review.selection.selectAll");
            NotifySelectionChanged();
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            InvalidateRectangleCache();
            RebuildChangeBarHits();
            base.OnSizeChanged(e);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && HasSelectableText)
                Focus();
            if (e.Button == MouseButtons.Left && HasSelectableText)
            {
                Focus();
                int bar = HitTestChangeBar(e.Location);
                if (bar >= 0)
                {
                    EndDrag();
                    if (_selection.SelectRange(_changeBarHits[bar].Start,
                        _changeBarHits[bar].End))
                    {
                        _feedback = PdfReviewSurfaceFeedback.None;
                        NotifySelectionChanged();
                    }
                    return;
                }
                int hit = HitTestWord(e.Location, true);
                if (hit < 0)
                {
                    ClearSelection();
                }
                else
                {
                    bool extend = (ModifierKeys & Keys.Shift) != 0;
                    _selection.Select(hit, extend);
                    _feedback = PdfReviewSurfaceFeedback.None;
                    _dragging = true;
                    Capture = true;
                    _autoScrollTimer.Start();
                    NotifySelectionChanged();
                }
            }
            base.OnMouseDown(e);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (_dragging && (e.Button & MouseButtons.Left) != 0 && _selection != null)
            {
                int hit = HitTestWord(e.Location, true);
                if (hit >= 0 && _selection.ExtendTo(hit))
                {
                    _feedback = PdfReviewSurfaceFeedback.None;
                    NotifySelectionChanged();
                }
            }
            base.OnMouseMove(e);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
                EndDrag();
            base.OnMouseUp(e);
        }

        protected override void OnMouseCaptureChanged(EventArgs e)
        {
            if (!Capture)
            {
                _dragging = false;
                _autoScrollTimer.Stop();
            }
            base.OnMouseCaptureChanged(e);
        }

        protected override void OnGotFocus(EventArgs e)
        {
            Invalidate();
            base.OnGotFocus(e);
        }

        protected override void OnLostFocus(EventArgs e)
        {
            Invalidate();
            base.OnLostFocus(e);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (HasSelectableText && keyData == (Keys.Control | Keys.A))
            {
                SelectAllWords();
                return true;
            }
            if (HasSelectableText && keyData == (Keys.Control | Keys.C))
            {
                CopySelection();
                return true;
            }
            if (keyData == Keys.Escape && HasSelection)
            {
                ClearSelection();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
            if (_selection != null && _selection.HasSelection)
            {
                EnsureRectangleCache();
                int start = _selection.SelectionStart;
                int end = _selection.SelectionEnd;
                bool highContrast = SystemInformation.HighContrast;
                Color outline = SystemColors.Highlight;
                using (var pen = new Pen(outline, highContrast ? 3f : 2f))
                {
                    pen.DashStyle = highContrast ? DashStyle.Dash : DashStyle.Solid;
                    if (!highContrast)
                    {
                        using (var brush = new SolidBrush(Color.FromArgb(62, outline)))
                            for (int i = start; i <= end && i < _wordRectangles.Count; i++)
                                if (!_wordRectangles[i].IsEmpty)
                                    pe.Graphics.FillRectangle(brush, _wordRectangles[i]);
                    }
                    for (int i = start; i <= end && i < _wordRectangles.Count; i++)
                        if (!_wordRectangles[i].IsEmpty)
                            pe.Graphics.DrawRectangle(pen, Rectangle.Round(_wordRectangles[i]));
                }
            }
            if (Focused && HasSelectableText)
                ControlPaint.DrawFocusRectangle(pe.Graphics,
                    new Rectangle(2, 2, Math.Max(0, ClientSize.Width - 5),
                        Math.Max(0, ClientSize.Height - 5)),
                    SystemColors.HighlightText, SystemColors.Highlight);
        }

        private bool IsChangeBarHitX(int x)
        {
            if (_highlight == null)
                return false;
            bool outer = _highlight.ChangeBarSide == PdfReviewChangeBarSide.Left
                ? x <= ChangeBarHitWidth
                : _highlight.ChangeBarSide == PdfReviewChangeBarSide.Right &&
                    x >= ClientSize.Width - ChangeBarHitWidth;
            if (!outer)
                return false;
            EnsureRectangleCache();
            for (int i = 0; i < _wordRectangles.Count; i++)
            {
                RectangleF rect = _wordRectangles[i];
                if (!rect.IsEmpty && x >= rect.Left && x <= rect.Right)
                    return false;
            }
            return true;
        }

        private int HitTestChangeBar(Point point)
        {
            if (_highlight == null || _highlight.ChangeBarSide == PdfReviewChangeBarSide.None ||
                _changeBarHits.Count == 0 || !IsChangeBarHitX(point.X))
                return -1;
            for (int i = 0; i < _changeBarHits.Count; i++)
            {
                ChangeBarHit hit = _changeBarHits[i];
                if (point.Y >= hit.Top - 4f && point.Y <= hit.Bottom + 4f)
                    return i;
            }
            return -1;
        }

        private void RebuildChangeBarHits()
        {
            _changeBarHits.Clear();
            if (_selection == null || _highlight == null || _highlight.Words.Count == 0)
                return;
            EnsureRectangleCache();
            var changed = new List<Tuple<int, RectangleF>>();
            foreach (PdfReviewWord changedWord in _highlight.Words)
            {
                if (changedWord == null)
                    continue;
                int index;
                if (!_wordIndex.TryGetValue(changedWord, out index))
                    continue;
                RectangleF rect = index < _wordRectangles.Count
                    ? _wordRectangles[index] : RectangleF.Empty;
                if (!rect.IsEmpty)
                    changed.Add(Tuple.Create(index, rect));
            }
            changed.Sort(delegate(Tuple<int, RectangleF> left, Tuple<int, RectangleF> right)
            {
                return left.Item1.CompareTo(right.Item1);
            });
            for (int i = 0; i < changed.Count; i++)
            {
                Tuple<int, RectangleF> item = changed[i];
                if (_changeBarHits.Count > 0)
                {
                    ChangeBarHit previous = _changeBarHits[_changeBarHits.Count - 1];
                    if (SameTextLine(previous.Top, previous.Bottom, item.Item2))
                    {
                        previous.End = item.Item1;
                        previous.Top = Math.Min(previous.Top, item.Item2.Top);
                        previous.Bottom = Math.Max(previous.Bottom, item.Item2.Bottom);
                        continue;
                    }
                }
                _changeBarHits.Add(new ChangeBarHit
                {
                    Top = item.Item2.Top,
                    Bottom = item.Item2.Bottom,
                    Start = item.Item1,
                    End = item.Item1
                });
            }
        }

        private static bool SameTextLine(float top, float bottom, RectangleF next)
        {
            float height = Math.Min(Math.Max(0f, bottom - top), next.Height);
            if (height <= 0f)
                return false;
            float overlap = Math.Min(bottom, next.Bottom) - Math.Max(top, next.Top);
            return overlap >= height * 0.35f ||
                Math.Abs((top + bottom) / 2f - (next.Top + next.Bottom) / 2f) <=
                height * 0.45f;
        }

        private int HitTestWord(Point point, bool nearest)
        {
            EnsureRectangleCache();
            if (IsChangeBarHitX(point.X))
                return -1;
            int bestExact = -1;
            double bestExactDistance = double.MaxValue;
            for (int index = 0; index < _wordRectangles.Count; index++)
            {
                RectangleF rect = _wordRectangles[index];
                if (rect.IsEmpty)
                    continue;
                RectangleF hit = rect;
                hit.Inflate(SelectionHitPadding, SelectionLinePadding);
                if (!hit.Contains(point))
                    continue;
                double distance = RectangleDistanceSquared(rect, point);
                if (distance < bestExactDistance ||
                    (Math.Abs(distance - bestExactDistance) < 0.0001 &&
                     (bestExact < 0 || index < bestExact)))
                {
                    bestExactDistance = distance;
                    bestExact = index;
                }
            }
            if (bestExact >= 0 || !nearest)
                return bestExact;

            double best = (double)PdfReviewTextSelection.NearestDragDistance *
                PdfReviewTextSelection.NearestDragDistance;
            int bestIndex = -1;
            for (int index = 0; index < _wordRectangles.Count; index++)
            {
                RectangleF rect = _wordRectangles[index];
                if (rect.IsEmpty)
                    continue;
                double distance = RectangleDistanceSquared(rect, point);
                if (distance < best || (Math.Abs(distance - best) < 0.0001 &&
                    (bestIndex < 0 || index < bestIndex)))
                {
                    best = distance;
                    bestIndex = index;
                }
            }
            return bestIndex;
        }

        private static double RectangleDistanceSquared(RectangleF rect, Point point)
        {
            double dx = point.X < rect.Left ? rect.Left - point.X :
                point.X > rect.Right ? point.X - rect.Right : 0;
            double dy = point.Y < rect.Top ? rect.Top - point.Y :
                point.Y > rect.Bottom ? point.Y - rect.Bottom : 0;
            return dx * dx + dy * dy;
        }

        private void EnsureRectangleCache()
        {
            if (_selection == null)
            {
                ClearRectangleCache();
                return;
            }
            if (_rectangleSize == ClientSize && _wordRectangles.Count == _selection.Count)
                return;
            ClearRectangleCache();
            for (int i = 0; i < _selection.Count; i++)
            {
                RectangleF rect = _selection.WordRectangle(i, ClientSize);
                _wordRectangles.Add(rect);
                PdfReviewSelectableWord selectable = _selection.Words[i];
                if (selectable != null && selectable.Word != null &&
                    !_wordIndex.ContainsKey(selectable.Word))
                    _wordIndex.Add(selectable.Word, i);
                if (rect.IsEmpty)
                    continue;
            }
            _rectangleSize = ClientSize;
        }

        private void ClearRectangleCache()
        {
            _rectangleSize = Size.Empty;
            _wordRectangles.Clear();
            _wordIndex.Clear();
        }

        private void InvalidateRectangleCache()
        {
            ClearRectangleCache();
            Invalidate();
        }

        private void AutoScrollTick()
        {
            if (!_dragging || !Capture || _selection == null || _scrollViewport == null ||
                _scrollViewport.IsDisposed || !_scrollViewport.IsHandleCreated)
            {
                _autoScrollTimer.Stop();
                return;
            }
            Point cursor = _scrollViewport.PointToClient(System.Windows.Forms.Cursor.Position);
            int dx = cursor.X < AutoScrollMargin ? -AutoScrollStep :
                cursor.X >= _scrollViewport.ClientSize.Width - AutoScrollMargin
                    ? AutoScrollStep : 0;
            int dy = cursor.Y < AutoScrollMargin ? -AutoScrollStep :
                cursor.Y >= _scrollViewport.ClientSize.Height - AutoScrollMargin
                    ? AutoScrollStep : 0;
            if (dx != 0 || dy != 0)
            {
                EventHandler<PdfReviewAutoScrollEventArgs> handler = AutoScrollRequested;
                if (handler != null)
                    handler(this, new PdfReviewAutoScrollEventArgs(dx, dy));
                Point local = PointToClient(System.Windows.Forms.Cursor.Position);
                int hit = HitTestWord(local, true);
                if (hit >= 0 && _selection.ExtendTo(hit))
                {
                    _feedback = PdfReviewSurfaceFeedback.None;
                    NotifySelectionChanged();
                }
            }
        }

        private void EndDrag()
        {
            _dragging = false;
            _autoScrollTimer.Stop();
            if (Capture)
                Capture = false;
        }

        private void NotifySelectionChanged()
        {
            Invalidate();
            if (IsHandleCreated)
            {
                AccessibilityNotifyClients(AccessibleEvents.Selection, -1);
                AccessibilityNotifyClients(AccessibleEvents.DescriptionChange, -1);
            }
            EventHandler handler = SelectionStateChanged;
            if (handler != null)
                handler(this, EventArgs.Empty);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Loc.Changed -= RefreshLocalization;
                EndDrag();
                _autoScrollTimer.Dispose();
                _selectionMenu.Dispose();
                _selection = null;
                _highlight = null;
                _changeBarHits.Clear();
            }
            base.Dispose(disposing);
        }
    }
}
