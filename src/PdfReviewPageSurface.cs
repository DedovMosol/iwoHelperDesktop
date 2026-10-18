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
        private const float ChangeBarHitTolerance = 2f;
        private const int SelectionHitPadding = 6;
        private const int SelectionLinePadding = 10;

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
        // Нарисованные полосы изменений в пикселях растра и размер этого растра: клик
        // масштабируется к ним, поэтому зона попадания там же, где рейка, при любом zoom.
        private readonly List<PdfReviewChangeBar> _changeBars = new List<PdfReviewChangeBar>();
        private Size _changeBarPixels;
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
            AttachPage(page, highlight, null);
        }

        /// <summary>
        /// bars — полосы, фактически нарисованные на Image (пиксели растра). Без них
        /// (растр без композиции) полосы раскладываются той же LayoutChangeBars по размеру Image.
        /// </summary>
        internal void AttachPage(PdfReviewPage page, PdfReviewHighlight highlight,
            IList<PdfReviewChangeBar> bars)
        {
            EndDrag();
            _selection = page == null ? null : new PdfReviewTextSelection(page);
            _highlight = highlight;
            _feedback = PdfReviewSurfaceFeedback.None;
            InvalidateRectangleCache();
            _changeBars.Clear();
            _changeBarPixels = Image != null ? Image.Size : ClientSize;
            if (_selection != null && highlight != null)
                _changeBars.AddRange(bars ?? PdfReviewPageView.LayoutChangeBars(highlight,
                    _changeBarPixels.Width, _changeBarPixels.Height, null));
            TabStop = HasSelectableText;
            Cursor = HasSelectableText ? Cursors.IBeam : Cursors.Default;
            NotifySelectionChanged();
        }

        internal void DetachPage()
        {
            EndDrag();
            _selection = null;
            _highlight = null;
            _changeBars.Clear();
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
                    if (_selection.SelectIndices(ChangeBarSelection(_changeBars[bar])))
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
                bool highContrast = SystemInformation.HighContrast;
                Color outline = SystemColors.Highlight;
                using (var pen = new Pen(outline, highContrast ? 3f : 2f))
                {
                    pen.DashStyle = highContrast ? DashStyle.Dash : DashStyle.Solid;
                    List<int> selected = _selection.SelectedIndices();
                    if (!highContrast)
                    {
                        using (var brush = new SolidBrush(Color.FromArgb(62, outline)))
                            foreach (int i in selected)
                                if (i < _wordRectangles.Count && !_wordRectangles[i].IsEmpty)
                                    pe.Graphics.FillRectangle(brush, _wordRectangles[i]);
                    }
                    foreach (int i in selected)
                        if (i < _wordRectangles.Count && !_wordRectangles[i].IsEmpty)
                            pe.Graphics.DrawRectangle(pen, Rectangle.Round(_wordRectangles[i]));
                }
            }
            if (Focused && HasSelectableText)
                ControlPaint.DrawFocusRectangle(pe.Graphics,
                    new Rectangle(2, 2, Math.Max(0, ClientSize.Width - 5),
                        Math.Max(0, ClientSize.Height - 5)),
                    SystemColors.HighlightText, SystemColors.Highlight);
        }

        /// <summary>
        /// Индекс полосы изменений под точкой клиента или -1. Зона — нарисованная рейка со
        /// знаком (HitBounds в пикселях растра, масштабированные к клиенту). Точка над
        /// самим словом остаётся выбором слова: рейка не крадёт клики по тексту у края.
        /// </summary>
        internal int HitTestChangeBar(Point point)
        {
            if (_selection == null || _changeBars.Count == 0 ||
                _changeBarPixels.Width <= 0 || _changeBarPixels.Height <= 0 ||
                ClientSize.Width <= 0 || ClientSize.Height <= 0)
                return -1;
            EnsureRectangleCache();
            foreach (RectangleF rect in _wordRectangles)
                if (!rect.IsEmpty && rect.Contains(point))
                    return -1;
            float sx = ClientSize.Width / (float)_changeBarPixels.Width;
            float sy = ClientSize.Height / (float)_changeBarPixels.Height;
            for (int i = 0; i < _changeBars.Count; i++)
            {
                PdfReviewChangeBar bar = _changeBars[i];
                if (bar.Words.Count == 0)
                    continue; // пробельная пометка: выбирать нечего
                var hit = new RectangleF(bar.HitBounds.X * sx, bar.HitBounds.Y * sy,
                    bar.HitBounds.Width * sx, bar.HitBounds.Height * sy);
                hit.Inflate(ChangeBarHitTolerance, ChangeBarHitTolerance);
                if (hit.Contains(point))
                    return i;
            }
            return -1;
        }

        /// <summary>
        /// Слова, которые выбирает клик по полосе: изменённые слова её строки и неизменённые
        /// слова МЕЖДУ ними, только если те стоят на той же строке. Две колонки на одной
        /// линии дают отдельные фрагменты, а не диапазон чтения через полколонки.
        /// </summary>
        private List<int> ChangeBarSelection(PdfReviewChangeBar bar)
        {
            var changed = new List<int>();
            foreach (PdfReviewWord word in bar.Words)
            {
                int index;
                if (word != null && _wordIndex.TryGetValue(word, out index))
                    changed.Add(index);
            }
            changed.Sort();
            var result = new List<int>();
            for (int n = 0; n < changed.Count; n++)
            {
                if (n > 0 && changed[n] == changed[n - 1])
                    continue;
                if (n > 0 && changed[n] - changed[n - 1] > 1 &&
                    WordsOnLine(changed[n - 1] + 1, changed[n] - 1, bar.Band))
                    for (int k = changed[n - 1] + 1; k < changed[n]; k++)
                        result.Add(k);
                result.Add(changed[n]);
            }
            return result;
        }

        private bool WordsOnLine(int from, int to, RectangleF band)
        {
            for (int k = from; k <= to; k++)
            {
                RectangleF rect = _selection.WordRectangle(k, _changeBarPixels);
                if (rect.IsEmpty || !PdfReviewGeometry.SameTextLine(band, rect))
                    return false;
            }
            return true;
        }

        internal int HitTestWord(Point point, bool nearest)
        {
            EnsureRectangleCache();
            if (HitTestChangeBar(point) >= 0)
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
                _changeBars.Clear();
            }
            base.Dispose(disposing);
        }
    }
}
