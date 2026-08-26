using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelMerger
{
    internal enum PdfReviewPageViewState
    {
        Empty,
        DropTarget,
        Loading,
        Ready,
        MissingCounterpart,
        Unavailable
    }

    internal enum PdfReviewWheelResult
    {
        NotHandled,
        Zoomed,
        Scrolled,
        AtPreviousBoundary,
        AtNextBoundary
    }

    internal enum PdfReviewPagePosition
    {
        Default,
        Top,
        Bottom
    }

    /// <summary>
    /// Одна read-only страница для бокового просмотра Review. Рендерит в фоне,
    /// масштабирует существующий bitmap без повторного чтения PDF, отбрасывает
    /// поздний результат по generation/content revision и совмещает растр с
    /// семантической подсветкой и отдельным trusted-слоем выбора текста.
    /// </summary>
    internal sealed class PdfReviewPageView : UserControl
    {
        private const int WheelNotch = 120;
        private const int ScrollPerNotch = 60;

        private readonly Panel _viewport;
        private readonly PdfReviewPageSurface _picture;
        private readonly Label _message;
        private readonly Label _status;
        private readonly RoundedButton _minus, _plus, _fit;
        private Bitmap _bitmap;
        private PdfPageRef _page;
        private PdfReviewPage _reviewPage;
        private PdfPageRef _targetPage;
        private long _targetContentRevision;
        private string _caption = "";
        private string _highlightDescription = "";
        private double _scale = 1.0;
        private int _generation;
        private PdfReviewPageViewState _state = PdfReviewPageViewState.Empty;
        private bool _dropTarget;
        private readonly object _renderGate = new object();
        private RenderRequest _pending;
        private bool _renderWorker;

        private sealed class RenderRequest
        {
            public int Generation;
            public long ContentRevision;
            public PdfPageRef Page;
            public PdfReviewPage ReviewPage;
            public string Caption;
            public PdfReviewHighlight Highlight;
            public PdfReviewPagePosition Position;
        }

        public PdfReviewPageView()
        {
            BackColor = Theme.DarkBarFill;
            AccessibleDescription = Loc.T("review.source.interactions");
            _minus = Button("−", 8);
            _plus = Button("+", 46);
            _fit = Button(Loc.T("preview.fit"), 84, 100);
            _minus.AccessibleName = Loc.T("preview.tip.zoomOut");
            _plus.AccessibleName = Loc.T("preview.tip.zoomIn");
            _minus.Click += delegate { Step(-1); };
            _plus.Click += delegate { Step(1); };
            _fit.Click += delegate { Fit(); };

            _status = new Label();
            _status.ForeColor = Color.White;
            _status.BackColor = Color.Transparent;
            _status.AutoEllipsis = true;
            _status.TextAlign = ContentAlignment.MiddleLeft;
            _status.SetBounds(194, 8, 200, 28);
            _status.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(_status);

            _viewport = new Panel();
            _viewport.AutoScroll = true;
            _viewport.BackColor = Color.FromArgb(55, 55, 58);
            _viewport.SetBounds(0, 44, Width, Height - 44);
            _viewport.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _viewport.Resize += delegate { PlacePicture(); };
            Controls.Add(_viewport);

            // Белое полотно существует только вместе с готовым bitmap. Нулевой размер и
            // Visible=false не оставляют случайный белый квадрат на начальном сером фоне.
            _picture = new PdfReviewPageSurface();
            _picture.BackColor = Color.White;
            _picture.Size = Size.Empty;
            _picture.Visible = false;
            _picture.SizeMode = PictureBoxSizeMode.Zoom;
            _picture.ScrollViewport = _viewport;
            _picture.SelectionStateChanged += delegate
            {
                UpdateStatusText();
                UpdateSelectionAccessibility();
            };
            _picture.AutoScrollRequested += delegate(object sender,
                PdfReviewAutoScrollEventArgs args)
            {
                ScrollBy(args.DeltaX, args.DeltaY);
            };
            _picture.DoubleClick += delegate
            {
                if (_page != null)
                    PagePreviewForm.Show(FindForm(), _page,
                        string.Format(Loc.T("preview.title"), _page.PageIndex + 1), null);
            };
            _viewport.Controls.Add(_picture);

            _message = new Label();
            _message.Dock = DockStyle.Fill;
            _message.TextAlign = ContentAlignment.MiddleCenter;
            _message.ForeColor = Color.Gainsboro;
            _message.BackColor = _viewport.BackColor;
            _message.Padding = new Padding(24);
            _message.AccessibleDescription = Loc.T("review.source.interactions");
            _viewport.Controls.Add(_message);
            ApplyVisualState();
        }

        internal PdfReviewPageViewState ViewState
        {
            get { return _dropTarget ? PdfReviewPageViewState.DropTarget : _state; }
        }

        internal bool HasVisiblePage { get { return _state == PdfReviewPageViewState.Ready && _bitmap != null; } }
        internal double ZoomScale { get { return _scale; } }
        internal Point ScrollOffset
        {
            get { return new Point(-_viewport.AutoScrollPosition.X, -_viewport.AutoScrollPosition.Y); }
        }

        /// <summary>
        /// true, если pane уже показывает, рендерит либо признал недоступной именно эту
        /// логическую страницу и эту публикацию результата. null означает явное состояние
        /// «парной страницы нет»; поворот является частью физической идентичности.
        /// </summary>
        internal bool IsShowing(PdfPageRef page, long contentRevision)
        {
            if (_targetContentRevision != contentRevision)
                return false;
            if (page == null)
                return _state == PdfReviewPageViewState.MissingCounterpart;
            return (_state == PdfReviewPageViewState.Loading ||
                    _state == PdfReviewPageViewState.Ready ||
                    _state == PdfReviewPageViewState.Unavailable) &&
                   SamePage(_targetPage, page);
        }

        // Сохранён для узких state-тестов и старых внутренних вызовов: production-путь
        // всегда передаёт revision и тем самым не принимает старый semantic projection.
        internal bool IsShowing(PdfPageRef page)
        {
            if (page == null)
                return _state == PdfReviewPageViewState.MissingCounterpart;
            return (_state == PdfReviewPageViewState.Loading ||
                    _state == PdfReviewPageViewState.Ready ||
                    _state == PdfReviewPageViewState.Unavailable) &&
                   SamePage(_targetPage, page);
        }

        private static bool SamePage(PdfPageRef left, PdfPageRef right)
        {
            return left != null && right != null &&
                left.PageIndex == right.PageIndex &&
                PdfPageRef.ComposeRotation(0, left.Rotation) ==
                    PdfPageRef.ComposeRotation(0, right.Rotation) &&
                string.Equals(left.SourcePath, right.SourcePath,
                    StringComparison.OrdinalIgnoreCase);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (_status == null || _viewport == null)
                return; // OnResize в середине конструктора (контролы ещё не созданы)
            // На узкой панели (разделение сплиттера) строка статуса может не поместиться
            // после кнопок — прижимаем её к правому краю и не даём вылезти за ширину.
            int statusX = Math.Min(194, Math.Max(0, Width - 10));
            _status.SetBounds(statusX, 8, Math.Max(1, Width - statusX - 8), 28);
            _viewport.SetBounds(0, 44, Width, Math.Max(1, Height - 44));
            PlacePicture();
        }

        /// <summary>Пустая сторона до сравнения; это не «нет парной страницы».</summary>
        public void ShowEmpty(string caption)
        {
            BeginState(caption, PdfReviewPageViewState.Empty);
        }

        /// <summary>
        /// Показать страницу; null — «парной страницы нет». ReviewPage и revision
        /// публикуются только вместе с bitmap подходящего render-generation.
        /// </summary>
        public void ShowPage(PdfPageRef page, PdfReviewPage reviewPage, long contentRevision,
            string caption, PdfReviewHighlight highlight,
            PdfReviewPagePosition position = PdfReviewPagePosition.Default)
        {
            int generation = ++_generation;
            CancelPendingRender();
            DisposeBitmap();
            _page = null;
            _reviewPage = null;
            _targetPage = page == null ? null : page.Clone();
            _targetContentRevision = contentRevision;
            _caption = caption ?? "";
            AccessibleName = _caption;
            ApplyHighlightAccessibility(highlight);
            if (page == null)
            {
                _state = PdfReviewPageViewState.MissingCounterpart;
                ApplyVisualState();
                return;
            }

            bool start = false;
            lock (_renderGate)
            {
                // Только последний запрос заслуживает рендера. Один worker на pane не даёт
                // быстрому листанию плодить десятки WinRT-документов и растров параллельно.
                _pending = new RenderRequest
                {
                    Generation = generation,
                    ContentRevision = contentRevision,
                    Page = _targetPage.Clone(),
                    ReviewPage = reviewPage,
                    Caption = caption,
                    Highlight = highlight,
                    Position = position
                };
                if (!_renderWorker)
                {
                    _renderWorker = true;
                    start = true;
                }
            }
            _state = PdfReviewPageViewState.Loading;
            ApplyVisualState();
            if (start)
                Ui.RunWorker(RenderLoop);
        }

        internal void ShowPage(PdfPageRef page, string caption, PdfReviewHighlight highlight,
            PdfReviewPagePosition position = PdfReviewPagePosition.Default)
        {
            ShowPage(page, null, 0L, caption, highlight, position);
        }

        /// <summary>Включить/снять временную подсказку цели перетаскивания.</summary>
        internal void SetDropTarget(bool active)
        {
            if (_dropTarget == active)
                return;
            _dropTarget = active;
            ApplyVisualState();
        }

        private void BeginState(string caption, PdfReviewPageViewState state)
        {
            _generation++;
            CancelPendingRender();
            DisposeBitmap();
            _page = null;
            _reviewPage = null;
            _targetPage = null;
            _targetContentRevision = 0L;
            _caption = caption ?? "";
            AccessibleName = _caption;
            ApplyHighlightAccessibility(null);
            _state = state;
            ApplyVisualState();
        }

        private void ApplyHighlightAccessibility(PdfReviewHighlight highlight)
        {
            var descriptions = new List<string>();
            descriptions.Add(Loc.T("review.source.interactions"));
            if (highlight != null)
            {
                descriptions.Add(Loc.T(highlight.Style == PdfReviewHighlightStyle.Added
                    ? "review.legend.added" : "review.legend.removed"));
                var seen = new HashSet<string>(StringComparer.CurrentCulture);
                foreach (PdfReviewWhitespaceMarker marker in highlight.WhitespaceMarkers)
                    if (marker != null && !string.IsNullOrWhiteSpace(
                        marker.AccessibleDescription) &&
                        seen.Add(marker.AccessibleDescription))
                        descriptions.Add(marker.AccessibleDescription);
            }
            _highlightDescription = string.Join(". ", descriptions.ToArray());
            UpdateSelectionAccessibility();
        }

        private void UpdateSelectionAccessibility()
        {
            string interaction = _picture == null ? "" : _picture.InteractionStatus;
            string description = string.IsNullOrEmpty(interaction)
                ? _highlightDescription
                : string.Join(". ", new[] { _highlightDescription, interaction });
            AccessibleDescription = description;
            if (_picture != null)
            {
                _picture.AccessibleName = _caption;
                _picture.AccessibleDescription = description;
            }
            if (_status != null)
                _status.AccessibleDescription = description;
            if (_message != null)
                _message.AccessibleDescription = description;
        }

        private void CancelPendingRender()
        {
            lock (_renderGate)
                _pending = null;
        }

        private void RenderLoop()
        {
            while (true)
            {
                RenderRequest request;
                lock (_renderGate)
                {
                    request = _pending;
                    _pending = null;
                    if (request == null)
                    {
                        _renderWorker = false;
                        return;
                    }
                }

                Bitmap rendered = null;
                try
                {
                    using (var renderer = new PdfThumbnailRenderer())
                        rendered = renderer.Render(request.Page.SourcePath,
                            request.Page.PageIndex, 1200, 20000); // ≤24 млн пикселей
                    if (rendered != null && request.Page.Rotation != 0)
                        rendered.RotateFlip(PageRotation.FlipFor(request.Page.Rotation));
                    // Подсветка рисуется СРАЗУ на копии растра в воркере: переключение пар
                    // и зум тогда ничего не перерисовывают (картинка уже готова).
                    if (rendered != null)
                        rendered = DrawHighlight(rendered, request.Highlight);
                }
                catch
                {
                    // Сбой рендера одной страницы не должен ронять воркер: явное состояние
                    // Unavailable покажет ApplyRendered, а следующая пара всё равно отрисуется.
                }
                Bitmap ready = rendered;
                if (!Ui.OnUi(this, delegate
                {
                    ApplyRenderedWithHighlight(request.Generation, request.ContentRevision,
                        request.Page, request.ReviewPage, request.Highlight, ready, request.Caption,
                        request.Position);
                }))
                    if (ready != null) ready.Dispose();
            }
        }

        /// <summary>
        /// В normal mode превращает paper внутри authoritative word-box в Word-подобный
        /// фон, сохраняя тёмный PDF ink. Whitespace markers остаются отдельными. В high
        /// contrast используется системная outline/pattern-грамматика. Успешный результат
        /// забирает и освобождает source; при сбое source возвращается неизменённым.
        /// </summary>
        internal static Bitmap DrawHighlight(Bitmap source, PdfReviewHighlight highlight)
        {
            return DrawHighlight(source, highlight, SystemInformation.HighContrast);
        }

        /// <summary>Перегрузка с явным high-contrast флагом нужна для детерминированной проверки.</summary>
        internal static Bitmap DrawHighlight(Bitmap source, PdfReviewHighlight highlight,
            bool highContrast)
        {
            if (source == null || highlight == null)
                return source;
            List<RectangleF> rectangles = HighlightRectangles(highlight,
                source.Width, source.Height);
            List<WhitespaceDrawInfo> whitespace = WhitespaceDrawInfos(highlight,
                source.Width, source.Height);
            if (rectangles.Count == 0 && whitespace.Count == 0)
                return source;
            return highContrast
                ? DrawHighContrastHighlight(source, highlight, rectangles, whitespace)
                : DrawNormalHighlight(source, highlight, rectangles, whitespace);
        }

        private static Bitmap DrawNormalHighlight(Bitmap source,
            PdfReviewHighlight highlight, IList<RectangleF> rectangles,
            IList<WhitespaceDrawInfo> whitespace)
        {
            Bitmap copy = null;
            try
            {
                copy = PdfReviewHighlightRenderer.Create32BppCopy(source);
                PdfReviewHighlightRenderer.ApplyWordFills(copy, rectangles, highlight.Color);
                using (Graphics graphics = Graphics.FromImage(copy))
                {
                    graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    if (whitespace.Count > 0)
                        PrepareWhitespacePlacements(graphics, copy, rectangles, whitespace);
                    Color color, edgeColor;
                    ResolveHighlightColors(highlight, false, out color, out edgeColor);
                    var railAnchors = new List<RectangleF>(
                        rectangles.Count + whitespace.Count);
                    railAnchors.AddRange(rectangles);
                    foreach (WhitespaceDrawInfo item in whitespace)
                        railAnchors.Add(item.Anchor);
                    DrawChangeBars(graphics, copy, MergeLineBands(railAnchors), highlight,
                        edgeColor, color);
                    DrawWhitespaceMarkers(graphics, whitespace, edgeColor, color);
                }
                source.Dispose();
                return copy;
            }
            catch
            {
                if (copy != null)
                    copy.Dispose();
                return source;
            }
        }

        private static Bitmap DrawHighContrastHighlight(Bitmap source,
            PdfReviewHighlight highlight, IList<RectangleF> rectangles,
            IList<WhitespaceDrawInfo> whitespace)
        {
            Bitmap copy = null;
            try
            {
                copy = PdfReviewHighlightRenderer.Create32BppCopy(source);
                Color color, edgeColor;
                ResolveHighlightColors(highlight, true, out color, out edgeColor);
                using (Graphics g = Graphics.FromImage(copy))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    PrepareWhitespacePlacements(g, copy, rectangles, whitespace);
                    using (var edge = new Pen(Color.FromArgb(255, edgeColor), 3f))
                    using (var primary = new Pen(Color.FromArgb(255, color), 1.5f))
                    {
                        ConfigurePattern(edge, highlight.Style);
                        ConfigurePattern(primary, highlight.Style);
                        edge.LineJoin = LineJoin.Round;
                        primary.LineJoin = LineJoin.Round;
                        foreach (RectangleF rect in rectangles)
                        {
                            RectangleF outline = rect;
                            outline.Inflate(3f, 3f);
                            g.DrawRectangle(edge, outline.X, outline.Y,
                                outline.Width, outline.Height);
                            g.DrawRectangle(primary, outline.X, outline.Y,
                                outline.Width, outline.Height);
                        }

                        var railAnchors = new List<RectangleF>(
                            rectangles.Count + whitespace.Count);
                        railAnchors.AddRange(rectangles);
                        foreach (WhitespaceDrawInfo item in whitespace)
                            railAnchors.Add(item.Anchor);
                        DrawChangeBars(g, copy, MergeLineBands(railAnchors), highlight,
                            edgeColor, color);
                    }
                    DrawWhitespaceMarkers(g, whitespace, edgeColor, color);
                }
                source.Dispose();
                return copy;
            }
            catch
            {
                if (copy != null)
                    copy.Dispose();
                return source;
            }
        }

        /// <summary>
        /// Цвет в high contrast выбирается только из системной пары Window/WindowText,
        /// причём берётся более контрастный к белому PDF-холсту. Pattern и знак остаются.
        /// </summary>
        internal static void ResolveHighlightColors(PdfReviewHighlight highlight,
            bool highContrast, out Color color, out Color edgeColor)
        {
            color = highlight == null ? Color.Empty : highlight.Color;
            edgeColor = highlight == null || highlight.EdgeColor.IsEmpty
                ? color : highlight.EdgeColor;
            if (!highContrast)
                return;
            Color system = ContrastRatio(SystemColors.WindowText, Color.White) >=
                ContrastRatio(SystemColors.Window, Color.White)
                ? SystemColors.WindowText : SystemColors.Window;
            color = system;
            edgeColor = system;
        }

        /// <summary>WCAG relative-luminance contrast ratio; чистая функция для UI-проверок.</summary>
        internal static double ContrastRatio(Color left, Color right)
        {
            double l1 = RelativeLuminance(left);
            double l2 = RelativeLuminance(right);
            return (Math.Max(l1, l2) + 0.05) / (Math.Min(l1, l2) + 0.05);
        }

        private static double RelativeLuminance(Color color)
        {
            return 0.2126 * LinearChannel(color.R / 255.0) +
                0.7152 * LinearChannel(color.G / 255.0) +
                0.0722 * LinearChannel(color.B / 255.0);
        }

        private static double LinearChannel(double value)
        {
            return value <= 0.04045 ? value / 12.92 :
                Math.Pow((value + 0.055) / 1.055, 2.4);
        }

        private static void ConfigurePattern(Pen pen, PdfReviewHighlightStyle style)
        {
            if (pen == null || style != PdfReviewHighlightStyle.Added)
                return;
            // DashPattern измеряется в ширинах пера. Пересчёт сохраняет одинаковые
            // абсолютные штрихи у тёмного края и цветной внутренней линии.
            pen.DashStyle = DashStyle.Custom;
            pen.DashPattern = new[] { 8f / pen.Width, 5f / pen.Width };
            pen.DashCap = DashCap.Flat;
        }

        /// <summary>
        /// Полосы затронутых строк в пикселях. Несколько изменённых слов одной строки
        /// дают одну пометку, разнесённые строки — независимые пометки.
        /// </summary>
        internal static List<RectangleF> ChangeLineBands(PdfReviewHighlight highlight,
            int pixelWidth, int pixelHeight)
        {
            return MergeLineBands(HighlightRectangles(highlight, pixelWidth, pixelHeight));
        }

        private static List<RectangleF> HighlightRectangles(PdfReviewHighlight highlight,
            int pixelWidth, int pixelHeight)
        {
            var result = new List<RectangleF>();
            if (highlight == null || highlight.Boxes.Count == 0 || pixelWidth <= 0 ||
                pixelHeight <= 0 || highlight.ViewWidthPt <= 0 ||
                highlight.ViewHeightPt <= 0)
                return result;
            foreach (PdfReviewBox box in highlight.Boxes)
            {
                RectangleF rect = PdfReviewGeometry.ToPixelRect(box,
                    highlight.ViewWidthPt, highlight.ViewHeightPt, pixelWidth, pixelHeight);
                if (rect.Width >= 1 && rect.Height >= 1)
                    result.Add(rect);
            }
            return result;
        }

        private sealed class WhitespaceDrawInfo
        {
            public PdfReviewWhitespaceMarker Marker;
            public RectangleF Anchor;
            public RectangleF Bounds;
            public float FontSize;
        }

        private static List<WhitespaceDrawInfo> WhitespaceDrawInfos(
            PdfReviewHighlight highlight, int pixelWidth, int pixelHeight)
        {
            var result = new List<WhitespaceDrawInfo>();
            if (highlight == null || highlight.WhitespaceMarkers == null ||
                highlight.ViewWidthPt <= 0 || highlight.ViewHeightPt <= 0 ||
                pixelWidth <= 0 || pixelHeight <= 0)
                return result;
            foreach (PdfReviewWhitespaceMarker marker in highlight.WhitespaceMarkers)
            {
                if (marker == null || string.IsNullOrEmpty(marker.Text))
                    continue;
                RectangleF anchor = PdfReviewGeometry.ToPixelRect(marker.Box,
                    highlight.ViewWidthPt, highlight.ViewHeightPt,
                    pixelWidth, pixelHeight);
                if (anchor.IsEmpty || anchor.Width <= 0 || anchor.Height <= 0)
                    continue;
                result.Add(new WhitespaceDrawInfo
                {
                    Marker = marker,
                    Anchor = anchor
                });
            }
            return result;
        }

        private static void PrepareWhitespacePlacements(Graphics graphics, Bitmap source,
            IList<RectangleF> wordRectangles, IList<WhitespaceDrawInfo> whitespace)
        {
            if (graphics == null || source == null || whitespace == null)
                return;
            var occupied = new List<RectangleF>();
            float preferred = Math.Max(10f, Math.Min(18f, source.Width / 70f));
            foreach (WhitespaceDrawInfo item in whitespace)
            {
                if (item == null || item.Marker == null)
                    continue;
                for (float size = preferred; size >= 8f; size -= 2f)
                {
                    RectangleF bounds;
                    using (var font = new Font(FontFamily.GenericSansSerif, size,
                        FontStyle.Bold, GraphicsUnit.Pixel))
                    {
                        SizeF measured = graphics.MeasureString(item.Marker.Text, font,
                            PointF.Empty, StringFormat.GenericTypographic);
                        bounds = FindWhitespaceBounds(source, item.Anchor,
                            (float)Math.Ceiling(measured.Width) + 6f,
                            (float)Math.Ceiling(measured.Height) + 4f,
                            wordRectangles, occupied);
                    }
                    if (!bounds.IsEmpty)
                    {
                        item.Bounds = bounds;
                        item.FontSize = size;
                        occupied.Add(bounds);
                        break;
                    }
                }
                // Если на этой строке нет чистого участка бумаги, текстовый token не
                // рисуется поверх исходного ink. Rail/знак и AccessibleDescription остаются.
            }
        }

        private static RectangleF FindWhitespaceBounds(Bitmap source, RectangleF anchor,
            float width, float height, IList<RectangleF> words, IList<RectangleF> occupied)
        {
            if (source == null || width <= 0 || height <= 0 ||
                width > source.Width - 4 || height > source.Height - 4)
                return RectangleF.Empty;
            float y = Math.Max(2f, Math.Min(source.Height - height - 2f,
                anchor.Top + anchor.Height / 2f - height / 2f));
            var inline = new RectangleF(
                Math.Max(2f, Math.Min(source.Width - width - 2f,
                    anchor.Left + anchor.Width / 2f - width / 2f)),
                y, width, height);
            if (MarkerAreaAvailable(source, inline, words, occupied))
                return inline;

            // Inline часто узок: ищем чистое место той же строки от внешнего поля.
            // Сторона определяется ближайшим краем marker-anchor и не меняет семантику.
            bool fromLeft = anchor.Left + anchor.Width / 2f <= source.Width / 2f;
            int max = Math.Max(2, (int)Math.Floor(source.Width - width - 2f));
            if (fromLeft)
            {
                for (int x = 2; x <= max; x += 4)
                {
                    var candidate = new RectangleF(x, y, width, height);
                    if (MarkerAreaAvailable(source, candidate, words, occupied))
                        return candidate;
                }
            }
            else
            {
                for (int x = max; x >= 2; x -= 4)
                {
                    var candidate = new RectangleF(x, y, width, height);
                    if (MarkerAreaAvailable(source, candidate, words, occupied))
                        return candidate;
                }
            }
            return RectangleF.Empty;
        }

        private static bool MarkerAreaAvailable(Bitmap source, RectangleF candidate,
            IList<RectangleF> words, IList<RectangleF> occupied)
        {
            if (!IsClearPaperArea(source, candidate))
                return false;
            if (IntersectsInflated(candidate, words, 4f) ||
                IntersectsInflated(candidate, occupied, 2f))
                return false;
            return true;
        }

        private static bool IntersectsInflated(RectangleF candidate,
            IList<RectangleF> rectangles, float inflate)
        {
            if (rectangles == null)
                return false;
            foreach (RectangleF value in rectangles)
            {
                RectangleF blocked = value;
                blocked.Inflate(inflate, inflate);
                if (candidate.IntersectsWith(blocked))
                    return true;
            }
            return false;
        }

        private static bool IsClearPaperArea(Bitmap bitmap, RectangleF area)
        {
            if (bitmap == null || area.Left < 0 || area.Top < 0 ||
                area.Right > bitmap.Width || area.Bottom > bitmap.Height)
                return false;
            int left = Math.Max(0, (int)Math.Floor(area.Left));
            int top = Math.Max(0, (int)Math.Floor(area.Top));
            int right = Math.Min(bitmap.Width - 1, (int)Math.Ceiling(area.Right));
            int bottom = Math.Min(bitmap.Height - 1, (int)Math.Ceiling(area.Bottom));
            for (int y = top; y <= bottom; y++)
                for (int x = left; x <= right; x++)
                {
                    Color pixel = bitmap.GetPixel(x, y);
                    if (pixel.A >= 24 && (pixel.R < 238 || pixel.G < 238 || pixel.B < 238))
                        return false;
                }
            return true;
        }

        private static List<RectangleF> MergeLineBands(IList<RectangleF> rectangles)
        {
            var ordered = new List<RectangleF>();
            if (rectangles != null)
                foreach (RectangleF rect in rectangles)
                    ordered.Add(rect);
            ordered.Sort(delegate(RectangleF left, RectangleF right)
            {
                int top = left.Top.CompareTo(right.Top);
                return top != 0 ? top : left.Left.CompareTo(right.Left);
            });

            var bands = new List<RectangleF>();
            foreach (RectangleF rect in ordered)
            {
                if (bands.Count > 0 && SameTextLine(bands[bands.Count - 1], rect))
                    bands[bands.Count - 1] = RectangleF.Union(bands[bands.Count - 1], rect);
                else
                    bands.Add(rect);
            }
            return bands;
        }

        private static bool SameTextLine(RectangleF left, RectangleF right)
        {
            float smallerHeight = Math.Min(left.Height, right.Height);
            if (smallerHeight <= 0)
                return false;
            float overlap = Math.Min(left.Bottom, right.Bottom) -
                Math.Max(left.Top, right.Top);
            if (overlap >= smallerHeight * 0.35f)
                return true;
            float leftCenter = left.Top + left.Height / 2f;
            float rightCenter = right.Top + right.Height / 2f;
            return Math.Abs(leftCenter - rightCenter) <= smallerHeight * 0.45f;
        }

        internal static string HighlightSymbol(PdfReviewHighlightStyle style)
        {
            return style == PdfReviewHighlightStyle.Added ? "+" : "−";
        }

        private static void DrawChangeBars(Graphics graphics, Bitmap raster,
            IList<RectangleF> bands, PdfReviewHighlight highlight, Color edgeColor,
            Color color)
        {
            if (graphics == null || raster == null || bands == null || bands.Count == 0 ||
                highlight.ChangeBarSide == PdfReviewChangeBarSide.None)
                return;
            float inset = Math.Max(5f, Math.Min(12f, raster.Width * 0.008f));
            string symbol = HighlightSymbol(highlight.Style);
            float symbolSize = Math.Max(10f, Math.Min(18f, raster.Width / 70f));
            using (var edge = new Pen(Color.FromArgb(255, edgeColor), 5f))
            using (var primary = new Pen(Color.FromArgb(255, color), 2.25f))
            using (var font = new Font(FontFamily.GenericSansSerif, symbolSize,
                FontStyle.Bold, GraphicsUnit.Pixel))
            {
                ConfigurePattern(edge, highlight.Style);
                ConfigurePattern(primary, highlight.Style);
                edge.StartCap = edge.EndCap = LineCap.Round;
                primary.StartCap = primary.EndCap = LineCap.Round;
                foreach (RectangleF band in bands)
                {
                    float padding = Math.Max(2f, band.Height * 0.18f);
                    float top = Math.Max(5f, band.Top - padding);
                    float bottom = Math.Min(raster.Height - 5f,
                        band.Bottom + padding);
                    if (bottom - top < 5f)
                    {
                        float middle = (top + bottom) / 2f;
                        top = Math.Max(5f, middle - 2.5f);
                        bottom = Math.Min(raster.Height - 5f, middle + 2.5f);
                    }
                    if (bottom <= top)
                        continue;
                    float x = FindClearRailX(raster, highlight.ChangeBarSide,
                        inset, top, bottom);
                    graphics.DrawLine(edge, x, top, x, bottom);
                    graphics.DrawLine(primary, x, top, x, bottom);

                    SizeF measured = graphics.MeasureString(symbol, font,
                        PointF.Empty, StringFormat.GenericTypographic);
                    float sx = highlight.ChangeBarSide == PdfReviewChangeBarSide.Left
                        ? x + 5f : x - measured.Width - 5f;
                    sx = Math.Max(1f, Math.Min(raster.Width - measured.Width - 1f, sx));
                    float sy = Math.Max(1f, Math.Min(raster.Height - measured.Height - 1f,
                        (top + bottom - measured.Height) / 2f));
                    DrawOutlinedText(graphics, symbol, font, new PointF(sx, sy),
                        edgeColor, color);
                }
            }
        }

        private static float FindClearRailX(Bitmap raster, PdfReviewChangeBarSide side,
            float preferredInset, float top, float bottom)
        {
            int outerLimit = Math.Max(5, Math.Min(80,
                (int)Math.Round(raster.Width * 0.10)));
            int preferred = side == PdfReviewChangeBarSide.Left
                ? (int)Math.Round(preferredInset)
                : raster.Width - 1 - (int)Math.Round(preferredInset);
            for (int offset = 0; offset <= outerLimit; offset += 2)
            {
                int x = side == PdfReviewChangeBarSide.Left ? 3 + offset :
                    raster.Width - 4 - offset;
                if (Math.Abs(x - preferred) > outerLimit)
                    continue;
                var strip = new RectangleF(x - 3f, top, 6f, bottom - top);
                if (IsClearPaperArea(raster, strip))
                    return x;
            }
            return Math.Max(3f, Math.Min(raster.Width - 4f, preferred));
        }

        private static void DrawWhitespaceMarkers(Graphics graphics,
            IList<WhitespaceDrawInfo> whitespace, Color edgeColor, Color color)
        {
            if (graphics == null || whitespace == null)
                return;
            foreach (WhitespaceDrawInfo item in whitespace)
            {
                if (item == null || item.Marker == null || item.Bounds.IsEmpty ||
                    item.FontSize <= 0)
                    continue;
                using (var font = new Font(FontFamily.GenericSansSerif, item.FontSize,
                    FontStyle.Bold, GraphicsUnit.Pixel))
                    DrawOutlinedText(graphics, item.Marker.Text, font,
                        new PointF(item.Bounds.Left + 3f, item.Bounds.Top + 2f),
                        edgeColor, color);
            }
        }

        private static void DrawOutlinedText(Graphics graphics, string text, Font font,
            PointF location, Color edgeColor, Color color)
        {
            using (var edge = new SolidBrush(Color.FromArgb(255, edgeColor)))
            using (var primary = new SolidBrush(Color.FromArgb(255, color)))
            {
                graphics.DrawString(text, font, edge,
                    new PointF(location.X - 1f, location.Y));
                graphics.DrawString(text, font, edge,
                    new PointF(location.X + 1f, location.Y));
                graphics.DrawString(text, font, edge,
                    new PointF(location.X, location.Y - 1f));
                graphics.DrawString(text, font, edge,
                    new PointF(location.X, location.Y + 1f));
                graphics.DrawString(text, font, primary, location);
            }
        }

        private void ApplyRendered(int generation, long contentRevision, PdfPageRef request,
            PdfReviewPage reviewPage, Bitmap rendered, string caption,
            PdfReviewPagePosition position)
        {
            ApplyRenderedWithHighlight(generation, contentRevision, request, reviewPage,
                null, rendered, caption, position);
        }

        private void ApplyRenderedWithHighlight(int generation, long contentRevision,
            PdfPageRef request, PdfReviewPage reviewPage, PdfReviewHighlight highlight,
            Bitmap rendered, string caption, PdfReviewPagePosition position)
        {
            if (generation != _generation || contentRevision != _targetContentRevision ||
                IsDisposed)
            {
                if (rendered != null) rendered.Dispose();
                return;
            }
            DisposeBitmap();
            _bitmap = rendered;
            _page = request;
            _reviewPage = null;
            _targetPage = request == null ? null : request.Clone();
            _caption = caption ?? "";
            if (_bitmap == null)
                _state = PdfReviewPageViewState.Unavailable;
            else
            {
                _reviewPage = reviewPage;
                _picture.Image = _bitmap;
                _picture.AttachPage(_reviewPage, highlight);
                _state = PdfReviewPageViewState.Ready;
                // AutoScroll не учитывает невидимый дочерний контрол. Страница уже имеет
                // готовый bitmap, поэтому включаем её до расчёта позиции Top/Bottom.
                _picture.Visible = true;
                if (position == PdfReviewPagePosition.Default)
                    Fit();
                else
                {
                    PlacePicture();
                    ApplyPosition(position);
                }
            }
            ApplyVisualState();
        }

        private void ApplyPosition(PdfReviewPagePosition position)
        {
            if (_bitmap == null || _state != PdfReviewPageViewState.Ready)
                return;
            _viewport.PerformLayout();
            int y = position == PdfReviewPagePosition.Bottom
                ? Math.Max(0, _picture.Height - _viewport.ClientSize.Height)
                : 0;
            _viewport.AutoScrollPosition = new Point(0, y);
            PlacePicture();
        }

        private void ApplyVisualState()
        {
            if (_message == null)
                return;
            bool ready = _state == PdfReviewPageViewState.Ready && _bitmap != null;
            _picture.Visible = ready;
            if (!ready)
            {
                _picture.Size = Size.Empty;
                _picture.Location = Point.Empty;
            }

            _message.Visible = _dropTarget || !ready;
            if (_dropTarget)
            {
                _message.Text = Loc.T("review.source.dropHere");
                _message.BackColor = Theme.ReviewBlueDark;
                _message.ForeColor = Color.White;
            }
            else
            {
                _message.BackColor = _viewport.BackColor;
                _message.ForeColor = Color.Gainsboro;
                switch (_state)
                {
                    case PdfReviewPageViewState.Loading:
                        _message.Text = Loc.T("preview.loading");
                        break;
                    case PdfReviewPageViewState.MissingCounterpart:
                        _message.Text = Loc.T("review.source.missing");
                        break;
                    case PdfReviewPageViewState.Unavailable:
                        _message.Text = Loc.T("preview.unavailable");
                        break;
                    default:
                        _message.Text = Loc.T("review.source.empty");
                        break;
                }
            }
            if (_message.Visible)
                _message.BringToFront();
            UpdateStatusText();
            UpdateButtons();
        }

        private void UpdateStatusText()
        {
            if (_state == PdfReviewPageViewState.Ready && _page != null)
            {
                string pageStatus = string.Format(Loc.T("review.source.ready"), _caption,
                    _page.PageIndex + 1, PreviewZoom.Percent(_scale));
                _status.Text = JoinStatus(pageStatus, _picture.InteractionStatus);
            }
            else if (_state == PdfReviewPageViewState.Loading)
                _status.Text = JoinStatus(_caption, Loc.T("preview.loading"));
            else if (_state == PdfReviewPageViewState.Unavailable)
                _status.Text = JoinStatus(_caption, Loc.T("preview.unavailable"));
            else
                _status.Text = _caption;
        }

        private static string JoinStatus(string caption, string state)
        {
            if (string.IsNullOrEmpty(caption)) return state ?? "";
            if (string.IsNullOrEmpty(state)) return caption;
            return caption + " · " + state;
        }

        private RoundedButton Button(string text, int x, int width = 32)
        {
            var button = new RoundedButton(ButtonLook.OnDark);
            button.Text = text;
            button.SetBounds(x, 8, width, 28);
            button.AccessibleName = text;
            Controls.Add(button);
            return button;
        }

        private void Step(int direction)
        {
            if (_bitmap == null) return;
            Point center = new Point(_viewport.ClientSize.Width / 2, _viewport.ClientSize.Height / 2);
            ApplyScale(PreviewZoom.Next(_scale, direction), center);
        }

        private void Fit()
        {
            if (_bitmap == null) return;
            double fit = PreviewZoom.Fit(_bitmap.Size,
                new Size(Math.Max(1, _viewport.ClientSize.Width - 24),
                         Math.Max(1, _viewport.ClientSize.Height - 24)));
            _scale = fit;
            _viewport.AutoScrollPosition = Point.Empty;
            PlacePicture();
        }

        private bool ApplyScale(double scale, Point anchor)
        {
            if (_bitmap == null || Math.Abs(scale - _scale) < 0.000001)
                return false;
            double old = _scale;
            Point before = ScrollOffset;
            _scale = scale;
            PlacePicture();
            _viewport.AutoScrollPosition = new Point(
                PreviewZoom.Anchor(before.X, anchor.X, old, _scale),
                PreviewZoom.Anchor(before.Y, anchor.Y, old, _scale));
            UpdateStatusText();
            return true;
        }

        private void PlacePicture()
        {
            if (_bitmap == null || _state != PdfReviewPageViewState.Ready)
                return;
            Size size = PreviewZoom.Scaled(_bitmap.Size, _scale);
            _picture.Size = size;
            _picture.Location = PreviewZoom.Centered(size, _viewport.ClientSize,
                _viewport.AutoScrollPosition);
            _viewport.PerformLayout();
            UpdateStatusText();
            UpdateButtons();
        }

        internal bool ContainsViewport(Point screenPoint)
        {
            return !IsDisposed && _viewport.IsHandleCreated &&
                _viewport.RectangleToScreen(_viewport.ClientRectangle).Contains(screenPoint);
        }

        /// <summary>
        /// Колесо действует только в этой области. Результат отделяет успешный scroll/zoom
        /// от границы страницы, чтобы форму — владельца строк viewer — можно было попросить
        /// продолжить чтение на соседней физической странице активной стороны.
        /// </summary>
        internal PdfReviewWheelResult HandleWheel(Point screenPoint, int delta, bool controlDown)
        {
            if (!ContainsViewport(screenPoint))
                return PdfReviewWheelResult.NotHandled;
            return HandleWheelAt(_viewport.PointToClient(screenPoint), delta, controlDown);
        }

        internal PdfReviewWheelResult HandleWheelAt(Point inViewport, int delta, bool controlDown)
        {
            if (_bitmap == null || _state != PdfReviewPageViewState.Ready || delta == 0 ||
                !_viewport.ClientRectangle.Contains(inViewport))
                return PdfReviewWheelResult.NotHandled;
            if (controlDown)
                return ApplyScale(PreviewZoom.Next(_scale, delta > 0 ? +1 : -1), inViewport)
                    ? PdfReviewWheelResult.Zoomed : PdfReviewWheelResult.NotHandled;

            int dy = -delta * ScrollPerNotch / WheelNotch;
            if (dy == 0)
                dy = delta > 0 ? -1 : 1;
            if (ScrollBy(dy))
                return PdfReviewWheelResult.Scrolled;
            return delta > 0
                ? PdfReviewWheelResult.AtPreviousBoundary
                : PdfReviewWheelResult.AtNextBoundary;
        }

        private bool ScrollBy(int dy)
        {
            return ScrollBy(0, dy);
        }

        private bool ScrollBy(int dx, int dy)
        {
            Point before = ScrollOffset;
            _viewport.AutoScrollPosition = new Point(before.X + dx, before.Y + dy);
            Point after = ScrollOffset;
            return after != before;
        }

        private void UpdateButtons()
        {
            bool ready = _bitmap != null && _state == PdfReviewPageViewState.Ready;
            _minus.Enabled = ready && _scale > PreviewZoom.Min;
            _plus.Enabled = ready && _scale < PreviewZoom.Max;
            _fit.Enabled = ready;
        }

        private void DisposeBitmap()
        {
            _picture.DetachPage();
            _picture.Image = null;
            _picture.Visible = false;
            _picture.Size = Size.Empty;
            if (_bitmap != null) _bitmap.Dispose();
            _bitmap = null;
            _reviewPage = null;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _generation++;
                CancelPendingRender();
                DisposeBitmap();
            }
            base.Dispose(disposing);
        }
    }
}
