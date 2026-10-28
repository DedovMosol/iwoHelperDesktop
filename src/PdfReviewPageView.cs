using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Automation;

namespace ExcelMerger
{
    /// <summary>
    /// Нарисованная полоса изменений одной строки (пиксели растра страницы): по ней же
    /// поверхность выбирает изменённые слова строки при клике.
    /// </summary>
    internal sealed class PdfReviewChangeBar
    {
        public RectangleF Band;
        public float X;
        public float Top;
        public float Bottom;
        public RectangleF HitBounds;
        public readonly List<PdfReviewWord> Words = new List<PdfReviewWord>();
    }

    internal enum PdfReviewPageViewState
    {
        Empty,
        DropTarget,
        Loading,
        Ready,
        MissingCounterpart,
        Unavailable,
        // Файл заведомо не помещается в бюджет показа этой разрядности и ширины
        // (см. PdfReviewRenderSession.MaxShowableFileBytes).
        TooLargeToShow
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
        internal const int WmMouseWheel = 0x020A;
        private const long MkControl = 0x0008;

        private readonly Panel _viewport;
        private readonly PdfReviewPageSurface _picture;
        private readonly Label _message;
        private readonly Label _status;
        private readonly RoundedButton _minus, _plus, _fit;
        private Bitmap _bitmap;
        private BudgetedBitmap _ownedBitmap;
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
        // Свой рендерер у одиночного вида; окно сравнения подставляет общий (RenderSession).
        private PdfReviewRenderSession _session = new PdfReviewRenderSession();
        private bool _ownsSession = true;
        // Повторы после MemoryReleased ограничены: страница, которой не хватает памяти даже
        // после освобождения, не должна перерендериваться на каждом простое рендерера.
        private const int MaxMemoryRetries = 3;
        private int _memoryRetries;

        private PdfReviewViewContent _currentContent;
        private PdfReviewPagePosition _currentPosition = PdfReviewPagePosition.Default;
        internal bool AllowFullScreen { get; set; } = true;
        internal int RenderWidthLimit { get; set; } = 1200;

        private sealed class RenderRequest
        {
            public int Generation;
            public PdfReviewViewContent Content;
            public PdfReviewPagePosition Position;
            public int RenderWidth;
            public PdfReviewRenderSession Session;
        }

        /// <summary>
        /// Общий рендерер окна; назначается до первого показа. Подставленная сессия не
        /// принадлежит виду и не закрывается им; собственная закрывается при замене и в Dispose.
        /// </summary>
        internal PdfReviewRenderSession RenderSession
        {
            get { return _session; }
            set
            {
                if (value == null || ReferenceEquals(value, _session))
                    return;
                if (_ownsSession)
                    _session.Dispose();
                _session = value;
                _ownsSession = false;
            }
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
                if (AllowFullScreen && _currentContent != null &&
                    _currentContent.BasePage != null)
                    PdfReviewFullScreenForm.Show(FindForm(), _currentContent, _session);
            };
            _viewport.Controls.Add(_picture);
            PdfMemoryBudget.MemoryReleased += OnMemoryReleased;

            _message = new Label();
            _message.Dock = DockStyle.Fill;
            _message.TextAlign = ContentAlignment.MiddleCenter;
            _message.ForeColor = Color.Gainsboro;
            _message.BackColor = _viewport.BackColor;
            _message.Padding = new Padding(24);
            // Живая область: «недоступно» и «слишком большой» диктор объявляет так же, как
            // их видно на экране, а не только при переходе фокуса на панель.
            _message.LiveSetting = AutomationLiveSetting.Polite;
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

        // Наблюдаемое состояние для тестов (TS-11): чтение по имени приватного поля ломалось бы
        // молча при переименовании и читало флаг рендера мимо его блокировки.
        internal Bitmap ShownBitmap { get { return _bitmap; } }   // только UI-поток
        internal PdfPageRef ShownPage { get { return _page; } }   // только UI-поток
        internal bool RenderWorkerActive { get { lock (_renderGate) return _renderWorker; } }
        internal int ShowGeneration { get { return _generation; } } // только UI-поток
        internal PdfReviewPage ShownReviewPage { get { return _reviewPage; } } // только UI-поток
        internal long TargetContentRevision { get { return _targetContentRevision; } } // только UI-поток
        internal Panel ViewportPanel { get { return _viewport; } }
        internal PdfReviewPageSurface SurfaceControl { get { return _picture; } }
        internal Label MessageLabel { get { return _message; } }
        internal Label StatusLabel { get { return _status; } }
        internal Point ScrollOffset
        {
            get { return new Point(-_viewport.AutoScrollPosition.X, -_viewport.AutoScrollPosition.Y); }
        }

        /// <summary>
        /// true, если pane уже показывает, рендерит либо признал слишком большой именно эту
        /// логическую страницу и эту публикацию результата. «Недоступно» (сбой рендера,
        /// нехватка памяти) показом не считается: повторный выбор строки пробует снова.
        /// null означает явное состояние «парной страницы нет»; поворот — часть идентичности.
        /// </summary>
        internal bool IsShowing(PdfPageRef page, long contentRevision)
        {
            return _targetContentRevision == contentRevision && IsShowing(page);
        }

        // Сохранён для узких state-тестов и старых внутренних вызовов: production-путь
        // всегда передаёт revision и тем самым не принимает старый semantic projection.
        internal bool IsShowing(PdfPageRef page)
        {
            if (page == null)
                return _state == PdfReviewPageViewState.MissingCounterpart;
            return (_state == PdfReviewPageViewState.Loading ||
                    _state == PdfReviewPageViewState.Ready ||
                    _state == PdfReviewPageViewState.TooLargeToShow) &&
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
            ShowContent(new PdfReviewViewContent(page, reviewPage, highlight, null,
                contentRevision, caption), position);
        }

        internal void ShowContent(PdfReviewViewContent content,
            PdfReviewPagePosition position = PdfReviewPagePosition.Default)
        {
            _memoryRetries = 0;
            ShowContentCore(content, position);
        }

        private void ShowContentCore(PdfReviewViewContent content,
            PdfReviewPagePosition position)
        {
            int generation = ++_generation;
            CancelPendingRender();
            DisposeBitmap();
            _currentContent = content;
            _currentPosition = position;
            PdfPageRef page = content == null ? null : content.BasePage;
            _page = null;
            _reviewPage = null;
            _targetPage = page == null ? null : page.Clone();
            _targetContentRevision = content == null ? 0L : content.Revision;
            _caption = content == null ? "" : content.Caption;
            AccessibleName = _caption;
            ApplyHighlightAccessibility(content == null ? null : content.BaseHighlight,
                content == null ? null : content.OverlayHighlight);
            if (page == null)
            {
                _state = PdfReviewPageViewState.MissingCounterpart;
                ApplyVisualState();
                return;
            }

            bool start = false;
            lock (_renderGate)
            {
                _pending = new RenderRequest
                {
                    Generation = generation,
                    Content = content,
                    Position = position,
                    RenderWidth = RenderWidthFor(_viewport.ClientSize.Width,
                        RenderWidthLimit),
                    Session = _session
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
            _currentContent = null;
            _page = null;
            _reviewPage = null;
            _targetPage = null;
            _targetContentRevision = 0L;
            _caption = caption ?? "";
            AccessibleName = _caption;
            ApplyHighlightAccessibility(null, null);
            _state = state;
            ApplyVisualState();
        }

        private void ApplyHighlightAccessibility(PdfReviewHighlight primary,
            PdfReviewHighlight overlay)
        {
            var descriptions = new List<string>
                { Loc.T("review.source.interactions") };
            var seen = new HashSet<string>(StringComparer.CurrentCulture);
            AppendHighlightAccessibility(descriptions, seen, primary);
            AppendHighlightAccessibility(descriptions, seen, overlay);
            _highlightDescription = string.Join(". ", descriptions.ToArray());
            UpdateSelectionAccessibility();
        }

        private static void AppendHighlightAccessibility(List<string> descriptions,
            HashSet<string> seen, PdfReviewHighlight highlight)
        {
            if (highlight == null)
                return;
            string ownership = Loc.T(highlight.Style == PdfReviewHighlightStyle.Added
                ? "review.legend.added" : "review.legend.removed");
            if (seen.Add(ownership))
                descriptions.Add(ownership);
            foreach (PdfReviewWhitespaceMarker marker in highlight.WhitespaceMarkers)
                if (marker != null && !string.IsNullOrWhiteSpace(
                    marker.AccessibleDescription) && seen.Add(marker.AccessibleDescription))
                    descriptions.Add(marker.AccessibleDescription);
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

        internal static int RenderWidthFor(int viewportWidth)
        {
            return RenderWidthFor(viewportWidth, 1200);
        }

        internal static int RenderWidthFor(int viewportWidth, int limit)
        {
            int width = viewportWidth > 0 ? viewportWidth * 2 : 640;
            int cap = Math.Max(640, Math.Min(RasterBudget.MaxRenderDimension, limit));
            if (width < 640) return 640;
            return width > cap ? cap : width;
        }

        private void RenderLoop()
        {
            try
            {
                while (true)
                {
                    RenderRequest request;
                    lock (_renderGate)
                    {
                        request = _pending;
                        _pending = null;
                        if (request == null)
                            return; // _renderWorker меняется только в finally
                    }

                    RenderOutcome rendered = null;
                    PdfReviewViewContent content = request.Content;
                    try
                    {
                        if (content != null && content.BasePage != null)
                            rendered = RenderContent(request.Session, content,
                                request.RenderWidth);
                    }
                    catch
                    {
                        // Сбой композиции не публикует страницу без пометок: «недоступно».
                        if (rendered != null) rendered.Dispose();
                        rendered = null;
                    }
                    RenderOutcome ready = rendered;
                    if (!Ui.OnUi(this, delegate
                    {
                        ApplyRender(request.Generation, content, ready, request.Position);
                    }))
                    {
                        if (ready != null) ready.Dispose();
                    }
                }
            }
            finally
            {
                bool restart;
                lock (_renderGate)
                {
                    restart = _pending != null;
                    _renderWorker = restart;
                }
                if (restart && !IsDisposed)
                    Ui.RunWorker(RenderLoop);
            }
        }

        /// <summary>Итог фонового рендера: один владелец растра и полосы, нарисованные на нём.</summary>
        internal sealed class RenderOutcome : IDisposable
        {
            public BudgetedBitmap Bitmap;
            public List<PdfReviewChangeBar> Bars = new List<PdfReviewChangeBar>();
            // Файл больше бюджета показа: растра нет, но и повторять бессмысленно.
            public bool TooLarge;

            public void Dispose()
            {
                if (Bitmap != null)
                    Bitmap.Dispose();
                Bitmap = null;
            }
        }

        /// <summary>
        /// Рендер через общую сессию: под её замком — только обращение к рендереру, а
        /// композиция подсветки идёт уже без замка, не задерживая соседнюю панель.
        /// </summary>
        private static RenderOutcome RenderContent(PdfReviewRenderSession session,
            PdfReviewViewContent content, int width)
        {
            if (session == null)
                return null;
            PdfPageRef page = content.BasePage;
            if (session.IsTooLargeToShow(page.SourcePath, width))
                return new RenderOutcome { TooLarge = true };
            BudgetedBitmap raster = session.Use(page.SourcePath, delegate(PdfThumbnailRenderer renderer)
            {
                return renderer.RenderOwned(page.SourcePath, page.PageIndex, width,
                    RasterBudget.MaxRenderDimension, RasterBudget.PreviewPixels);
            });
            RenderOutcome outcome = ComposeHighlighted(raster, page, content.BaseHighlight);
            if (outcome != null && content.IsComposite)
            {
                try
                {
                    PdfReviewUnifiedRenderer.DrawDeletedMarkers(outcome.Bitmap.Bitmap,
                        content.OverlayHighlight);
                }
                catch
                {
                    outcome.Dispose();
                    throw;
                }
            }
            return outcome;
        }

        private static RenderOutcome ComposeHighlighted(BudgetedBitmap rendered,
            PdfPageRef page, PdfReviewHighlight highlight)
        {
            if (rendered == null)
                return null;
            try
            {
                Bitmap bitmap = rendered.Bitmap;
                if (page.Rotation != 0)
                    bitmap.RotateFlip(PageRotation.FlipFor(page.Rotation));
                var outcome = new RenderOutcome { Bitmap = rendered };
                bool hasMarks = highlight != null &&
                    (highlight.Boxes.Count > 0 || highlight.WhitespaceMarkers.Count > 0);
                if (!hasMarks)
                    return outcome;
                long extra = PdfMemoryBudget.EstimateBitmapBytes(bitmap.Width, bitmap.Height);
                if (!rendered.TryGrow(extra))
                {
                    rendered.Dispose();
                    return null;
                }
                Bitmap drawn = DrawHighlight(bitmap, highlight, SystemInformation.HighContrast,
                    out outcome.Bars);
                if (drawn == null)
                {
                    // Растр без подсветки скрыл бы правки: показываем «недоступно».
                    rendered.Dispose();
                    return null;
                }
                rendered.ReplaceAfterTransform(drawn);
                return outcome;
            }
            catch
            {
                rendered.Dispose();
                throw;
            }
        }

        /// <summary>Перегрузка с явным high-contrast флагом нужна для детерминированной проверки.</summary>
        internal static Bitmap DrawHighlight(Bitmap source, PdfReviewHighlight highlight,
            bool highContrast)
        {
            List<PdfReviewChangeBar> bars;
            return DrawHighlight(source, highlight, highContrast, out bars);
        }

        /// <summary>
        /// Композиция подсветки. Успех освобождает source и возвращает новую 32bpp-копию.
        /// Сбой GDI+/памяти возвращает null и оставляет source вызывающему: страница без
        /// пометок выглядела бы «правок нет», поэтому вызывающий показывает «недоступно».
        /// bars — фактически нарисованные полосы изменений (пиксели растра) для клика мыши.
        /// </summary>
        internal static Bitmap DrawHighlight(Bitmap source, PdfReviewHighlight highlight,
            bool highContrast, out List<PdfReviewChangeBar> bars)
        {
            bars = new List<PdfReviewChangeBar>();
            if (source == null || highlight == null)
                return source;
            List<RectangleF> rectangles = HighlightRectangles(highlight,
                source.Width, source.Height);
            List<WhitespaceDrawInfo> whitespace = WhitespaceDrawInfos(highlight,
                source.Width, source.Height);
            if (rectangles.Count == 0 && whitespace.Count == 0)
                return source;
            Bitmap copy = null;
            try
            {
                copy = PdfReviewHighlightRenderer.Create32BppCopy(source);
                if (!highContrast)
                    PdfReviewHighlightRenderer.ApplyWordFills(copy, rectangles, highlight.Color);
                // Раскладка полос — до Graphics: поиск чистого поля читает пиксели LockBits.
                List<PdfReviewChangeBar> laidOut = LayoutChangeBars(highlight,
                    copy.Width, copy.Height, copy);
                Color color, edgeColor;
                ResolveHighlightColors(highlight, highContrast, out color, out edgeColor);
                using (Graphics graphics = Graphics.FromImage(copy))
                {
                    graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    PrepareWhitespacePlacements(graphics, copy, whitespace);
                    if (highContrast)
                        DrawHighContrastOutlines(graphics, rectangles, highlight.Style,
                            color, edgeColor);
                    DrawChangeBars(graphics, copy, laidOut, highlight, edgeColor, color);
                    DrawWhitespaceMarkers(graphics, highlight, whitespace, edgeColor,
                        highContrast);
                }
                source.Dispose();
                bars = laidOut;
                return copy;
            }
            catch
            {
                if (copy != null)
                    copy.Dispose();
                return null;
            }
        }

        private static void DrawHighContrastOutlines(Graphics g, IList<RectangleF> rectangles,
            PdfReviewHighlightStyle style, Color color, Color edgeColor)
        {
            using (var edge = new Pen(Color.FromArgb(255, edgeColor), 3f))
            using (var primary = new Pen(Color.FromArgb(255, color), 1.5f))
            {
                ConfigurePattern(edge, style);
                ConfigurePattern(primary, style);
                edge.LineJoin = LineJoin.Round;
                primary.LineJoin = LineJoin.Round;
                foreach (RectangleF rect in rectangles)
                {
                    RectangleF outline = rect;
                    outline.Inflate(3f, 3f);
                    g.DrawRectangle(edge, outline.X, outline.Y, outline.Width, outline.Height);
                    g.DrawRectangle(primary, outline.X, outline.Y, outline.Width, outline.Height);
                }
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
            var bands = new List<RectangleF>();
            foreach (PdfReviewChangeBar bar in LayoutChangeBars(highlight, pixelWidth,
                pixelHeight, null))
                bands.Add(bar.Band);
            return bands;
        }

        private static List<RectangleF> HighlightRectangles(PdfReviewHighlight highlight,
            int pixelWidth, int pixelHeight)
        {
            var result = new List<RectangleF>();
            foreach (BandItem item in WordItems(highlight, pixelWidth, pixelHeight))
                result.Add(item.Rect);
            return result;
        }

        private struct BandItem
        {
            public RectangleF Rect;
            public PdfReviewWord Word;
        }

        /// <summary>Пиксельные рамки слов подсветки; Words[i] — владелец Boxes[i].</summary>
        private static List<BandItem> WordItems(PdfReviewHighlight highlight,
            int pixelWidth, int pixelHeight)
        {
            var result = new List<BandItem>();
            if (highlight == null || highlight.Boxes.Count == 0 || pixelWidth <= 0 ||
                pixelHeight <= 0 || highlight.ViewWidthPt <= 0 ||
                highlight.ViewHeightPt <= 0)
                return result;
            for (int i = 0; i < highlight.Boxes.Count; i++)
            {
                RectangleF rect = PdfReviewGeometry.ToPixelRect(highlight.Boxes[i],
                    highlight.ViewWidthPt, highlight.ViewHeightPt, pixelWidth, pixelHeight);
                if (rect.Width >= 1 && rect.Height >= 1)
                    result.Add(new BandItem
                    {
                        Rect = rect,
                        Word = i < highlight.Words.Count ? highlight.Words[i] : null
                    });
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
            IList<WhitespaceDrawInfo> whitespace)
        {
            if (graphics == null || source == null || whitespace == null)
                return;
            float preferred = Math.Max(10f, Math.Min(18f, source.Width / 70f));
            var placed = new List<RectangleF>();
            foreach (WhitespaceDrawInfo item in whitespace)
            {
                if (item == null || item.Marker == null)
                    continue;
                item.Bounds = AwayFromPlaced(VisibleWhitespaceBounds(graphics, source,
                    item.Marker.Text, item.Anchor, preferred), placed, source.Height);
                item.FontSize = preferred;
                placed.Add(item.Bounds);
            }
        }

        /// <summary>
        /// «−» и «+» замены пробела стоят на одной границе слов: следующий маркер уходит над
        /// уже нарисованным (у верхнего края — под него), иначе символы ложатся друг на друга.
        /// Попыток не больше, чем маркеров: плотная колонка не зацикливает раскладку.
        /// </summary>
        private static RectangleF AwayFromPlaced(RectangleF bounds, IList<RectangleF> placed,
            int height)
        {
            for (int attempt = 0; attempt <= placed.Count; attempt++)
            {
                int hit = -1;
                for (int i = 0; i < placed.Count && hit < 0; i++)
                    if (placed[i].IntersectsWith(bounds))
                        hit = i;
                if (hit < 0)
                    break;
                float above = placed[hit].Top - bounds.Height - 1f;
                bounds.Y = above >= 2f ? above
                    : Math.Max(2f, Math.Min(height - bounds.Height - 2f, placed[hit].Bottom + 1f));
            }
            return bounds;
        }

        private static RectangleF VisibleWhitespaceBounds(Graphics graphics, Bitmap source,
            string text, RectangleF anchor, float fontSize)
        {
            float width, height;
            using (var font = new Font(FontFamily.GenericSansSerif, fontSize,
                FontStyle.Bold, GraphicsUnit.Pixel))
            {
                SizeF measured = graphics.MeasureString(text, font,
                    PointF.Empty, StringFormat.GenericTypographic);
                width = (float)Math.Ceiling(measured.Width) + 6f;
                height = (float)Math.Ceiling(measured.Height) + 4f;
            }
            width = Math.Min(Math.Max(8f, width), Math.Max(8f, source.Width - 4f));
            height = Math.Min(Math.Max(8f, height), Math.Max(8f, source.Height - 4f));
            float x = anchor.Left + anchor.Width / 2f - width / 2f;
            float y = anchor.Top + anchor.Height / 2f - height / 2f;
            x = Math.Max(2f, Math.Min(source.Width - width - 2f, x));
            y = Math.Max(2f, Math.Min(source.Height - height - 2f, y));
            return new RectangleF(x, y, width, height);
        }

        /// <summary>
        /// Единственная раскладка полос изменений: строки из рамок слов и пробельных
        /// маркеров (одна реализация SameTextLine), рейка у внешнего поля в чистом месте.
        /// Её рисует композитор, и по ней же поверхность ловит клик, поэтому зона попадания
        /// совпадает с нарисованной рейкой при любом масштабе. clearance — растр для поиска
        /// чистого поля (null — поле считается чистым).
        /// </summary>
        internal static List<PdfReviewChangeBar> LayoutChangeBars(PdfReviewHighlight highlight,
            int pixelWidth, int pixelHeight, Bitmap clearance)
        {
            var bars = new List<PdfReviewChangeBar>();
            if (highlight == null || highlight.ChangeBarSide == PdfReviewChangeBarSide.None)
                return bars;
            List<BandItem> items = WordItems(highlight, pixelWidth, pixelHeight);
            foreach (WhitespaceDrawInfo info in WhitespaceDrawInfos(highlight,
                pixelWidth, pixelHeight))
                items.Add(new BandItem { Rect = info.Anchor });
            if (items.Count == 0)
                return bars;
            items.Sort(delegate(BandItem left, BandItem right)
            {
                int top = left.Rect.Top.CompareTo(right.Rect.Top);
                return top != 0 ? top : left.Rect.Left.CompareTo(right.Rect.Left);
            });

            PdfReviewChangeBar current = null;
            foreach (BandItem item in items)
            {
                if (current != null && PdfReviewGeometry.SameTextLine(current.Band, item.Rect))
                    current.Band = RectangleF.Union(current.Band, item.Rect);
                else
                {
                    current = new PdfReviewChangeBar { Band = item.Rect };
                    bars.Add(current);
                }
                if (item.Word != null)
                    current.Words.Add(item.Word);
            }

            PaperProbe probe = PaperProbe.ForSide(clearance, highlight.ChangeBarSide);
            float inset = Math.Max(5f, Math.Min(12f, pixelWidth * 0.008f));
            float symbolSize = RailSymbolSize(pixelWidth);
            for (int i = bars.Count - 1; i >= 0; i--)
                if (!PlaceRail(bars[i], highlight.ChangeBarSide, pixelWidth, pixelHeight,
                    inset, symbolSize, probe))
                    bars.RemoveAt(i);
            return bars;
        }

        private static float RailSymbolSize(int pixelWidth)
        {
            return Math.Max(10f, Math.Min(18f, pixelWidth / 70f));
        }

        private static bool PlaceRail(PdfReviewChangeBar bar, PdfReviewChangeBarSide side,
            int width, int height, float inset, float symbolSize, PaperProbe probe)
        {
            RectangleF band = bar.Band;
            float padding = Math.Max(2f, band.Height * 0.18f);
            float top = Math.Max(5f, band.Top - padding);
            float bottom = Math.Min(height - 5f, band.Bottom + padding);
            if (bottom - top < 5f)
            {
                float middle = (top + bottom) / 2f;
                top = Math.Max(5f, middle - 2.5f);
                bottom = Math.Min(height - 5f, middle + 2.5f);
            }
            if (bottom <= top)
                return false;
            bar.Top = top;
            bar.Bottom = bottom;
            bar.X = FindClearRailX(probe, side, width, inset, top, bottom);
            // Зона клика — нарисованная рейка (перо 5 px) вместе со знаком −/+ рядом с ней.
            float glyph = symbolSize * 0.75f + 2f;
            float middleY = (top + bottom) / 2f;
            float hitTop = Math.Min(top, middleY - symbolSize * 0.6f);
            float hitBottom = Math.Max(bottom, middleY + symbolSize * 0.6f);
            bar.HitBounds = side == PdfReviewChangeBarSide.Left
                ? RectangleF.FromLTRB(bar.X - 3f, hitTop, bar.X + 5f + glyph, hitBottom)
                : RectangleF.FromLTRB(bar.X - 5f - glyph, hitTop, bar.X + 3f, hitBottom);
            return true;
        }

        private static int OuterRailLimit(int width)
        {
            return Math.Max(5, Math.Min(80, (int)Math.Round(width * 0.10)));
        }

        private static float FindClearRailX(PaperProbe probe, PdfReviewChangeBarSide side,
            int width, float preferredInset, float top, float bottom)
        {
            int outerLimit = OuterRailLimit(width);
            int preferred = side == PdfReviewChangeBarSide.Left
                ? (int)Math.Round(preferredInset)
                : width - 1 - (int)Math.Round(preferredInset);
            for (int offset = 0; offset <= outerLimit; offset += 2)
            {
                int x = side == PdfReviewChangeBarSide.Left ? 3 + offset :
                    width - 4 - offset;
                if (Math.Abs(x - preferred) > outerLimit)
                    continue;
                if (probe == null || probe.IsClear(new RectangleF(x - 3f, top, 6f, bottom - top)))
                    return x;
            }
            return Math.Max(3f, Math.Min(width - 4f, preferred));
        }

        /// <summary>
        /// Пиксели полосы у внешнего поля, прочитанные одним LockBits. Поиск чистого места
        /// под рейку раньше звал GetPixel на каждый пиксель каждого кандидата — до ~10⁶
        /// вызовов на страницу; теперь это чтение из готового буфера.
        /// </summary>
        private sealed class PaperProbe
        {
            private const int InkAlphaMinimum = 24;
            private readonly byte[] _pixels;
            private readonly int _left;
            private readonly int _span;
            private readonly int _bitmapWidth;
            private readonly int _height;

            private PaperProbe(byte[] pixels, int left, int span, int bitmapWidth, int height)
            {
                _pixels = pixels;
                _left = left;
                _span = span;
                _bitmapWidth = bitmapWidth;
                _height = height;
            }

            internal static PaperProbe ForSide(Bitmap bitmap, PdfReviewChangeBarSide side)
            {
                if (bitmap == null || bitmap.Width <= 0 || bitmap.Height <= 0 ||
                    side == PdfReviewChangeBarSide.None)
                    return null;
                // Кандидаты X лежат в OuterRailLimit от края, полоса проверки — ±3 px.
                int span = Math.Min(bitmap.Width, OuterRailLimit(bitmap.Width) + 8);
                int left = side == PdfReviewChangeBarSide.Left ? 0 : bitmap.Width - span;
                int rowBytes = checked(span * 4);
                var pixels = new byte[checked(rowBytes * bitmap.Height)];
                BitmapData data = bitmap.LockBits(new Rectangle(left, 0, span, bitmap.Height),
                    ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                try
                {
                    for (int y = 0; y < bitmap.Height; y++)
                        Marshal.Copy(IntPtr.Add(data.Scan0, checked(y * data.Stride)),
                            pixels, y * rowBytes, rowBytes);
                }
                finally { bitmap.UnlockBits(data); }
                return new PaperProbe(pixels, left, span, bitmap.Width, bitmap.Height);
            }

            /// <summary>Нет ли в области чернил (тот же порог, что у прежнего GetPixel-обхода).</summary>
            internal bool IsClear(RectangleF area)
            {
                if (area.Left < 0 || area.Top < 0 || area.Right > _bitmapWidth ||
                    area.Bottom > _height)
                    return false;
                int left = Math.Max(0, (int)Math.Floor(area.Left));
                int top = Math.Max(0, (int)Math.Floor(area.Top));
                int right = Math.Min(_bitmapWidth - 1, (int)Math.Ceiling(area.Right));
                int bottom = Math.Min(_height - 1, (int)Math.Ceiling(area.Bottom));
                if (left < _left || right >= _left + _span)
                    return false;
                for (int y = top; y <= bottom; y++)
                {
                    int row = y * _span * 4;
                    for (int x = left; x <= right; x++)
                    {
                        int offset = row + (x - _left) * 4;
                        if (_pixels[offset + 3] >= InkAlphaMinimum &&
                            (_pixels[offset + 2] < PdfReviewHighlightRenderer.PaperMinimum ||
                             _pixels[offset + 1] < PdfReviewHighlightRenderer.PaperMinimum ||
                             _pixels[offset] < PdfReviewHighlightRenderer.PaperMinimum))
                            return false;
                    }
                }
                return true;
            }
        }

        internal static string HighlightSymbol(PdfReviewHighlightStyle style)
        {
            return style == PdfReviewHighlightStyle.Added ? "+" : "−";
        }

        /// <summary>Рисует ровно разложенные LayoutChangeBars полосы: рейку и знак −/+.</summary>
        internal static void DrawChangeBars(Graphics graphics, Bitmap raster,
            IList<PdfReviewChangeBar> bars, PdfReviewHighlight highlight, Color edgeColor,
            Color color)
        {
            if (graphics == null || raster == null || bars == null || bars.Count == 0 ||
                highlight == null || highlight.ChangeBarSide == PdfReviewChangeBarSide.None)
                return;
            string symbol = HighlightSymbol(highlight.Style);
            using (var edge = new Pen(Color.FromArgb(255, edgeColor), 5f))
            using (var primary = new Pen(Color.FromArgb(255, color), 2.25f))
            using (var font = new Font(FontFamily.GenericSansSerif, RailSymbolSize(raster.Width),
                FontStyle.Bold, GraphicsUnit.Pixel))
            {
                ConfigurePattern(edge, highlight.Style);
                ConfigurePattern(primary, highlight.Style);
                edge.StartCap = edge.EndCap = LineCap.Round;
                primary.StartCap = primary.EndCap = LineCap.Round;
                foreach (PdfReviewChangeBar bar in bars)
                {
                    graphics.DrawLine(edge, bar.X, bar.Top, bar.X, bar.Bottom);
                    graphics.DrawLine(primary, bar.X, bar.Top, bar.X, bar.Bottom);

                    SizeF measured = graphics.MeasureString(symbol, font,
                        PointF.Empty, StringFormat.GenericTypographic);
                    float sx = highlight.ChangeBarSide == PdfReviewChangeBarSide.Left
                        ? bar.X + 5f : bar.X - measured.Width - 5f;
                    sx = Math.Max(1f, Math.Min(raster.Width - measured.Width - 1f, sx));
                    float sy = Math.Max(1f, Math.Min(raster.Height - measured.Height - 1f,
                        (bar.Top + bar.Bottom - measured.Height) / 2f));
                    DrawOutlinedText(graphics, symbol, font, new PointF(sx, sy),
                        edgeColor, color);
                }
            }
        }

        private static void DrawWhitespaceMarkers(Graphics graphics,
            PdfReviewHighlight highlight, IList<WhitespaceDrawInfo> whitespace,
            Color edgeColor, bool highContrast)
        {
            if (graphics == null || whitespace == null)
                return;
            foreach (WhitespaceDrawInfo item in whitespace)
            {
                if (item == null || item.Marker == null || item.Bounds.IsEmpty ||
                    item.FontSize <= 0)
                    continue;
                PdfReviewHighlightStyle style = MarkerStyle(item.Marker,
                    highlight == null ? PdfReviewHighlightStyle.Added : highlight.Style);
                Color markerColor = style == PdfReviewHighlightStyle.Removed
                    ? Theme.ReviewDeleteMarker : Theme.ReviewInsertMarker;
                if (highContrast)
                    markerColor = edgeColor;
                using (var font = new Font(FontFamily.GenericSansSerif, item.FontSize,
                    FontStyle.Bold, GraphicsUnit.Pixel))
                    DrawOutlinedText(graphics, item.Marker.Text, font,
                        new PointF(item.Bounds.Left + 3f, item.Bounds.Top + 2f),
                        markerColor, markerColor);
            }
        }

        private static PdfReviewHighlightStyle MarkerStyle(
            PdfReviewWhitespaceMarker marker, PdfReviewHighlightStyle fallback)
        {
            if (marker == null || string.IsNullOrEmpty(marker.Text))
                return fallback;
            if (marker.Text.StartsWith("−", StringComparison.Ordinal))
                return PdfReviewHighlightStyle.Removed;
            if (marker.Text.StartsWith("+", StringComparison.Ordinal))
                return PdfReviewHighlightStyle.Added;
            return marker.Style;
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

        /// <summary>
        /// Публикация завершённого рендера на UI-потоке: растр, trusted-слой и полосы
        /// принимаются вместе и только для текущих generation и content revision.
        /// </summary>
        private void ApplyRender(int generation, PdfReviewViewContent content,
            RenderOutcome outcome, PdfReviewPagePosition position)
        {
            long contentRevision = content == null ? 0L : content.Revision;
            if (generation != _generation || contentRevision != _targetContentRevision ||
                IsDisposed)
            {
                if (outcome != null) outcome.Dispose();
                return;
            }
            DisposeBitmap();
            BudgetedBitmap rendered = outcome == null ? null : outcome.Bitmap;
            PdfPageRef request = content == null ? null : content.BasePage;
            _ownedBitmap = rendered;
            _bitmap = rendered == null ? null : rendered.Bitmap;
            _page = request;
            _reviewPage = null;
            _targetPage = request == null ? null : request.Clone();
            _caption = content == null ? "" : content.Caption;
            if (_bitmap == null)
                _state = outcome != null && outcome.TooLarge
                    ? PdfReviewPageViewState.TooLargeToShow : PdfReviewPageViewState.Unavailable;
            else
            {
                _reviewPage = content.BaseReviewPage;
                _picture.Image = _bitmap;
                _picture.AttachPage(_reviewPage, content.BaseHighlight, outcome.Bars);
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
                    case PdfReviewPageViewState.TooLargeToShow:
                        _message.Text = Loc.T("review.source.tooLarge");
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
            else if (_state == PdfReviewPageViewState.TooLargeToShow)
                _status.Text = JoinStatus(_caption, Loc.T("review.source.tooLarge"));
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
        /// Разобрать WM_MOUSEWHEEL, адресованный окну owner или его дочернему контролу.
        /// Фильтр Application глобален для UI-потока: без проверки владельца колесо над
        /// полноэкранным окном, диалогом или другим окном приложения листало бы скрытую
        /// под ним панель. Точка и Ctrl берутся из самого сообщения (lParam — экранные
        /// координаты со знаком, MK_CONTROL в wParam), а не из текущего состояния мыши.
        /// </summary>
        internal static bool TryReadWheel(Form owner, ref Message m, out Point screenPoint,
            out int delta, out bool controlDown)
        {
            screenPoint = Point.Empty;
            delta = 0;
            controlDown = false;
            if (m.Msg != WmMouseWheel || owner == null || owner.IsDisposed ||
                !owner.IsHandleCreated)
                return false;
            if (m.HWnd != owner.Handle)
            {
                Control target = Control.FromChildHandle(m.HWnd);
                if (target == null || target.FindForm() != owner)
                    return false;
            }
            long wParam = m.WParam.ToInt64();
            long lParam = m.LParam.ToInt64();
            delta = (short)(wParam >> 16);
            controlDown = (wParam & MkControl) != 0;
            screenPoint = new Point((short)(lParam & 0xFFFF), (short)((lParam >> 16) & 0xFFFF));
            return true;
        }

        /// <summary>
        /// Поглощает ли окно колесо после HandleWheel. Ctrl+колесо над готовой страницей —
        /// всегда жест масштаба: на пределе 25%/400% ступень не меняется, но отпущенное
        /// сообщение ушло бы в AutoScroll панели и молча прокрутило бы страницу.
        /// </summary>
        internal bool ConsumesWheel(PdfReviewWheelResult outcome, bool controlDown)
        {
            return outcome == PdfReviewWheelResult.Zoomed ||
                outcome == PdfReviewWheelResult.Scrolled ||
                (controlDown && HasVisiblePage);
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
            if (_ownedBitmap != null)
                _ownedBitmap.Dispose();
            _ownedBitmap = null;
            _bitmap = null;
            _reviewPage = null;
        }

        /// <summary>
        /// Бюджет освободился (закрыт документ, вытеснен растр): страница, которой не хватило
        /// памяти, пробует ещё раз сама — не дожидаясь, пока пользователь кликнет по строке.
        /// Событие приходит с любого потока; повтор — на UI-потоке и только из «недоступно».
        /// </summary>
        private void OnMemoryReleased()
        {
            Ui.OnUi(this, RetryUnavailable);
        }

        private void RetryUnavailable()
        {
            if (IsDisposed || _state != PdfReviewPageViewState.Unavailable ||
                _currentContent == null || _currentContent.BasePage == null ||
                _memoryRetries >= MaxMemoryRetries)
                return;
            _memoryRetries++;
            ShowContentCore(_currentContent, _currentPosition);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                PdfMemoryBudget.MemoryReleased -= OnMemoryReleased;
                _generation++;
                CancelPendingRender();
                DisposeBitmap();
                if (_ownsSession)
                    _session.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
