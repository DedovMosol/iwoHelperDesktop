using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    internal enum PdfReviewViewMode { Unified, SideBySide }

    /// <summary>
    /// Сравнение двух выбранных пользователем born-digital PDF. Инструмент read-only:
    /// ранняя версия всегда слева с красным фоном удалений, поздняя справа с зелёным
    /// фоном добавлений. Подсветка и счётчики строятся из одного document-wide результата.
    /// </summary>
    public sealed class PdfReviewForm : PdfToolFormBase, IFileAcceptor, IUnsavedStateAware, IMessageFilter
    {
        private const int WmMouseWheel = 0x020A;
        private static string Title { get { return Loc.T("hub.review.name"); } }

        private TextBox _leftPath, _rightPath;
        private Label _leftPathLabel, _rightPathLabel;
        private TextBox _leftPageInput, _rightPageInput;
        private Label _leftPageRange, _rightPageRange;
        private Label _leftLegend, _rightLegend;
        private Label _leftWhitespaceLegend, _rightWhitespaceLegend;
        private Button _pickLeft, _pickRight, _swap, _compare;
        private SplitContainer _body, _sourceSplit;
        private Panel _viewHost, _unifiedHost;
        private RoundedButton _unifiedModeButton, _sideModeButton;
        private PdfReviewViewMode _viewMode = PdfReviewViewMode.Unified;
        private ListBox _pairs;
        private PdfReviewPageView _leftSource, _rightSource, _unifiedSource;
        private Label _summary, _position;
        private Button _previous, _next, _manual;
        private PdfReviewResult _result;
        private long _contentRevision;
        private int _pairIndex = -1;
        private int _leftRowIndex = -1, _rightRowIndex = -1;
        private int _pathTop; // ордината ряда выбора файлов: пересчитывается в OnResize
        private string _leftFile, _rightFile;
        private bool _pathTextSync;
        private bool _pageTextSync, _leftPageDirty, _rightPageDirty;
        private int _leftSourceGeneration, _rightSourceGeneration;
        private bool _leftSourceChecking, _rightSourceChecking;
        private bool _leftProbeActive, _rightProbeActive;
        private bool _leftProbeQueued, _rightProbeQueued;
        private PdfReviewSourceError _leftSourceError, _rightSourceError;
        private bool _sourceCallbacksStopped;
        private bool _wheelFilterRegistered;
        private int _navigationSide; // 0 — обе, 1 — только left, 2 — только right
        private PdfReviewPagePosition _navigationPosition = PdfReviewPagePosition.Default;
        private readonly HashSet<Control> _dropWired = new HashSet<Control>();

        public PdfReviewForm() : this(null) { }

        public PdfReviewForm(Action showHub) : base(showHub)
        {
            BuildUi();
            SyncControls();
        }

        public bool HasUncommittedState
        {
            get
            {
                return !string.IsNullOrWhiteSpace(_leftPath == null ? null : _leftPath.Text) ||
                       !string.IsNullOrWhiteSpace(_rightPath == null ? null : _rightPath.Text) ||
                       _result != null;
            }
        }

        private void BuildReviewHeader()
        {
            MenuStrip menu = HelpMenu.Create(this, ShowHelp);
            MainMenuStrip = menu;
            Controls.Add(menu);
            var header = new HeaderBand(Title, Loc.T("review.header.subtitle"),
                Theme.ReviewBlue, Theme.ReviewBlueDark);
            header.SetBounds(0, HelpMenu.Height, ClientSize.Width, 76);
            header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(header);
            Ui.HomeOnHeader(header, _showHub, _tips, 22);
        }

        protected override string ToolTitle { get { return Title; } }

        protected override string IdleStatusText()
        {
            if (_result != null)
                return string.Format(Loc.T("review.status.ready"), _result.Pairs.Count,
                    _result.Stats == null ? 0 : _result.Stats.ChangedPages);
            return Loc.T("review.status.pickBoth");
        }

        private void BuildUi()
        {
            InitShell(Title, new Size(1120, 760), new Size(900, 640), Theme.ReviewBlue);
            BuildReviewHeader();
            int menu = HelpMenu.Height;
            int top = menu + 106;
            _pathTop = top;
            int right = ClientSize.Width - 20;

            // Границы ряда задаёт LayoutPathRow (единая точка на построение и ресайз):
            // якоря ширину полей не меняют, и при сжатии окна колонки наезжали друг на друга.
            _leftPathLabel = Ui.Label(this, Loc.T("review.left"), 20, top - 22, Font, Theme.TextPrimary);
            _rightPathLabel = Ui.Label(this, Loc.T("review.right"), 20, top - 22, Font, Theme.TextPrimary);
            _leftPathLabel.AutoSize = _rightPathLabel.AutoSize = false;
            _leftPathLabel.TextAlign = _rightPathLabel.TextAlign = ContentAlignment.MiddleLeft;
            _leftPath = PathBox(20, top, 300, Loc.T("review.left"));
            WirePathBox(_leftPath, true);
            _leftPath.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            _pickLeft = Secondary(20, top - 1, 100, Loc.T("common.browse"));
            _pickLeft.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            _pickLeft.Click += delegate { PickSide(true); };
            _rightPath = PathBox(20, top, 300, Loc.T("review.right"));
            WirePathBox(_rightPath, false);
            _rightPath.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _pickRight = Secondary(20, top - 1, 100, Loc.T("common.browse"));
            _pickRight.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _pickRight.Click += delegate { PickSide(false); };
            LayoutPathRow();

            _swap = Secondary(20, top + 34, 190, Loc.T("review.swap"));
            _swap.Click += delegate { SwapSides(); };
            _compare = new RoundedButton(true);
            _compare.Text = Loc.T("review.compare");
            _compare.SetBounds(right - 190, top + 32, 190, 34);
            _compare.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _compare.Click += delegate { StartComparison(); };
            Controls.Add(_compare);
            RegisterActionButton(_compare);
            AcceptButton = _compare;

            // Строка статуса — МЕЖДУ кнопками ряда: раньше она начиналась на 184 и залезала
            // под кнопку «Поменять местами» (та доходит до 210). Оба края считаются от
            // кнопок, поэтому наложение невозможно при любой ширине окна.
            int summaryLeft = _swap.Right + 10;
            _summary = Ui.Ellipsize(Ui.Label(this, Loc.T("review.status.pickBoth"),
                summaryLeft, top + 39, Font, Theme.TextMuted),
                _compare.Left - 10 - summaryLeft,
                AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
            _lblStatus = _summary;

            _body = new SplitContainer();
            _body.SetBounds(20, top + 80, right - 20, ClientSize.Height - (top + 80) - 112);
            _body.Orientation = Orientation.Vertical;
            _body.FixedPanel = FixedPanel.None;
            _body.IsSplitterFixed = false;
            _body.Panel1MinSize = 160;
            _body.Panel2MinSize = 420;
            _body.SplitterWidth = 8;
            _body.SplitterIncrement = 1;
            _body.SplitterDistance = 250;
            _body.BackColor = SystemInformation.HighContrast
                ? SystemColors.ControlDark : Theme.Border;
            _body.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(_body);
            _tips.SetToolTip(_body, Loc.T("review.pairs.resizeTip"));

            _pairs = new ListBox();
            _pairs.Dock = DockStyle.Fill;
            _pairs.IntegralHeight = false;
            _pairs.Font = Font;
            _pairs.SelectedIndexChanged += delegate { SelectPair(_pairs.SelectedIndex); };
            // После независимой навигации список показывает строку активной стороны.
            // Повторный клик по уже выбранной строке снова синхронизирует обе стороны.
            _pairs.MouseClick += delegate
            {
                if (_pairs.SelectedIndex >= 0)
                    SelectPair(_pairs.SelectedIndex);
            };
            _body.Panel1.Controls.Add(_pairs);

            BuildViews();
            BuildProgress(right);
            BuildNavigation(right);
            WireReviewDropTree(this);
        }

        private void BuildProgress(int right)
        {
            int y = ClientSize.Height - 88;
            _progress = new ProgressBar();
            _progress.SetBounds(20, y, right - 20 - 52, 16);
            _progress.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _progress.Minimum = 0;
            _progress.Maximum = 100;
            _progress.Visible = false;
            Controls.Add(_progress);

            _progressPct = Ui.Label(this, "", right - 46, y, Font, Theme.TextMuted);
            _progressPct.AutoSize = false;
            _progressPct.SetBounds(right - 46, y, 46, 16);
            _progressPct.TextAlign = ContentAlignment.MiddleRight;
            _progressPct.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _progressPct.Visible = false;
        }

        private TextBox PathBox(int x, int y, int width, string accessible)
        {
            var box = new TextBox();
            box.BackColor = Color.White;
            box.AccessibleName = accessible;
            box.AccessibleDescription = Loc.T("review.source.inputDescription");
            box.SetBounds(x, y, width, 27);
            box.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _tips.SetToolTip(box, Loc.T("review.source.inputDescription"));
            Controls.Add(box);
            return box;
        }

        private void WirePathBox(TextBox box, bool left)
        {
            box.TextChanged += delegate
            {
                if (!_pathTextSync)
                    InvalidateSource(left);
            };
            box.Leave += delegate
            {
                if (!Working)
                    CommitSource(left, false);
            };
        }

        private Button Secondary(int x, int y, int width, string text)
        {
            var button = new RoundedButton(false);
            button.Text = text;
            button.SetBounds(x, y, width, 30);
            Controls.Add(button);
            return button;
        }

        /// <summary>
        /// Unified redline is the default; side-by-side remains one click away. Both modes use
        /// the same semantic result and page renderer—switching never recomputes the diff.
        /// </summary>
        private void BuildViews()
        {
            var modeBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 44,
                BackColor = SystemInformation.HighContrast
                    ? SystemColors.Control : Color.White
            };
            _unifiedModeButton = new RoundedButton(false);
            _unifiedModeButton.SetBounds(8, 6, 150, 32);
            _unifiedModeButton.Text = Loc.T("review.mode.unified");
            _unifiedModeButton.AccessibleName = Loc.T("review.mode.unified");
            _unifiedModeButton.Click += delegate
            {
                SetViewMode(PdfReviewViewMode.Unified);
            };
            modeBar.Controls.Add(_unifiedModeButton);

            _sideModeButton = new RoundedButton(false);
            _sideModeButton.SetBounds(166, 6, 150, 32);
            _sideModeButton.Text = Loc.T("review.mode.sideBySide");
            _sideModeButton.AccessibleName = Loc.T("review.mode.sideBySide");
            _sideModeButton.Click += delegate
            {
                SetViewMode(PdfReviewViewMode.SideBySide);
            };
            modeBar.Controls.Add(_sideModeButton);

            _viewHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Size = new Size(600, 400)
            };
            _viewHost.Resize += delegate { LayoutViewHosts(); };
            _sourceSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 8,
                IsSplitterFixed = false,
                Size = new Size(600, 400),
                SplitterDistance = 300,
                Panel1MinSize = 220,
                Panel2MinSize = 220,
                BackColor = SystemInformation.HighContrast
                    ? SystemColors.ControlDark : Theme.Border
            };
            _leftSource = new PdfReviewPageView { Dock = DockStyle.Fill };
            _rightSource = new PdfReviewPageView { Dock = DockStyle.Fill };
            _sourceSplit.Panel1.Controls.Add(BuildPageHost(_leftSource, true));
            _sourceSplit.Panel2.Controls.Add(BuildPageHost(_rightSource, false));
            _viewHost.Controls.Add(_sourceSplit);

            _unifiedHost = BuildUnifiedHost();
            _viewHost.Controls.Add(_unifiedHost);

            _body.Panel2.Controls.Add(_viewHost);
            _body.Panel2.Controls.Add(modeBar);
            modeBar.BringToFront();
            _leftSource.ShowEmpty(Loc.T("review.left"));
            _rightSource.ShowEmpty(Loc.T("review.right"));
            _unifiedSource.ShowEmpty(Loc.T("review.mode.unified"));
            _tips.SetToolTip(_leftSource, Loc.T("review.source.interactions"));
            _tips.SetToolTip(_rightSource, Loc.T("review.source.interactions"));
            _tips.SetToolTip(_unifiedSource, Loc.T("review.unified.interactions"));
            LayoutViewHosts();
            SetViewMode(PdfReviewViewMode.Unified);
        }

        private Panel BuildUnifiedHost()
        {
            var host = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Size = new Size(600, 400)
            };
            var legend = new Panel
            {
                Dock = DockStyle.Top,
                Size = new Size(600, 72),
                Height = 72,
                BackColor = SystemInformation.HighContrast
                    ? SystemColors.Window : Color.White
            };
            Label title = Ui.Label(legend, Loc.T("review.unified.legend"), 10, 8,
                Ui.Font(Font.Size, FontStyle.Bold), Theme.TextPrimary);
            title.AutoSize = false;
            title.SetBounds(10, 7, 520, 24);
            title.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Label hint = Ui.Label(legend, Loc.T("review.unified.hint"), 10, 34,
                Ui.Font(Math.Max(8f, Font.Size - 0.5f)), Theme.TextMuted);
            hint.AutoSize = false;
            hint.SetBounds(10, 33, 520, 32);
            hint.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            _unifiedSource = new PdfReviewPageView { Dock = DockStyle.Fill };
            host.Controls.Add(_unifiedSource);
            host.Controls.Add(legend);
            legend.BringToFront();
            return host;
        }

        private void LayoutViewHosts()
        {
            if (_viewHost == null)
                return;
            Rectangle bounds = _viewHost.ClientRectangle;
            if (_sourceSplit != null) _sourceSplit.Bounds = bounds;
            if (_unifiedHost != null) _unifiedHost.Bounds = bounds;
        }

        private void SetViewMode(PdfReviewViewMode mode)
        {
            _viewMode = mode;
            bool unified = mode == PdfReviewViewMode.Unified;
            if (_unifiedHost != null) _unifiedHost.Visible = unified;
            if (_sourceSplit != null) _sourceSplit.Visible = !unified;
            if (_unifiedModeButton != null)
            {
                _unifiedModeButton.Selected = unified;
                _unifiedModeButton.AccessibleDescription = unified
                    ? Loc.T("review.mode.selected") : "";
            }
            if (_sideModeButton != null)
            {
                _sideModeButton.Selected = !unified;
                _sideModeButton.AccessibleDescription = !unified
                    ? Loc.T("review.mode.selected") : "";
            }
            if (_result == null || _pairIndex < 0)
                return;
            if (unified)
                RenderUnified(PdfReviewPagePosition.Default);
            else
            {
                CenterSourceSplitter();
                RenderSource(true, PdfReviewPagePosition.Default);
                RenderSource(false, PdfReviewPagePosition.Default);
            }
        }

        private void CenterSourceSplitter()
        {
            if (_sourceSplit == null || _sourceSplit.ClientSize.Width <=
                _sourceSplit.SplitterWidth)
                return;
            int available = _sourceSplit.ClientSize.Width - _sourceSplit.SplitterWidth;
            int wanted = available / 2;
            int min = _sourceSplit.Panel1MinSize;
            int max = available - _sourceSplit.Panel2MinSize;
            if (max >= min)
                _sourceSplit.SplitterDistance = Math.Max(min, Math.Min(max, wanted));
        }


        private Panel BuildPageHost(PdfReviewPageView source, bool leftSide)
        {
            var host = new Panel { Dock = DockStyle.Fill, TabStop = false };
            var top = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 116,
                RowCount = 2,
                ColumnCount = 1,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = SystemInformation.HighContrast
                    ? SystemColors.Window : Color.White,
                TabStop = false
            };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            top.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            top.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var navigator = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                Padding = new Padding(8, 6, 8, 5),
                BackColor = top.BackColor,
                TabStop = false
            };
            navigator.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            navigator.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 62));
            navigator.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            navigator.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var label = new Label
            {
                Text = Loc.T("review.page.label"),
                AutoSize = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = Font,
                ForeColor = SystemInformation.HighContrast
                    ? SystemColors.WindowText : Theme.TextPrimary,
                Margin = new Padding(0, 3, 6, 0),
                TabStop = false
            };
            var input = new TextBox
            {
                BackColor = SystemInformation.HighContrast
                    ? SystemColors.Window : Color.White,
                ForeColor = SystemInformation.HighContrast
                    ? SystemColors.WindowText : Theme.TextPrimary,
                Dock = DockStyle.Fill,
                MaxLength = 10,
                TextAlign = HorizontalAlignment.Center,
                AccessibleName = Loc.T(leftSide
                    ? "review.page.left.accessible" : "review.page.right.accessible"),
                AccessibleDescription = Loc.T("review.page.description"),
                Margin = new Padding(0, 0, 6, 0),
                TabIndex = 0
            };
            var range = new Label
            {
                AutoEllipsis = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = Font,
                ForeColor = SystemInformation.HighContrast
                    ? SystemColors.WindowText : Theme.TextMuted,
                Margin = new Padding(0, 3, 0, 0),
                TabStop = false
            };
            input.TextChanged += delegate
            {
                if (!_pageTextSync)
                {
                    if (leftSide) _leftPageDirty = true;
                    else _rightPageDirty = true;
                }
            };
            input.Leave += delegate
            {
                if (!Working && (leftSide ? _leftPageDirty : _rightPageDirty))
                    NavigateToPhysicalPage(leftSide);
            };
            _tips.SetToolTip(input, Loc.T("review.page.description"));
            navigator.Controls.Add(label, 0, 0);
            navigator.Controls.Add(input, 1, 0);
            navigator.Controls.Add(range, 2, 0);

            Label ownership;
            Label whitespace;
            TableLayoutPanel legend = BuildLegend(leftSide, out ownership, out whitespace);
            top.Controls.Add(navigator, 0, 0);
            top.Controls.Add(legend, 0, 1);

            if (leftSide)
            {
                _leftPageInput = input;
                _leftPageRange = range;
                _leftLegend = ownership;
                _leftWhitespaceLegend = whitespace;
            }
            else
            {
                _rightPageInput = input;
                _rightPageRange = range;
                _rightLegend = ownership;
                _rightWhitespaceLegend = whitespace;
            }
            source.TabIndex = 1;
            host.Controls.Add(source);
            host.Controls.Add(top);
            top.BringToFront();
            return host;
        }

        private TableLayoutPanel BuildLegend(bool leftSide, out Label ownership,
            out Label whitespace)
        {
            Color background = SystemInformation.HighContrast
                ? SystemColors.Window : Color.White;
            Color foreground = SystemInformation.HighContrast
                ? SystemColors.WindowText
                : (leftSide ? Theme.ReviewDeleteMarker : Theme.ReviewInsertMarker);
            var legend = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                Padding = new Padding(8, 2, 8, 4),
                BackColor = background,
                TabStop = false
            };
            legend.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            legend.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            legend.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            ownership = new Label
            {
                Text = Loc.T(leftSide ? "review.legend.removed" :
                    "review.legend.added"),
                Dock = DockStyle.Fill,
                AutoSize = false,
                Font = Ui.Font(Font.Size, FontStyle.Bold),
                ForeColor = foreground,
                BackColor = background,
                TextAlign = ContentAlignment.MiddleLeft,
                AccessibleName = Loc.T(leftSide ? "review.legend.removed" :
                    "review.legend.added"),
                TabStop = false
            };
            whitespace = new Label
            {
                Text = Loc.T("review.legend.whitespace"),
                Dock = DockStyle.Fill,
                AutoSize = false,
                Font = Ui.Font(Math.Max(8f, Font.Size - 0.5f)),
                ForeColor = SystemInformation.HighContrast
                    ? SystemColors.WindowText : Theme.TextMuted,
                BackColor = background,
                TextAlign = ContentAlignment.MiddleLeft,
                AccessibleName = Loc.T("review.legend.whitespace"),
                TabStop = false
            };
            legend.Controls.Add(ownership, 0, 0);
            legend.Controls.Add(whitespace, 0, 1);
            return legend;
        }

        /// <summary>
        /// Ряд выбора файлов: два поля с кнопками делят окно пополам, у каждой половины
        /// кнопка «Обзор…» внутри. Одна точка расчёта на построение и на ресайз (см.
        /// OnResize): раньше ширина задавалась один раз при создании и якоря её не трогали,
        /// поэтому при сжатии окна колонки наезжали друг на друга.
        /// </summary>
        private void LayoutPathRow()
        {
            int right = ClientSize.Width - 20;
            int gap = 24;
            int browseWidth = 100;
            int browseGap = 6;
            // Ширина окна может временно упасть ниже минимальной (пересборка, смена
            // масштаба) — не даём расчёту уйти в отрицательную ширину.
            int half = Math.Max(140, (right - 20 - gap) / 2);
            int leftPathWidth = Math.Max(60, half - browseGap - browseWidth);

            _leftPath.SetBounds(20, _pathTop, leftPathWidth, _leftPath.Height);
            _pickLeft.SetBounds(20 + leftPathWidth + browseGap, _pathTop - 1, browseWidth, _pickLeft.Height);
            if (_leftPathLabel != null)
                _leftPathLabel.SetBounds(20, _pathTop - 22, leftPathWidth, 18);

            int rightStart = 20 + half + gap;
            int rightPathWidth = Math.Max(60, right - rightStart - browseGap - browseWidth);
            _rightPath.SetBounds(rightStart, _pathTop, rightPathWidth, _rightPath.Height);
            _pickRight.SetBounds(rightStart + rightPathWidth + browseGap, _pathTop - 1, browseWidth, _pickRight.Height);
            if (_rightPathLabel != null)
                _rightPathLabel.SetBounds(rightStart, _pathTop - 22, rightPathWidth, 18);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (_leftPath == null)
                return; // InitShell задал ClientSize до того, как контролы построены
            LayoutPathRow();
        }

        private void BuildNavigation(int right)
        {
            int y = ClientSize.Height - 52;
            _previous = Secondary(20, y, 150, Loc.T("review.previous"));
            _previous.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            _previous.Click += delegate { MoveChange(-1); };
            _next = Secondary(180, y, 150, Loc.T("review.next"));
            _next.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            _next.Click += delegate { MoveChange(1); };
            _manual = Secondary(right - 230, y, 230, Loc.T("review.manualPair"));
            _manual.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _manual.Click += delegate { ManualPair(); };
            _position = Ui.Label(this, "", 342, y + 7, Font, Theme.TextMuted);
            _position.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
        }

        public void AcceptFiles(string[] paths)
        {
            if (Working || paths == null || paths.Length == 0) return;
            string[] pdfs = ChoiceCard.FilterByExtension(paths, PdfDrop.PdfOnly);
            ExecuteDropPlan(PdfReviewInput.PlanDrop(pdfs, PdfReviewDropTarget.Neutral,
                HasSourceText(true), HasSourceText(false)));
        }

        private void PickSide(bool left)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = Loc.T("common.pdfFilter");
                dialog.Title = left ? Loc.T("review.pickLeft") : Loc.T("review.pickRight");
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    AssignSource(left, dialog.FileName);
            }
        }

        /// <summary>
        /// Любой ввод — клавиатура, browse или drop — сначала сбрасывает прежний resolved path.
        /// Поэтому видимый текст никогда не расходится со скрытым источником Compare.
        /// </summary>
        private void InvalidateSource(bool left)
        {
            if (left)
            {
                _leftSourceGeneration++;
                _leftSourceChecking = false;
                _leftSourceError = PdfReviewSourceError.None;
                _leftFile = null;
            }
            else
            {
                _rightSourceGeneration++;
                _rightSourceChecking = false;
                _rightSourceError = PdfReviewSourceError.None;
                _rightFile = null;
            }
            ClearResult();
            SyncControls();
        }

        private void AssignSource(bool left, string text)
        {
            SetSourceText(left, text);
            InvalidateSource(left);
            CommitSource(left, false);
        }

        private void AssignSources(string left, string right)
        {
            _pathTextSync = true;
            try
            {
                _leftPath.Text = left ?? "";
                _rightPath.Text = right ?? "";
            }
            finally { _pathTextSync = false; }
            InvalidateSource(true);
            InvalidateSource(false);
            if (!string.IsNullOrWhiteSpace(_leftPath.Text)) CommitSource(true, false);
            if (!string.IsNullOrWhiteSpace(_rightPath.Text)) CommitSource(false, false);
        }

        private void SetSourceText(bool left, string text)
        {
            TextBox box = left ? _leftPath : _rightPath;
            _pathTextSync = true;
            try
            {
                box.Text = text ?? "";
                box.SelectionStart = box.TextLength;
            }
            finally { _pathTextSync = false; }
        }

        private void CommitSource(bool left, bool showEmptyError)
        {
            if (Working || _sourceCallbacksStopped) return;
            TextBox box = left ? _leftPath : _rightPath;
            string current = left ? _leftFile : _rightFile;
            if (!string.IsNullOrEmpty(current) &&
                string.Equals(box.Text, current, StringComparison.OrdinalIgnoreCase))
                return;

            int generation;
            if (left)
            {
                generation = ++_leftSourceGeneration;
                _leftSourceChecking = false;
                _leftSourceError = PdfReviewSourceError.None;
                _leftFile = null;
            }
            else
            {
                generation = ++_rightSourceGeneration;
                _rightSourceChecking = false;
                _rightSourceError = PdfReviewSourceError.None;
                _rightFile = null;
            }
            ClearResult();

            PdfReviewSourceResult source = PdfReviewInput.Resolve(box.Text);
            if (!source.IsValid)
            {
                ShowSourceError(left, source.Error, showEmptyError);
                SyncControls();
                return;
            }

            // Probe is deliberately single-flight per side. A fast sequence of edits/Leave
            // events queues only the latest source instead of reading several large PDFs in
            // parallel; the generation check still rejects every stale completion.
            bool active = left ? _leftProbeActive : _rightProbeActive;
            if (active)
            {
                if (left) _leftProbeQueued = true;
                else _rightProbeQueued = true;
                if (left) _leftSourceChecking = true;
                else _rightSourceChecking = true;
                SetStatus(Loc.T("review.status.checkingSource"), Theme.TextMuted);
                SyncControls();
                return;
            }
            if (left) _leftProbeActive = true;
            else _rightProbeActive = true;
            if (left) _leftSourceChecking = true;
            else _rightSourceChecking = true;
            SetStatus(Loc.T("review.status.checkingSource"), Theme.TextMuted);
            SyncControls();
            Ui.RunWorker(delegate
            {
                PdfReviewSourceResult probed = PdfReviewInput.Probe(source);
                OnUi(delegate { CompleteSource(left, generation, probed); });
            });
        }

        private void CompleteSource(bool left, int generation, PdfReviewSourceResult source)
        {
            // Probe завершается в фоне. Закрытие формы инвалидирует поколения до Dispose,
            // поэтому уже поставленный в очередь UI-callback не трогает закрытые контролы.
            bool currentGeneration = (left ? _leftSourceGeneration : _rightSourceGeneration) == generation;
            if (left) _leftProbeActive = false;
            else _rightProbeActive = false;
            if (_sourceCallbacksStopped || IsDisposed)
                return;
            if (!currentGeneration)
            {
                bool queued = left ? _leftProbeQueued : _rightProbeQueued;
                if (queued)
                {
                    if (left) _leftProbeQueued = false;
                    else _rightProbeQueued = false;
                    CommitSource(left, false);
                }
                return;
            }
            if (left) _leftSourceChecking = false;
            else _rightSourceChecking = false;

            if (source == null || !source.IsValid)
            {
                if (left) _leftFile = null;
                else _rightFile = null;
                ShowSourceError(left,
                    source == null ? PdfReviewSourceError.Unreadable : source.Error, true);
                SyncControls();
                return;
            }

            if (left)
            {
                _leftFile = source.Path;
                _leftSourceError = PdfReviewSourceError.None;
            }
            else
            {
                _rightFile = source.Path;
                _rightSourceError = PdfReviewSourceError.None;
            }
            SetSourceText(left, source.Path);
            UpdateSourceStatus();
            SyncControls();
        }

        private void ShowSourceError(bool left, PdfReviewSourceError error, bool showEmpty)
        {
            if (left) _leftSourceError = error;
            else _rightSourceError = error;
            if (error == PdfReviewSourceError.Empty && !showEmpty)
            {
                UpdateSourceStatus();
                return;
            }
            SetStatus(SourceErrorText(error), Theme.ErrRed);
        }

        private static string SourceErrorText(PdfReviewSourceError error)
        {
            switch (error)
            {
                case PdfReviewSourceError.Empty: return Loc.T("review.err.source.empty");
                case PdfReviewSourceError.InvalidPath: return Loc.T("review.err.source.invalidPath");
                case PdfReviewSourceError.Missing: return Loc.T("review.err.source.missing");
                case PdfReviewSourceError.NotPdf: return Loc.T("review.err.source.notPdf");
                default: return Loc.T("review.err.source.unreadable");
            }
        }

        private void UpdateSourceStatus()
        {
            if (_result != null)
                return;
            PdfReviewSourceError error = VisibleSourceError;
            if (error != PdfReviewSourceError.None)
                SetStatus(SourceErrorText(error), Theme.ErrRed);
            else if (_leftSourceChecking || _rightSourceChecking)
                SetStatus(Loc.T("review.status.checkingSource"), Theme.TextMuted);
            else if (SameSourceSelected)
                SetStatus(Loc.T("review.err.sameFile"), Theme.ErrRed);
            else if (!string.IsNullOrEmpty(_leftFile) && !string.IsNullOrEmpty(_rightFile))
                SetStatus(Loc.T("review.status.sourcesReady"), Theme.OkGreen);
            else
                SetStatus(Loc.T("review.status.pickBoth"), Theme.TextMuted);
        }

        private PdfReviewSourceError VisibleSourceError
        {
            get
            {
                if (_leftSourceError != PdfReviewSourceError.None &&
                    _leftSourceError != PdfReviewSourceError.Empty)
                    return _leftSourceError;
                if (_rightSourceError != PdfReviewSourceError.None &&
                    _rightSourceError != PdfReviewSourceError.Empty)
                    return _rightSourceError;
                return PdfReviewSourceError.None;
            }
        }

        private bool HasSourceText(bool left)
        {
            TextBox box = left ? _leftPath : _rightPath;
            return box != null && !string.IsNullOrWhiteSpace(box.Text);
        }

        private bool SameSourceSelected
        {
            get
            {
                return !string.IsNullOrEmpty(_leftFile) && !string.IsNullOrEmpty(_rightFile) &&
                    OutputFile.IsSameFile(_leftFile, _rightFile);
            }
        }

        private void SwapSides()
        {
            if (Working) return;
            string left = _leftPath.Text;
            string right = _rightPath.Text;
            AssignSources(right, left);
        }

        /// <summary>
        /// WinForms посылает drag-событие самому глубокому контролу под указателем. Поэтому
        /// подписываем всё дерево (и поздно добавленные контролы), а сторону определяем по
        /// экранной точке: drop над реальной кнопкой/полем/viewport не теряется.
        /// </summary>
        private void WireReviewDropTree(Control root)
        {
            if (root == null || _dropWired.Contains(root))
                return;
            _dropWired.Add(root);
            root.AllowDrop = true;
            root.DragEnter += OnReviewDragEnter;
            root.DragOver += OnReviewDragOver;
            root.DragLeave += OnReviewDragLeave;
            root.DragDrop += OnReviewDragDrop;
            root.ControlAdded += delegate(object sender, ControlEventArgs e)
            {
                WireReviewDropTree(e.Control);
            };
            foreach (Control child in root.Controls)
                WireReviewDropTree(child);
        }

        private void OnReviewDragEnter(object sender, DragEventArgs e)
        {
            UpdateReviewDrag(e);
        }

        private void OnReviewDragOver(object sender, DragEventArgs e)
        {
            UpdateReviewDrag(e);
        }

        private void UpdateReviewDrag(DragEventArgs e)
        {
            string[] paths = Working ? new string[0] : PdfDrop.ExtractPaths(e, PdfDrop.PdfOnly);
            e.Effect = paths.Length == 0 ? DragDropEffects.None : DragDropEffects.Copy;
            PdfReviewDropPlan plan = paths.Length == 0 ? null :
                PdfReviewInput.PlanDrop(paths, ReviewDropTargetAt(new Point(e.X, e.Y)),
                    HasSourceText(true), HasSourceText(false));
            SetDropCue(plan);
        }

        private void OnReviewDragLeave(object sender, EventArgs e)
        {
            ClearDropCueOutside(Cursor.Position);
        }

        /// <summary>
        /// DragLeave дочернего control не означает уход с окна: при переходе между вложенными
        /// слоями сохраняем cue до немедленного DragEnter/DragOver нового слоя. Снимаем его
        /// только когда указатель действительно покинул клиентскую область Review.
        /// </summary>
        internal void ClearDropCueOutside(Point screenPoint)
        {
            if (!ContainsScreenPoint(this, screenPoint))
                SetDropCue(null);
        }

        private void OnReviewDragDrop(object sender, DragEventArgs e)
        {
            PdfReviewDropTarget target = ReviewDropTargetAt(new Point(e.X, e.Y));
            string[] paths = Working ? new string[0] : PdfDrop.ExtractPaths(e, PdfDrop.PdfOnly);
            PdfReviewDropPlan plan = paths.Length == 0 ? null :
                PdfReviewInput.PlanDrop(paths, target, HasSourceText(true), HasSourceText(false));
            SetDropCue(null);
            ExecuteDropPlan(plan);
        }

        internal PdfReviewDropTarget ReviewDropTargetAt(Point screenPoint)
        {
            if (ContainsScreenPoint(_leftPath, screenPoint) ||
                ContainsScreenPoint(_pickLeft, screenPoint) ||
                ContainsScreenPoint(_sourceSplit == null ? null : _sourceSplit.Panel1, screenPoint))
                return PdfReviewDropTarget.Left;
            if (ContainsScreenPoint(_rightPath, screenPoint) ||
                ContainsScreenPoint(_pickRight, screenPoint) ||
                ContainsScreenPoint(_sourceSplit == null ? null : _sourceSplit.Panel2, screenPoint))
                return PdfReviewDropTarget.Right;
            return PdfReviewDropTarget.Neutral;
        }

        private static bool ContainsScreenPoint(Control control, Point screenPoint)
        {
            return control != null && !control.IsDisposed && control.Visible &&
                control.IsHandleCreated &&
                control.RectangleToScreen(control.ClientRectangle).Contains(screenPoint);
        }

        /// <summary>
        /// Подсказка показывает фактический план, а не просто control под мышью: два файла
        /// подсвечивают обе стороны, neutral single-drop — первую пустую, неоднозначный drop —
        /// ни одну (иначе интерфейс обещал бы молчаливую замену, которой не будет).
        /// </summary>
        internal void SetDropCue(PdfReviewDropPlan plan)
        {
            if (_leftSource == null || _rightSource == null)
                return;
            PdfReviewDropAction action = plan == null ? PdfReviewDropAction.None : plan.Action;
            _leftSource.SetDropTarget(action == PdfReviewDropAction.AssignLeft ||
                action == PdfReviewDropAction.AssignBoth);
            _rightSource.SetDropTarget(action == PdfReviewDropAction.AssignRight ||
                action == PdfReviewDropAction.AssignBoth);
        }

        private void ExecuteDropPlan(PdfReviewDropPlan plan)
        {
            if (plan == null)
                return;
            switch (plan.Action)
            {
                case PdfReviewDropAction.AssignLeft:
                    AssignSource(true, plan.LeftPath);
                    break;
                case PdfReviewDropAction.AssignRight:
                    AssignSource(false, plan.RightPath);
                    break;
                case PdfReviewDropAction.AssignBoth:
                    AssignSources(plan.LeftPath, plan.RightPath);
                    break;
                case PdfReviewDropAction.NeedExplicitSide:
                    Dialogs.Info(this, Title, Loc.T("review.err.dropSide.title"),
                        Loc.T("review.err.dropSide.body"));
                    break;
                case PdfReviewDropAction.TooMany:
                    Dialogs.Info(this, Title, Loc.T("review.err.tooMany.title"),
                        Loc.T("review.err.tooMany.body"));
                    break;
            }
        }

        internal int DropWiredControlCount { get { return _dropWired.Count; } }

        private void StartComparison()
        {
            if (Working || string.IsNullOrEmpty(_leftFile) || string.IsNullOrEmpty(_rightFile)) return;
            if (OutputFile.IsSameFile(_leftFile, _rightFile))
            {
                Dialogs.Error(this, Title, Loc.T("review.err.sameFile.title"), Loc.T("review.err.sameFile"));
                return;
            }
            string left = _leftFile, right = _rightFile;
            // Два документа могут содержать сотни страниц: Cancel предлагаем всегда, несмотря
            // на две стороны, потому что единица долгой работы здесь — страница, а не файл.
            BeginOperation(Loc.T("review.status.comparing"), 5);
            Action<int, int> progress = UiProgress();
            Func<bool> cancel = CancelToken();
            Ui.RunWorker(delegate
            {
                PdfReviewResult result = null;
                Exception error = null;
                try { result = PdfReviewService.Compare(left, right, progress, cancel); }
                catch (Exception ex) { error = ex; }
                OnUi(delegate { CompleteComparison(error, result); });
            });
        }

        private void CompleteComparison(Exception error, PdfReviewResult result)
        {
            var reviewError = error as PdfReviewException;
            if (reviewError != null && reviewError.Reason == PdfReviewFailure.PasswordRequired &&
                !string.IsNullOrEmpty(reviewError.FilePath))
            {
                EndOperation();
                bool retry = !string.IsNullOrEmpty(PdfPasswords.For(reviewError.FilePath));
                string password = PasswordPromptDialog.Show(this,
                    Path.GetFileName(reviewError.FilePath), retry);
                if (!string.IsNullOrEmpty(password))
                {
                    PdfPasswords.Remember(reviewError.FilePath, password);
                    StartComparison();
                }
                else
                {
                    PdfPasswords.Remember(reviewError.FilePath, null);
                    SetStatus(Loc.T("review.err.password"), Theme.WarnOrange);
                }
                return;
            }
            if (!FinishOperation(error, Loc.T("review.status.failed"), Loc.T("review.err.failed")))
                return;
            UsageStats.RecordPdfCompare();
            ApplyResult(result);
        }

        private void ApplyResult(PdfReviewResult result)
        {
            unchecked { _contentRevision++; }
            _result = result;
            _pairIndex = _leftRowIndex = _rightRowIndex = -1;
            _navigationSide = 0;
            _navigationPosition = PdfReviewPagePosition.Default;
            _pairs.BeginUpdate();
            _pairs.Items.Clear();
            for (int i = 0; i < result.Pairs.Count; i++)
                _pairs.Items.Add(PairLabel(result.Pairs[i]));
            _pairs.EndUpdate();
            string stats = StatsText(result.Stats);
            _summary.Text = stats;
            _summary.AccessibleDescription = stats;
            _tips.SetToolTip(_summary, stats);
            if (_pairs.Items.Count > 0)
                _pairs.SelectedIndex = 0;
            else
            {
                ClearViews();
                SyncPageInputs();
            }
            RefreshRestingStatus();
            SyncControls();
        }

        private string PairLabel(PdfReviewPagePair pair)
        {
            string left = pair.LeftPageIndex >= 0 ? (pair.LeftPageIndex + 1).ToString() : "—";
            string right = pair.RightPageIndex >= 0 ? (pair.RightPageIndex + 1).ToString() : "—";
            string marker = pair.Status == PdfReviewPairStatus.Unchanged ? "=" :
                pair.Status == PdfReviewPairStatus.LeftOnly ? "−" :
                pair.Status == PdfReviewPairStatus.RightOnly ? "+" : "≠";
            return marker + "  " + left + "  ↔  " + right;
        }

        private string StatsText(PdfReviewStats s)
        {
            if (s == null) return "";
            return string.Format(Loc.T("review.stats"), s.ChangedPages,
                s.LeftOnlyPages, s.RightOnlyPages, s.DeletedWords, s.InsertedWords,
                s.ChangedPercent, s.WhitespaceChanges, s.DeletedWhitespaceAtoms,
                s.InsertedWhitespaceAtoms);
        }

        private void SelectPair(int index)
        {
            int side = _navigationSide;
            PdfReviewPagePosition position = _navigationPosition;
            _navigationSide = 0;
            _navigationPosition = PdfReviewPagePosition.Default;
            _pairIndex = index;
            if (_result == null || index < 0 || index >= _result.Pairs.Count)
            {
                _leftRowIndex = _rightRowIndex = -1;
                ClearViews();
                SyncPageInputs();
                SyncControls();
                return;
            }

            if (_viewMode == PdfReviewViewMode.Unified)
            {
                _leftRowIndex = _rightRowIndex = index;
                RenderUnified(position);
            }
            else if (side == 1)
            {
                _leftRowIndex = index;
                RenderSource(true, position);
            }
            else if (side == 2)
            {
                _rightRowIndex = index;
                RenderSource(false, position);
            }
            else
            {
                _leftRowIndex = _rightRowIndex = index;
                RenderSource(true, PdfReviewPagePosition.Default);
                RenderSource(false, PdfReviewPagePosition.Default);
            }
            _position.Text = string.Format(Loc.T("review.position"), index + 1,
                _result.Pairs.Count);
            SyncPageInputs();
            SyncControls();
        }

        /// <summary>
        /// Рендерит только запрошенную сторону. Wheel и поля страницы не перезагружают
        /// соседний pane, поэтому его физическая страница, масштаб и offset сохраняются.
        /// Клик по alignment-row передаёт side=0 через SelectPair и синхронизирует обе стороны.
        /// </summary>
        private void RenderSource(bool leftSide, PdfReviewPagePosition position)
        {
            if (_result == null)
                return;
            int rowIndex = leftSide ? _leftRowIndex : _rightRowIndex;
            if (rowIndex < 0 || rowIndex >= _result.Pairs.Count)
                return;
            PdfReviewPagePair pair = _result.Pairs[rowIndex];
            if (pair == null)
                return;
            int pageIndex = leftSide ? pair.LeftPageIndex : pair.RightPageIndex;
            PdfReviewDocument document = leftSide ? _result.Left : _result.Right;
            PdfPageRef page = PageRef(document, pageIndex);
            PdfReviewPage reviewPage = PdfReviewDiff.PageAt(document, pageIndex);
            PdfReviewPageView source = leftSide ? _leftSource : _rightSource;
            if (!source.IsShowing(page, _contentRevision))
                source.ShowPage(page, reviewPage, _contentRevision,
                    Loc.T(leftSide ? "review.left" : "review.right"),
                    BuildHighlight(_result, pair, leftSide), position);
        }

        private void RenderUnified(PdfReviewPagePosition position)
        {
            if (_result == null || _pairIndex < 0 ||
                _pairIndex >= _result.Pairs.Count)
                return;
            PdfReviewViewContent content = BuildUnifiedContent(_result,
                _result.Pairs[_pairIndex], _contentRevision);
            if (content == null)
            {
                _unifiedSource.ShowEmpty(Loc.T("review.mode.unified"));
                return;
            }
            if (!_unifiedSource.IsShowing(content.BasePage, content.Revision))
                _unifiedSource.ShowContent(content, position);
        }

        internal static PdfReviewViewContent BuildUnifiedContent(PdfReviewResult result,
            PdfReviewPagePair pair, long revision)
        {
            if (result == null || pair == null)
                return null;
            bool hasRight = pair.RightPageIndex >= 0;
            bool baseLeft = !hasRight;
            PdfReviewDocument baseDocument = baseLeft ? result.Left : result.Right;
            int baseIndex = baseLeft ? pair.LeftPageIndex : pair.RightPageIndex;
            PdfPageRef basePage = PageRef(baseDocument, baseIndex);
            PdfReviewPage baseReviewPage = PdfReviewDiff.PageAt(baseDocument, baseIndex);
            if (basePage == null || baseReviewPage == null)
                return null;
            PdfReviewHighlight baseHighlight = BuildHighlight(result, pair, baseLeft);

            PdfPageRef overlayPage = null;
            PdfReviewHighlight overlayHighlight = null;
            if (!baseLeft && pair.LeftPageIndex >= 0)
            {
                PdfReviewHighlight deleted = BuildHighlight(result, pair, true);
                if (deleted.Boxes.Count > 0 || deleted.WhitespaceMarkers.Count > 0)
                {
                    overlayPage = PageRef(result.Left, pair.LeftPageIndex);
                    overlayHighlight = deleted;
                }
            }
            string caption = string.Format(Loc.T("review.unified.caption"),
                pair.LeftPageIndex < 0 ? "—" : (pair.LeftPageIndex + 1).ToString(),
                pair.RightPageIndex < 0 ? "—" : (pair.RightPageIndex + 1).ToString());
            return new PdfReviewViewContent(basePage, baseReviewPage, baseHighlight,
                overlayPage, overlayHighlight, revision, caption);
        }


        private bool SelectViewerRow(int index, bool leftSide,
            PdfReviewPagePosition position)
        {
            if (_result == null || index < 0 || index >= _result.Pairs.Count)
                return false;
            _navigationSide = leftSide ? 1 : 2;
            _navigationPosition = position;
            if (_pairs.SelectedIndex == index)
                SelectPair(index);
            else
                _pairs.SelectedIndex = index;

            bool selected = (leftSide ? _leftRowIndex : _rightRowIndex) == index;
            // На случай, если WinForms не послал событие из-за уничтожения handle/очистки.
            _navigationSide = 0;
            _navigationPosition = PdfReviewPagePosition.Default;
            return selected;
        }

        private void SyncPageInputs()
        {
            if (_leftPageInput == null || _rightPageInput == null)
                return;
            _pageTextSync = true;
            try
            {
                _leftPageInput.Text = PhysicalPageText(true, _leftRowIndex);
                _rightPageInput.Text = PhysicalPageText(false, _rightRowIndex);
                _leftPageRange.Text = PageRangeText(_result == null || _result.Left == null
                    ? 0 : _result.Left.Pages.Count);
                _rightPageRange.Text = PageRangeText(_result == null || _result.Right == null
                    ? 0 : _result.Right.Pages.Count);
                _leftPageDirty = _rightPageDirty = false;
            }
            finally { _pageTextSync = false; }
        }

        private string PhysicalPageText(bool leftSide, int rowIndex)
        {
            if (_result == null || rowIndex < 0 || rowIndex >= _result.Pairs.Count)
                return "";
            PdfReviewPagePair pair = _result.Pairs[rowIndex];
            if (pair == null)
                return "";
            int pageIndex = leftSide ? pair.LeftPageIndex : pair.RightPageIndex;
            return pageIndex < 0 ? "" : (pageIndex + 1).ToString();
        }

        private static string PageRangeText(int pageCount)
        {
            return pageCount <= 0 ? "" : string.Format(Loc.T("review.page.of"), pageCount);
        }

        /// <summary>
        /// Подсветка — обратная проекция одного глобального semantic result на физическую
        /// страницу: Delete и подтверждённые пробельные удаления принадлежат ранней/левой
        /// стороне, Insert и пробельные добавления — поздней/правой. Viewer не диффит заново.
        /// </summary>
        internal static PdfReviewHighlight BuildHighlight(PdfReviewResult result,
            PdfReviewPagePair pair, bool leftSide)
        {
            var highlight = new PdfReviewHighlight
            {
                Color = leftSide ? Theme.ReviewDeleteFill : Theme.ReviewInsertFill,
                EdgeColor = leftSide ? Theme.ReviewDeleteMarker : Theme.ReviewInsertMarker,
                Style = leftSide ? PdfReviewHighlightStyle.Removed :
                    PdfReviewHighlightStyle.Added,
                // Каждая пометка живёт у ВНЕШНЕГО поля своей панели. Центральный просвет
                // не образует двусмысленную общую ось удаления/добавления.
                ChangeBarSide = leftSide
                    ? PdfReviewChangeBarSide.Left : PdfReviewChangeBarSide.Right
            };
            if (result == null || pair == null)
                return highlight;
            int pageIndex = leftSide ? pair.LeftPageIndex : pair.RightPageIndex;
            PdfReviewPage page = leftSide
                ? PdfReviewDiff.PageAt(result.Left, pageIndex)
                : PdfReviewDiff.PageAt(result.Right, pageIndex);
            if (page == null)
                return highlight;
            highlight.ViewWidthPt = page.ViewWidthPt;
            highlight.ViewHeightPt = page.ViewHeightPt;

            Dictionary<int, List<PdfReviewWord>> index = leftSide
                ? result.DeletedWordsByPage : result.InsertedWordsByPage;
            List<PdfReviewWord> words;
            if (index.TryGetValue(pageIndex, out words))
                foreach (PdfReviewWord word in words)
                    if (word != null && !IsStableFooterPageNumber(result, pair,
                        leftSide, page, word))
                    {
                        highlight.Boxes.Add(word.Box);
                        highlight.Words.Add(word);
                    }

            Dictionary<int, List<PdfReviewWhitespaceMarker>> whitespaceIndex = leftSide
                ? result.DeletedWhitespaceByPage : result.InsertedWhitespaceByPage;
            List<PdfReviewWhitespaceMarker> markers;
            if (whitespaceIndex.TryGetValue(pageIndex, out markers))
                foreach (PdfReviewWhitespaceMarker marker in markers)
                    if (marker != null)
                        highlight.WhitespaceMarkers.Add(marker);
            return highlight;
        }

        private static bool IsStableFooterPageNumber(PdfReviewResult result,
            PdfReviewPagePair pair, bool leftSide, PdfReviewPage page,
            PdfReviewWord word)
        {
            if (result == null || pair == null || page == null || word == null ||
                page.ViewHeightPt <= 0 || word.Box.Bottom < page.ViewHeightPt * 0.82)
                return false;
            string pageNumber = (page.PageIndex + 1).ToString(
                System.Globalization.CultureInfo.InvariantCulture);
            if (!string.Equals(word.Key, pageNumber, StringComparison.Ordinal) ||
                !string.Equals(word.Text, pageNumber, StringComparison.Ordinal))
                return false;

            int otherPageIndex = leftSide ? pair.RightPageIndex : pair.LeftPageIndex;
            PdfReviewDocument otherDocument = leftSide ? result.Right : result.Left;
            PdfReviewPage other = PdfReviewDiff.PageAt(otherDocument, otherPageIndex);
            if (other == null || other.ViewHeightPt <= 0)
                return false;
            double toleranceX = Math.Max(2.0,
                Math.Max(page.ViewWidthPt, other.ViewWidthPt) * 0.015);
            double toleranceY = Math.Max(2.0,
                Math.Max(page.ViewHeightPt, other.ViewHeightPt) * 0.015);
            foreach (PdfReviewWord candidate in other.Words)
            {
                if (candidate == null || !string.Equals(candidate.Key, pageNumber,
                    StringComparison.Ordinal) || !string.Equals(candidate.Text, pageNumber,
                    StringComparison.Ordinal) || candidate.Box.Bottom < other.ViewHeightPt * 0.82)
                    continue;
                return Math.Abs(word.Box.Left - candidate.Box.Left) <= toleranceX &&
                    Math.Abs(word.Box.Right - candidate.Box.Right) <= toleranceX &&
                    Math.Abs(word.Box.Bottom - candidate.Box.Bottom) <= toleranceY &&
                    Math.Abs(word.Box.Top - candidate.Box.Top) <= toleranceY;
            }
            return false;
        }

        private static PdfPageRef PageRef(PdfReviewDocument doc, int pageIndex)
        {
            return doc == null || pageIndex < 0 ? null : new PdfPageRef { SourcePath = doc.Path, PageIndex = pageIndex };
        }

        private void MoveChange(int direction)
        {
            if (_result == null || _result.Pairs.Count == 0) return;
            int start = _pairIndex;
            for (int step = 1; step <= _result.Pairs.Count; step++)
            {
                int index = (start + direction * step) % _result.Pairs.Count;
                if (index < 0) index += _result.Pairs.Count;
                if (_result.Pairs[index].Status != PdfReviewPairStatus.Unchanged)
                {
                    if (_pairs.SelectedIndex == index)
                        SelectPair(index);
                    else
                        _pairs.SelectedIndex = index;
                    return;
                }
            }
        }

        private void ManualPair()
        {
            if (_result == null || _pairIndex < 0) return;
            PdfReviewPagePair current = _result.Pairs[_pairIndex];
            int leftDefault = current.LeftPageIndex >= 0 ? current.LeftPageIndex + 1 : 1;
            int rightDefault = current.RightPageIndex >= 0 ? current.RightPageIndex + 1 : 1;
            int left = NumberPromptDialog.Show(this, Loc.T("review.manualPair"),
                string.Format(Loc.T("review.manual.left"), _result.Left.Pages.Count),
                Loc.T("common.ok"), 1, _result.Left.Pages.Count, leftDefault) - 1;
            if (left < 0) return;
            int right = NumberPromptDialog.Show(this, Loc.T("review.manualPair"),
                string.Format(Loc.T("review.manual.right"), _result.Right.Pages.Count),
                Loc.T("common.ok"), 1, _result.Right.Pages.Count, rightDefault) - 1;
            if (right < 0) return;
            List<PdfReviewPagePair> remapped = PdfReviewDiff.ApplyManualPair(_result.Pairs,
                _result.Left, _result.Right, left, right, PdfReviewLimits.Default());
            _result.Pairs.Clear(); _result.Pairs.AddRange(remapped);
            PdfReviewDiff.Project(_result);
            ApplyResult(_result);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (!_wheelFilterRegistered)
            {
                Application.AddMessageFilter(this);
                _wheelFilterRegistered = true;
            }
            if (_viewMode == PdfReviewViewMode.SideBySide)
                CenterSourceSplitter();
        }

        bool IMessageFilter.PreFilterMessage(ref Message m)
        {
            if (m.Msg != WmMouseWheel || !_wheelFilterRegistered || IsDisposed || !Visible)
                return false;
            int delta = (short)(m.WParam.ToInt64() >> 16);
            return RouteWheel(Cursor.Position, delta,
                (Control.ModifierKeys & Keys.Control) != 0);
        }

        internal bool RouteWheel(Point screenPoint, int delta, bool controlDown)
        {
            if (_viewMode == PdfReviewViewMode.Unified)
            {
                if (_unifiedSource == null ||
                    !_unifiedSource.ContainsViewport(screenPoint))
                    return false;
                PdfReviewWheelResult unifiedOutcome = _unifiedSource.HandleWheel(
                    screenPoint, delta, controlDown);
                if (unifiedOutcome == PdfReviewWheelResult.Zoomed ||
                    unifiedOutcome == PdfReviewWheelResult.Scrolled)
                    return true;
                if (unifiedOutcome != PdfReviewWheelResult.AtPreviousBoundary &&
                    unifiedOutcome != PdfReviewWheelResult.AtNextBoundary)
                    return false;
                int unifiedDirection = unifiedOutcome ==
                    PdfReviewWheelResult.AtPreviousBoundary ? -1 : 1;
                int unifiedNext = FindUnifiedRow(_result == null ? null : _result.Pairs,
                    _pairIndex, unifiedDirection);
                if (unifiedNext < 0)
                    return false;
                _navigationPosition = unifiedDirection < 0
                    ? PdfReviewPagePosition.Bottom : PdfReviewPagePosition.Top;
                if (_pairs.SelectedIndex == unifiedNext)
                    SelectPair(unifiedNext);
                else
                    _pairs.SelectedIndex = unifiedNext;
                return true;
            }

            bool leftSide;
            PdfReviewPageView view;
            if (_leftSource != null && _leftSource.ContainsViewport(screenPoint))
            {
                leftSide = true;
                view = _leftSource;
            }
            else if (_rightSource != null && _rightSource.ContainsViewport(screenPoint))
            {
                leftSide = false;
                view = _rightSource;
            }
            else
            {
                return false;
            }

            PdfReviewWheelResult outcome = view.HandleWheel(screenPoint, delta, controlDown);
            if (outcome == PdfReviewWheelResult.Zoomed ||
                outcome == PdfReviewWheelResult.Scrolled)
                return true;
            if (outcome != PdfReviewWheelResult.AtPreviousBoundary &&
                outcome != PdfReviewWheelResult.AtNextBoundary)
                return false;

            int direction = outcome == PdfReviewWheelResult.AtPreviousBoundary ? -1 : 1;
            int current = leftSide ? _leftRowIndex : _rightRowIndex;
            int next = FindViewerRow(_result == null ? null : _result.Pairs,
                current, leftSide, direction);
            if (next < 0)
                return false;
            return SelectViewerRow(next, leftSide, direction < 0
                ? PdfReviewPagePosition.Bottom : PdfReviewPagePosition.Top);
        }

        internal static int FindUnifiedRow(IList<PdfReviewPagePair> pairs, int current,
            int direction)
        {
            if (pairs == null || current < 0 || current >= pairs.Count || direction == 0)
                return -1;
            int step = direction < 0 ? -1 : 1;
            for (int index = current + step; index >= 0 && index < pairs.Count;
                index += step)
            {
                PdfReviewPagePair pair = pairs[index];
                if (pair != null && (pair.RightPageIndex >= 0 || pair.LeftPageIndex >= 0))
                    return index;
            }
            return -1;
        }


        internal static int FindViewerRow(IList<PdfReviewPagePair> pairs, int current,
            bool leftSide, int direction)
        {
            if (pairs == null || current < 0 || current >= pairs.Count || direction == 0)
                return -1;
            int step = direction < 0 ? -1 : 1;
            for (int index = current + step; index >= 0 && index < pairs.Count; index += step)
            {
                PdfReviewPagePair pair = pairs[index];
                if (pair != null && (leftSide ? pair.LeftPageIndex : pair.RightPageIndex) >= 0)
                    return index;
            }
            return -1;
        }

        /// <summary>
        /// Находит уже существующую строку, содержащую физическую страницу выбранной стороны.
        /// Ничего не сопоставляет и не меняет: семантический diff и presentation rows неизменны.
        /// </summary>
        internal static int FindViewerRowForPhysicalPage(IList<PdfReviewPagePair> pairs,
            bool leftSide, int physicalPageIndex)
        {
            if (pairs == null || physicalPageIndex < 0)
                return -1;
            for (int index = 0; index < pairs.Count; index++)
            {
                PdfReviewPagePair pair = pairs[index];
                if (pair != null && (leftSide ? pair.LeftPageIndex : pair.RightPageIndex) ==
                    physicalPageIndex)
                    return index;
            }
            return -1;
        }

        /// <summary>
        /// Переходит к введённой one-based физической странице только на выбранной стороне.
        /// Использует готовую alignment-row; ApplyManualPair/Project/DiffWords здесь невозможны.
        /// </summary>
        internal bool NavigateToPhysicalPage(bool leftSide)
        {
            if (Working || _result == null)
                return false;
            TextBox input = leftSide ? _leftPageInput : _rightPageInput;
            PdfReviewDocument document = leftSide ? _result.Left : _result.Right;
            int pageCount = document == null ? 0 : document.Pages.Count;
            int pageNumber;
            if (input == null || !int.TryParse((input.Text ?? "").Trim(), out pageNumber))
            {
                SetStatus(Loc.T("review.page.err.number"), Theme.ErrRed);
                return false;
            }
            if (pageNumber < 1 || pageNumber > pageCount)
            {
                SetStatus(string.Format(Loc.T("review.page.err.range"), pageCount),
                    Theme.ErrRed);
                return false;
            }
            int row = FindViewerRowForPhysicalPage(_result.Pairs, leftSide,
                pageNumber - 1);
            if (row < 0)
            {
                SetStatus(Loc.T("review.page.err.unavailable"), Theme.ErrRed);
                return false;
            }
            if (!SelectViewerRow(row, leftSide, PdfReviewPagePosition.Default))
                return false;
            SetStatus(StatsText(_result.Stats), Theme.TextMuted);
            return true;
        }

        internal bool WheelFilterRegistered { get { return _wheelFilterRegistered; } }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if (!e.Cancel)
                StopSourceCallbacks();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            StopSourceCallbacks();
            RemoveWheelFilter();
            base.OnFormClosed(e);
        }

        private void StopSourceCallbacks()
        {
            if (_sourceCallbacksStopped)
                return;
            _sourceCallbacksStopped = true;
            _leftSourceGeneration++;
            _rightSourceGeneration++;
            _leftSourceChecking = false;
            _rightSourceChecking = false;
            _leftProbeQueued = _rightProbeQueued = false;
        }

        private void RemoveWheelFilter()
        {
            if (!_wheelFilterRegistered)
                return;
            Application.RemoveMessageFilter(this);
            _wheelFilterRegistered = false;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter &&
                (_leftPageInput.Focused || _rightPageInput.Focused))
            {
                NavigateToPhysicalPage(_leftPageInput.Focused);
                return true; // поле страницы не запускает AcceptButton «Сравнить»
            }
            if (keyData == Keys.Enter && (_leftPath.Focused || _rightPath.Focused))
            {
                CommitSource(_leftPath.Focused, true);
                return true; // не отдаём Enter AcceptButton до успешной проверки обоих путей
            }
            if (keyData == Keys.F3) { MoveChange(1); return true; }
            if (keyData == (Keys.Shift | Keys.F3)) { MoveChange(-1); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void ClearResult()
        {
            unchecked { _contentRevision++; }
            _result = null;
            _pairIndex = _leftRowIndex = _rightRowIndex = -1;
            _navigationSide = 0;
            _navigationPosition = PdfReviewPagePosition.Default;
            _pairs.Items.Clear();
            _summary.Text = Loc.T("review.status.pickBoth");
            _summary.AccessibleDescription = _summary.Text;
            _tips.SetToolTip(_summary, _summary.Text);
            ClearViews();
            SyncPageInputs();
        }

        private void ClearViews()
        {
            _leftSource.ShowEmpty(Loc.T("review.left"));
            _rightSource.ShowEmpty(Loc.T("review.right"));
            _unifiedSource.ShowEmpty(Loc.T("review.mode.unified"));
            _position.Text = "";
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, Title, Loc.T("menu.howTo"), Loc.T("review.help.body"));
        }

        protected override void SyncControls()
        {
            bool checking = _leftSourceChecking || _rightSourceChecking;
            bool files = !string.IsNullOrEmpty(_leftFile) && !string.IsNullOrEmpty(_rightFile);
            _leftPath.Enabled = _rightPath.Enabled = !Working;
            _pickLeft.Enabled = _pickRight.Enabled = !Working;
            _swap.Enabled = !Working && (HasSourceText(true) || HasSourceText(false));
            _compare.Enabled = !Working && !checking && files && !SameSourceSelected;
            _pairs.Enabled = !Working && _result != null;
            _previous.Enabled = _next.Enabled = !Working && _result != null && _result.Pairs.Count > 0;
            _manual.Enabled = !Working && _result != null && _pairIndex >= 0;
            if (_unifiedModeButton != null) _unifiedModeButton.Enabled = !Working;
            if (_sideModeButton != null) _sideModeButton.Enabled = !Working;
            bool canNavigate = !Working && _result != null && _result.Pairs.Count > 0;
            _leftPageInput.Enabled = canNavigate && _result.Left != null &&
                _result.Left.Pages.Count > 0;
            _rightPageInput.Enabled = canNavigate && _result.Right != null &&
                _result.Right.Pages.Count > 0;
            _leftPageRange.Enabled = _leftPageInput.Enabled;
            _rightPageRange.Enabled = _rightPageInput.Enabled;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                StopSourceCallbacks();
                RemoveWheelFilter();
            }
            base.Dispose(disposing);
        }
    }
}
