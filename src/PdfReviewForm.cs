using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Сравнение двух выбранных пользователем born-digital PDF. Инструмент read-only:
    /// страницы бок о бок в исходном виде, удалённое подсвечено красным на левой странице,
    /// добавленное — зелёным на правой. Подсветка и счётчики строятся из одного ворд-диффа.
    /// </summary>
    public sealed class PdfReviewForm : PdfToolFormBase, IFileAcceptor, IUnsavedStateAware
    {
        private static string Title { get { return Loc.T("hub.review.name"); } }

        private TextBox _leftPath, _rightPath;
        private Button _pickLeft, _pickRight, _swap, _compare;
        private SplitContainer _body, _sourceSplit;
        private ListBox _pairs;
        private PdfReviewPageView _leftSource, _rightSource;
        private Label _summary, _position;
        private Button _previous, _next, _manual;
        private PdfReviewResult _result;
        private int _pairIndex = -1;
        private int _pathTop; // ордината ряда выбора файлов: пересчитывается в OnResize
        private string _leftFile, _rightFile;

        public PdfReviewForm() : this(null) { }

        public PdfReviewForm(Action showHub) : base(showHub)
        {
            BuildUi();
            SyncControls();
        }

        public bool HasUncommittedState
        {
            get { return !string.IsNullOrEmpty(_leftFile) || !string.IsNullOrEmpty(_rightFile) || _result != null; }
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
            WireFileDrop(AcceptFiles);
            int menu = HelpMenu.Height;
            int top = menu + 88;
            _pathTop = top;
            int right = ClientSize.Width - 20;

            // Границы ряда задаёт LayoutPathRow (единая точка на построение и ресайз):
            // якоря ширину полей не меняют, и при сжатии окна колонки наезжали друг на друга.
            _leftPath = PathBox(20, top, 300, Loc.T("review.left"));
            _leftPath.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            _pickLeft = Secondary(20, top - 1, 100, Loc.T("common.browse"));
            _pickLeft.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            _pickLeft.Click += delegate { PickSide(true); };
            _rightPath = PathBox(20, top, 300, Loc.T("review.right"));
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
            _body.FixedPanel = FixedPanel.Panel1;
            _body.Panel1MinSize = 190;
            _body.Panel2MinSize = 500;
            _body.SplitterDistance = 250;
            _body.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(_body);

            _pairs = new ListBox();
            _pairs.Dock = DockStyle.Fill;
            _pairs.IntegralHeight = false;
            _pairs.Font = Font;
            _pairs.SelectedIndexChanged += delegate { SelectPair(_pairs.SelectedIndex); };
            _body.Panel1.Controls.Add(_pairs);

            BuildViews();
            BuildProgress(right);
            BuildNavigation(right);
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
            box.ReadOnly = true;
            box.BackColor = Color.White;
            box.AccessibleName = accessible;
            box.SetBounds(x, y, width, 27);
            box.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(box);
            return box;
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
        /// Единственный вид сравнения — исходные страницы бок о бок с подсветкой:
        /// удалённое красным на левой, добавленное зелёным на правой.
        /// </summary>
        private void BuildViews()
        {
            _sourceSplit = new SplitContainer();
            _sourceSplit.Dock = DockStyle.Fill;
            _sourceSplit.Orientation = Orientation.Vertical;
            _leftSource = new PdfReviewPageView { Dock = DockStyle.Fill };
            _rightSource = new PdfReviewPageView { Dock = DockStyle.Fill };
            _sourceSplit.Panel1.Controls.Add(_leftSource);
            _sourceSplit.Panel2.Controls.Add(_rightSource);
            _body.Panel2.Controls.Add(_sourceSplit);
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

            int rightStart = 20 + half + gap;
            int rightPathWidth = Math.Max(60, right - rightStart - browseGap - browseWidth);
            _rightPath.SetBounds(rightStart, _pathTop, rightPathWidth, _rightPath.Height);
            _pickRight.SetBounds(rightStart + rightPathWidth + browseGap, _pathTop - 1, browseWidth, _pickRight.Height);
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
            if (pdfs.Length == 0) return;
            if (pdfs.Length > 2)
            {
                Dialogs.Info(this, Title, Loc.T("review.err.tooMany.title"),
                    Loc.T("review.err.tooMany.body"));
                return;
            }
            if (pdfs.Length == 2)
            {
                SetSide(true, pdfs[0]);
                SetSide(false, pdfs[1]);
            }
            else if (string.IsNullOrEmpty(_leftFile))
                SetSide(true, pdfs[0]);
            else
                SetSide(false, pdfs[0]);
        }

        private void PickSide(bool left)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = Loc.T("common.pdfFilter");
                dialog.Title = left ? Loc.T("review.pickLeft") : Loc.T("review.pickRight");
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    SetSide(left, dialog.FileName);
            }
        }

        private void SetSide(bool left, string path)
        {
            if (left) { _leftFile = path; _leftPath.Text = path; _leftPath.SelectionStart = path.Length; }
            else { _rightFile = path; _rightPath.Text = path; _rightPath.SelectionStart = path.Length; }
            ClearResult();
            SyncControls();
        }

        private void SwapSides()
        {
            string path = _leftFile;
            _leftFile = _rightFile;
            _rightFile = path;
            _leftPath.Text = _leftFile ?? "";
            _rightPath.Text = _rightFile ?? "";
            ClearResult();
            SyncControls();
        }

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
            ApplyResult(result);
        }

        private void ApplyResult(PdfReviewResult result)
        {
            _result = result;
            _pairs.BeginUpdate();
            _pairs.Items.Clear();
            for (int i = 0; i < result.Pairs.Count; i++)
                _pairs.Items.Add(PairLabel(result.Pairs[i]));
            _pairs.EndUpdate();
            _summary.Text = StatsText(result.Stats);
            if (_pairs.Items.Count > 0) _pairs.SelectedIndex = 0;
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
                s.LeftOnlyPages, s.RightOnlyPages, s.DeletedWords, s.InsertedWords, s.ChangedPercent);
        }

        private void SelectPair(int index)
        {
            _pairIndex = index;
            if (_result == null || index < 0 || index >= _result.Pairs.Count)
            {
                ClearViews();
                SyncControls();
                return;
            }
            RenderSources();
            _position.Text = string.Format(Loc.T("review.position"), index + 1, _result.Pairs.Count);
            SyncControls();
        }

        /// <summary>
        /// Показать пару страниц с подсветкой из ТОГО ЖЕ ворд-диффа, что и счётчики:
        /// слева — удалённые слова красным, справа — добавленные зелёным.
        /// </summary>
        private void RenderSources()
        {
            if (_result == null || _pairIndex < 0 || _pairIndex >= _result.Pairs.Count) return;
            PdfReviewPagePair pair = _result.Pairs[_pairIndex];
            _leftSource.ShowPage(PageRef(_result.Left, pair.LeftPageIndex),
                Loc.T("review.left"), HighlightFor(pair, true));
            _rightSource.ShowPage(PageRef(_result.Right, pair.RightPageIndex),
                Loc.T("review.right"), HighlightFor(pair, false));
        }

        /// <summary>
        /// Подсветка одной стороны пары. Страница без пары целиком того же цвета
        /// (вся удалена / вся добавлена); неизменённая пара — без подсветки.
        /// </summary>
        private PdfReviewHighlight HighlightFor(PdfReviewPagePair pair, bool leftSide)
        {
            var highlight = new PdfReviewHighlight
            {
                Color = leftSide ? Theme.ErrRed : Theme.OkGreen
            };
            if (pair == null || pair.Status == PdfReviewPairStatus.Unchanged)
                return highlight;
            PdfReviewPage page = leftSide
                ? PdfReviewDiff.PageAt(_result.Left, pair.LeftPageIndex)
                : PdfReviewDiff.PageAt(_result.Right, pair.RightPageIndex);
            if (page == null) return highlight;
            highlight.ViewWidthPt = page.ViewWidthPt;
            highlight.ViewHeightPt = page.ViewHeightPt;
            if (pair.Status == PdfReviewPairStatus.Changed)
            {
                PdfReviewDiffKind wanted = leftSide
                    ? PdfReviewDiffKind.Delete : PdfReviewDiffKind.Insert;
                foreach (PdfReviewWordOp op in pair.Operations)
                    if (op.Kind == wanted)
                        foreach (PdfReviewWord word in op.Words)
                            highlight.Boxes.Add(word.Box);
                return highlight;
            }
            // Страница без пары: всё содержимое — удаление (слева) или вставка (справа).
            if ((leftSide && pair.Status == PdfReviewPairStatus.LeftOnly) ||
                (!leftSide && pair.Status == PdfReviewPairStatus.RightOnly))
                foreach (PdfReviewWord word in page.Words)
                    highlight.Boxes.Add(word.Box);
            return highlight;
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
            _result.Stats = PdfReviewDiff.Statistics(_result);
            ApplyResult(_result);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (_sourceSplit.ClientSize.Width > _sourceSplit.SplitterWidth)
                _sourceSplit.SplitterDistance = (_sourceSplit.ClientSize.Width - _sourceSplit.SplitterWidth) / 2;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.F3) { MoveChange(1); return true; }
            if (keyData == (Keys.Shift | Keys.F3)) { MoveChange(-1); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void ClearResult()
        {
            _result = null; _pairIndex = -1;
            _pairs.Items.Clear(); _summary.Text = Loc.T("review.status.pickBoth");
            ClearViews();
        }

        private void ClearViews()
        {
            _leftSource.ShowPage(null, null, null); _rightSource.ShowPage(null, null, null);
            _position.Text = "";
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, Title, Loc.T("menu.howTo"), Loc.T("review.help.body"));
        }

        protected override void SyncControls()
        {
            bool files = !string.IsNullOrEmpty(_leftFile) && !string.IsNullOrEmpty(_rightFile);
            _pickLeft.Enabled = _pickRight.Enabled = !Working;
            _swap.Enabled = !Working && (!string.IsNullOrEmpty(_leftFile) || !string.IsNullOrEmpty(_rightFile));
            _compare.Enabled = !Working && files;
            _pairs.Enabled = !Working && _result != null;
            _previous.Enabled = _next.Enabled = _result != null && _result.Pairs.Count > 0;
            _manual.Enabled = !Working && _result != null && _pairIndex >= 0;
        }
    }
}
