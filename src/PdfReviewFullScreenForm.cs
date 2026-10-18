using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Maximized Review canvas using the exact same highlighted render request.</summary>
    internal sealed class PdfReviewFullScreenForm : Form, IMessageFilter
    {
        private readonly PdfReviewPageView _view;
        private readonly PdfReviewViewContent _content;
        private bool _wheelFilterRegistered;

        internal PdfReviewFullScreenForm(PdfReviewViewContent content)
        {
            Text = content == null ? Loc.T("hub.review.name") : content.Caption;
            Icon icon = Ui.AppIcon();
            if (icon != null) Icon = icon;
            BackColor = Theme.DarkBarFill;
            Font = Ui.Font(9.75f);
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = true;
            MaximizeBox = true;
            ShowInTaskbar = true;
            StartPosition = FormStartPosition.CenterParent;
            WindowState = FormWindowState.Maximized;
            MinimumSize = new Size(640, 480);
            KeyPreview = true;
            WindowChrome.Enable(this, Theme.ReviewBlueDark);

            _content = content;
            _view = new PdfReviewPageView
            {
                Dock = DockStyle.Fill,
                AllowFullScreen = false,
                RenderWidthLimit = IntPtr.Size == 8 ? 2400 : 1600
            };
            Controls.Add(_view);
        }

        /// <summary>session — рендерер окна сравнения: документы уже загружены, второй копии нет.</summary>
        internal static void Show(IWin32Window owner, PdfReviewViewContent content,
            PdfReviewRenderSession session)
        {
            if (content == null || content.BasePage == null)
                return;
            using (var form = new PdfReviewFullScreenForm(content))
            {
                if (session != null)
                    form.View.RenderSession = session;
                form.ShowDialog(owner);
            }
        }

        internal PdfReviewPageView View { get { return _view; } }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            // Свой обработчик колеса: фильтр окна сравнения намеренно не видит сообщения
            // этого окна, поэтому прокрутку и Ctrl+масштаб здесь ведёт собственный фильтр.
            if (!_wheelFilterRegistered)
            {
                Application.AddMessageFilter(this);
                _wheelFilterRegistered = true;
            }
            _view.ShowContent(_content);
        }

        bool IMessageFilter.PreFilterMessage(ref Message m)
        {
            Point screenPoint;
            int delta;
            bool controlDown;
            if (!_wheelFilterRegistered || IsDisposed || !Visible ||
                !PdfReviewPageView.TryReadWheel(this, ref m, out screenPoint, out delta,
                    out controlDown) || !_view.ContainsViewport(screenPoint))
                return false;
            // Соседних строк сопоставления здесь нет: граница страницы просто конец прокрутки.
            return _view.ConsumesWheel(_view.HandleWheel(screenPoint, delta, controlDown),
                controlDown);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape || e.KeyCode == Keys.F11)
            {
                e.Handled = true;
                Close();
                return;
            }
            base.OnKeyDown(e);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            RemoveWheelFilter();
            base.OnFormClosed(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                RemoveWheelFilter();
            base.Dispose(disposing);
        }

        private void RemoveWheelFilter()
        {
            if (!_wheelFilterRegistered)
                return;
            Application.RemoveMessageFilter(this);
            _wheelFilterRegistered = false;
        }
    }
}
