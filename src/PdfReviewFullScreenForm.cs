using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Maximized Review canvas using the exact same highlighted render request.</summary>
    internal sealed class PdfReviewFullScreenForm : Form
    {
        private readonly PdfReviewPageView _view;
        private readonly PdfReviewViewContent _content;

        private PdfReviewFullScreenForm(PdfReviewViewContent content)
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

        internal static void Show(IWin32Window owner, PdfReviewViewContent content)
        {
            if (content == null || content.BasePage == null)
                return;
            using (var form = new PdfReviewFullScreenForm(content))
                form.ShowDialog(owner);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _view.ShowContent(_content);
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
    }
}
