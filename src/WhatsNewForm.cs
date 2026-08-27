using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Version-aware, dismissible release notes with an unobtrusive support section.</summary>
    internal sealed class WhatsNewForm : Form
    {
        private const int WidthPx = 620;
        private const int Pad = 24;
        private readonly string _version;
        private readonly Panel _scroll;
        private readonly List<Panel> _cards = new List<Panel>();
        private readonly LinkLabel _supportLink;
        private readonly Panel _supportPanel;
        private readonly AccentCheckBox _dontShow;
        private bool _persisted;

        internal WhatsNewForm(string version)
        {
            _version = version;
            Ui.InitDialog(this, string.Format(Loc.T("whatsnew.title"), version));
            ClientSize = new Size(WidthPx, 620);
            MinimumSize = Size;
            WindowChrome.Enable(this, Theme.HubBlue);
            BuildHeader();

            _scroll = new Panel
            {
                AutoScroll = true,
                BackColor = Color.FromArgb(248, 249, 251)
            };
            _scroll.SetBounds(0, 102, WidthPx, 446);
            _scroll.Anchor = AnchorStyles.Top | AnchorStyles.Bottom |
                AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(_scroll);

            int number = 0;
            foreach (string item in WhatsNewCatalog.Items(version, Loc.Code(Loc.Current)))
                _cards.Add(FeatureCard(++number, item));

            _supportLink = Ui.Link(_scroll, Loc.T("whatsnew.support.link"), Pad, 0);
            _supportLink.AccessibleDescription = Loc.T("whatsnew.support.hint");
            _supportLink.LinkClicked += delegate
            {
                _supportPanel.Visible = !_supportPanel.Visible;
                _supportLink.Text = Loc.T(_supportPanel.Visible
                    ? "whatsnew.support.hide" : "whatsnew.support.link");
                LayoutBody();
            };

            _supportPanel = BuildSupportPanel();
            _supportPanel.Visible = false;
            _scroll.Controls.Add(_supportPanel);

            _dontShow = new AccentCheckBox
            {
                Text = Loc.T("whatsnew.dontShow"),
                Checked = !UserSettings.Load().ShowWhatsNewOnStart
            };
            Controls.Add(_dontShow);
            Size optionSize = _dontShow.GetPreferredSize(Size.Empty);
            _dontShow.SetBounds(Pad, ClientSize.Height - 50,
                Math.Min(optionSize.Width, WidthPx - 180), optionSize.Height);
            _dontShow.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

            var close = new RoundedButton(true)
            {
                Text = Loc.T("common.close")
            };
            close.SetBounds(WidthPx - Pad - 110, ClientSize.Height - 56, 110, 36);
            close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            close.Click += delegate { Close(); };
            Controls.Add(close);
            AcceptButton = close;
            CancelButton = close;
            LayoutBody();
        }

        private void BuildHeader()
        {
            var header = new Panel
            {
                BackColor = Theme.HubBlue
            };
            header.SetBounds(0, 0, WidthPx, 102);
            header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(header);
            Label spark = Ui.Label(header, "✦", Pad, 22,
                Ui.Font(28f, FontStyle.Bold), Color.White);
            spark.AccessibleName = Loc.T("whatsnew.accessible");
            Label title = Ui.Label(header,
                string.Format(Loc.T("whatsnew.header"), _version), 82, 20,
                Ui.Font(16f, FontStyle.Bold), Color.White);
            title.AutoSize = false;
            title.SetBounds(82, 16, WidthPx - 108, 34);
            title.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Label subtitle = Ui.Label(header, Loc.T("whatsnew.subtitle"), 84, 54,
                Ui.Font(9.5f), Color.FromArgb(226, 237, 250));
            subtitle.AutoSize = false;
            subtitle.SetBounds(84, 52, WidthPx - 110, 38);
            subtitle.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        }

        private Panel FeatureCard(int number, string text)
        {
            var card = new Panel
            {
                BackColor = Color.White,
                AccessibleName = string.Format(Loc.T("whatsnew.item.accessible"), number)
            };
            var numberLabel = new Label
            {
                Text = number.ToString(),
                Font = Ui.Font(10f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Theme.HubBlue,
                TextAlign = ContentAlignment.MiddleCenter
            };
            numberLabel.SetBounds(14, 14, 30, 30);
            card.Controls.Add(numberLabel);
            var body = new Label
            {
                Text = text,
                Font = Font,
                ForeColor = Theme.TextPrimary,
                BackColor = Color.White,
                AutoSize = true,
                MaximumSize = new Size(WidthPx - 2 * Pad - 76, 0)
            };
            body.Location = new Point(58, 14);
            card.Controls.Add(body);
            card.Height = Math.Max(58, body.Bottom + 14);
            _scroll.Controls.Add(card);
            return card;
        }

        private Panel BuildSupportPanel()
        {
            var panel = new Panel
            {
                BackColor = Color.FromArgb(239, 246, 253),
                Width = WidthPx - 2 * Pad,
                Height = 128
            };
            Label note = Ui.Label(panel, Loc.T("whatsnew.support.body"), 14, 10,
                Ui.Font(9f), Theme.TextPrimary);
            note.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            note.AutoSize = false;
            note.SetBounds(14, 9, WidthPx - 2 * Pad - 28, 38);
            Label account = Ui.Label(panel, Loc.T("about.account"), 14, 56,
                Font, Theme.TextMuted);
            TextBox accountValue = Selectable(AboutForm.DonationAccount);
            accountValue.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            accountValue.SetBounds(112, 54, WidthPx - 2 * Pad - 128, 22);
            panel.Controls.Add(accountValue);
            Label bank = Ui.Label(panel, Loc.T("about.bank"), 14, 88,
                Font, Theme.TextMuted);
            TextBox bankValue = Selectable(AboutForm.DonationBank);
            bankValue.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            bankValue.SetBounds(112, 86, WidthPx - 2 * Pad - 128, 22);
            panel.Controls.Add(bankValue);
            return panel;
        }

        private TextBox Selectable(string text)
        {
            return new TextBox
            {
                Text = text,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = Color.FromArgb(239, 246, 253),
                ForeColor = Theme.TextPrimary,
                Font = Font,
                TabStop = true
            };
        }

        private void LayoutBody()
        {
            int y = 18;
            int width = Math.Max(280, _scroll.ClientSize.Width - 2 * Pad -
                SystemInformation.VerticalScrollBarWidth);
            foreach (Panel card in _cards)
            {
                card.SetBounds(Pad, y, width, card.Height);
                y = card.Bottom + 10;
            }
            _supportLink.Location = new Point(Pad, y + 4);
            y = _supportLink.Bottom + 10;
            _supportPanel.SetBounds(Pad, y, width, _supportPanel.Height);
            if (_supportPanel.Visible)
                y = _supportPanel.Bottom + 16;
            _scroll.AutoScrollMinSize = new Size(0, y);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (_scroll != null)
                LayoutBody();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!_persisted)
            {
                _persisted = true;
                if (!UserSettings.SaveWhatsNew(!_dontShow.Checked, _version))
                    Dialogs.Error(this, Text, Loc.T("settings.err.save.title"),
                        Loc.T("settings.err.save.body"));
            }
            base.OnFormClosing(e);
        }
    }

    internal static class WhatsNewUi
    {
        private static WhatsNewForm _open;

        internal static void ShowIfNeeded(Form owner)
        {
            string version = WhatsNewCatalog.CurrentVersion;
            if (!WhatsNewCatalog.ShouldShow(UserSettings.Load(), version))
                return;
            ShowModeless(owner, version);
        }

        internal static void ShowModeless(Form owner, string version)
        {
            if (_open != null && !_open.IsDisposed)
            {
                _open.Activate();
                return;
            }
            _open = new WhatsNewForm(version);
            _open.FormClosed += delegate { _open = null; };
            _open.Show(owner);
        }

        internal static void ShowDialog(IWin32Window owner)
        {
            if (_open != null && !_open.IsDisposed)
            {
                _open.Activate();
                return;
            }
            using (var form = new WhatsNewForm(WhatsNewCatalog.CurrentVersion))
                form.ShowDialog(owner);
        }
    }
}
