using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Компактное окно «О программе»: главное — версия, руководство и приватность;
    /// редкие контакты и поддержка проекта раскрываются по запросу.
    /// </summary>
    public class AboutForm : Form
    {
        internal const string DonationAccount = "40817810354405296071";
        internal const string DonationBank = "ПОВОЛЖСКИЙ БАНК ПАО СБЕРБАНК";
        internal const string ProjectSectionName = "aboutProject";
        internal const string SupportSectionName = "aboutSupport";

        private const int Pad = 24;
        private const int WidthPx = 500;
        private const int FooterH = 60;

        private Bitmap _iconBitmap;
        private Panel _scroll;
        private PictureBox _iconBox;
        private Label _product, _version;
        private TextBoxBase _description;
        private Panel _manualRow, _privacyRow;
        private Label _privacyNote;
        private Button _projectToggle, _supportToggle;
        private Panel _projectPanel, _supportPanel;
        private Label _license;
        private Button _ok;
        internal Button ProjectToggle { get { return _projectToggle; } }
        internal Button SupportToggle { get { return _supportToggle; } }
        internal Panel ProjectPanel { get { return _projectPanel; } }
        internal Panel SupportPanel { get { return _supportPanel; } }

        public AboutForm()
        {
            Ui.InitDialog(this, Loc.T("hub.about"));
            ClientSize = new Size(WidthPx, 450);
            WindowChrome.Enable(this, Theme.HubBlue);
            Ui.AccentBar(this, 0, Theme.HubBlue);

            _scroll = new Panel();
            _scroll.AutoScroll = true;
            _scroll.BackColor = Color.White;
            Controls.Add(_scroll);

            BuildIdentity();
            BuildPrimaryActions();
            BuildProjectSection();
            BuildSupportSection();
            BuildFooter();
            LayoutContent();
        }

        private void BuildIdentity()
        {
            _iconBox = new PictureBox();
            _iconBox.SizeMode = PictureBoxSizeMode.StretchImage;
            Icon appIcon = Ui.AppIcon();
            if (appIcon != null)
            {
                _iconBitmap = appIcon.ToBitmap();
                _iconBox.Image = _iconBitmap;
            }
            _scroll.Controls.Add(_iconBox);

            _product = Ui.Label(_scroll, "iwo Helper Desktop", 0, 0,
                Ui.Font(14f, FontStyle.Bold), Color.FromArgb(40, 40, 40));
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            _version = Ui.Label(_scroll, string.Format(Loc.T("about.version"), version.ToString(3)),
                0, 0, Font, Theme.TextMuted);

            _description = JustifiedText.Paragraph(_scroll, Loc.T("about.desc"),
                0, 0, WidthPx - 2 * Pad, Theme.TextPrimary);
            _description.Name = "aboutDescription";
            _description.AccessibleName = Loc.T("about.desc");
        }

        private void BuildPrimaryActions()
        {
            _manualRow = new Panel { BackColor = Color.White };
            Label manual = Ui.Label(_manualRow, Loc.T("about.manual"), 0, 2, Font, Theme.TextPrimary);
            LinkLabel openManual = Ui.Link(_manualRow, Loc.T("about.manual.open"), manual.Right + 6, 2);
            openManual.Name = "aboutManual";
            openManual.LinkClicked += delegate { UserManual.Open(this, Loc.T("hub.about")); };
            _scroll.Controls.Add(_manualRow);

            _privacyRow = new Panel { BackColor = Color.White };
            LinkLabel privacy = Ui.UrlLink(_privacyRow, Loc.T("about.privacy"), 0, 2,
                "https://github.com/DedovMosol/iwoHelperDesktop/blob/main/docs/PRIVACY.md");
            privacy.Name = "aboutPrivacy";
            _privacyRow.Controls.Add(privacy);
            LinkLabel notices = Ui.Link(_privacyRow, Loc.T("about.thirdParty"),
                privacy.Right + 16, 2);
            notices.Name = "aboutThirdParty";
            notices.LinkClicked += delegate
            {
                ThirdPartyNotices.Open(this, Loc.T("hub.about"));
            };
            _privacyRow.Controls.Add(notices);
            _scroll.Controls.Add(_privacyRow);

            _privacyNote = new Label();
            _privacyNote.Text = Loc.T("about.privacyNote");
            _privacyNote.ForeColor = Theme.TextMuted;
            _privacyNote.BackColor = Color.White;
            _privacyNote.AutoSize = true;
            _privacyNote.MaximumSize = new Size(WidthPx - 2 * Pad, 0);
            _scroll.Controls.Add(_privacyNote);
        }

        private void BuildProjectSection()
        {
            _projectToggle = DisclosureButton(ProjectSectionName, Loc.T("about.project"));
            _projectToggle.Click += delegate
            {
                _projectPanel.Visible = !_projectPanel.Visible;
                UpdateDisclosure(_projectToggle, Loc.T("about.project"), _projectPanel.Visible);
                LayoutContent();
            };
            _scroll.Controls.Add(_projectToggle);

            _projectPanel = new Panel { Name = ProjectSectionName + "Panel", BackColor = Color.White, Visible = false };
            _projectPanel.SetBounds(Pad, 0, WidthPx - 2 * Pad, 92);
            Ui.Label(_projectPanel, Loc.T("about.author"), 12, 6, Font, Theme.TextPrimary);
            Label github = Ui.Label(_projectPanel, "GitHub:", 12, 34, Font, Theme.TextPrimary);
            Ui.UrlLink(_projectPanel, "DedovMosol/iwoHelperDesktop", github.Right + 6, 34,
                "https://github.com/DedovMosol/iwoHelperDesktop");
            Label telegram = Ui.Label(_projectPanel, "Telegram:", 12, 62, Font, Theme.TextPrimary);
            Ui.UrlLink(_projectPanel, "t.me/i_wantout", telegram.Right + 6, 62,
                "https://t.me/i_wantout");
            _scroll.Controls.Add(_projectPanel);
        }

        private void BuildSupportSection()
        {
            _supportToggle = DisclosureButton(SupportSectionName, Loc.T("about.donate"));
            _supportToggle.Click += delegate
            {
                _supportPanel.Visible = !_supportPanel.Visible;
                UpdateDisclosure(_supportToggle, Loc.T("about.donate"), _supportPanel.Visible);
                LayoutContent();
            };
            _scroll.Controls.Add(_supportToggle);

            _supportPanel = new Panel { Name = SupportSectionName + "Panel", BackColor = Color.White, Visible = false };
            _supportPanel.SetBounds(Pad, 0, WidthPx - 2 * Pad, 104);
            Label hint = Ui.Label(_supportPanel, Loc.T("about.copyHint"), 12, 4, Font, Theme.TextMuted);
            Label account = Ui.Label(_supportPanel, Loc.T("about.account"), 12, hint.Bottom + 10, Font, Theme.TextPrimary);
            TextBox accountValue = SelectableText(DonationAccount);
            accountValue.Name = "aboutAccount";
            accountValue.SetBounds(108, account.Top - 1, WidthPx - 2 * Pad - 120, 22);
            _supportPanel.Controls.Add(accountValue);
            Label bank = Ui.Label(_supportPanel, Loc.T("about.bank"), 12, account.Bottom + 12, Font, Theme.TextPrimary);
            TextBox bankValue = SelectableText(DonationBank);
            bankValue.Name = "aboutBank";
            bankValue.SetBounds(108, bank.Top - 1, WidthPx - 2 * Pad - 120, 22);
            _supportPanel.Controls.Add(bankValue);
            _scroll.Controls.Add(_supportPanel);
        }

        private Button DisclosureButton(string name, string text)
        {
            var button = new RoundedButton(false);
            button.Name = name;
            button.AccessibleRole = AccessibleRole.PushButton;
            button.AccessibleName = text;
            button.SetBounds(0, 0, WidthPx - 2 * Pad, 34);
            UpdateDisclosure(button, text, false);
            return button;
        }

        private static void UpdateDisclosure(Button button, string text, bool expanded)
        {
            button.Text = (expanded ? "▾ " : "▸ ") + text;
            button.AccessibleDescription = expanded ? "expanded" : "collapsed";
        }

        private void BuildFooter()
        {
            _license = Ui.Label(this, Loc.T("about.license"), Pad, 0, Font, Theme.TextMuted);
            _ok = new RoundedButton(true);
            _ok.Text = Loc.T("common.ok");
            _ok.Click += delegate { Close(); };
            Controls.Add(_ok);
            AcceptButton = _ok;
            CancelButton = _ok;
        }

        /// <summary>Одна reflow-точка: ни локализация, ни раскрытие секций не держат свои Y.</summary>
        private void LayoutContent()
        {
            int contentW = WidthPx - 2 * Pad;
            int y = 22;
            _iconBox.SetBounds(Pad, y, 48, 48);
            _product.Location = new Point(Pad + 62, y);
            _version.Location = new Point(Pad + 64, y + 32);
            y += 70;

            _description.SetBounds(Pad, y, contentW, _description.Height);
            y = _description.Bottom + 14;
            _manualRow.SetBounds(Pad, y, contentW, 24);
            y += 30;
            _privacyRow.SetBounds(Pad, y, contentW, 24);
            y += 26;
            _privacyNote.Location = new Point(Pad, y);
            y = _privacyNote.Bottom + 16;

            _projectToggle.SetBounds(Pad, y, contentW, 34);
            _projectPanel.SetBounds(Pad, _projectPanel.Visible ? y + 40 : 0, contentW, 92);
            y = _projectToggle.Bottom + 6;
            if (_projectPanel.Visible)
            {
                y = _projectPanel.Bottom + 8;
            }

            _supportToggle.SetBounds(Pad, y, contentW, 34);
            _supportPanel.SetBounds(Pad, _supportPanel.Visible ? y + 40 : 0, contentW, 104);
            y = _supportToggle.Bottom + 6;
            if (_supportPanel.Visible)
            {
                _supportPanel.SetBounds(Pad, y, contentW, 104);
                y = _supportPanel.Bottom + 8;
            }

            int maxClient = Math.Max(440, Screen.PrimaryScreen.WorkingArea.Height - 100);
            int desired = y + FooterH + 10;
            int clientH = Math.Min(maxClient, desired);
            ClientSize = new Size(WidthPx, clientH);
            _scroll.SetBounds(0, 3, WidthPx, clientH - FooterH);
            _scroll.AutoScrollMinSize = new Size(0, y + 8);

            _ok.SetBounds(WidthPx - Pad - 100, clientH - 52, 100, 36);
            _license.Location = new Point(Pad,
                _ok.Top + (_ok.Height - _license.PreferredHeight) / 2);
        }

        private TextBox SelectableText(string text)
        {
            return new TextBox
            {
                Text = text,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White,
                ForeColor = Theme.TextPrimary,
                Font = Font,
                TabStop = true
            };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _iconBitmap != null)
                _iconBitmap.Dispose();
            base.Dispose(disposing);
        }
    }
}
