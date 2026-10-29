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
    public class AboutForm : ScaledForm
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
        private RichTextBox _privacyNote;
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
            // Ряды и панели — в окно ДО подписей: подпись, созданная в ещё не вставленной
            // панели, мерилась шрифтом по умолчанию, а показывалась шрифтом окна, и ссылку
            // справа от неё ставили по чужой ширине.
            _manualRow = new Panel { BackColor = Color.White };
            _scroll.Controls.Add(_manualRow);
            Label manual = Ui.Label(_manualRow, Loc.T("about.manual"), 0, 2, Font, Theme.TextPrimary);
            // Правый край подписи AutoSize — в пикселях экрана, раскладка здесь логическая (DpiScale).
            LinkLabel openManual = Ui.Link(_manualRow, Loc.T("about.manual.open"),
                DpiScale.ToLogical(manual.Right) + 6, 2);
            openManual.Name = "aboutManual";
            openManual.LinkClicked += delegate { UserManual.Open(this, Loc.T("hub.about")); };

            _privacyRow = new Panel { BackColor = Color.White };
            _scroll.Controls.Add(_privacyRow);
            LinkLabel privacy = Ui.UrlLink(_privacyRow, Loc.T("about.privacy"), 0, 2,
                "https://github.com/DedovMosol/iwoHelperDesktop/blob/main/docs/PRIVACY.md");
            privacy.Name = "aboutPrivacy";
            _privacyRow.Controls.Add(privacy);
            LinkLabel notices = Ui.Link(_privacyRow, Loc.T("about.thirdParty"),
                DpiScale.ToLogical(privacy.Right) + 16, 2);
            notices.Name = "aboutThirdParty";
            notices.LinkClicked += delegate
            {
                ThirdPartyNotices.Open(this, Loc.T("hub.about"));
            };
            _privacyRow.Controls.Add(notices);

            _privacyNote = JustifiedText.Paragraph(_scroll, Loc.T("about.privacyNote"),
                Pad, 0, WidthPx - 2 * Pad, Theme.TextMuted);
            _privacyNote.Name = "aboutPrivacyNote";
            _privacyNote.TabStop = false;
            _privacyNote.AccessibleName = Loc.T("about.privacyNote");
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
            _scroll.Controls.Add(_projectPanel); // до подписей — см. BuildPrimaryActions
            Ui.Label(_projectPanel, Loc.T("about.author"), 12, 6, Font, Theme.TextPrimary);
            Label github = Ui.Label(_projectPanel, "GitHub:", 12, 34, Font, Theme.TextPrimary);
            Ui.UrlLink(_projectPanel, "DedovMosol/iwoHelperDesktop", DpiScale.ToLogical(github.Right) + 6, 34,
                "https://github.com/DedovMosol/iwoHelperDesktop");
            Label telegram = Ui.Label(_projectPanel, "Telegram:", 12, 62, Font, Theme.TextPrimary);
            Ui.UrlLink(_projectPanel, "t.me/i_wantout", DpiScale.ToLogical(telegram.Right) + 6, 62,
                "https://t.me/i_wantout");
            // Высота по содержимому (низ последней строки + нижний отступ), без магических чисел.
            _projectPanel.SetBounds(Pad, 0, WidthPx - 2 * Pad, DpiScale.LogicalBottom(telegram) + 8);
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
            _scroll.Controls.Add(_supportPanel); // до подписей — см. BuildPrimaryActions
            Label account = Ui.Label(_supportPanel, Loc.T("about.account"), 12, 10, Font, Theme.TextPrimary);
            TextBox accountValue = SelectableText(DonationAccount);
            accountValue.Name = "aboutAccount";
            accountValue.SetBounds(108, account.Top - 1, WidthPx - 2 * Pad - 120, 22);
            _supportPanel.Controls.Add(accountValue);
            Label bank = Ui.Label(_supportPanel, Loc.T("about.bank"), 12, DpiScale.LogicalBottom(account) + 12, Font, Theme.TextPrimary);
            TextBox bankValue = SelectableText(DonationBank);
            bankValue.Name = "aboutBank";
            bankValue.SetBounds(108, bank.Top - 1, WidthPx - 2 * Pad - 120, 22);
            _supportPanel.Controls.Add(bankValue);
            // Высота по содержимому (низ последнего поля + нижний отступ), без магических чисел.
            _supportPanel.SetBounds(Pad, 0, WidthPx - 2 * Pad, DpiScale.LogicalBottom(bankValue) + 8);
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

        /// <summary>
        /// Одна reflow-точка: ни локализация, ни раскрытие секций не держат свои Y. Зовётся из
        /// конструктора (окно ещё логическое) и при раскрытии секций (уже в пикселях экрана),
        /// поэтому константы проходят через <see cref="DpiScale.Px"/>, замеры — через
        /// <see cref="DpiScale.FromDevice"/>, а ширина переноса — через MeasureWidth.
        /// </summary>
        private void LayoutContent()
        {
            Func<int, int> d = delegate(int v) { return DpiScale.Px(this, v); };
            Func<int, int> m = delegate(int v) { return DpiScale.FromDevice(this, v); };
            int pad = d(Pad), width = d(WidthPx);
            int contentW = width - 2 * pad;
            int y = d(22);
            _iconBox.SetBounds(pad, y, d(48), d(48));
            _product.Location = new Point(pad + d(62), y);
            _version.Location = new Point(pad + d(64), y + d(32));
            y += d(70);

            _description.SetBounds(pad, y, contentW,
                m(JustifiedText.Height(_description.Text, _description.Font, DpiScale.MeasureWidth(this, contentW))));
            y = _description.Bottom + d(14);
            _manualRow.SetBounds(pad, y, contentW, d(24));
            y += d(30);
            _privacyRow.SetBounds(pad, y, contentW, d(24));
            y += d(26);
            _privacyNote.SetBounds(pad, y, contentW,
                m(JustifiedText.Height(_privacyNote.Text, _privacyNote.Font, DpiScale.MeasureWidth(this, contentW))));
            y = _privacyNote.Bottom + d(16);

            _projectToggle.SetBounds(pad, y, contentW, d(34));
            y = _projectToggle.Bottom + d(6);
            if (_projectPanel.Visible)
            {
                _projectPanel.SetBounds(pad, y, contentW, _projectPanel.Height);
                y = _projectPanel.Bottom + d(8);
            }

            _supportToggle.SetBounds(pad, y, contentW, d(34));
            y = _supportToggle.Bottom + d(6);
            if (_supportPanel.Visible)
            {
                _supportPanel.SetBounds(pad, y, contentW, _supportPanel.Height);
                y = _supportPanel.Bottom + d(8);
            }

            // Экран окна, а не основной: на втором мониторе другого размера потолок был чужим.
            int maxClient = Math.Max(d(440), m(ScaledForm.WorkingArea(Bounds).Height) - d(100));
            int desired = y + d(FooterH) + d(10);
            int clientH = Math.Min(maxClient, desired);
            ClientSize = new Size(width, clientH);
            _scroll.SetBounds(0, d(3), width, clientH - d(FooterH));
            _scroll.AutoScrollMinSize = new Size(0, y + d(8));

            _ok.SetBounds(width - pad - d(100), clientH - d(52), d(100), d(36));
            _license.Location = new Point(pad,
                _ok.Top + (_ok.Height - m(_license.PreferredHeight)) / 2);
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
