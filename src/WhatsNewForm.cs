using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Version-aware, dismissible release notes with an unobtrusive support section.</summary>
    internal sealed class WhatsNewForm : Form, IMessageFilter
    {
        private const int WmMouseWheel = 0x020A;
        // Дефолт — замеренный у пользователя удобный размер (1193×699 клиент): карточки
        // версии читаются без прокрутки. WindowPlacement привязывается ниже и при наличии
        // сохранённого размера восстановит его вместо дефолта.
        private const int DefaultWidth = 1193;
        private const int DefaultHeight = 699;
        private const int HeaderH = 76;
        private const int FooterH = 88;
        private const int Pad = 24;
        private const int MinTextPercent = 80;
        private const int MaxTextPercent = 160;
        private const int TextStepPercent = 10;
        private const float BaseTextSize = 9.75f;

        private readonly string _version;
        private readonly Panel _scroll;
        private readonly List<Panel> _cards = new List<Panel>();
        private readonly List<NonSelectableText> _cardBodies = new List<NonSelectableText>();
        private readonly LinkLabel _supportLink;
        private readonly Panel _supportPanel;
        private NonSelectableText _supportNote;
        private Label _supportAccountLabel, _supportBankLabel;
        private TextBox _supportAccount, _supportBank;
        private readonly AccentCheckBox _dontShow;
        private readonly RoundedButton _close;
        private readonly ToolTip _zoomTip = new ToolTip();
        private bool _filterRegistered;
        private bool _persisted;
        private bool _positioned;
        private int _textPercent;

        internal WhatsNewForm(string version)
        {
            _version = version;
            _textPercent = NormalizeTextPercent(UserSettings.Load().WhatsNewTextPercent);
            Ui.InitDialog(this, string.Format(Loc.T("whatsnew.title"), version));
            // Это самостоятельное окно: его размер нужен для длинных заметок, а состояние
            // должно возвращаться так же, как у рабочих окон приложения.
            StartPosition = FormStartPosition.Manual;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            ShowInTaskbar = true;
            // Дефолт не должен превышать рабочую область: на маленьком экране большое окно
            // уезжало бы за край ещё до CenterOnOwner, и центрирование ломалось клампом.
            Rectangle workArea = Screen.FromPoint(Cursor.Position).WorkingArea;
            int chromeW = Width - ClientSize.Width;   // рамки (0 до создания хэндла)
            int chromeH = Height - ClientSize.Height; // рамка + заголовок
            if (chromeH <= 0) chromeH = 39;           // типичный оверхед до создания хэндла
            ClientSize = new Size(
                Math.Min(DefaultWidth, Math.Max(520, workArea.Width - chromeW)),
                Math.Min(DefaultHeight, Math.Max(400, workArea.Height - chromeH)));
            MinimumSize = SizeFromClientSize(new Size(520, 400));
            WindowChrome.Enable(this, Theme.HubBlue);
            WindowPlacement.Attach(this);

            BuildHeader();

            _scroll = new Panel
            {
                AutoScroll = true,
                BackColor = Color.FromArgb(248, 249, 251)
            };
            _scroll.SetBounds(0, HeaderH, ClientSize.Width,
                ClientSize.Height - HeaderH - FooterH);
            _scroll.Anchor = AnchorStyles.Top | AnchorStyles.Bottom |
                AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(_scroll);

            int number = 0;
            foreach (string item in WhatsNewCatalog.Items(version, Loc.Code(Loc.Current)))
                _cards.Add(FeatureCard(++number, item));

            _supportLink = Ui.Link(this, Loc.T("whatsnew.support.link"), Pad, 0);
            _supportLink.AccessibleDescription = Loc.T("whatsnew.support.hint");
            _supportLink.LinkClicked += delegate { ToggleSupport(); };

            _supportPanel = BuildSupportPanel();
            _supportPanel.Visible = false;
            _scroll.Controls.Add(_supportPanel);

            // Ctrl+колесо — масштаб текста. Фильтр снимает зависимость от фокуса:
            // колесо над любым местом окна работает одинаково (как в браузерах).
            Application.AddMessageFilter(this);
            _filterRegistered = true;
            _zoomTip.SetToolTip(this, Loc.T("whatsnew.zoomTip"));
            _zoomTip.SetToolTip(_scroll, Loc.T("whatsnew.zoomTip"));

            _dontShow = new AccentCheckBox
            {
                Text = Loc.T("whatsnew.dontShow"),
                Checked = !UserSettings.Load().ShowWhatsNewOnStart
            };
            Controls.Add(_dontShow);

            _close = new RoundedButton(true)
            {
                Text = Loc.T("common.close"),
                TabIndex = 0
            };
            _close.Click += delegate { Close(); };
            Controls.Add(_close);
            AcceptButton = _close;
            CancelButton = _close;

            // Adding child controls can make AutoScroll retain a stale offset. The initial
            // publication is always the beginning of the notes; later relayouts preserve it.
            LayoutBody(0);
            LayoutFooter();
        }

        private void BuildHeader()
        {
            var header = new Panel
            {
                BackColor = Theme.HubBlue
            };
            header.SetBounds(0, 0, DefaultWidth, HeaderH);
            header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(header);
            Label spark = Ui.Label(header, "✦", Pad, 20,
                Ui.Font(28f, FontStyle.Bold), Color.White);
            spark.AutoSize = false;
            spark.TextAlign = ContentAlignment.MiddleCenter;
            spark.SetBounds(Pad, 18, 46, 42);
            spark.AccessibleName = Loc.T("whatsnew.accessible");
            Label title = Ui.Label(header,
                string.Format(Loc.T("whatsnew.header"), _version), 82, 0,
                Ui.Font(16f, FontStyle.Bold), Color.White);
            title.AutoSize = false;
            title.SetBounds(82, 20, DefaultWidth - 108, 36);
            title.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        }

        private Panel FeatureCard(int number, string text)
        {
            var card = new Panel
            {
                BackColor = Color.White,
                Font = Font,
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
            // Номер центрируется по вертикали относительно текста карточки, а не сверху.
            numberLabel.SetBounds(14, 0, 30, 30); // Y пересчитается после body
            card.Controls.Add(numberLabel);
            var body = new NonSelectableText
            {
                Text = text,
                Font = TextFont(),
                ForeColor = Theme.TextPrimary,
                BackColor = Color.White
            };
            body.SetBounds(58, 14,
                Math.Max(100, DefaultWidth - 2 * Pad - 76),
                NonSelectableText.MeasureHeight(text, body.Font,
                    Math.Max(100, DefaultWidth - 2 * Pad - 76)));
            body.Name = "whatsNewItem" + number;
            _cardBodies.Add(body);
            card.Height = Math.Max(58, body.Bottom + 14);
            numberLabel.Top = (card.Height - numberLabel.Height) / 2;
            _scroll.Controls.Add(card);
            return card;
        }

        private void ToggleSupport()
        {
            int previousOffset = ScrollOffsetY();
            bool expanding = !_supportPanel.Visible;
            _supportPanel.Visible = expanding;
            _supportLink.Text = Loc.T(expanding
                ? "whatsnew.support.hide" : "whatsnew.support.link");
            _close.Focus();
            // При сворачивании нельзя оставлять старую позицию прокрутки: после уменьшения
            // содержимого она могла указывать за пределы списка и создавать пустоту над пунктом 1.
            int target = expanding ? previousOffset : Math.Min(previousOffset, MaxScrollOffset());
            LayoutBody(target);
            if (expanding)
            {
                // Keep the clicked link at the same screen position. The panel grows below it;
                // WinForms must not auto-scroll a focused child back to the top of the list.
                SetScrollOffsetY(Math.Min(MaxScrollOffset(), previousOffset));
            }
            else
            {
                SetScrollOffsetY(Math.Min(target, MaxScrollOffset()));
            }
        }

        private int ScrollOffsetY()
        {
            return Math.Max(0, -_scroll.AutoScrollPosition.Y);
        }

        private int MaxScrollOffset()
        {
            return Math.Max(0, _scroll.AutoScrollMinSize.Height - _scroll.ClientSize.Height);
        }

        private void SetScrollOffsetY(int offset)
        {
            _scroll.AutoScrollPosition = new Point(0,
                Math.Max(0, Math.Min(MaxScrollOffset(), offset)));
        }

        private void LayoutBody(int preservedOffset)
        {
            if (_scroll == null || _supportLink == null || _supportPanel == null)
                return;
            _scroll.SuspendLayout();
            try
            {
                int y = 18;
                int width = Math.Max(280, _scroll.ClientSize.Width - 2 * Pad -
                    (_scroll.VerticalScroll.Visible
                        ? SystemInformation.VerticalScrollBarWidth : 0));
                for (int i = 0; i < _cards.Count; i++)
                {
                    Panel card = _cards[i];
                    NonSelectableText body = _cardBodies[i];
                    int bodyWidth = Math.Max(120, width - 72);
                    body.Font = TextFont();
                    body.SetBounds(58, 14, bodyWidth,
                        NonSelectableText.MeasureHeight(body.Text, body.Font, bodyWidth));
                    card.SetBounds(Pad, y, width, Math.Max(58, body.Bottom + 14));
                    // Номер центрируется по вертикали относительно текста.
                    card.Controls[0].Top = (card.Height - card.Controls[0].Height) / 2;
                    y = card.Bottom + 10;
                }
                LayoutSupportPanel(width);
                _supportPanel.SetBounds(Pad, y, width, _supportPanel.Height);
                if (_supportPanel.Visible)
                    y = _supportPanel.Bottom + 16;
                _scroll.AutoScrollMinSize = new Size(0, y);
            }
            finally
            {
                _scroll.ResumeLayout();
            }
            SetScrollOffsetY(preservedOffset);
        }

        private void LayoutSupportPanel(int width)
        {
            int innerWidth = Math.Max(160, width - 28);
            _supportNote.Font = TextFont();
            _supportNote.SetBounds(14, 10, innerWidth,
                NonSelectableText.MeasureHeight(_supportNote.Text, _supportNote.Font, innerWidth));

            int y = _supportNote.Bottom + 12;
            _supportAccountLabel.Font = TextFont();
            _supportAccountLabel.Location = new Point(14, y);
            _supportAccount.Font = TextFont();
            _supportAccount.SetBounds(112, y - 1, Math.Max(80, width - 128), 22);
            y = _supportAccount.Bottom + 10;
            _supportBankLabel.Font = TextFont();
            _supportBankLabel.Location = new Point(14, y);
            _supportBank.Font = TextFont();
            _supportBank.SetBounds(112, y - 1, Math.Max(80, width - 128), 22);
            _supportPanel.Height = _supportBank.Bottom + 14;
        }

        private Panel BuildSupportPanel()
        {
            var panel = new Panel
            {
                BackColor = Color.FromArgb(239, 246, 253),
                Width = DefaultWidth - 2 * Pad,
                Height = 128,
                Font = Font
            };
            _supportNote = new NonSelectableText
            {
                Text = Loc.T("whatsnew.support.body"),
                Font = TextFont(),
                ForeColor = Theme.TextPrimary,
                BackColor = panel.BackColor
            };
            _supportNote.SetBounds(14, 10, DefaultWidth - 2 * Pad - 28,
                NonSelectableText.MeasureHeight(_supportNote.Text, _supportNote.Font,
                    DefaultWidth - 2 * Pad - 28));
            panel.Controls.Add(_supportNote);
            _supportAccountLabel = Ui.Label(panel, Loc.T("about.account"), 14, 56,
                Font, Theme.TextMuted);
            _supportAccount = Selectable(AboutForm.DonationAccount);
            _supportAccount.SetBounds(112, 54, DefaultWidth - 2 * Pad - 128, 22);
            panel.Controls.Add(_supportAccount);
            _supportBankLabel = Ui.Label(panel, Loc.T("about.bank"), 14, 88,
                Font, Theme.TextMuted);
            _supportBank = Selectable(AboutForm.DonationBank);
            _supportBank.SetBounds(112, 86, DefaultWidth - 2 * Pad - 128, 22);
            panel.Controls.Add(_supportBank);
            return panel;
        }

        private void LayoutFooter()
        {
            if (_close == null)
                return;
            // Чекбокс всегда по центру окна по горизонтали, независимо от ширины.
            Size optionSize = _dontShow.GetPreferredSize(Size.Empty);
            int bottomY = ClientSize.Height - 48;
            _close.SetBounds(ClientSize.Width - Pad - 110, ClientSize.Height - 52,
                110, 36);
            // Ссылка поддержки — слева, по центру вертикали относительно кнопки «Закрыть».
            _supportLink.Location = new Point(Pad,
                _close.Top + (_close.Height - _supportLink.Height) / 2);
            // При узком окне чекбокс сдвигается правее ссылки, не перекрывая её.
            int optionX = Math.Max(Pad, (ClientSize.Width - optionSize.Width) / 2);
            if (optionX < _supportLink.Right + 8)
                optionX = _supportLink.Right + 8;
            // И не заходя на кнопку «Закрыть».
            if (optionX + optionSize.Width > _close.Left - 8)
                optionX = Math.Max(_supportLink.Right + 8, _close.Left - 8 - optionSize.Width);
            _dontShow.SetBounds(optionX, bottomY, optionSize.Width, optionSize.Height);
        }

        private int NormalizeTextPercent(int value)
        {
            if (value < MinTextPercent || value > MaxTextPercent)
                return 100;
            return value;
        }

        private Font TextFont()
        {
            return Ui.Font(BaseTextSize * _textPercent / 100f);
        }

        /// <summary>Изменить масштаб текста с клампом в [80..160] и пересчётом layout.</summary>
        internal void SetTextPercent(int value)
        {
            int next = Math.Max(MinTextPercent, Math.Min(MaxTextPercent, value));
            if (next == _textPercent)
                return;
            int offset = ScrollOffsetY();
            _textPercent = next;
            LayoutBody(offset);
            LayoutFooter();
        }

        /// <summary>Текущий масштаб текста (для тестов и отладки).</summary>
        internal int TextPercent => _textPercent;

        // ---- IMessageFilter: Ctrl+колесо над любым местом окна — масштаб текста ----

        bool IMessageFilter.PreFilterMessage(ref Message m)
        {
            if (m.Msg != WmMouseWheel || !_filterRegistered || IsDisposed || !Visible)
                return false;
            if ((ModifierKeys & Keys.Control) == 0)
                return false; // без Ctrl — обычная прокрутка, не трогаем
            // Фильтр глобален (Application.AddMessageFilter): реагируем, только если сообщение
            // адресовано ЭТОМУ окну или его дочернему контролу — чужой Ctrl+колесо в другой
            // форме не должен менять наш текст, а фильтр не должен мешать чужой активации.
            if (m.HWnd != Handle)
            {
                Control target = FromChildHandle(m.HWnd);
                if (target == null || !Contains(target))
                    return false;
            }
            int delta = (short)((long)m.WParam >> 16);
            SetTextPercent(_textPercent + (delta > 0 ? TextStepPercent : -TextStepPercent));
            return true;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (_filterRegistered)
            {
                Application.RemoveMessageFilter(this);
                _filterRegistered = false;
            }
            base.OnFormClosed(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_filterRegistered)
                {
                    Application.RemoveMessageFilter(this);
                    _filterRegistered = false;
                }
                if (_zoomTip != null)
                    _zoomTip.Dispose();
            }
            base.Dispose(disposing);
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
                Font = TextFont(),
                TabStop = true
            };
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (!_positioned)
            {
                _positioned = true;
                CenterOnOwner();
                SetScrollOffsetY(0);
            }
        }

        private void CenterOnOwner()
        {
            Rectangle area;
            if (Owner != null && Owner.IsHandleCreated && Owner.Visible &&
                Owner.Width > 0 && Owner.Height > 0)
                area = Owner.Bounds;
            else
                area = Screen.FromControl(this).WorkingArea;
            int x = area.Left + (area.Width - Width) / 2;
            int y = area.Top + (area.Height - Height) / 2;
            Rectangle work = Screen.FromRectangle(area).WorkingArea;
            x = Math.Max(work.Left, Math.Min(work.Right - Width, x));
            y = Math.Max(work.Top, Math.Min(work.Bottom - Height, y));
            Location = new Point(x, y);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (_scroll != null && _supportLink != null)
            {
                int offset = ScrollOffsetY();
                LayoutBody(offset);
                LayoutFooter();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!_persisted)
            {
                _persisted = true;
                if (!UserSettings.SaveWhatsNew(!_dontShow.Checked, _version, _textPercent))
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
