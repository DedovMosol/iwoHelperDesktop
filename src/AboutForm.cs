using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Окно «О программе»: версия, автор, лицензия, ссылки и пункты для доната.</summary>
    public class AboutForm : Form
    {
        // Пункты доната — выделяемые read-only поля (правый клик / Ctrl+C).
        internal const string DonationAccount = "40817810354405296071";
        internal const string DonationBank = "ПОВОЛЖСКИЙ БАНК ПАО СБЕРБАНК";

        private Bitmap _iconBitmap; // PictureBox своё Image не освобождает — держим и диспозим сами

        public AboutForm()
        {
            Ui.InitDialog(this, Loc.T("hub.about"));
            ClientSize = new Size(460, 340); // высота уточняется в конце под контент
            WindowChrome.Enable(this, Theme.HubBlue); // синий заголовок на Windows 11 — как на главной

            Ui.AccentBar(this, 0, Theme.HubBlue);

            var iconBox = new PictureBox();
            iconBox.SetBounds(24, 26, 48, 48);
            iconBox.SizeMode = PictureBoxSizeMode.StretchImage;
            Icon appIcon = Ui.AppIcon();
            if (appIcon != null)
            {
                _iconBitmap = appIcon.ToBitmap();
                iconBox.Image = _iconBitmap;
            }
            Controls.Add(iconBox);

            Ui.Label(this, "iwo Helper Desktop", 86, 26,
                Ui.Font(14f, FontStyle.Bold), Color.FromArgb(40, 40, 40));
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            Ui.Label(this, string.Format(Loc.T("about.version"), version.ToString(3)), 88, 58, Font, Theme.TextMuted);

            // Описание — выделяемое и копируемое, поэтому read-only поле, а не подпись.
            // Ширина ограничена окном (текст не вылезает за край), высота считается под перенос.
            TextBoxBase desc = SelectableText(Loc.T("about.desc"), 24, 96, ClientSize.Width - 48, true);

            int y = desc.Bottom + 14;
            Ui.Label(this, Loc.T("about.author"), 24, y, Font, Theme.TextPrimary); y += 24;

            // Руководство пользователя лежит ресурсом в exe и открывается отсюда: держать его
            // только в интернете для офлайнового приложения было бы обманом.
            Label manual = Ui.Label(this, Loc.T("about.manual"), 24, y, Font, Theme.TextPrimary);
            LinkLabel openManual = Ui.Link(this, Loc.T("about.manual.open"), manual.Right + 6, y);
            openManual.LinkClicked += delegate { UserManual.Open(this, Loc.T("hub.about")); };
            y += 24;

            Label tg = Ui.Label(this, "Telegram:", 24, y, Font, Theme.TextPrimary);
            Ui.UrlLink(this, "t.me/i_wantout", tg.Right + 6, y, "https://t.me/i_wantout"); y += 24;
            Label gh = Ui.Label(this, "GitHub:", 24, y, Font, Theme.TextPrimary);
            Ui.UrlLink(this, "DedovMosol/iwoHelperDesktop", gh.Right + 6, y,
                "https://github.com/DedovMosol/iwoHelperDesktop"); y += 24;
            LinkLabel pp = Ui.UrlLink(this, Loc.T("about.privacy"), 24, y,
                "https://github.com/DedovMosol/iwoHelperDesktop/blob/main/docs/PRIVACY.md");
            Ui.Label(this, Loc.T("about.privacyNote"), pp.Right + 6, y, Font, Theme.TextMuted); y += 34;

            // --- Донаты: пункты можно выделить и скопировать (read-only TextBox) ---
            Ui.Label(this, Loc.T("about.donate"), 24, y,
                Ui.Font(9.75f, FontStyle.Bold), Theme.TextPrimary); y += 26;

            // Пункты — выделяемые read-only поля: копируются правым кликом или Ctrl+C
            // (отдельная кнопка «копировать» не нужна).
            Label accCap = Ui.Label(this, Loc.T("about.account"), 24, y, Font, Theme.TextPrimary);
            SelectableText(DonationAccount, accCap.Right + 6, y - 1, ClientSize.Width - (accCap.Right + 6) - 24, false);
            y += 26;

            Label bankCap = Ui.Label(this, Loc.T("about.bank"), 24, y, Font, Theme.TextPrimary);
            SelectableText(DonationBank, bankCap.Right + 6, y - 1, ClientSize.Width - (bankCap.Right + 6) - 24, false);
            y += 30;

            // Высота окна — под весь контент плюс нижняя строка с кнопкой.
            ClientSize = new Size(ClientSize.Width, y + 16 + 36 + 16);

            var ok = new RoundedButton(true);
            ok.Text = Loc.T("common.ok");
            ok.SetBounds(ClientSize.Width - 124, ClientSize.Height - 52, 100, 36);
            ok.Click += delegate { Close(); };
            Controls.Add(ok);
            AcceptButton = ok;
            CancelButton = ok; // Esc тоже закрывает

            // Копирайт и лицензия — в самом низу слева, на одной линии с кнопкой справа.
            Label license = Ui.Label(this, Loc.T("about.license"), 24, ok.Top, Font, Theme.TextMuted);
            license.Top = ok.Top + (ok.Height - license.Height) / 2; // по центру относительно кнопки
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _iconBitmap != null)
                _iconBitmap.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>
        /// Текст только для чтения, но выделяемый и копируемый (Ctrl+C), без рамки и на белом —
        /// внешне подпись, на деле поле. paragraph=true включает перенос по словам и считает
        /// высоту под весь текст, иначе это одна строка фиксированной высоты.
        /// </summary>
        private TextBoxBase SelectableText(string text, int x, int y, int width, bool paragraph)
        {
            // Абзац описания — RichTextBox: обычный TextBox умеет выключку только влево,
            // вправо и по центру, а описание читается заметно ровнее выключенным по ширине.
            // Выделять и копировать текст по-прежнему можно — это тот же TextBoxBase.
            if (paragraph)
                return JustifiedText.Paragraph(this, text, x, y, width, Theme.TextPrimary);
            var tb = new TextBox();
            tb.Multiline = false;
            tb.WordWrap = false;
            tb.ScrollBars = ScrollBars.None; // высота подобрана под текст, полосы не нужны
            tb.Text = text;
            tb.ReadOnly = true;
            tb.BorderStyle = BorderStyle.None;
            tb.BackColor = Color.White;
            tb.ForeColor = Theme.TextPrimary;
            tb.Font = Font;
            tb.TabStop = false;
            tb.SetBounds(x, y, width, 20);
            Controls.Add(tb);
            return tb;
        }

    }
}
