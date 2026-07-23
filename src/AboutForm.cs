using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Окно «О программе»: версия, автор, лицензия, ссылки и реквизиты для доната.</summary>
    public class AboutForm : Form
    {
        // Реквизиты доната — выделяемые read-only поля (правый клик / Ctrl+C).
        internal const string DonationAccount = "40817810354405296071";
        internal const string DonationBank = "ПОВОЛЖСКИЙ БАНК ПАО СБЕРБАНК";

        public AboutForm()
        {
            Text = Loc.T("hub.about");
            Font = new Font("Segoe UI", 9.75f);
            BackColor = Color.White;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(460, 340); // высота уточняется в конце под контент
            WindowChrome.Enable(this, Theme.HubBlue); // синий заголовок на Windows 11 — как на главной

            Ui.AccentBar(this, 0, Theme.HubBlue);

            var iconBox = new PictureBox();
            iconBox.SetBounds(24, 26, 48, 48);
            iconBox.SizeMode = PictureBoxSizeMode.StretchImage;
            Icon appIcon = Ui.AppIcon();
            if (appIcon != null)
                iconBox.Image = appIcon.ToBitmap();
            Controls.Add(iconBox);

            Ui.Label(this, "iwo Helper Desktop", 86, 26,
                new Font("Segoe UI", 14f, FontStyle.Bold), Color.FromArgb(40, 40, 40));
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            Ui.Label(this, string.Format(Loc.T("about.version"), version.ToString(3)), 88, 58, Font, Theme.TextMuted);

            // Описание — выравнивание по ширине; ширина ограничена окном (не вылезает за край).
            var desc = new JustifiedLabel();
            desc.Font = Font;
            desc.ForeColor = Theme.TextPrimary;
            desc.SetBounds(24, 96, ClientSize.Width - 48, 10);
            desc.Text = Loc.T("about.desc");
            desc.Height = desc.GetPreferredHeight();
            Controls.Add(desc);

            int y = desc.Bottom + 14;
            Ui.Label(this, Loc.T("about.author"), 24, y, Font, Theme.TextPrimary); y += 24;
            Ui.Label(this, Loc.T("about.license"), 24, y, Font, Theme.TextMuted); y += 30;

            Label tg = Ui.Label(this, "Telegram:", 24, y, Font, Theme.TextPrimary);
            Ui.UrlLink(this, "t.me/i_wantout", tg.Right + 6, y, "https://t.me/i_wantout"); y += 24;
            Label gh = Ui.Label(this, "GitHub:", 24, y, Font, Theme.TextPrimary);
            Ui.UrlLink(this, "DedovMosol/iwoHelperDesktop", gh.Right + 6, y,
                "https://github.com/DedovMosol/iwoHelperDesktop"); y += 24;
            LinkLabel pp = Ui.UrlLink(this, Loc.T("about.privacy"), 24, y,
                "https://github.com/DedovMosol/iwoHelperDesktop/blob/main/docs/PRIVACY.md");
            Ui.Label(this, Loc.T("about.privacyNote"), pp.Right + 6, y, Font, Theme.TextMuted); y += 34;

            // --- Донаты: реквизиты можно выделить и скопировать (read-only TextBox) ---
            Ui.Label(this, Loc.T("about.donate"), 24, y,
                new Font("Segoe UI", 9.75f, FontStyle.Bold), Theme.TextPrimary); y += 26;

            // Реквизиты — выделяемые read-only поля: копируются правым кликом или Ctrl+C
            // (отдельная кнопка «копировать» не нужна).
            Label accCap = Ui.Label(this, Loc.T("about.account"), 24, y, Font, Theme.TextPrimary);
            SelectableValue(DonationAccount, accCap.Right + 6, y - 1, ClientSize.Width - (accCap.Right + 6) - 24);
            y += 26;

            Label bankCap = Ui.Label(this, Loc.T("about.bank"), 24, y, Font, Theme.TextPrimary);
            SelectableValue(DonationBank, bankCap.Right + 6, y - 1, ClientSize.Width - (bankCap.Right + 6) - 24);
            y += 30;

            // Высота окна — под весь контент плюс кнопка снизу.
            ClientSize = new Size(ClientSize.Width, y + 16 + 36 + 16);

            var ok = new RoundedButton(true);
            ok.Text = "OK";
            ok.SetBounds(ClientSize.Width - 124, ClientSize.Height - 52, 100, 36);
            ok.Click += delegate { Close(); };
            Controls.Add(ok);
            AcceptButton = ok;
            CancelButton = ok; // Esc тоже закрывает
        }

        /// <summary>Значение только для чтения, но выделяемое и копируемое (Ctrl+C), без рамки.</summary>
        private TextBox SelectableValue(string text, int x, int y, int width)
        {
            var tb = new TextBox();
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
