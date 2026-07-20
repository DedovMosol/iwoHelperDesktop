using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Окно «О программе»: версия, автор, лицензия, ссылки и реквизиты для доната.</summary>
    public class AboutForm : Form
    {
        // Реквизиты доната — можно выделить и скопировать, а также скопировать по кнопке.
        internal const string DonationAccount = "40817810354405296071";
        internal const string DonationBank = "ПОВОЛЖСКИЙ БАНК ПАО СБЕРБАНК";

        private Label _copyStatus;

        public AboutForm()
        {
            Text = "О программе";
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
            Ui.Label(this, "Версия " + version.ToString(3), 88, 58, Font, Theme.TextMuted);

            // Описание — выравнивание по ширине; ширина ограничена окном (не вылезает за край).
            var desc = new JustifiedLabel();
            desc.Font = Font;
            desc.ForeColor = Theme.TextPrimary;
            desc.SetBounds(24, 96, ClientSize.Width - 48, 10);
            desc.Text = "Офисные инструменты: свод листов Excel, объединение и разделение PDF со сжатием.";
            desc.Height = desc.GetPreferredHeight();
            Controls.Add(desc);

            int y = desc.Bottom + 14;
            Ui.Label(this, "Автор: Dodonov Andrey (DedovMosol)", 24, y, Font, Theme.TextPrimary); y += 24;
            Ui.Label(this, "© 2026 · Лицензия MIT", 24, y, Font, Theme.TextMuted); y += 30;

            Label tg = Ui.Label(this, "Telegram:", 24, y, Font, Theme.TextPrimary);
            Ui.UrlLink(this, "t.me/i_wantout", tg.Right + 6, y, "https://t.me/i_wantout"); y += 24;
            Label gh = Ui.Label(this, "GitHub:", 24, y, Font, Theme.TextPrimary);
            Ui.UrlLink(this, "DedovMosol/iwoHelperDesktop", gh.Right + 6, y,
                "https://github.com/DedovMosol/iwoHelperDesktop"); y += 24;
            Label pp = Ui.Label(this, "Конфиденциальность:", 24, y, Font, Theme.TextPrimary);
            Ui.UrlLink(this, "политика (данные не покидают ваш ПК)", pp.Right + 6, y,
                "https://github.com/DedovMosol/iwoHelperDesktop/blob/main/docs/PRIVACY.md"); y += 34;

            // --- Донаты: реквизиты можно выделить и скопировать (read-only TextBox) ---
            Ui.Label(this, "Поддержать проект (донаты):", 24, y,
                new Font("Segoe UI", 9.75f, FontStyle.Bold), Theme.TextPrimary); y += 26;

            Label accCap = Ui.Label(this, "Счёт:", 24, y, Font, Theme.TextPrimary);
            TextBox accBox = SelectableValue(DonationAccount, accCap.Right + 6, y - 1, 168);
            LinkLabel copy = Ui.Link(this, "копировать", accBox.Right + 12, y);
            copy.LinkClicked += delegate { Copy(DonationAccount); };
            y += 26;

            Label bankCap = Ui.Label(this, "Банк:", 24, y, Font, Theme.TextPrimary);
            SelectableValue(DonationBank, bankCap.Right + 6, y - 1, ClientSize.Width - (bankCap.Right + 6) - 24);
            y += 26;

            _copyStatus = Ui.Label(this, "", 24, y, Font, Theme.OkGreen); y += 24;

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

        private void Copy(string text)
        {
            try
            {
                Clipboard.SetText(text);
                _copyStatus.ForeColor = Theme.OkGreen;
                _copyStatus.Text = "✓ Скопировано в буфер обмена";
            }
            catch
            {
                _copyStatus.ForeColor = Theme.ErrRed;
                _copyStatus.Text = "Не удалось скопировать (буфер занят)";
            }
        }
    }
}
