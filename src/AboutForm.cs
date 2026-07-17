using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Окно «О программе»: версия, автор, лицензия, ссылки.</summary>
    public class AboutForm : Form
    {
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
            ClientSize = new Size(460, 306);

            Ui.AccentBar(this, 0);

            var iconBox = new PictureBox();
            iconBox.SetBounds(24, 26, 48, 48);
            iconBox.SizeMode = PictureBoxSizeMode.StretchImage;
            try
            {
                Icon appIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                if (appIcon != null)
                    iconBox.Image = appIcon.ToBitmap();
            }
            catch { } // без картинки, с одним текстом
            Controls.Add(iconBox);

            Ui.Label(this, "iwo Helper Desktop", 86, 26,
                new Font("Segoe UI", 14f, FontStyle.Bold), Color.FromArgb(40, 40, 40));
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            Ui.Label(this, "Версия " + version.ToString(3), 88, 58, Font, Theme.TextMuted);

            Ui.Label(this, "Офисные инструменты: свод листов Excel и объединение PDF.",
                24, 96, Font, Theme.TextPrimary);
            Ui.Label(this, "Автор: DedovMosol", 24, 128, Font, Theme.TextPrimary);
            Ui.Label(this, "© 2026 · Лицензия MIT", 24, 152, Font, Theme.TextMuted);

            Ui.UrlLink(this, "Telegram: t.me/i_wantout", 24, 186, "https://t.me/i_wantout");
            Ui.UrlLink(this, "GitHub: DedovMosol/iwoHelperDesktop", 24, 212,
                "https://github.com/DedovMosol/iwoHelperDesktop");

            // Кнопка OK — в самом низу, ниже ссылок (иначе длинная ссылка налезала).
            var ok = new RoundedButton(true);
            ok.Text = "OK";
            ok.SetBounds(ClientSize.Width - 124, ClientSize.Height - 52, 100, 36);
            ok.Click += delegate { Close(); };
            Controls.Add(ok);
            AcceptButton = ok;
            CancelButton = ok; // Esc тоже закрывает
        }
    }
}
