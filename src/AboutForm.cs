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
            ClientSize = new Size(460, 340);
            WindowChrome.Enable(this, Theme.HubBlue); // синий заголовок на Windows 11 — как на главной

            Ui.AccentBar(this, 0, Theme.HubBlue); // синяя полоса, как на стартовом экране

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

            // Описание — с выравниванием по ширине (justify); ширина ограничена окном,
            // поэтому за край не выходит. Высота — под число строк; остальные строки
            // позиционируются относительно его низа (устойчиво к 1–2 строкам).
            var desc = new JustifiedLabel();
            desc.Font = Font;
            desc.ForeColor = Theme.TextPrimary;
            desc.SetBounds(24, 96, ClientSize.Width - 48, 10);
            desc.Text = "Офисные инструменты: свод листов Excel, объединение и разделение PDF со сжатием.";
            desc.Height = desc.GetPreferredHeight();
            Controls.Add(desc);
            int y = desc.Bottom + 14;
            Ui.Label(this, "Автор: DedovMosol", 24, y, Font, Theme.TextPrimary); y += 24;
            Ui.Label(this, "© 2026 · Лицензия MIT", 24, y, Font, Theme.TextMuted); y += 34;

            // Кликабельна только ссылка, подпись слева — обычный текст.
            Label tg = Ui.Label(this, "Telegram:", 24, y, Font, Theme.TextPrimary);
            Ui.UrlLink(this, "t.me/i_wantout", tg.Right + 6, y, "https://t.me/i_wantout"); y += 26;
            Label gh = Ui.Label(this, "GitHub:", 24, y, Font, Theme.TextPrimary);
            Ui.UrlLink(this, "DedovMosol/iwoHelperDesktop", gh.Right + 6, y,
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
