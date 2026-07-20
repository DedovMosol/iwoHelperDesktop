using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Единый фирменный диалог-сообщение (замена нативных MessageBox/TaskDialog):
    /// цветной значок по типу, заголовок, текст и кнопки. Одна кнопка — по центру,
    /// две — разнесены по краям (слева/справа). Вызывается через <see cref="Dialogs"/>.
    /// </summary>
    public class MessageForm : Form
    {
        public enum Kind { Info, Warning, Error }

        private const int W = 440;
        private const int Pad = 20;
        private const int IconSize = 40;
        private const int BtnW = 112;
        private const int BtnH = 36;

        private Bitmap _icon;

        public static void ShowMessage(IWin32Window owner, Kind kind, string title, string header, string body)
        {
            using (var f = new MessageForm(kind, title, header, body, false, null, null))
                f.ShowDialog(owner);
        }

        /// <summary>Сообщение с кликабельной ссылкой под текстом (например, страница загрузки).</summary>
        public static void ShowMessage(IWin32Window owner, Kind kind, string title, string header, string body,
            string linkText, string linkUrl)
        {
            using (var f = new MessageForm(kind, title, header, body, false, linkText, linkUrl))
                f.ShowDialog(owner);
        }

        public static bool ShowConfirm(IWin32Window owner, string title, string header, string body)
        {
            using (var f = new MessageForm(Kind.Warning, title, header, body, true, null, null))
                return f.ShowDialog(owner) == DialogResult.Yes;
        }

        private MessageForm(Kind kind, string title, string header, string body, bool confirm,
            string linkText, string linkUrl)
        {
            Text = title ?? "iwo Helper Desktop";
            Font = new Font("Segoe UI", 9.75f);
            BackColor = Color.White;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            Icon appIcon = Ui.AppIcon();
            if (appIcon != null)
                Icon = appIcon;

            int textLeft = Pad + IconSize + 16;
            int textW = W - textLeft - Pad;

            _icon = MakeIcon(kind, IconSize);
            var iconBox = new PictureBox();
            iconBox.Image = _icon;
            iconBox.SetBounds(Pad, Pad, IconSize, IconSize);
            Controls.Add(iconBox);

            var hdr = new Label();
            hdr.Font = new Font("Segoe UI", 11f, FontStyle.Bold);
            hdr.ForeColor = Color.FromArgb(40, 40, 40);
            hdr.MaximumSize = new Size(textW, 0);
            hdr.AutoSize = true;
            hdr.Text = header ?? "";
            hdr.Location = new Point(textLeft, Pad);
            Controls.Add(hdr);

            int y = hdr.Bottom;
            if (!string.IsNullOrEmpty(body))
            {
                var text = new Label();
                text.ForeColor = Theme.TextPrimary;
                text.MaximumSize = new Size(textW, 0);
                text.AutoSize = true;
                text.Text = body;
                text.Location = new Point(textLeft, y + 8);
                Controls.Add(text);
                y = text.Bottom;
            }

            // Кликабельная ссылка (например, страница загрузки Ghostscript). Высота
            // учитывается в y ДО расчёта кнопок/ClientSize — иначе ссылка легла бы на кнопки.
            if (!string.IsNullOrEmpty(linkText) && !string.IsNullOrEmpty(linkUrl))
            {
                LinkLabel link = Ui.UrlLink(this, linkText, textLeft, y + 10, linkUrl);
                y = link.Bottom;
            }

            int textBottom = Math.Max(y, Pad + IconSize);
            int btnY = textBottom + 24;

            if (confirm)
            {
                var no = new RoundedButton(false);
                no.Text = "Нет";
                no.SetBounds(ButtonX(0, 2, W, BtnW, Pad), btnY, BtnW, BtnH);
                no.DialogResult = DialogResult.No;
                Controls.Add(no);

                var yes = new RoundedButton(true);
                yes.Text = "Да";
                yes.SetBounds(ButtonX(1, 2, W, BtnW, Pad), btnY, BtnW, BtnH);
                yes.DialogResult = DialogResult.Yes;
                Controls.Add(yes);

                AcceptButton = yes;
                CancelButton = no;
            }
            else
            {
                var ok = new RoundedButton(true);
                ok.Text = "OK";
                ok.SetBounds(ButtonX(0, 1, W, BtnW, Pad), btnY, BtnW, BtnH);
                ok.DialogResult = DialogResult.OK;
                Controls.Add(ok);

                AcceptButton = ok;
                CancelButton = ok;
            }

            ClientSize = new Size(W, btnY + BtnH + Pad);
        }

        /// <summary>X кнопки: одна — по центру; две — по краям (0 слева, 1 справа). Чистая — под тест.</summary>
        internal static int ButtonX(int index, int count, int width, int btnW, int margin)
        {
            if (count == 1)
                return (width - btnW) / 2;
            return index == 0 ? margin : width - margin - btnW;
        }

        private static Bitmap MakeIcon(Kind kind, int size)
        {
            Color color = kind == Kind.Error ? Theme.ErrRed
                : kind == Kind.Warning ? Theme.WarnOrange
                : Theme.HubBlue;
            string glyph = kind == Kind.Error ? "✕" : kind == Kind.Warning ? "!" : "i";

            var bmp = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                using (var b = new SolidBrush(color))
                    g.FillEllipse(b, 0.5f, 0.5f, size - 1f, size - 1f);
                using (var font = new Font("Segoe UI", size * 0.5f, FontStyle.Bold))
                using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                    g.DrawString(glyph, font, Brushes.White, new RectangleF(0, -1, size, size), sf);
            }
            return bmp;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _icon != null)
                _icon.Dispose();
            base.Dispose(disposing);
        }
    }
}
