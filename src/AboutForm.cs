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
                return JustifiedParagraph(text, x, y, width);
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

        /// <summary>
        /// Абзац, выключенный ПО ШИРИНЕ. WinForms такого выравнивания не предлагает ни у
        /// TextBox, ни у RichTextBox (в перечислении только лево/право/центр), поэтому
        /// выставляем его формату абзаца напрямую сообщением EM_SETPARAFORMAT — это штатный
        /// способ, которым пользуется и сам элемент для остальных выравниваний.
        /// </summary>
        private RichTextBox JustifiedParagraph(string text, int x, int y, int width)
        {
            var rtb = new RichTextBox();
            rtb.Multiline = true;
            rtb.WordWrap = true;
            rtb.ScrollBars = RichTextBoxScrollBars.None;
            rtb.BorderStyle = BorderStyle.None;
            rtb.BackColor = Color.White;
            rtb.ForeColor = Theme.TextPrimary;
            rtb.Font = Font;
            rtb.TabStop = false;
            rtb.ReadOnly = true;
            rtb.Text = text;
            rtb.SetBounds(x, y, width, ParagraphHeight(text, width));
            // Формат абзаца хранится в самом окне элемента, а WinForms окно пересоздаёт (смена
            // родителя, пересборка интерфейса) — поэтому ставим выключку на КАЖДОЕ создание
            // хэндла, иначе разовая установка в конструкторе однажды потерялась бы.
            rtb.HandleCreated += delegate { Justify(rtb); };
            Controls.Add(rtb);
            return rtb;
        }

        private const int WmUser = 0x0400;
        private const int EmGetParaFormat = WmUser + 61;
        private const int EmSetParaFormat = WmUser + 71;
        private const int PfmAlignment = 0x00000008;
        private const short PfaJustify = 4;  // выключка по ширине

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct ParaFormat2
        {
            public int cbSize, dwMask;
            public short wNumbering, wReserved;
            public int dxStartIndent, dxRightIndent, dxOffset;
            public short wAlignment, cTabCount;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 32)]
            public int[] rgxTabs;
            public int dySpaceBefore, dySpaceAfter, dyLineSpacing;
            public short sStyle;
            public byte bLineSpacingRule, bOutlineLevel;
            public short wShadingWeight, wShadingStyle;
            public short wNumberingStart, wNumberingStyle, wNumberingTab, wBorderSpace, wBorderWidth, wBorders;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref ParaFormat2 lParam);

        /// <summary>Выключить весь текст по ширине. Сбой не критичен — останется обычная выключка влево.</summary>
        private static void Justify(RichTextBox rtb)
        {
            try
            {
                var fmt = new ParaFormat2();
                fmt.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(ParaFormat2));
                fmt.dwMask = PfmAlignment;
                fmt.wAlignment = PfaJustify;
                fmt.rgxTabs = new int[32];
                // Сообщение действует на абзацы ВЫДЕЛЕНИЯ (wParam обязан быть нулём), поэтому
                // выделяем всё, применяем и снимаем выделение — на экране это не мелькает.
                rtb.SelectAll();
                SendMessage(rtb.Handle, EmSetParaFormat, IntPtr.Zero, ref fmt);
                rtb.Select(0, 0);
            }
            catch { } // выключка — оформление, а не работа: не получилось, текст всё равно читается
        }

        /// <summary>Выключен ли абзац по ширине (спрашиваем сам элемент). Нужно тесту раскладки.</summary>
        internal static bool IsJustified(RichTextBox rtb)
        {
            var fmt = new ParaFormat2();
            fmt.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(ParaFormat2));
            fmt.rgxTabs = new int[32];
            SendMessage(rtb.Handle, EmGetParaFormat, IntPtr.Zero, ref fmt);
            return (fmt.dwMask & PfmAlignment) != 0 && fmt.wAlignment == PfaJustify;
        }

        /// <summary>
        /// Высота абзаца под заданную ширину. Меряем чуть более узкой строкой и добавляем запас:
        /// TextBox переносит слова по своей внутренней ширине, которая на пару пикселей меньше
        /// заданной, и без запаса последняя строка обрезалась бы.
        /// </summary>
        private int ParagraphHeight(string text, int width)
        {
            const int Inset = 4; // внутренние поля TextBox с обеих сторон
            Size measured = TextRenderer.MeasureText(text, Font,
                new Size(Math.Max(width - Inset, 1), int.MaxValue), TextFormatFlags.WordBreak);
            return measured.Height + Inset;
        }

    }
}
