using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Стартовый экран — хаб выбора инструмента (свод Excel и PDF-инструменты) 2×2.
    /// Только представление: открытие инструментов, дедупликацию и жизненный цикл
    /// окон ведёт <see cref="ShellContext"/>. Закрытие хаба не закрывает уже
    /// открытые инструменты; кнопка «Главная» в инструменте снова покажет этот экран.
    /// </summary>
    public class StartForm : Form
    {
        private const string AppTitle = "iwo Helper Desktop";
        private readonly ShellContext _context;
        private ToolTip _langTip;         // подсказка кнопки-глобуса (компонент — освобождаем вручную)
        private ContextMenuStrip _langMenu; // меню выбора языка (одно на окно; окно пересоздаётся при смене языка)

        public StartForm() : this(null) { } // для смоук-теста; открытие инструментов недоступно

        internal StartForm(ShellContext context)
        {
            _context = context;

            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            Text = AppTitle + " " + version.ToString(3);
            Icon startIcon = Ui.AppIcon();
            if (startIcon != null)
                Icon = startIcon;
            Font = new Font("Segoe UI", 9.75f);
            BackColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(546, 692); // 2×2 карточки
            WindowChrome.Enable(this, Theme.HubBlue); // синий заголовок на Windows 11

            var header = new HeaderBand(AppTitle, Loc.T("hub.subtitle"),
                Theme.HubBlue, Theme.HubBlueDark);
            header.Centered = true; // на стартовом экране заголовок и подпись по центру
            header.SetBounds(0, 0, ClientSize.Width, 78);
            header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(header);

            // Выбор языка — белый глиф-глобус в правом верхнем углу шапки (на синем, без рамки).
            _langMenu = HelpMenu.LanguageContextMenu(); // одно меню на жизнь окна (окно пересоздаётся при смене языка)
            var globe = new GlyphButton("", 15f); // U+E774 — «глобус» (Segoe MDL2 Assets)
            globe.ForeColor = Color.White;
            globe.SetBounds(header.Width - 42, 10, 30, 30);
            globe.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            globe.Click += delegate { _langMenu.Show(globe, new Point(globe.Width, globe.Height), ToolStripDropDownDirection.BelowLeft); };
            _langTip = new ToolTip();
            _langTip.SetToolTip(globe, Loc.T("lang.tooltip"));
            header.Controls.Add(globe);

            var excel = new ChoiceCard(CardGlyph.Excel, Loc.T("hub.excel.name"), Loc.T("hub.excel.desc"));
            excel.SetBounds(24, 96, 240, 250);
            excel.Click += delegate
            {
                if (_context != null)
                    _context.OpenTool("excel", Loc.T("hub.excel.name"), delegate(Action back) { return new MainForm(back); });
            };
            Controls.Add(excel);

            var pdf = new ChoiceCard(CardGlyph.Pdf, Loc.T("hub.pdf.name"), Loc.T("hub.pdf.desc"));
            pdf.SetBounds(282, 96, 240, 250);
            pdf.Click += delegate
            {
                if (_context != null)
                    _context.OpenTool("pdf", Loc.T("hub.pdf.name"), delegate(Action back) { return new PdfMergeForm(back); });
            };
            Controls.Add(pdf);

            var split = new ChoiceCard(CardGlyph.PdfSplit, Loc.T("hub.split.name"), Loc.T("hub.split.desc"));
            split.SetBounds(24, 364, 240, 250);
            split.Click += delegate
            {
                if (_context != null)
                    _context.OpenTool("split", Loc.T("hub.split.name"), delegate(Action back) { return new PdfSplitForm(back); });
            };
            Controls.Add(split);

            var ocr = new ChoiceCard(CardGlyph.Ocr, Loc.T("hub.ocr.name"), Loc.T("hub.ocr.desc"));
            ocr.SetBounds(282, 364, 240, 250);
            ocr.Click += delegate
            {
                if (_context != null)
                    _context.OpenTool("ocr", Loc.T("hub.ocr.name"), delegate(Action back) { return new OcrForm(back); });
            };
            Controls.Add(ocr);

            // Нижний ряд: «Проверить обновления» слева (на месте версии), «О программе» справа.
            const int rowY = 632, rowH = 36;
            var update = new RoundedButton(false);
            update.Text = Loc.T("hub.update");
            update.SetBounds(24, rowY, 224, rowH);
            update.Click += delegate { UpdateUi.Check(this); };
            Controls.Add(update);

            var about = new RoundedButton(false);
            about.Text = Loc.T("hub.about");
            about.SetBounds(ClientSize.Width - 24 - 168, rowY, 168, rowH);
            about.Click += delegate { using (var f = new AboutForm()) f.ShowDialog(this); };
            Controls.Add(about);

            AcceptButton = null; // Enter активирует карточку в фокусе
        }

        protected override void Dispose(bool disposing)
        {
            // ToolTip и ContextMenuStrip — компоненты (не дочерние контролы): освобождаем вручную.
            if (disposing)
            {
                if (_langTip != null) _langTip.Dispose();
                if (_langMenu != null) _langMenu.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Иконка-глиф без рамки на прозрачном фоне (для белого глобуса на синей шапке).
        /// SupportsTransparentBackColor + BackColor=Transparent показывает фон родителя
        /// (градиент шапки); глиф рисуется по центру. Кликается как кнопка.
        /// </summary>
        private sealed class GlyphButton : Control
        {
            public GlyphButton(string glyph, float size)
            {
                SetStyle(ControlStyles.SupportsTransparentBackColor | ControlStyles.UserPaint |
                         ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
                BackColor = Color.Transparent;
                Font = new Font("Segoe MDL2 Assets", size);
                Text = glyph;
                Cursor = Cursors.Hand;
                TabStop = false;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, ForeColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
            }
        }
    }
}
