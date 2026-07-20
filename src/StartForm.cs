using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Стартовый экран — хаб выбора инструмента (свод Excel или объединение PDF).
    /// Только представление: открытие инструментов, дедупликацию и жизненный цикл
    /// окон ведёт <see cref="ShellContext"/>. Закрытие хаба не закрывает уже
    /// открытые инструменты; кнопка «Главная» в инструменте снова покажет этот экран.
    /// </summary>
    public class StartForm : Form
    {
        private const string AppTitle = "iwo Helper Desktop";
        private readonly ShellContext _context;

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
            ClientSize = new Size(804, 416);
            WindowChrome.Enable(this, Theme.HubBlue); // синий заголовок на Windows 11

            var header = new HeaderBand(AppTitle, "Выберите инструмент — свод Excel, объединение или разделение PDF",
                Theme.HubBlue, Theme.HubBlueDark);
            header.Centered = true; // на стартовом экране заголовок и подпись по центру
            header.SetBounds(0, 0, ClientSize.Width, 78);
            header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(header);

            var excel = new ChoiceCard(CardGlyph.Excel, "Свод Excel",
                "Объединить листы из нескольких файлов Excel в один свод: оглавление, замена формул значениями, сопроводительная записка Word.");
            excel.SetBounds(24, 96, 240, 250);
            excel.Click += delegate
            {
                if (_context != null)
                    _context.OpenTool("excel", "Свод Excel", delegate(Action back) { return new MainForm(back); });
            };
            Controls.Add(excel);

            var pdf = new ChoiceCard(CardGlyph.Pdf, "Объединение PDF",
                "Собрать один PDF из нескольких файлов: выбрать нужные страницы и задать их порядок. Страницы копируются без искажений.");
            pdf.SetBounds(282, 96, 240, 250);
            pdf.Click += delegate
            {
                if (_context != null)
                    _context.OpenTool("pdf", "Объединение PDF", delegate(Action back) { return new PdfMergeForm(back); });
            };
            Controls.Add(pdf);

            var split = new ChoiceCard(CardGlyph.PdfSplit, "Разделение PDF",
                "Извлечь выбранные страницы в один PDF или разбить документ на несколько: по диапазонам, каждые N страниц или по закладкам.");
            split.SetBounds(540, 96, 240, 250);
            split.Click += delegate
            {
                if (_context != null)
                    _context.OpenTool("split", "Разделение PDF", delegate(Action back) { return new PdfSplitForm(back); });
            };
            Controls.Add(split);

            // Нижний ряд: «Проверить обновления» слева (на месте версии), «О программе»
            // справа. Чуть выше и крупнее — аккуратный ряд под карточками.
            const int rowY = 360, rowH = 36;
            var update = new RoundedButton(false);
            update.Text = "⟳ Проверить обновления";
            update.SetBounds(24, rowY, 224, rowH);
            update.Click += delegate { UpdateUi.Check(this); };
            Controls.Add(update);

            var about = new RoundedButton(false);
            about.Text = "О программе";
            about.SetBounds(ClientSize.Width - 24 - 168, rowY, 168, rowH);
            about.Click += delegate { using (var f = new AboutForm()) f.ShowDialog(this); };
            Controls.Add(about);

            AcceptButton = null; // Enter активирует карточку в фокусе
        }
    }
}
