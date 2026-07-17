using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Стартовый экран: выбор инструмента (свод Excel или объединение PDF).
    /// Выбранный инструмент открывается модально; после закрытия снова
    /// показывается выбор — можно перейти к другому инструменту без перезапуска.
    /// </summary>
    public class StartForm : Form
    {
        public StartForm()
        {
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            Text = "iwo Helper Desktop " + version.ToString(2);
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
            catch { }
            Font = new Font("Segoe UI", 9.75f);
            BackColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(600, 380);

            Ui.AccentBar(this, 0);
            Ui.Label(this, "Выберите инструмент", 30, 26,
                new Font("Segoe UI", 15f, FontStyle.Bold), Color.FromArgb(40, 40, 40));
            Ui.Label(this, "Что нужно сделать?", 32, 58, Font, Theme.TextMuted);

            var excel = new ChoiceCard(CardGlyph.Excel, "Свод Excel",
                "Объединить листы из нескольких файлов Excel в один свод: оглавление, замена формул значениями, сопроводительная записка Word.");
            excel.SetBounds(30, 96, 262, 250);
            excel.Click += delegate { OpenTool(new MainForm()); };
            Controls.Add(excel);

            var pdf = new ChoiceCard(CardGlyph.Pdf, "Объединение PDF",
                "Собрать один PDF из нескольких файлов: выбрать нужные страницы и задать их порядок. Страницы копируются без искажений.");
            pdf.SetBounds(308, 96, 262, 250);
            pdf.Click += delegate { OpenTool(new PdfMergeForm()); };
            Controls.Add(pdf);

            AcceptButton = null; // Enter активирует карточку в фокусе, а не «первую кнопку»
        }

        private void OpenTool(Form tool)
        {
            Hide();
            try
            {
                tool.ShowInTaskbar = true;
                tool.ShowDialog();
            }
            finally
            {
                tool.Dispose();
                if (!IsDisposed)
                    Show();
            }
        }
    }
}
