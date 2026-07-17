using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Стартовый экран — хаб выбора инструмента (свод Excel или объединение PDF).
    /// Инструменты открываются немодально: можно держать открытыми несколько
    /// разных одновременно; повторное открытие того же — с уведомлением и
    /// переводом фокуса на уже открытое окно. Кнопка «Назад в меню» в каждом
    /// инструменте возвращает фокус сюда.
    /// </summary>
    public class StartForm : Form
    {
        private const string AppTitle = "iwo Helper Desktop";
        private readonly ToolRegistry _tools = new ToolRegistry();

        public StartForm()
        {
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            Text = AppTitle + " " + version.ToString(2);
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
            WindowChrome.Enable(this, Theme.HubBlue); // синий заголовок на Windows 11

            var header = new HeaderBand(AppTitle, "Выберите инструмент — свод Excel или объединение PDF",
                Theme.HubBlue, Theme.HubBlueDark);
            header.SetBounds(0, 0, ClientSize.Width, 78);
            header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(header);

            var excel = new ChoiceCard(CardGlyph.Excel, "Свод Excel",
                "Объединить листы из нескольких файлов Excel в один свод: оглавление, замена формул значениями, сопроводительная записка Word.");
            excel.SetBounds(30, 96, 262, 250);
            excel.Click += delegate
            {
                OpenTool("excel", "Свод Excel", delegate(Action back) { return new MainForm(back); });
            };
            Controls.Add(excel);

            var pdf = new ChoiceCard(CardGlyph.Pdf, "Объединение PDF",
                "Собрать один PDF из нескольких файлов: выбрать нужные страницы и задать их порядок. Страницы копируются без искажений.");
            pdf.SetBounds(308, 96, 262, 250);
            pdf.Click += delegate
            {
                OpenTool("pdf", "Объединение PDF", delegate(Action back) { return new PdfMergeForm(back); });
            };
            Controls.Add(pdf);

            AcceptButton = null; // Enter активирует карточку в фокусе
        }

        private void OpenTool(string key, string name, Func<Action, Form> factory)
        {
            Form existing;
            if (_tools.TryGetOpen(key, out existing))
            {
                Dialogs.Info(this, AppTitle, "Инструмент уже открыт",
                    "«" + name + "» уже запущен — открыто его окно.");
                BringToFront(existing);
                return;
            }

            Form tool = factory(BackToMenu);
            _tools.Add(key, tool);
            tool.FormClosed += delegate { _tools.Remove(key); };
            tool.Show(); // немодально, без владельца — отдельное окно в панели задач
            BringToFront(tool);
        }

        private void BackToMenu()
        {
            BringToFront(this);
        }

        private static void BringToFront(Form form)
        {
            if (form == null || form.IsDisposed)
                return;
            if (form.WindowState == FormWindowState.Minimized)
                form.WindowState = FormWindowState.Normal;
            form.Activate();
            form.BringToFront();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Закрытие хаба закрывает открытые инструменты. Если инструмент занят
            // (идёт слияние) и отменяет своё закрытие — отменяем закрытие хаба.
            foreach (Form tool in _tools.OpenForms())
            {
                tool.Close();
                if (!tool.IsDisposed)
                {
                    BringToFront(tool);
                    e.Cancel = true;
                    return;
                }
            }
            base.OnFormClosing(e);
        }
    }
}
