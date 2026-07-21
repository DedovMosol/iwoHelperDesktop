using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Окно «Статистика»: локальные счётчики операций, выбор авто-очистки
    /// (выкл / раз в день / 7 / 30 дней) и ручная очистка. Данные — в UsageStats.
    /// </summary>
    public class StatsForm : Form
    {
        private static readonly int[] AutoDays = { 0, 1, 7, 30 }; // индекс combo -> дни

        private Label _since;
        private Label _excel, _merge, _extract, _ranges, _everyN, _bookmarks, _pdftoword, _compress, _total;
        private ComboBox _cmbAuto;
        private bool _loading;

        public StatsForm()
        {
            Text = "Статистика";
            Font = new Font("Segoe UI", 9.75f);
            BackColor = Color.White;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(440, 444);
            Icon appIcon = Ui.AppIcon();
            if (appIcon != null)
                Icon = appIcon;
            WindowChrome.Enable(this, Theme.HubBlue);

            Ui.AccentBar(this, 0, Theme.HubBlue);
            Ui.Label(this, "Статистика", 24, 22, new Font("Segoe UI", 14f, FontStyle.Bold), Color.FromArgb(40, 40, 40));
            _since = Ui.Label(this, "", 24, 56, Font, Theme.TextMuted);

            _excel = Row("Своды Excel", 92);
            _merge = Row("Объединения PDF", 118);
            _extract = Row("Извлечения страниц (PDF)", 144);
            _ranges = Row("Разбиение по диапазонам", 170);
            _everyN = Row("Разбиение: каждые N страниц", 196);
            _bookmarks = Row("Разбиение по закладкам", 222);
            _pdftoword = Row("Конвертации PDF → Word", 248);
            _compress = Row("Сжатия PDF (файлов)", 274);
            _total = Ui.Label(this, "", 24, 306, new Font("Segoe UI", 10.5f, FontStyle.Bold), Theme.TextPrimary);

            Ui.Label(this, "Автоочистка:", 24, 352, Font, Theme.TextPrimary);
            _cmbAuto = new ComboBox();
            _cmbAuto.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbAuto.Items.AddRange(new object[] { "Выключена", "Раз в день", "Раз в 7 дней", "Раз в 30 дней" });
            _cmbAuto.SetBounds(120, 349, 180, 27);
            _cmbAuto.SelectedIndexChanged += OnAutoChanged;
            Controls.Add(_cmbAuto);
            _tips = new ToolTip();
            _tips.SetToolTip(_cmbAuto, "Счётчики будут автоматически обнуляться с выбранной периодичностью");

            var clear = new RoundedButton(false);
            clear.Text = "Очистить";
            clear.SetBounds(24, ClientSize.Height - 52, 130, 36);
            clear.Click += OnClear;
            Controls.Add(clear);

            var close = new RoundedButton(true);
            close.Text = "Закрыть";
            close.SetBounds(ClientSize.Width - 124, ClientSize.Height - 52, 100, 36);
            close.Click += delegate { Close(); };
            Controls.Add(close);
            AcceptButton = close;
            CancelButton = close;

            LoadAndShow();
        }

        private ToolTip _tips;

        private Label Row(string caption, int y)
        {
            Ui.Label(this, caption, 24, y, Font, Theme.TextPrimary);
            var value = Ui.Label(this, "0", 300, y, new Font("Segoe UI", 9.75f, FontStyle.Bold), Theme.TextPrimary);
            return value;
        }

        private void LoadAndShow()
        {
            UsageStats s = UsageStats.Load();
            _loading = true;
            _since.Text = "Считается с " + s.SinceUtc.ToLocalTime().ToString("dd.MM.yyyy") + ".";
            _excel.Text = s.ExcelDigests.ToString();
            _merge.Text = s.PdfMerges.ToString();
            _extract.Text = s.PdfExtracts.ToString();
            _ranges.Text = s.PdfSplitRanges.ToString();
            _everyN.Text = s.PdfSplitEveryN.ToString();
            _bookmarks.Text = s.PdfSplitBookmarks.ToString();
            _pdftoword.Text = s.PdfToWord.ToString();
            _compress.Text = s.PdfCompressions.ToString();
            _total.Text = "Всего операций: " + s.Total;
            _cmbAuto.SelectedIndex = Array.IndexOf(AutoDays, s.AutoClearDays) >= 0 ? Array.IndexOf(AutoDays, s.AutoClearDays) : 0;
            _loading = false;
        }

        private void OnAutoChanged(object sender, EventArgs e)
        {
            if (_loading)
                return;
            int days = AutoDays[_cmbAuto.SelectedIndex];
            UsageStats.SetAutoClear(days);
            LoadAndShow(); // применённая автоочистка могла обнулить счётчики
        }

        private void OnClear(object sender, EventArgs e)
        {
            if (Dialogs.ConfirmWarning(this, "Статистика", "Очистить счётчики?",
                    "Все накопленные числа обнулятся. Действие необратимо."))
            {
                UsageStats.ClearCounters();
                LoadAndShow();
            }
        }
    }
}
