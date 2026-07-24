using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Маленький модальный диалог «Перейти к странице» (Ctrl+G в сетке миниатюр):
    /// номер в пределах 1..max, Enter — перейти, Esc — отмена. Возвращает номер
    /// (с единицы) или -1. Освобождается вызывающим (using в <see cref="Show"/>).
    /// </summary>
    internal sealed class GoToPageDialog : Form
    {
        private readonly NumericUpDown _num;

        private GoToPageDialog(int maxPage)
        {
            Text = Loc.T("goto.title");
            Font = new Font("Segoe UI", 9.75f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = MaximizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(280, 108);
            BackColor = Color.White;

            var label = new Label();
            label.AutoSize = true;
            label.Location = new Point(16, 16);
            label.Text = string.Format(Loc.T("goto.prompt"), maxPage);
            Controls.Add(label);

            _num = new NumericUpDown();
            _num.Minimum = 1;
            _num.Maximum = maxPage;
            _num.Value = 1;
            _num.SetBounds(16, 40, 100, 27);
            Controls.Add(_num);

            var ok = new RoundedButton(true);
            ok.Text = Loc.T("goto.ok");
            ok.SetBounds(120, 72, 70, 28);
            ok.DialogResult = DialogResult.OK;
            Controls.Add(ok);

            var cancel = new RoundedButton(false);
            cancel.Text = Loc.T("common.cancel");
            cancel.SetBounds(196, 72, 70, 28);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _num.Select(0, _num.Value.ToString().Length); // сразу печатать номер без очистки поля
            _num.Focus();
        }

        /// <summary>Показать диалог; вернуть номер страницы (1..maxPage) или -1 при отмене.</summary>
        public static int Show(IWin32Window owner, int maxPage)
        {
            if (maxPage < 1)
                return -1;
            using (var dialog = new GoToPageDialog(maxPage))
                return dialog.ShowDialog(owner) == DialogResult.OK ? (int)dialog._num.Value : -1;
        }
    }
}
