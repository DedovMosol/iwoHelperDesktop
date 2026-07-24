using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Маленький модальный диалог выбора числа — общий для «Перейти к странице» (Ctrl+G)
    /// и «Переместить после страницы N…» (DRY). Enter — подтвердить, Esc — отмена.
    /// Возвращает число в [min..max] или -1 при отмене. Освобождается вызывающим
    /// (using в <see cref="Show"/>).
    /// </summary>
    internal sealed class NumberPromptDialog : Form
    {
        private readonly NumericUpDown _num;

        private NumberPromptDialog(string title, string prompt, string okText, int min, int max, int initial)
        {
            Text = title;
            Font = new Font("Segoe UI", 9.75f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = MaximizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(300, 108);
            BackColor = Color.White;

            var label = new Label();
            label.AutoSize = true;
            label.Location = new Point(16, 16);
            label.Text = prompt;
            Controls.Add(label);

            _num = new NumericUpDown();
            _num.Minimum = min;
            _num.Maximum = max;
            _num.Value = initial < min ? min : (initial > max ? max : initial);
            _num.SetBounds(16, 40, 100, 27);
            Controls.Add(_num);

            var ok = new RoundedButton(true);
            ok.Text = okText;
            ok.SetBounds(120, 72, 84, 28);
            ok.DialogResult = DialogResult.OK;
            Controls.Add(ok);

            var cancel = new RoundedButton(false);
            cancel.Text = Loc.T("common.cancel");
            cancel.SetBounds(210, 72, 74, 28);
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

        /// <summary>Показать диалог; вернуть число в [min..max] или -1 при отмене.</summary>
        public static int Show(IWin32Window owner, string title, string prompt, string okText, int min, int max, int initial)
        {
            if (max < min)
                return -1;
            using (var dialog = new NumberPromptDialog(title, prompt, okText, min, max, initial))
                return dialog.ShowDialog(owner) == DialogResult.OK ? (int)dialog._num.Value : -1;
        }
    }
}
