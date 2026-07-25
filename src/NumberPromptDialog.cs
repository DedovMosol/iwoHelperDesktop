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
            const int margin = 16, btnW = 112, btnH = 30;
            // База 300 вмещает кнопки по краям; длинная подпись (например, 5-значное число
            // страниц в «до {0}») расширяет окно по замеру — статичная константа обрезала бы её.
            int w = DialogWidth(TextRenderer.MeasureText(prompt, Font).Width, margin, 300);
            BackColor = Color.White;

            var label = new Label();
            label.AutoSize = true;
            label.Location = new Point(margin, 16);
            label.Text = prompt;
            Controls.Add(label);

            _num = new NumericUpDown();
            _num.Minimum = min;
            _num.Maximum = max;
            _num.Value = initial < min ? min : (initial > max ? max : initial);
            _num.SetBounds(margin, 42, 100, 27);
            Controls.Add(_num);

            int btnY = 76;
            // Кнопки разнесены по краям (как в MessageForm): «Отмена» слева, действие справа.
            // Ширина 112 вмещает самую длинную подпись («Переместить» ≈ 85 px) без обрезки.
            var cancel = new RoundedButton(false);
            cancel.Text = Loc.T("common.cancel");
            cancel.SetBounds(MessageForm.ButtonX(0, 2, w, btnW, margin), btnY, btnW, btnH);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);

            var ok = new RoundedButton(true);
            ok.Text = okText;
            ok.SetBounds(MessageForm.ButtonX(1, 2, w, btnW, margin), btnY, btnW, btnH);
            ok.DialogResult = DialogResult.OK;
            Controls.Add(ok);

            ClientSize = new Size(w, btnY + btnH + margin);
            AcceptButton = ok;
            CancelButton = cancel;
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _num.Select(0, _num.Value.ToString().Length); // сразу печатать номер без очистки поля
            _num.Focus();
        }

        /// <summary>Ширина диалога: не уже базовой и не уже подписи с полями. Чистая — под тест.</summary>
        internal static int DialogWidth(int promptWidth, int margin, int minWidth)
        {
            return Math.Max(minWidth, promptWidth + 2 * margin);
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
