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
            Ui.InitDialog(this, title);
            const int margin = 16, btnH = 30;

            // Кнопки создаём ДО раскладки: ширину задаёт самая длинная подпись, а мерить её
            // надо шрифтом самой кнопки — главная кнопка набрана крупнее и полужирным, и от
            // прибитой константы 112 подпись «Переместить» (120 px) уходила в многоточие.
            var cancel = new RoundedButton(false);
            cancel.Text = Loc.T("common.cancel");
            cancel.DialogResult = DialogResult.Cancel;
            var ok = new RoundedButton(true);
            ok.Text = okText;
            ok.DialogResult = DialogResult.OK;
            int btnW = ButtonWidth(Ui.TextWidth(cancel.Text, cancel.Font), Ui.TextWidth(ok.Text, ok.Font));

            // База 300 вмещает кнопки по краям; длинная подпись (например, 5-значное число
            // страниц в «до {0}») расширяет окно по замеру — статичная константа обрезала бы её.
            int w = Math.Max(DialogWidth(TextRenderer.MeasureText(prompt, Font).Width, margin, 300),
                2 * btnW + 3 * margin); // разнесённым по краям кнопкам нужно не сойтись вплотную

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
            cancel.SetBounds(MessageForm.ButtonX(0, 2, w, btnW, margin), btnY, btnW, btnH);
            Controls.Add(cancel);
            ok.SetBounds(MessageForm.ButtonX(1, 2, w, btnW, margin), btnY, btnW, btnH);
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

        /// <summary>
        /// Ширина обеих кнопок: обе одинаковы (они стоят по краям одной строки) и вмещают
        /// самую длинную подпись с полями, которые рисует <see cref="RoundedButton"/>.
        /// Чистая — под тест.
        /// </summary>
        internal static int ButtonWidth(int cancelWidth, int okWidth)
        {
            const int minWidth = 112;
            int text = Math.Max(cancelWidth, okWidth);
            int need = text + 2 * RoundedButton.TextPadFor(minWidth);
            return need < minWidth ? minWidth : need;
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
