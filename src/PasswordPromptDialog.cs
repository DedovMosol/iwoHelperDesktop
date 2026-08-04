using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Ввод пароля к защищённому PDF. Enter — подтвердить, Esc — отказаться.
    ///
    /// Отдельный диалог, а не поле в общем: у пароля свои правила — он не показывается на
    /// экране и не попадает ни в какой журнал. Метрики (ширина окна и кнопок) берутся у
    /// <see cref="NumberPromptDialog"/>: оба диалога встречаются в одном окне и обязаны
    /// выглядеть одинаково, а считать это дважды — значит однажды разойтись.
    ///
    /// О последствиях отказа предупреждает САМ диалог, а не отдельное окно следом. При
    /// пакетном добавлении защищённых файлов подтверждение каждого отказа удвоило бы число
    /// модальных окон — человек, бросивший десяток файлов, закрывал бы двадцать. Сказать
    /// заранее в том же окне и честнее, и дешевле.
    /// </summary>
    internal sealed class PasswordPromptDialog : Form
    {
        private readonly TextBox _input;

        private PasswordPromptDialog(string title, string prompt)
        {
            Ui.InitDialog(this, title);
            const int margin = 16, btnH = 30, promptH = 52;

            var cancel = new RoundedButton(false);
            cancel.Text = Loc.T("common.cancel");
            cancel.DialogResult = DialogResult.Cancel;
            var ok = new RoundedButton(true);
            ok.Text = Loc.T("pdf.password.open");
            ok.DialogResult = DialogResult.OK;
            int btnW = NumberPromptDialog.ButtonWidth(Ui.TextWidth(cancel.Text, cancel.Font),
                Ui.TextWidth(ok.Text, ok.Font));

            // Ширина считается по САМОЙ ДЛИННОЙ строке подписи, а не по всему тексту: подпись
            // многострочная (имя файла плюс предупреждение), и мерить её целиком значило бы
            // растянуть окно на всю ширину экрана.
            int w = Math.Max(NumberPromptDialog.DialogWidth(WidestLine(prompt, Font), margin, 380),
                2 * btnW + 3 * margin);

            var label = new Label();
            label.AutoSize = false;
            label.SetBounds(margin, 16, w - 2 * margin, promptH);
            label.Text = prompt;
            Controls.Add(label);

            _input = new TextBox();
            _input.UseSystemPasswordChar = true;   // пароль не показывается на экране
            _input.SetBounds(margin, 16 + promptH + 8, w - 2 * margin, 27);
            Controls.Add(_input);

            int btnY = _input.Bottom + 16;
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
            _input.Focus();
        }

        /// <summary>Ширина самой длинной строки многострочной подписи. Чистая — под тест.</summary>
        internal static int WidestLine(string text, Font font)
        {
            if (string.IsNullOrEmpty(text) || font == null)
                return 0;
            int widest = 0;
            foreach (string line in text.Split('\n'))
                widest = Math.Max(widest, TextRenderer.MeasureText(line.TrimEnd('\r'), font).Width);
            return widest;
        }

        /// <summary>
        /// Спросить пароль к файлу. retry — этот пароль уже пробовали и он не подошёл.
        /// Возвращает введённое или null при отказе (пустая строка тоже считается отказом:
        /// пустой пароль защищённый файл не откроет).
        /// </summary>
        public static string Show(IWin32Window owner, string fileName, bool retry)
        {
            string prompt = string.Format(Loc.T(retry ? "pdf.password.again" : "pdf.password.prompt"), fileName);
            using (var dialog = new PasswordPromptDialog(Loc.T("pdf.password.title"), prompt))
            {
                if (dialog.ShowDialog(owner) != DialogResult.OK)
                    return null;
                string entered = dialog._input.Text;
                return string.IsNullOrEmpty(entered) ? null : entered;
            }
        }
    }
}
