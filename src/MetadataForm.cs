using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Свойства документа, которые видит любой просмотрщик PDF в «Свойствах файла».</summary>
    public sealed class PdfMetadata
    {
        public string Title = "";
        public string Author = "";
        public string Subject = "";
        public string Keywords = "";
    }

    /// <summary>
    /// Правка свойств документа: заголовок, автор, тема, ключевые слова. Только сбор ввода —
    /// записывает их <see cref="PdfMetadataService"/>, форма о PDF ничего не знает.
    /// Каркас окна общий (<see cref="Ui.InitDialog"/>).
    /// </summary>
    internal sealed class MetadataForm : Form
    {
        private readonly TextBox _title, _author, _subject, _keywords;

        private MetadataForm(PdfMetadata current)
        {
            Ui.InitDialog(this, Loc.T("meta.title"));
            ClientSize = new Size(460, 250);

            _title = AddField(Loc.T("meta.field.title"), 20, current.Title);
            _author = AddField(Loc.T("meta.field.author"), 66, current.Author);
            _subject = AddField(Loc.T("meta.field.subject"), 112, current.Subject);
            _keywords = AddField(Loc.T("meta.field.keywords"), 158, current.Keywords);

            var ok = new RoundedButton(true);
            ok.Text = Loc.T("common.ok");
            ok.SetBounds(ClientSize.Width - 232, 200, 100, 32);
            ok.DialogResult = DialogResult.OK;
            Controls.Add(ok);

            var cancel = new RoundedButton(false);
            cancel.Text = Loc.T("common.cancel");
            cancel.SetBounds(ClientSize.Width - 124, 200, 100, 32);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
        }

        private TextBox AddField(string caption, int y, string value)
        {
            Ui.Label(this, caption, 20, y, Font, Theme.TextMuted);
            var box = new TextBox();
            box.SetBounds(20, y + 18, ClientSize.Width - 44, 24);
            box.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            box.Text = value ?? "";
            Controls.Add(box);
            return box;
        }

        /// <summary>
        /// Показать окно правки. Возвращает изменённые свойства или null, если пользователь
        /// отказался. Значения обрезаются по краям: пробел в конце автора — не свойство файла.
        /// </summary>
        public static PdfMetadata Edit(IWin32Window owner, PdfMetadata current)
        {
            using (var f = new MetadataForm(current ?? new PdfMetadata()))
            {
                if (f.ShowDialog(owner) != DialogResult.OK)
                    return null;
                return new PdfMetadata
                {
                    Title = f._title.Text.Trim(),
                    Author = f._author.Text.Trim(),
                    Subject = f._subject.Text.Trim(),
                    Keywords = f._keywords.Text.Trim()
                };
            }
        }
    }
}
