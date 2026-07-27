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
    ///
    /// Раскладка считается по РЕАЛЬНЫМ высотам созданных контролов, а не по литералам:
    /// высоту однострочного TextBox WinForms назначает сама по шрифту (заданные 24 превращались
    /// в 25), и посчитанная по литералам сетка разъезжалась — кнопки наезжали на рамку
    /// последнего поля на 1 px, а между группами оставалось 3 px вместо просвета.
    /// </summary>
    internal sealed class MetadataForm : Form
    {
        private const int Pad = 24;       // поля окна — как в «О программе» (имя Margin занято Control.Margin)
        private const int LabelGap = 4;   // между подписью и её полем
        private const int GroupGap = 14;  // между соседними полями
        private const int ButtonW = 100, ButtonH = 32;

        private readonly TextBox _title, _author, _subject, _keywords;

        private MetadataForm(PdfMetadata current, string fileName)
        {
            Ui.InitDialog(this, Loc.T("meta.title"));
            ClientSize = new Size(460, 200); // высота уточняется в конце под контент

            int y = Pad;
            if (!string.IsNullOrEmpty(fileName))
            {
                // Чьи свойства правим. Имя файла бывает длинным, поэтому с многоточием и
                // подсказкой на весь текст, а не «уезжает за край».
                Label file = Ui.Label(this, fileName, Pad, y, Font, Theme.TextMuted);
                Ui.Ellipsize(file, ClientSize.Width - 2 * Pad, AnchorStyles.Top | AnchorStyles.Left);
                y = file.Bottom + GroupGap;
            }

            _title = AddField(Loc.T("meta.field.title"), ref y, current.Title);
            _author = AddField(Loc.T("meta.field.author"), ref y, current.Author);
            _subject = AddField(Loc.T("meta.field.subject"), ref y, current.Subject);
            _keywords = AddField(Loc.T("meta.field.keywords"), ref y, current.Keywords);
            y -= GroupGap; // просвет после последнего поля не нужен — дальше своя дистанция

            // Пустое поле стирает свойство — это осознанный способ убрать имя автора перед
            // отправкой файла, и об этом надо сказать прямо: иначе выглядит как «не заполнил».
            Label hint = Ui.Label(this, Loc.T("meta.hint.empty"), Pad, y + 16, Font, Theme.TextMuted);

            int buttonsTop = hint.Bottom + 16;
            ClientSize = new Size(ClientSize.Width, buttonsTop + ButtonH + Pad);

            var ok = new RoundedButton(true);
            ok.Text = Loc.T("common.ok");
            ok.SetBounds(ClientSize.Width - Pad - 2 * ButtonW - 8, buttonsTop, ButtonW, ButtonH);
            ok.DialogResult = DialogResult.OK;
            Controls.Add(ok);

            var cancel = new RoundedButton(false);
            cancel.Text = Loc.T("common.cancel");
            cancel.SetBounds(ClientSize.Width - Pad - ButtonW, buttonsTop, ButtonW, ButtonH);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
        }

        /// <summary>
        /// Подпись и поле под ней, курсор y сдвигается на следующую группу. Высоту поля читаем
        /// у него самого: WinForms ставит её по шрифту и заданную попросту игнорирует.
        /// </summary>
        private TextBox AddField(string caption, ref int y, string value)
        {
            Label label = Ui.Label(this, caption, Pad, y, Font, Theme.TextMuted);
            var box = new TextBox();
            box.SetBounds(Pad, label.Bottom + LabelGap, ClientSize.Width - 2 * Pad, 24);
            box.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            box.Text = value ?? "";
            Controls.Add(box);
            y = box.Bottom + GroupGap;
            return box;
        }

        /// <summary>
        /// Показать окно правки. Возвращает изменённые свойства или null, если пользователь
        /// отказался. Значения обрезаются по краям: пробел в конце автора — не свойство файла.
        /// fileName — чьи свойства правим (показывается сверху), пустое имя строку убирает.
        /// </summary>
        public static PdfMetadata Edit(IWin32Window owner, PdfMetadata current, string fileName)
        {
            using (var f = new MetadataForm(current ?? new PdfMetadata(), fileName))
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
