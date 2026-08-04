using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// «Так это ляжет на бумагу» — предпросмотр перед печатью.
    ///
    /// Своё окно, а не штатный <c>PrintPreviewDialog</c>, по двум причинам, и обе весомые.
    /// Во-первых, его кнопка печати печатает САМА, тут же и на потоке интерфейса: пропали бы
    /// и полоса хода, и отмена, и фоновая работа — всё, что у печати уже есть. Во-вторых, он
    /// строит страницы ВСЕ СРАЗУ и держит их в памяти: сотня листов — это сотня растров.
    /// Здесь окно только показывает, а печатает по-прежнему общий путь с прогрессом и отменой.
    ///
    /// Показываются лишь первые несколько листов (<see cref="PreviewPages"/>). Предпросмотр
    /// отвечает на вопрос «как ляжет», а не заменяет чтение документа, и за этот ответ не
    /// стоит платить рендером всего задания.
    /// </summary>
    internal sealed class PrintPreviewForm : Form
    {
        /// <summary>Сколько листов рисуем. Хватает, чтобы увидеть раскладку и поля.</summary>
        internal const int PreviewPages = 8;

        private PrintPreviewForm(PrintDocument document, int shown, int total)
        {
            Ui.InitDialog(this, Loc.T("common.print.preview"));
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            MaximizeBox = true;
            MinimumSize = new Size(520, 480);
            ClientSize = new Size(760, 620);
            StartPosition = FormStartPosition.CenterParent;

            const int margin = 12, btnH = 30, btnW = 132;
            var print = new RoundedButton(true);
            print.Text = Loc.T("common.print.doPrint");
            print.DialogResult = DialogResult.OK;
            var cancel = new RoundedButton(false);
            cancel.Text = Loc.T("common.cancel");
            cancel.DialogResult = DialogResult.Cancel;

            var note = new Label();
            note.AutoSize = false;
            note.TextAlign = ContentAlignment.MiddleLeft;
            note.ForeColor = Theme.TextMuted;
            // Молчать о том, что показана только часть, нельзя: иначе неполный предпросмотр
            // читается как «в задании всего столько листов».
            note.Text = shown < total ? string.Format(Loc.T("common.print.previewPart"), shown, total) : "";

            var preview = new PrintPreviewControl();
            preview.Document = document;
            preview.UseAntiAlias = true;
            preview.Zoom = 0.4;
            preview.Columns = 2;
            preview.BackColor = Theme.TextMuted;

            // Раскладка вручную (как в остальных наших окнах): предпросмотр тянется, нижняя
            // полоса прижата к низу и не меняет высоту.
            preview.Bounds = new Rectangle(margin, margin,
                ClientSize.Width - 2 * margin, ClientSize.Height - btnH - 3 * margin);
            preview.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            note.Bounds = new Rectangle(margin, ClientSize.Height - btnH - margin,
                ClientSize.Width - 2 * btnW - 4 * margin, btnH);
            note.Anchor = AnchorStyles.Left | AnchorStyles.Bottom;
            cancel.Bounds = new Rectangle(ClientSize.Width - 2 * btnW - 2 * margin,
                ClientSize.Height - btnH - margin, btnW, btnH);
            cancel.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
            print.Bounds = new Rectangle(ClientSize.Width - btnW - margin,
                ClientSize.Height - btnH - margin, btnW, btnH);
            print.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;

            Controls.Add(preview);
            Controls.Add(note);
            Controls.Add(cancel);
            Controls.Add(print);
            AcceptButton = print;
            CancelButton = cancel;
        }

        /// <summary>
        /// Сколько листов показать: не больше <see cref="PreviewPages"/> и не больше, чем есть.
        /// Чистая — под тест.
        /// </summary>
        internal static int PagesToShow(int total)
        {
            if (total <= 0)
                return 0;
            return total < PreviewPages ? total : PreviewPages;
        }

        /// <summary>Показать предпросмотр; true — человек подтвердил печать.</summary>
        public static bool Confirm(IWin32Window owner, PrintDocument document, int shown, int total)
        {
            using (var form = new PrintPreviewForm(document, shown, total))
                return form.ShowDialog(owner) == DialogResult.OK;
        }
    }
}
