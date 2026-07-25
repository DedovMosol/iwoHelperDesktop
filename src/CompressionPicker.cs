using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Общий выбор уровня сжатия PDF для обоих инструментов (объединение/разделение).
    /// Подпись + выпадающий список «Отлично/Хорошо/Нормально» (дефолт — без сжатия).
    /// Если Ghostscript не найден и выбирают сжатие — брендированный диалог со ссылкой
    /// на загрузку и возврат к «Отлично». Значение читается через <see cref="Level"/>.
    /// </summary>
    public sealed class CompressionPicker : UserControl
    {
        private readonly ComboBox _combo;
        private readonly ToolTip _tips;
        private bool _reverting; // защита от реентранси SelectedIndexChanged

        private const int CaptionGap = 8;  // подпись → список
        private const int ComboExtra = 40; // кнопка списка + внутренние отступы

        /// <summary>
        /// Ширина контрола (подпись + список под самый длинный пункт). Статична, чтобы экраны БЕЗ
        /// сжатия (PDF → Word) резервировали такое же место и регулятор масштаба был одинакового
        /// размера на всех PDF-инструментах (единая раскладка).
        /// </summary>
        public static int PreferredWidth()
        {
            var font = Ui.Font(9.75f);
            int widest = 0;
            foreach (string lbl in PdfCompression.LevelLabels())
                widest = Math.Max(widest, TextRenderer.MeasureText(lbl, font).Width);
            return TextRenderer.MeasureText(Loc.T("common.compression"), font).Width + CaptionGap + widest + ComboExtra;
        }

        public CompressionPicker()
        {
            // Без этого стиля UserControl игнорирует Transparent и красит себя серым
            // (SystemColors.Control) — серый прямоугольник на белой форме.
            SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            Font = Ui.Font(9.75f); // общий кэшированный шрифт (не освобождать)

            var caption = new Label();
            caption.Text = Loc.T("common.compression");
            caption.AutoSize = true;
            caption.ForeColor = Theme.TextPrimary;
            caption.BackColor = Color.Transparent;
            caption.Location = new Point(0, 5);
            Controls.Add(caption);

            _combo = new ComboBox();
            _combo.DropDownStyle = ComboBoxStyle.DropDownList;
            _combo.Items.AddRange(PdfCompression.LevelLabels());
            _combo.SelectedIndex = (int)CompressionLevel.None; // «Отлично — без сжатия»
            // Ширина контрола — PreferredWidth (под самый длинный пункт), ту же резервируют экраны
            // без сжатия, чтобы регулятор масштаба был одинакового размера везде.
            int total = PreferredWidth();
            _combo.SetBounds(caption.Right + CaptionGap, 1, total - caption.Right - CaptionGap, 27);
            _combo.SelectedIndexChanged += OnSelectedIndexChanged;
            Controls.Add(_combo);

            _tips = new ToolTip();
            _tips.SetToolTip(_combo, Loc.T("common.tip.compression"));

            Size = new Size(total, 29);
        }

        /// <summary>Выбранный уровень сжатия.</summary>
        public CompressionLevel Level
        {
            get
            {
                int i = _combo.SelectedIndex;
                return i < 0 ? CompressionLevel.None : (CompressionLevel)i;
            }
        }

        /// <summary>
        /// Восстановить сохранённый уровень (без показа диалога «нужен Ghostscript» — это не
        /// действие пользователя): вне диапазона ИЛИ сжатие при отсутствии Ghostscript тихо
        /// откатывается к «Отлично».
        /// </summary>
        public void SetLevel(CompressionLevel level)
        {
            int i = (int)level;
            if (i <= 0 || i >= _combo.Items.Count || !Ghostscript.Available)
                i = (int)CompressionLevel.None;
            bool prev = _reverting;
            _reverting = true; // подавить обработчик выбора (иначе всплыл бы диалог при старте)
            try { _combo.SelectedIndex = i; }
            finally { _reverting = prev; }
        }

        protected override void Dispose(bool disposing)
        {
            // ToolTip — компонент, а не дочерний контрол: авто-освобождение не сработает.
            if (disposing && _tips != null)
                _tips.Dispose();
            base.Dispose(disposing);
        }

        private void OnSelectedIndexChanged(object sender, EventArgs e)
        {
            if (_reverting)
                return;
            if (_combo.SelectedIndex > 0 && !Ghostscript.Available)
            {
                Dialogs.InfoWithLink(FindForm(),
                    Loc.T("gs.title"),
                    Loc.T("gs.heading"),
                    Loc.T("gs.body"),
                    Loc.T("gs.download"),
                    Ghostscript.DownloadPage);
                _reverting = true;
                try { _combo.SelectedIndex = (int)CompressionLevel.None; }
                finally { _reverting = false; }
            }
        }
    }
}
