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

        public CompressionPicker()
        {
            // Без этого стиля UserControl игнорирует Transparent и красит себя серым
            // (SystemColors.Control) — серый прямоугольник на белой форме.
            SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            Font = new Font("Segoe UI", 9.75f);

            var caption = new Label();
            caption.Text = "Сжатие:";
            caption.AutoSize = true;
            caption.ForeColor = Theme.TextPrimary;
            caption.BackColor = Color.Transparent;
            caption.Location = new Point(0, 5);
            Controls.Add(caption);

            _combo = new ComboBox();
            _combo.DropDownStyle = ComboBoxStyle.DropDownList;
            _combo.Items.AddRange(PdfCompression.LevelLabels);
            _combo.SelectedIndex = (int)CompressionLevel.None; // «Отлично — без сжатия»
            // Ширина — под самый длинный пункт (иначе «Нормально — минимальный размер» обрезается).
            int widest = 0;
            foreach (string lbl in PdfCompression.LevelLabels)
                widest = Math.Max(widest, TextRenderer.MeasureText(lbl, Font).Width);
            _combo.SetBounds(caption.Right + 8, 1, widest + 40, 27); // +кнопка списка и отступы
            _combo.SelectedIndexChanged += OnSelectedIndexChanged;
            Controls.Add(_combo);

            _tips = new ToolTip();
            _tips.SetToolTip(_combo,
                "«Хорошо»/«Нормально» уменьшают размер, снижая разрешение изображений (как в Acrobat).\n" +
                "Текст и вектор сохраняются. У подписанных PDF подпись станет недействительной.");

            Size = new Size(_combo.Right, 29);
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
                    "Сжатие недоступно",
                    "Нужен Ghostscript",
                    "Сжатие PDF использует Ghostscript — он не найден в системе. " +
                    "Установите его (бесплатно), затем перезапустите приложение — либо " +
                    "используйте установщик приложения, в него Ghostscript уже входит.",
                    "Скачать Ghostscript",
                    Ghostscript.DownloadPage);
                _reverting = true;
                try { _combo.SelectedIndex = (int)CompressionLevel.None; }
                finally { _reverting = false; }
            }
        }
    }
}
