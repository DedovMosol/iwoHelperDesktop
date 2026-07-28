using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Окно «Настройки» — место для того, что относится ко всей программе, а не к открытому
    /// документу. Настройки текущей операции (масштаб сетки, уровень сжатия) сюда НЕ
    /// переезжают: они меняются по ходу работы, и уводить их в отдельное окно значило бы
    /// заставлять ходить туда-обратно на каждый файл.
    ///
    /// Языка здесь нет намеренно. Он уже переключается глобусом на стартовом экране и в «☰
    /// Меню» каждого инструмента, а третья точка входа — это не удобство, а третье место,
    /// которое придётся не забыть при правке. Вдобавок смена языка ПЕРЕСОБИРАЕТ открытые
    /// окна, и запуск пересборки из модального окна поверх пересобираемого владельца — уже
    /// наступавшая ловушка.
    ///
    /// Раскладка считается СВЕРХУ ВНИЗ от реальных размеров контролов, а не литералами.
    /// Пояснение под флажком переносится по словам, и его высота зависит от шрифта и
    /// масштаба экрана: на 150% посчитанная на глаз координата кнопки оказалась бы под
    /// текстом. Ровно так уже разъезжались «Свойства документа» до 1.17.9.
    /// </summary>
    public class SettingsForm : Form
    {
        private const int Pad = 24;        // поле слева и справа
        private const int BtnH = 34;
        private const int BtnMinW = 190;
        private const int BtnTextPad = 28; // запас вокруг подписи внутри кнопки
        private const int Gap = 10;        // между соседними строками одной группы
        private const int GroupGap = 22;   // между группами

        /// <summary>Имя кнопки «снова напоминать» — по нему её находит проверка.</summary>
        internal const string UnskipName = "unskip";

        private readonly AccentCheckBox _checkOnStart;
        private readonly RoundedButton _unskip;
        private ToolTip _tips;
        private bool _loading;

        public SettingsForm()
        {
            Ui.InitDialog(this, Loc.T("settings.title"));
            int width = 460;
            ClientSize = new Size(width, 100); // высота считается ниже, по содержимому
            WindowChrome.Enable(this, Theme.HubBlue);

            Ui.AccentBar(this, 0, Theme.HubBlue);
            Label title = Ui.Label(this, Loc.T("settings.title"), Pad, 22,
                Ui.Font(14f, FontStyle.Bold), Color.FromArgb(40, 40, 40));

            // ---------- обновления ----------
            int y = title.Bottom + GroupGap;
            y = Section(Loc.T("settings.section.updates"), y).Bottom + Gap;

            _checkOnStart = new AccentCheckBox();
            _checkOnStart.Text = Loc.T("settings.chk.updateOnStart");
            Controls.Add(_checkOnStart); // шрифт наследуется от формы: мерить только после этого
            Size want = _checkOnStart.GetPreferredSize(Size.Empty);
            _checkOnStart.SetBounds(Pad, y, Math.Min(want.Width, width - 2 * Pad), want.Height);
            _checkOnStart.CheckedChanged += OnCheckOnStartChanged;

            Label hint = Ui.Label(this, Loc.T("settings.hint.updates"), Pad, _checkOnStart.Bottom + 6,
                Font, Theme.TextMuted);
            hint.MaximumSize = new Size(width - 2 * Pad, 0); // перенос по словам в пределах полосы
            hint.AutoSize = true;

            RoundedButton checkNow = AddButton(Loc.T("settings.btn.checkNow"), Pad, hint.Bottom + Gap + 4);
            checkNow.Click += delegate
            {
                // Кнопка гасится на время запроса: сеть с таймаутом 10 с, а нетерпеливые
                // нажатия плодили бы по воркеру и по окну с ответом на каждое.
                checkNow.Enabled = false;
                UpdateUi.Check(this, delegate { checkNow.Enabled = true; });
            };

            // «Снова напоминать» — СВОЕЙ строкой, а не рядом: подпись несёт номер версии и на
            // разных языках разной длины, и в паре кнопок она рано или поздно не поместилась
            // бы по ширине. Место под неё держим всегда, даже когда она спрятана, иначе окно
            // меняло бы высоту от того, отказывался ли человек от напоминаний.
            _unskip = AddButton("", Pad, checkNow.Bottom + Gap);
            _unskip.Name = UnskipName; // чтобы проверка находила кнопку по имени, а не по пустой подписи
            _unskip.Click += OnUnskip;

            // ---------- статистика ----------
            y = Section(Loc.T("settings.section.stats"), _unskip.Bottom + GroupGap).Bottom + Gap;
            RoundedButton stats = AddButton(Loc.T("settings.btn.stats"), Pad, y);
            stats.Click += delegate { using (var f = new StatsForm()) f.ShowDialog(this); };

            var close = new RoundedButton(true);
            close.SetBounds(width - Pad - 100, stats.Bottom + GroupGap, 100, 36);
            Controls.Add(close);
            close.Text = Loc.T("common.close");
            close.Click += delegate { Close(); };
            AcceptButton = close;
            CancelButton = close;

            ClientSize = new Size(width, close.Bottom + Pad);

            _tips = new ToolTip();
            _tips.SetToolTip(_checkOnStart, Loc.T("settings.tip.updateOnStart"));

            LoadAndShow();
        }

        /// <summary>Заголовок группы настроек.</summary>
        private Label Section(string text, int y)
        {
            return Ui.Label(this, text, Pad, y, Ui.Font(10.5f, FontStyle.Bold), Theme.TextPrimary);
        }

        /// <summary>
        /// Кнопка, ширина которой посчитана по её СОБСТВЕННОЙ подписи. Литеральная ширина
        /// обрезает перевод многоточием: «Remind me about 1.18.0 again» длиннее русского
        /// «Снова напоминать о 1.18.0», и та ширина, что подошла на одном языке, на другом
        /// съедает конец подписи молча.
        /// </summary>
        private RoundedButton AddButton(string text, int x, int y)
        {
            var b = new RoundedButton(false);
            b.SetBounds(x, y, BtnMinW, BtnH);
            Controls.Add(b); // шрифт наследуется от формы, до добавления мерить нечем
            SetButtonText(b, text);
            return b;
        }

        /// <summary>Задать подпись и подогнать ширину под неё (подпись меняется в LoadAndShow).</summary>
        private static void SetButtonText(RoundedButton button, string text)
        {
            button.Text = text;
            int want = TextRenderer.MeasureText(text ?? "", button.Font).Width + BtnTextPad;
            button.Width = Math.Max(BtnMinW, want);
        }

        /// <summary>
        /// Показать состояние с диска. Настройки читаются здесь, а не запоминаются полем:
        /// окно живёт долго, а «не напоминать» могли нажать в окне обновления, открытом
        /// поверх этого.
        /// </summary>
        private void LoadAndShow()
        {
            UserSettings s = UserSettings.Load();
            _loading = true; // не считать программную установку флажка за выбор человека
            _checkOnStart.Checked = s.UpdateCheckOnStart;
            _loading = false;

            bool skipped = !string.IsNullOrEmpty(s.SkippedVersion);
            _unskip.Visible = skipped;
            if (skipped)
                SetButtonText(_unskip, string.Format(Loc.T("settings.btn.unskip"), s.SkippedVersion));
        }

        private void OnCheckOnStartChanged(object sender, EventArgs e)
        {
            if (_loading)
                return;
            UserSettings.SaveUpdateCheckOnStart(_checkOnStart.Checked);
        }

        private void OnUnskip(object sender, EventArgs e)
        {
            UserSettings.SaveSkippedVersion(null);
            LoadAndShow(); // кнопка исчезает: напоминать больше нечего отменять
        }

        protected override void Dispose(bool disposing)
        {
            // ToolTip — компонент, а не дочерний контрол: авто-освобождение не срабатывает.
            if (disposing && _tips != null)
                _tips.Dispose();
            base.Dispose(disposing);
        }
    }
}
