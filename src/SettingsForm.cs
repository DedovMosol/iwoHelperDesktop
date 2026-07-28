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
    /// </summary>
    public class SettingsForm : Form
    {
        private readonly AccentCheckBox _checkOnStart;
        private readonly RoundedButton _unskip;
        private ToolTip _tips;
        private bool _loading;

        public SettingsForm()
        {
            Ui.InitDialog(this, Loc.T("settings.title"));
            ClientSize = new Size(460, 336);
            WindowChrome.Enable(this, Theme.HubBlue);

            Ui.AccentBar(this, 0, Theme.HubBlue);
            Ui.Label(this, Loc.T("settings.title"), 24, 22, Ui.Font(14f, FontStyle.Bold), Color.FromArgb(40, 40, 40));

            // ---------- обновления ----------
            Ui.Label(this, Loc.T("settings.section.updates"), 24, 70, Ui.Font(10.5f, FontStyle.Bold), Theme.TextPrimary);

            _checkOnStart = new AccentCheckBox();
            _checkOnStart.Text = Loc.T("settings.chk.updateOnStart");
            Size want = _checkOnStart.GetPreferredSize(Size.Empty);
            _checkOnStart.SetBounds(24, 98, Math.Min(want.Width, ClientSize.Width - 48), want.Height);
            _checkOnStart.CheckedChanged += OnCheckOnStartChanged;
            Controls.Add(_checkOnStart);

            Label hint = Ui.Label(this, Loc.T("settings.hint.updates"), 24, 126, Font, Theme.TextMuted);
            hint.MaximumSize = new Size(ClientSize.Width - 48, 0);
            hint.AutoSize = true;

            var checkNow = new RoundedButton(false);
            checkNow.Text = Loc.T("settings.btn.checkNow");
            checkNow.SetBounds(24, 178, 190, 34);
            checkNow.Click += delegate
            {
                // Кнопка гасится на время запроса: сеть с таймаутом 10 с, а нетерпеливые
                // нажатия плодили бы по воркеру и по окну с ответом на каждое.
                checkNow.Enabled = false;
                UpdateUi.Check(this, delegate { checkNow.Enabled = true; });
            };
            Controls.Add(checkNow);

            // «Снова напоминать» появляется, только если о какой-то версии просили молчать:
            // кнопка, отменяющая то, чего не было, — это вопрос без ответа.
            _unskip = new RoundedButton(false);
            _unskip.SetBounds(226, 178, 210, 34);
            _unskip.Click += OnUnskip;
            Controls.Add(_unskip);

            // ---------- статистика ----------
            Ui.Label(this, Loc.T("settings.section.stats"), 24, 234, Ui.Font(10.5f, FontStyle.Bold), Theme.TextPrimary);

            var stats = new RoundedButton(false);
            stats.Text = Loc.T("settings.btn.stats");
            stats.SetBounds(24, 262, 190, 34);
            stats.Click += delegate { using (var f = new StatsForm()) f.ShowDialog(this); };
            Controls.Add(stats);

            var close = new RoundedButton(true);
            close.Text = Loc.T("common.close");
            close.SetBounds(ClientSize.Width - 124, ClientSize.Height - 52, 100, 36);
            close.Click += delegate { Close(); };
            Controls.Add(close);
            AcceptButton = close;
            CancelButton = close;

            _tips = new ToolTip();
            _tips.SetToolTip(_checkOnStart, Loc.T("settings.tip.updateOnStart"));

            LoadAndShow();
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
                _unskip.Text = string.Format(Loc.T("settings.btn.unskip"), s.SkippedVersion);
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
