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

        private static readonly int[] AutoDays = { 0, 1, 7, 30 }; // индекс списка -> дни (как в статистике)

        private readonly AccentCheckBox _checkOnStart;
        private readonly AccentCheckBox _showWhatsNew;
        private readonly RoundedButton _unskip;
        private readonly AccentCheckBox _keepHistory;
        private readonly ComboBox _historyAuto;
        private readonly Label _historyCount;
        private ToolTip _tips;
        private bool _loading;

        public SettingsForm()
        {
            Ui.InitDialog(this, Loc.T("settings.title"));
            int width = 460;
            ClientSize = new Size(width, 100); // высота считается ниже, по содержимому
            WindowChrome.Enable(this, Theme.HubBlue);

            Ui.AccentBar(this, 0, Theme.HubBlue);
            // Заголовок по центру: окно узкое и без панели инструментов, и подпись у левого
            // края выглядела приклеенной к рамке.
            Label title = Ui.Label(this, Loc.T("settings.title"), Pad, 22,
                Ui.Font(14f, FontStyle.Bold), Color.FromArgb(40, 40, 40));
            title.AutoSize = false;
            title.TextAlign = ContentAlignment.MiddleCenter;
            title.SetBounds(Pad, 22, width - 2 * Pad, title.PreferredHeight);

            // ---------- обновления ----------
            int y = title.Bottom + GroupGap;
            y = Section(Loc.T("settings.section.updates"), y).Bottom + Gap;

            _checkOnStart = new AccentCheckBox();
            _checkOnStart.Text = Loc.T("settings.chk.updateOnStart");
            Controls.Add(_checkOnStart); // шрифт наследуется от формы: мерить только после этого
            Size want = _checkOnStart.GetPreferredSize(Size.Empty);
            _checkOnStart.SetBounds(Pad, y, Math.Min(want.Width, width - 2 * Pad), want.Height);
            _checkOnStart.CheckedChanged += OnCheckOnStartChanged;

            // Пояснение — выключкой ПО ШИРИНЕ: рваный правый край в узком окне бросается в
            // глаза сильнее, чем в тексте на всю страницу (см. JustifiedText).
            RichTextBox hint = JustifiedText.Paragraph(this, Loc.T("settings.hint.updates"),
                Pad, _checkOnStart.Bottom + 6, width - 2 * Pad, Theme.TextMuted);

            _showWhatsNew = new AccentCheckBox
            {
                Text = Loc.T("settings.chk.whatsNew")
            };
            Controls.Add(_showWhatsNew);
            Size whatsNewSize = _showWhatsNew.GetPreferredSize(Size.Empty);
            _showWhatsNew.SetBounds(Pad, hint.Bottom + 8,
                Math.Min(whatsNewSize.Width, width - 2 * Pad), whatsNewSize.Height);
            _showWhatsNew.CheckedChanged += OnShowWhatsNewChanged;

            RoundedButton[] updateActions = RowButtons(new[]
                {
                    Loc.T("settings.btn.checkNow"),
                    string.Format(Loc.T("settings.btn.whatsNew"),
                        WhatsNewCatalog.CurrentVersion)
                }, Pad, _showWhatsNew.Bottom + Gap + 2, width - 2 * Pad);
            RoundedButton checkNow = updateActions[0];
            checkNow.Click += delegate
            {
                checkNow.Enabled = false;
                UpdateUi.Check(this, delegate { checkNow.Enabled = true; });
            };
            updateActions[1].Click += delegate { WhatsNewUi.ShowDialog(this); };

            // «Снова напоминать» — СВОЕЙ строкой: подпись несёт номер версии.
            _unskip = AddButton("", Pad, checkNow.Bottom + Gap);
            _unskip.Name = UnskipName; // чтобы проверка находила кнопку по имени, а не по пустой подписи
            _unskip.Click += OnUnskip;

            // ---------- история и статистика ----------
            y = Section(Loc.T("settings.section.history"), _unskip.Bottom + GroupGap).Bottom + Gap;

            _keepHistory = new AccentCheckBox();
            _keepHistory.Text = Loc.T("settings.chk.history");
            Controls.Add(_keepHistory);
            Size hw = _keepHistory.GetPreferredSize(Size.Empty);
            _keepHistory.SetBounds(Pad, y, Math.Min(hw.Width, width - 2 * Pad), hw.Height);
            _keepHistory.CheckedChanged += OnKeepHistoryChanged;

            Ui.Label(this, Loc.T("settings.lbl.historyAuto"), Pad, _keepHistory.Bottom + 14, Font, Theme.TextPrimary);
            _historyAuto = new ComboBox();
            _historyAuto.DropDownStyle = ComboBoxStyle.DropDownList;
            _historyAuto.Items.AddRange(new object[]
            {
                Loc.T("stats.auto.off"), Loc.T("stats.auto.daily"),
                Loc.T("stats.auto.7days"), Loc.T("stats.auto.30days")
            });
            Controls.Add(_historyAuto);
            _historyAuto.SetBounds(width - Pad - 180, _keepHistory.Bottom + 10, 180, 27);
            _historyAuto.SelectedIndexChanged += OnHistoryAutoChanged;

            _historyCount = Ui.Label(this, "", Pad, _historyAuto.Bottom + 14, Font, Theme.TextMuted);

            // Три действия истории — ОДНИМ рядом в пределах той же полосы, что и текст выше:
            // столбиком они занимали треть окна, а делают вещи одного порядка. Ширина делится
            // поровну, подписи ужимаются шрифтом до тех пор, пока не поместятся (см. RowButtons).
            RoundedButton[] row = RowButtons(new[]
                {
                    Loc.T("settings.btn.historyShow"),
                    Loc.T("settings.btn.historyClear"),
                    Loc.T("settings.btn.stats")
                }, Pad, _historyCount.Bottom + Gap, width - 2 * Pad);
            row[0].Click += delegate { using (var f = new HistoryForm()) f.ShowDialog(this); };
            row[1].Click += OnClearHistory;
            row[2].Click += delegate { using (var f = new StatsForm()) f.ShowDialog(this); };

            var close = new RoundedButton(true);
            close.SetBounds(width - Pad - 100, row[0].Bottom + GroupGap, 100, 36);
            Controls.Add(close);
            close.Text = Loc.T("common.close");
            close.Click += delegate { Close(); };
            AcceptButton = close;
            CancelButton = close;

            ClientSize = new Size(width, close.Bottom + Pad);

            _tips = new ToolTip();
            _tips.SetToolTip(_checkOnStart, Loc.T("settings.tip.updateOnStart"));
            _tips.SetToolTip(_showWhatsNew, Loc.T("settings.chk.whatsNew"));
            _tips.SetToolTip(_keepHistory, Loc.T("settings.tip.history"));

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

        /// <summary>
        /// Ряд кнопок равной ширины в пределах полосы: делим место поровну, а подпись, которая
        /// не влезла, уменьшаем шрифтом — но не бесконечно, а до разумного предела. Обрезать
        /// подпись многоточием нельзя: «Показать историю» и «Очистить историю» тогда выглядят
        /// одинаково, а это разные и невосстановимые действия.
        /// </summary>
        private RoundedButton[] RowButtons(string[] captions, int x, int y, int totalWidth)
        {
            const int RowGap = 8;
            int each = (totalWidth - RowGap * (captions.Length - 1)) / captions.Length;
            var result = new RoundedButton[captions.Length];
            for (int i = 0; i < captions.Length; i++)
            {
                var b = new RoundedButton(false);
                b.SetBounds(x + i * (each + RowGap), y, each, BtnH);
                Controls.Add(b); // шрифт наследуется от формы, до добавления мерить нечем
                b.Text = captions[i];
                // Ширину под подпись считаем ТЕМ ЖЕ отступом, каким кнопка рисует текст,
                // иначе подпись «влезает» по расчёту и обрезается многоточием на экране.
                b.Font = FitFont(captions[i], b.Font, each - 2 * RoundedButton.TextPadFor(each));
                result[i] = b;
            }
            return result;
        }

        /// <summary>
        /// Шрифт, при котором подпись помещается в заданную ширину. Уменьшаем по половине
        /// пункта и не ниже нижнего предела — дальше подпись становится нечитаемой, и лучше
        /// уж пусть она чуть выйдет за край, чем превратится в мелкий шрифт.
        /// </summary>
        private static Font FitFont(string text, Font font, int maxWidth)
        {
            const float MinSize = 8f;
            float size = font.Size;
            while (size > MinSize && TextRenderer.MeasureText(text, Ui.Font(size, font.Style)).Width > maxWidth)
                size -= 0.5f;
            return Ui.Font(size, font.Style);
        }

        /// <summary>Задать подпись и подогнать ширину под неё (подпись меняется в LoadAndShow).</summary>
        private static void SetButtonText(RoundedButton button, string text)
        {
            button.Text = text;
            int want = TextRenderer.MeasureText(text ?? "", button.Font).Width + BtnTextPad;
            button.Width = Math.Max(BtnMinW, want);
        }

        /// <summary>
        /// Возврат фокуса — перечитать настройки. Окно модально только для СВОЕГО владельца
        /// (<c>ShowDialog(owner)</c>), а инструменты живут независимыми окнами: вторые
        /// «Настройки» из «☰ Меню» инструмента открываются поверх первых, и каждое показывало
        /// бы свой снимок. Поверх настроек может встать и окно обновления с флажком «больше не
        /// напоминать». Без перечитывания видимое состояние расходится с диском, а следующее
        /// переключение флажка вернуло бы на диск устаревшее значение.
        /// </summary>
        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);
            LoadAndShow();
        }

        /// <summary>Показать состояние с диска.</summary>
        private void LoadAndShow()
        {
            UserSettings s = UserSettings.Load();
            _loading = true; // не считать программную установку флажка за выбор человека
            _checkOnStart.Checked = s.UpdateCheckOnStart;
            _showWhatsNew.Checked = s.ShowWhatsNewOnStart;
            _loading = false;

            bool skipped = !string.IsNullOrEmpty(s.SkippedVersion);
            _unskip.Visible = skipped;
            if (skipped)
                SetButtonText(_unskip, string.Format(Loc.T("settings.btn.unskip"), s.SkippedVersion));

            OperationHistory.Data h = OperationHistory.Load();
            _loading = true;
            _keepHistory.Checked = h.Enabled;
            int index = Array.IndexOf(AutoDays, h.AutoClearDays);
            _historyAuto.SelectedIndex = index >= 0 ? index : 0;
            _loading = false;
            _historyAuto.Enabled = h.Enabled; // выключенной истории нечего чистить по расписанию
            _historyCount.Text = h.Entries.Count == 0
                ? Loc.T("settings.history.empty")
                : string.Format(Loc.T("settings.history.count"), h.Entries.Count);
        }

        /// <summary>
        /// Подтверждение стирания истории. Отдельным полем, чтобы проверка могла ответить за
        /// человека: ветка «согласился» без этого недостижима автоматически, а именно в ней
        /// и жила ошибка.
        /// </summary>
        internal Func<bool> ConfirmClearHistory { get; set; }

        private void OnShowWhatsNewChanged(object sender, EventArgs e)
        {
            if (_loading)
                return;
            UserSettings current = UserSettings.Load();
            if (!UserSettings.SaveWhatsNew(_showWhatsNew.Checked,
                current.LastWhatsNewVersion))
            {
                ShowSaveError();
                LoadAndShow();
            }
        }

        private void OnKeepHistoryChanged(object sender, EventArgs e)
        {
            if (_loading)
                return;
            // Намерение снимаем СРАЗУ, до всякого диалога. Пока модальное окно на экране, наше
            // теряет и возвращает фокус, а на возврате OnActivated перечитывает настройки и
            // ставит флажок обратно по диску. Прочитанное ПОСЛЕ диалога значение оказывалось
            // прежним — подтверждённое выключение молча не срабатывало.
            bool wanted = _keepHistory.Checked;

            // Выключение стирает накопленное — так решено в хранилище. Спрашиваем прямо, а не
            // молча: человек мог не ожидать, что снятие флажка удалит уже собранный список.
            if (!wanted && OperationHistory.Load().Entries.Count > 0 && !AskClearHistory())
            {
                _loading = true;
                _keepHistory.Checked = true; // отказался — возвращаем флажок, ничего не трогаем
                _loading = false;
                return;
            }
            if (!OperationHistory.SetEnabled(wanted))
                ShowSaveError();
            LoadAndShow();
        }

        private bool AskClearHistory()
        {
            if (ConfirmClearHistory != null)
                return ConfirmClearHistory();
            return Dialogs.ConfirmWarning(this, Loc.T("settings.title"),
                Loc.T("settings.confirm.clearHistory.title"), Loc.T("settings.confirm.clearHistory.body"));
        }

        private void OnHistoryAutoChanged(object sender, EventArgs e)
        {
            if (_loading)
                return;
            if (!OperationHistory.SetAutoClear(AutoDays[_historyAuto.SelectedIndex]))
                ShowSaveError();
            LoadAndShow(); // применённая давность могла убрать часть записей
        }

        private void OnClearHistory(object sender, EventArgs e)
        {
            if (AskClearHistory())
            {
                if (!OperationHistory.Clear())
                    ShowSaveError();
                LoadAndShow();
            }
        }

        private void OnCheckOnStartChanged(object sender, EventArgs e)
        {
            if (_loading)
                return;
            if (!UserSettings.SaveUpdateCheckOnStart(_checkOnStart.Checked))
            {
                ShowSaveError();
                LoadAndShow();
            }
        }

        private void OnUnskip(object sender, EventArgs e)
        {
            if (!UserSettings.SaveSkippedVersion(null))
                ShowSaveError();
            LoadAndShow(); // кнопка исчезает: напоминать больше нечего отменять
        }

        private void ShowSaveError()
        {
            Dialogs.Error(this, Loc.T("settings.title"), Loc.T("settings.err.save.title"),
                Loc.T("settings.err.save.body"));
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
