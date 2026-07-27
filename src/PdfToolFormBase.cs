using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Общая база всех инструментов работы с PDF («Объединение», «Разделение», «PDF → Word»):
    /// держит сетку миниатюр, регулятор масштаба с троттлингом, выбор сжатия, строку статуса
    /// и подсказки; единообразно закрывается (дожидаясь фоновой операции) и
    /// детерминированно освобождает ресурсы. Фоновый разбор PDF (добавление/открытие) идёт
    /// через <see cref="BeginLoad"/>/<see cref="EndLoad"/>, а доступность кнопок каждый
    /// наследник сводит в <see cref="SyncControls"/>. Наследник строит своё содержимое в
    /// собственном BuildUi и присваивает унаследованные поля; специфика (порядок и
    /// режимы, сохранение) остаётся в наследнике.
    ///
    /// База НЕ вызывает виртуальные методы из конструктора: поля наследника
    /// инициализируются позже, поэтому раскладку строит сам наследник, а база лишь
    /// хранит общее состояние и поведение (KISS, без анти-паттерна «virtual в ctor»).
    /// </summary>
    public abstract class PdfToolFormBase : Form, IBusyAware
    {
        protected readonly Action _showHub;
        protected PdfPageGrid _grid;
        protected TrackBar _zoom;
        protected NumericUpDown _zoomInput; // редактируемый масштаб в процентах (рядом со слайдером)
        protected Label _zoomPct; // подпись «%» справа от поля ввода масштаба
        private bool _zoomSync;   // идёт программная синхронизация поля % от регулятора — не применять обратно
        protected System.Windows.Forms.Timer _zoomTimer;
        protected CompressionPicker _compress;
        protected Label _lblStatus;
        protected ToolTip _tips;
        protected bool _busy; // идёт фоновая операция (только UI-поток)
        private bool _loading; // идёт фоновый разбор PDF (добавление/открытие) — только UI-поток

        /// <summary>Идёт фоновая операция ИЛИ фоновый разбор PDF: правки сетки и запуск заблокированы.</summary>
        protected bool Working { get { return _busy || _loading; } }
        // Отмена длинных операций: кнопка «Отмена» показывается поверх кнопки действия во
        // время операции от CancelPageThreshold единиц. Флаг ставит UI-поток, читает воркер.
        private Button _actionButton;   // кнопка действия (у окна «Прочие операции» её нет)
        private bool _cancelSlot;       // место для «Отмены» объявлено формой
        private Button _btnCancel;
        private volatile bool _cancelRequested;
        private bool _cancelOffered; // на эту операцию показана кнопка отмены (восстановить в EndProgress)
        private const int CancelPageThreshold = 5; // от скольких единиц предлагать отмену

        /// <summary>
        /// Идёт операция ИЛИ фоновый разбор PDF: смена языка это окно не пересоздаёт
        /// (см. IBusyAware). Разбор входит сюда наравне с операцией: пересоздание закрыло бы
        /// окно на полпути, а результат разбора ушёл бы в никуда — новое окно осталось бы
        /// пустым и без единого сообщения о том, что файл вообще-то загружался.
        /// </summary>
        public bool IsBusy
        {
            get { return Working; }
        }

        // Прогресс операции: полоса + проценты в свободной зоне нижнего строя (видны только во
        // время работы) и дублирование на кнопке панели задач Windows. Все обновления — на UI-потоке.
        protected ProgressBar _progress;
        protected Label _progressPct;
        private readonly TaskbarProgress _taskbar = new TaskbarProgress();
        private int _lastPct; // последний показанный процент (UI-поток) — чтобы не перерисовывать зря
        // «Страница N из M» в статусе во время операции: формат и всего единиц (0/null — только процент).
        private int _progTotal;
        private string _progFmt;

        protected PdfToolFormBase(Action showHub)
        {
            _showHub = showHub;
        }

        /// <summary>
        /// Базовая настройка окна PDF-инструмента (единая, DRY): шрифт, фон,
        /// центрирование, DPI-масштаб, размеры, цветной хром заголовка, AllowDrop,
        /// подсказки. Обработчики DragEnter/DragDrop наследник вешает сам (различаются).
        /// </summary>
        protected void InitShell(string title, Size clientSize, Size minSize, Color chromeColor)
        {
            Text = title;
            Icon icon = Ui.AppIcon();
            if (icon != null)
                Icon = icon;
            Font = Ui.Font(9.75f); // общий кэшированный шрифт (не освобождать)
            BackColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = clientSize;
            MinimumSize = minSize;
            ShowInTaskbar = true;
            WindowChrome.Enable(this, chromeColor); // цветной заголовок на Windows 11
            AllowDrop = true;
            _tips = new ToolTip();
            WindowPlacement.Attach(this); // размер и положение окна между запусками
        }

        /// <summary>
        /// Меню «Справка», брендовая шапка и кнопка «Главная» (единые, DRY). Вызывать
        /// после <see cref="InitShell"/>. Содержимое наследник кладёт ниже
        /// HelpMenu.Height + высоты шапки (76).
        /// </summary>
        protected void BuildHeaderWithHome(string title, string subtitle, Color colorTop, Color colorBottom,
            Action showHelp, params ToolStripMenuItem[] extras)
        {
            // Пункт «Горячие клавиши» — только у PDF-инструментов (у них сетка с клавишами).
            var shortcuts = new ToolStripMenuItem(Loc.T("menu.shortcuts"));
            shortcuts.Click += delegate { ShowShortcuts(); };
            var items = new List<ToolStripMenuItem>();
            items.Add(shortcuts);
            if (extras != null)
                items.AddRange(extras);
            MenuStrip menu = HelpMenu.Create(this, showHelp, items.ToArray());
            MainMenuStrip = menu;
            Controls.Add(menu);

            var header = new HeaderBand(title, subtitle, colorTop, colorBottom);
            header.SetBounds(0, HelpMenu.Height, ClientSize.Width, 76);
            header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(header);
            Ui.HomeOnHeader(header, _showHub, _tips, 22);
            // «Главная» уходит в конец обхода Tab не здесь, а в OnLoad: индекс, заданный
            // при добавлении, поставил бы шапку, наоборот, первой (см. Ui.HeaderLastInTabOrder).
        }

        /// <summary>
        /// Окно собрано: отправляем шапку с кнопкой «Главная» в конец обхода Tab, чтобы
        /// фокус при открытии достался сетке/полю, а не возврату в меню. Единая точка на
        /// все PDF-инструменты (наследники своих OnLoad не пишут).
        /// </summary>
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            Ui.HeaderLastInTabOrder(this);
        }

        /// <summary>Троттлинг пересборки плиток при перетаскивании регулятора масштаба.</summary>
        protected void ScheduleZoom()
        {
            _zoomTimer.Stop();
            _zoomTimer.Start();
        }

        /// <summary>
        /// Синхронизировать поле ввода % со значением регулятора (мгновенно, до пересборки плиток).
        /// Под _zoomSync, чтобы вызванное этим ValueChanged поля НЕ применилось обратно к регулятору
        /// (иначе округление WidthFromPercent(Percent(w)) дало бы петлю/дрейф).
        /// </summary>
        private void SyncZoomInput()
        {
            if (_zoomInput == null)
                return;
            _zoomSync = true;
            try { _zoomInput.Value = ThumbZoom.Percent(_zoom.Value); }
            finally { _zoomSync = false; }
        }

        /// <summary>Ввод масштаба в поле %: применить к регулятору (тот же троттлинг). Петли нет — см. _zoomSync.</summary>
        private void OnZoomInputChanged(object sender, EventArgs e)
        {
            if (_zoomSync || _zoomInput == null || _zoom == null)
                return;
            int width = ThumbZoom.WidthFromPercent((int)_zoomInput.Value);
            if (width != _zoom.Value)
                _zoom.Value = width; // → ValueChanged → ScheduleZoom + SyncZoomInput (под guard)
        }

        /// <summary>Сбросить масштаб к 100% (Ctrl+0, двойной клик по «%»).</summary>
        private void ResetZoom()
        {
            if (_zoom != null)
                _zoom.Value = ThumbZoom.DefaultWidth; // 100% → ValueChanged применит и синхронизирует поле
        }

        /// <summary>
        /// Enter в поле «%»: зафиксировать введённое значение немедленно. Чтение
        /// <c>NumericUpDown.Value</c> при незавершённом вводе само валидирует и клампит текст в
        /// [Min..Max] (и обновляет отображение к разрешённому значению — прямое присваивание
        /// Value вне диапазона бросило бы исключение), а изменение значения применит масштаб
        /// через ValueChanged. ProcessCmdKey заодно гасит Enter, чтобы он не нажал кнопку
        /// действия окна (Сохранить/Конвертировать) и не начал операцию.
        /// </summary>
        private void CommitZoomInput()
        {
            if (_zoomInput != null)
                _ = _zoomInput.Value; // геттер валидирует и клампит введённый пользователем текст
        }

        /// <summary>Ctrl+0 (основной ряд и цифровой блок) — сброс масштаба к 100%. Чистая — под тест.</summary>
        internal static bool IsResetZoomKey(Keys keyData)
        {
            return keyData == (Keys.Control | Keys.D0) || keyData == (Keys.Control | Keys.NumPad0);
        }

        /// <summary>Ctrl+колесо над регулятором масштаба меняет масштаб шагом сетки (без Ctrl — родное поведение TrackBar).</summary>
        private void OnZoomWheel(object sender, MouseEventArgs e)
        {
            if ((ModifierKeys & Keys.Control) == 0)
                return;
            var handled = e as HandledMouseEventArgs;
            if (handled != null)
                handled.Handled = true;
            int w = ThumbZoom.StepFromWheel(_zoom.Value, e.Delta);
            if (w != _zoom.Value)
                _zoom.Value = w; // ValueChanged → ScheduleZoom + UpdateZoomPercent
        }

        protected void SetStatus(string text, Color color)
        {
            _lblStatus.Text = text;
            _lblStatus.ForeColor = color;
        }

        /// <summary>
        /// Строка успешного завершения: «✓ » + непустые части через « · » + точка. Пунктуацию
        /// собираем здесь одинаково для всех инструментов, поэтому в каталоге строк части
        /// лежат без галочки, разделителей и точки, а null и пустые просто отбрасываются.
        /// Пустой набор частей даёт пустую строку. Чистая — под тест.
        /// </summary>
        internal static string SuccessStatus(params string[] parts)
        {
            var kept = new List<string>();
            if (parts != null)
                foreach (string part in parts)
                    if (!string.IsNullOrEmpty(part))
                        kept.Add(part);
            return kept.Count == 0 ? "" : "✓ " + string.Join(" · ", kept.ToArray()) + ".";
        }

        /// <summary>
        /// Часть статуса про сжатие для инструмента с ОДНИМ результатом: «сжато, изображения
        /// до 150 dpi». null, если сжатия не было — тогда часть просто не попадёт в статус.
        /// Разрешение берётся из <see cref="PdfCompression.ImageDpi"/> (единственный источник).
        /// Чистая — под тест.
        /// </summary>
        internal static string CompressedPart(bool compressed, CompressionLevel level)
        {
            return compressed ? string.Format(Loc.T("common.suffix.compressed"), PdfCompression.ImageDpi(level)) : null;
        }

        /// <summary>
        /// Текст «покоя» статуса: при выделении — «Выбрано N из M», иначе — idle. Чистая — под тест.
        /// </summary>
        internal static string RestingStatus(int selected, int total, string idle, string selectedFormat)
        {
            return selected > 0 ? string.Format(selectedFormat, selected, total) : idle;
        }

        /// <summary>Idle-текст статуса конкретного инструмента (без выделения). Переопределяют формы.</summary>
        protected virtual string IdleStatusText() { return ""; }

        /// <summary>
        /// Имя инструмента для заголовков диалогов (реализуют формы). Объявлено здесь, а не в
        /// промежуточных базах: заголовок нужен обеим ветвям — и инструментам, собирающим итог
        /// из нескольких файлов, и инструментам одного открытого документа.
        /// </summary>
        protected abstract string ToolTitle { get; }

        /// <summary>
        /// Обновить статус «покоя»: счётчик выделения или idle-текст инструмента. Во время
        /// операции не трогаем статус (там прогресс). Зовут формы на смене выделения и мутациях.
        /// </summary>
        protected void RefreshRestingStatus()
        {
            if (Working || _grid == null)
                return; // во время операции/загрузки статус занят прогрессом или «Загрузка…»
            SetStatus(RestingStatus(_grid.SelectedCount, _grid.Count, IdleStatusText(),
                Loc.T("common.status.selected")), Theme.TextMuted);
        }

        /// <summary>
        /// Синхронизировать доступность кнопок и блокировку сетки с текущим состоянием
        /// (Working, выделение, наличие содержимого). Реализуют формы. Зовётся базой после
        /// смены состояния загрузки/операции и формами на своих мутациях.
        /// </summary>
        protected abstract void SyncControls();

        /// <summary>
        /// Начать фоновый разбор PDF: пометить загрузку, отключить кнопки/заблокировать сетку
        /// (SyncControls), показать статус loadingText. false — уже идёт операция/загрузка
        /// (повторный вызов игнорируется). Пару составляет <see cref="EndLoad"/>. Только UI-поток.
        /// </summary>
        protected bool BeginLoad(string loadingText)
        {
            if (Working)
                return false;
            _loading = true;
            SyncControls();
            SetStatus(loadingText, Theme.TextMuted);
            return true;
        }

        /// <summary>Завершить фоновый разбор PDF: снять пометку и восстановить кнопки/сетку. Только UI-поток.</summary>
        protected void EndLoad()
        {
            _loading = false;
            SyncControls();
        }

        /// <summary>
        /// Единая обвязка приёма PDF-файлов, брошенных на окно (DRY: раньше по копии в каждой
        /// форме). Курсор-эффект только когда не идёт операция/загрузка и в дропе есть файлы;
        /// сам дроп передаёт непустой список в onFiles. Обработчик грида (дроп на сетку) — отдельно.
        /// </summary>
        protected void WireFileDrop(Action<string[]> onFiles)
        {
            DragEnter += delegate(object s, DragEventArgs e)
            {
                e.Effect = !Working && PdfDrop.ExtractPaths(e).Length > 0
                    ? DragDropEffects.Copy : DragDropEffects.None;
            };
            DragDrop += delegate(object s, DragEventArgs e)
            {
                if (Working)
                    return;
                string[] paths = PdfDrop.ExtractPaths(e);
                if (paths.Length > 0)
                    onFiles(paths);
            };
        }

        /// <summary>
        /// Общий нижний строй PDF-инструментов: подпись + регулятор масштаба (с
        /// троттлинг-таймером), выбор сжатия, полоса прогресса и строка статуса.
        /// Вызывать ПОСЛЕ создания сетки (_grid) — к ней привязан масштаб.
        /// right — правый край рабочей области; actionWidth — ширина кнопки действия
        /// в правом нижнем углу: статус обрезается до неё и не уходит под кнопку.
        /// </summary>
        protected void BuildBottomStrip(int right, string statusText, int actionWidth, bool withCompress = true)
        {
            int h = ClientSize.Height;
            // Восстановить масштаб прошлой сессии ДО создания регулятора (0 — не задан).
            UserSettings prefs = UserSettings.Load();
            if (prefs.ZoomWidth >= ThumbZoom.MinWidth)
                _grid.SetTileWidth(prefs.ZoomWidth);

            Ui.Label(this, Loc.T("common.zoom"), 20, h - 144, Font, Theme.TextMuted)
                .Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

            // Сжатие (если есть) создаём ПЕРВЫМ: от его левого края отмеряем подпись % и правый
            // край растяжимого регулятора. Подпись и сжатие оба на правом якоре → между ними
            // ПОСТОЯННЫЙ зазор на любой ширине окна, поэтому они не наезжают друг на друга даже
            // при минимальной ширине (замер: длинная метка сжатия иначе перекрывала бы «%»).
            const int pctW = 62;        // ширина блока масштаба: поле ввода «303» + подпись «%»
            const int compressGap = 12; // зазор между блоком «%» и выбором сжатия
            // Место под выбор сжатия резервируется ВСЕГДА — даже когда его нет (PDF → Word), — чтобы
            // регулятор масштаба и блок «%» были одинакового размера и положения на ВСЕХ PDF-экранах.
            int compressSlot = CompressionPicker.PreferredWidth();
            if (withCompress)
            {
                _compress = new CompressionPicker();
                _compress.SetLevel((CompressionLevel)prefs.CompressionLevel); // сжатие прошлой сессии
                _compress.Location = new Point(right - _compress.Width, h - 146);
                _compress.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
                Controls.Add(_compress);
            }
            int pctRight = right - compressSlot - compressGap; // единый правый край «%» на всех экранах
            int pctLeft = pctRight - pctW;
            const int sliderLeft = 85;

            _zoom = new TrackBar();
            // Реальная высота TrackBar — 45 (AutoSize), заданную меньшую он игнорирует.
            // Ширина — до подписи %; тянется с окном (левый+правый якорь), не сталкиваясь со сжатием.
            _zoom.SetBounds(sliderLeft, h - 148, pctLeft - 10 - sliderLeft, 45);
            _zoom.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _zoom.Minimum = ThumbZoom.MinWidth;
            _zoom.Maximum = ThumbZoom.MaxWidth;
            _zoom.Value = _grid.TileWidth; // до подписки на ValueChanged — не триггерит ScheduleZoom
            _zoom.TickFrequency = 32;
            _zoom.SmallChange = 16;
            _zoom.LargeChange = 32;
            _zoom.ValueChanged += delegate { ScheduleZoom(); SyncZoomInput(); };
            _zoom.MouseWheel += OnZoomWheel; // Ctrl+колесо над регулятором — как в сетке
            _tips.SetToolTip(_zoom, Loc.T("common.tip.zoom"));
            Controls.Add(_zoom);
            _grid.ZoomChanged += delegate(int w) { _zoom.Value = w; }; // Ctrl+колесо в сетке двигает регулятор (→ %)

            // Текущий масштаб — редактируемое поле процентов + подпись «%» справа от регулятора,
            // правым якорём (левее сжатия). Поле даёт точный ввод, которого не хватало регулятору;
            // регулятор остаётся. Двусторонняя синхронизация — через _zoomSync (без петли/дрейфа).
            const int suffixW = 16;            // подпись «%»
            int inputW = pctW - suffixW - 2;   // поле ввода процента (44)
            _zoomInput = new NumericUpDown();
            _zoomInput.Minimum = ThumbZoom.MinPercent;
            _zoomInput.Maximum = ThumbZoom.MaxPercent;
            _zoomInput.Value = ThumbZoom.Percent(_grid.TileWidth); // ∈ [MinPercent..MaxPercent]
            _zoomInput.TextAlign = HorizontalAlignment.Right;
            _zoomInput.SetBounds(pctLeft, h - 140, inputW, 22);
            _zoomInput.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _zoomInput.ValueChanged += OnZoomInputChanged;
            _tips.SetToolTip(_zoomInput, Loc.T("common.tip.zoomInput"));
            Controls.Add(_zoomInput);

            _zoomPct = Ui.Label(this, "%", pctLeft + inputW + 2, h - 138, Font, Theme.TextMuted);
            _zoomPct.AutoSize = false;
            _zoomPct.SetBounds(pctLeft + inputW + 2, h - 138, suffixW, 20);
            _zoomPct.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _zoomPct.Cursor = Cursors.Hand;                    // рука на «%»: видно, что кликабельно (сброс)
            _zoomPct.DoubleClick += delegate { ResetZoom(); }; // двойной клик по «%» → 100%

            _zoomTimer = new System.Windows.Forms.Timer();
            _zoomTimer.Interval = 60; // троттлинг пересборки плиток при перетаскивании регулятора
            _zoomTimer.Tick += delegate { _zoomTimer.Stop(); _grid.SetTileWidth(_zoom.Value); };

            // Статус делит нижний ряд с кнопкой действия справа: фиксированная ширина до
            // кнопки + многоточие, длинный текст целиком покажет подсказка AutoEllipsis.
            _lblStatus = Ui.Ellipsize(Ui.Label(this, statusText, 20, h - 50, Font, Theme.TextMuted),
                right - actionWidth - 12 - 20, AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right);

            // Полоса прогресса — отдельный ряд между «Масштаб/Сжатие» (низ h-103: регулятор
            // h-148 + 45) и «статус/кнопка» (верх h-58): зазоры 11 и 18 px, перекрытий нет.
            _progress = new ProgressBar();
            _progress.SetBounds(20, h - 92, right - 20 - 52, 16);
            _progress.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _progress.Minimum = 0;
            _progress.Maximum = 100;
            _progress.Style = ProgressBarStyle.Continuous;
            _progress.Visible = false;
            Controls.Add(_progress);

            _progressPct = Ui.Label(this, "", right - 46, h - 92, Font, Theme.TextMuted);
            _progressPct.AutoSize = false; // фикс. ширина + выключка вправо: «0 %»/«100 %» стоят одинаково
            _progressPct.SetBounds(right - 46, h - 92, 46, 16);
            _progressPct.TextAlign = ContentAlignment.MiddleRight;
            _progressPct.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _progressPct.Visible = false;
        }

        /// <summary>Проценты из «сделано/всего»: total ≤ 0 → 0; клампится в 0..100. Чистая — под тест.</summary>
        internal static int ProgressPercent(int completed, int total)
        {
            if (total <= 0 || completed <= 0)
                return 0;
            if (completed >= total)
                return 100;
            return (int)(100L * completed / total);
        }

        /// <summary>
        /// Номер текущей единицы (страницы) из процента и их общего числа: 1..total, 0 при
        /// total ≤ 0. Приблизительно (по проценту): у некоторых операций шкала двухфазная.
        /// Чистая — под тест.
        /// </summary>
        internal static int ProgressItem(int pct, int total)
        {
            if (total <= 0)
                return 0;
            int item = (int)Math.Round(pct / 100.0 * total, MidpointRounding.AwayFromZero);
            if (item < 1) return 1;
            if (item > total) return total;
            return item;
        }

        /// <summary>
        /// Показать полосу в начале операции. total &gt; 0 и runningFormat («… {0} из {1}»)
        /// включают счётчик «страница N из M» в статусе; иначе — только процент. Только UI-поток.
        /// </summary>
        protected void BeginProgress(int total = 0, string runningFormat = null)
        {
            _lastPct = -1;
            _progTotal = total;
            _progFmt = runningFormat;
            _cancelRequested = false;
            SetDeterminate(true); // прошлая операция могла оставить полосу бегущей
            _progress.Value = 0;
            _progressPct.Text = "0 %";
            _progress.Visible = true;
            _taskbar.SetState(Handle, TaskbarProgressState.Normal);
            _taskbar.SetValue(Handle, 0, 100);
            ShowCancelButton(ShouldOfferCancel(total));
        }

        /// <summary>
        /// Перевести индикатор в режим «идёт работа неизвестной длительности»: бегущая полоса
        /// без процентов и тот же режим на кнопке панели задач. Нужен фазе сжатия — Ghostscript
        /// работает отдельным процессом и о ходе не сообщает, поэтому полоса, застывшая на 100%
        /// со снятой кнопкой «Отмена», выглядела как зависшее окно, хотя работа шла ещё минуты.
        /// Счётчик страниц гасим: на этой фазе он больше не осмыслен. Только UI-поток.
        /// </summary>
        protected void BeginIndeterminate(string statusText)
        {
            _progFmt = null;
            SetDeterminate(false);
            _taskbar.SetState(Handle, TaskbarProgressState.Indeterminate);
            SetStatus(statusText, Theme.TextMuted);
        }

        /// <summary>Полоса измеримая (с процентами) или бегущая — одно место на оба перехода (DRY).</summary>
        private void SetDeterminate(bool determinate)
        {
            _progress.Style = determinate ? ProgressBarStyle.Continuous : ProgressBarStyle.Marquee;
            _progress.MarqueeAnimationSpeed = determinate ? 0 : 30;
            _progressPct.Visible = determinate;
        }

        /// <summary>Стоит ли предлагать отмену для операции такого размера (в единицах). Чистая — под тест.</summary>
        internal static bool ShouldOfferCancel(int total)
        {
            return total >= CancelPageThreshold;
        }

        /// <summary>
        /// Начать фоновую ОПЕРАЦИЮ (сохранение/конвертация/разделение) — единый пролог всех
        /// трёх инструментов (DRY): пометить занятость, синхронизировать кнопки, показать
        /// статус и полосу прогресса (счётчик «страница N из M» при заданном runningFormat).
        /// Пару составляет <see cref="EndOperation"/>. Только UI-поток.
        /// </summary>
        protected void BeginOperation(string statusText, int total, string runningFormat = null)
        {
            _busy = true;
            SyncControls();
            SetStatus(statusText, Theme.TextMuted);
            BeginProgress(total, runningFormat);
        }

        /// <summary>
        /// Начать операцию, ход которой измерить НЕЛЬЗЯ: внешний движок о прогрессе не
        /// сообщает, а разовая запись файла не делится на шаги. Полоса сразу становится
        /// бегущей — застывшая на нуле она читается как зависшее окно, и это не догадка:
        /// Ghostscript на большом скане работает минутами. Только UI-поток.
        /// </summary>
        protected void BeginUnmeasuredOperation(string statusText)
        {
            BeginOperation(statusText, 0);
            BeginIndeterminate(statusText);
        }

        /// <summary>Завершить фоновую операцию — единый эпилог (DRY): снять занятость, спрятать полосу, синхронизировать кнопки. Только UI-поток.</summary>
        protected void EndOperation()
        {
            _busy = false;
            EndProgress();
            SyncControls();
        }

        /// <summary>
        /// Завершить операцию и разобрать её исход. true — операция удалась и вызывающему есть
        /// что показать; false — отмена или ошибка, они уже отражены в статусе (а ошибка ещё и
        /// в диалоге). Все три инструмента повторяли этот разбор дословно, различаясь ровно
        /// двумя строками, поэтому он живёт здесь. Только UI-поток.
        /// </summary>
        protected bool FinishOperation(Exception error, string failStatus, string failDialogTitle)
        {
            EndOperation();
            if (error is OperationCanceledException)
            {
                SetStatus(Loc.T("common.status.canceled"), Theme.WarnOrange); // результат не создан
                return false;
            }
            if (error != null)
            {
                SetStatus(failStatus, Theme.ErrRed);
                Dialogs.Error(this, Text, failDialogTitle, error.Message);
                return false;
            }
            return true;
        }

        /// <summary>Токен кооперативной отмены для сервисов: читает volatile-флаг «нажата Отмена».</summary>
        protected Func<bool> CancelToken()
        {
            return delegate { return _cancelRequested; };
        }

        /// <summary>Форма сообщает базе кнопку действия — её база подменяет кнопкой «Отмена» на время операции.</summary>
        protected void RegisterActionButton(Button b)
        {
            _actionButton = b;
            _cancelSlot = true;
        }

        /// <summary>
        /// Где показывать «Отмену» окну БЕЗ единственной кнопки действия («Прочие операции»:
        /// там каждая операция — своя кнопка в панели, подменять нечего). Место то же, что у
        /// кнопки действия остальных инструментов — правый нижний угол, чтобы отмена всегда
        /// была там, где её станут искать.
        ///
        /// Кнопку создаём СРАЗУ, пусть и скрытой: ленивое создание к первой операции взяло бы
        /// координаты, посчитанные для прежнего размера окна, а так её держит в углу якорь.
        /// </summary>
        protected void RegisterCancelArea(Rectangle bounds, AnchorStyles anchor)
        {
            _cancelSlot = true;
            EnsureCancelButton(bounds, anchor);
        }

        /// <summary>Создать скрытую кнопку «Отмена» в заданном месте (один раз на окно).</summary>
        private void EnsureCancelButton(Rectangle bounds, AnchorStyles anchor)
        {
            if (_btnCancel != null)
                return;
            _btnCancel = new RoundedButton(false);
            _btnCancel.Bounds = bounds;
            _btnCancel.Anchor = anchor;
            _btnCancel.Visible = false;
            _btnCancel.Click += delegate
            {
                _cancelRequested = true;
                _btnCancel.Enabled = false;
                _btnCancel.Text = Loc.T("common.canceling");
            };
            Controls.Add(_btnCancel);
        }

        /// <summary>
        /// Убрать предложение отмены НЕ дожидаясь конца операции — когда пройдена точка
        /// невозврата (результат уже записан на диск и отменить его без следа нельзя, дальше
        /// идёт лишь дожатие/пост-обработка). Возвращает кнопку действия (пока отключённую).
        /// Маршалить на UI-поток. Идемпотентна; EndProgress затем ничего не делает.
        /// </summary>
        protected void StopOfferingCancel()
        {
            if (_cancelOffered)
                ShowCancelButton(false);
        }

        /// <summary>Показать/скрыть кнопку «Отмена» поверх кнопки действия (создаётся один раз по её геометрии).</summary>
        private void ShowCancelButton(bool show)
        {
            if (!_cancelSlot)
                return;
            if (!show)
            {
                _cancelOffered = false;
                if (_btnCancel != null)
                    _btnCancel.Visible = false;
                if (_actionButton != null)
                    _actionButton.Visible = true;
                return;
            }
            // У окна с кнопкой действия «Отмена» встаёт ровно на её место, и берём его ЖИВЫМ:
            // окно могли изменить в размере с момента постройки.
            if (_actionButton != null)
                EnsureCancelButton(_actionButton.Bounds, _actionButton.Anchor);
            if (_btnCancel == null)
                return;
            _btnCancel.Text = Loc.T("common.cancel");
            _btnCancel.Enabled = true;
            _btnCancel.Visible = true;
            _btnCancel.BringToFront();
            if (_actionButton != null)
                _actionButton.Visible = false; // окно без единственной кнопки действия прятать нечего
            _cancelOffered = true;
        }

        /// <summary>Спрятать полосу по завершении/ошибке. Только UI-поток.</summary>
        protected void EndProgress()
        {
            _progress.Visible = false;
            SetDeterminate(true); // не оставляем скрытую полосу «бегущей» после фазы сжатия
            _progressPct.Visible = false;
            _progFmt = null;
            if (_cancelOffered)
                ShowCancelButton(false); // вернуть кнопку действия на место // счётчик страниц больше не обновляем
            _taskbar.SetState(Handle, TaskbarProgressState.None);
        }

        /// <summary>Применить процент к полосе, панели задач и (при заданном формате) счётчику страниц. Только UI-поток.</summary>
        private void ApplyProgress(int pct)
        {
            if (!_busy || pct == _lastPct)
                return; // поздний вызов после завершения или без изменения — пропускаем
            _lastPct = pct;
            SetBarInstant(pct);
            _progressPct.Text = pct + " %";
            _taskbar.SetValue(Handle, pct, 100);
            if (_progFmt != null && _progTotal > 0)
                SetStatus(string.Format(_progFmt, ProgressItem(pct, _progTotal), _progTotal), Theme.TextMuted);
        }

        /// <summary>
        /// Нарисовать полосу на точном значении МГНОВЕННО, минуя «догоняющую» анимацию Vista+
        /// ProgressBar. Из-за этой анимации при частых обновлениях в кадре виден разрыв заливки
        /// (белый провал, будто полосу перекрыли). Приём: задать значение на 1 больше и тут же
        /// вернуть — уменьшение ProgressBar не анимируется, поэтому рисуется сразу и целиком.
        /// </summary>
        private void SetBarInstant(int pct)
        {
            if (pct >= _progress.Maximum)
                _progress.Value = _progress.Maximum;
            else
            {
                _progress.Value = pct + 1;
                _progress.Value = pct;
            }
        }

        /// <summary>Доставка действия в UI-поток этого окна из воркера (guard + BeginInvoke, см. Ui.OnUi).</summary>
        protected void OnUi(Action action)
        {
            Ui.OnUi(this, action);
        }

        /// <summary>
        /// Колбэк прогресса для сервиса (вызывается из воркера): считает процент, отсекает
        /// повторы и маршалит применение на UI-поток. Троттлинг по проценту — не чаще 100 обновлений.
        /// </summary>
        protected Action<int, int> UiProgress()
        {
            int workerLastPct = -1; // читается/пишется только воркером — без гонок
            return delegate(int completed, int total)
            {
                int pct = ProgressPercent(completed, total);
                if (pct == workerLastPct)
                    return;
                workerLastPct = pct;
                OnUi(delegate { ApplyProgress(pct); });
            };
        }

        /// <summary>Действие клавиатуры для сетки страниц. Чистая — под тест.</summary>
        internal enum PageKeyAction
        {
            None, Remove, MoveEarlier, MoveLater, SelectAll, Swallow,
            Cut, Copy, Paste, GoTo, RotateRight, RotateLeft, CancelClipboard, Undo, Redo
        }

        internal static PageKeyAction ClassifyPageKey(Keys keyData)
        {
            if (keyData == Keys.Delete) return PageKeyAction.Remove;
            if (keyData == (Keys.Alt | Keys.Left)) return PageKeyAction.MoveEarlier;
            if (keyData == (Keys.Alt | Keys.Right)) return PageKeyAction.MoveLater;
            if (keyData == (Keys.Control | Keys.A)) return PageKeyAction.SelectAll;
            if (keyData == (Keys.Control | Keys.X)) return PageKeyAction.Cut;
            if (keyData == (Keys.Control | Keys.C)) return PageKeyAction.Copy;
            if (keyData == (Keys.Control | Keys.V)) return PageKeyAction.Paste;
            if (keyData == (Keys.Control | Keys.Z)) return PageKeyAction.Undo;
            // Возврат отката — оба общепринятых сочетания.
            if (keyData == (Keys.Control | Keys.Y) ||
                keyData == (Keys.Control | Keys.Shift | Keys.Z)) return PageKeyAction.Redo;
            if (keyData == (Keys.Control | Keys.G)) return PageKeyAction.GoTo;
            // Поворот как в Acrobat: Ctrl+Shift+«+»/«−» (основной ряд и цифровой блок).
            if (keyData == (Keys.Control | Keys.Shift | Keys.Oemplus) ||
                keyData == (Keys.Control | Keys.Shift | Keys.Add)) return PageKeyAction.RotateRight;
            if (keyData == (Keys.Control | Keys.Shift | Keys.OemMinus) ||
                keyData == (Keys.Control | Keys.Shift | Keys.Subtract)) return PageKeyAction.RotateLeft;
            if (keyData == Keys.Escape) return PageKeyAction.CancelClipboard;
            if (keyData == Keys.Enter) return PageKeyAction.Swallow;
            return PageKeyAction.None;
        }

        /// <summary>
        /// Единые горячие клавиши сетки страниц (одна раскладка на все PDF-инструменты):
        /// Ctrl+A — выделить всё, Ctrl+G — перейти к странице; в редактируемых сетках
        /// (AllowReorder) Delete, Alt+←/→ и буфер страниц Ctrl+X/C/V правят порядок,
        /// Esc отменяет вырезание, а Enter гасится, чтобы не жать кнопку действия по
        /// AcceptButton; при AllowRotate Ctrl+Shift+«+»/«−» поворачивают выбранные.
        /// Остальное уходит в стандартную обработку, как и раньше.
        /// </summary>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (_zoom != null && IsResetZoomKey(keyData)) // Ctrl+0 — сброс масштаба к 100% (на уровне окна)
            {
                ResetZoom();
                return true;
            }
            // Enter в поле масштаба применяет введённый процент и НЕ запускает кнопку действия
            // окна (иначе редактирование «%» случайно начинало бы Сохранение/Конвертацию).
            if (keyData == Keys.Enter && _zoomInput != null && _zoomInput.Focused)
            {
                CommitZoomInput();
                return true;
            }
            if (_grid != null && _grid.ListFocused)
            {
                switch (ClassifyPageKey(keyData))
                {
                    case PageKeyAction.SelectAll:
                        _grid.SelectAll();
                        return true;
                    case PageKeyAction.Remove:
                        if (_grid.AllowReorder) { RemoveSelectedPages(); return true; }
                        break;
                    case PageKeyAction.MoveEarlier:
                        if (_grid.AllowReorder) { MoveSelectedPage(false); return true; }
                        break;
                    case PageKeyAction.MoveLater:
                        if (_grid.AllowReorder) { MoveSelectedPage(true); return true; }
                        break;
                    case PageKeyAction.Cut:
                        if (_grid.AllowReorder) { _grid.CutSelected(); return true; }
                        break;
                    case PageKeyAction.Copy:
                        if (_grid.AllowReorder) { _grid.CopySelected(); return true; }
                        break;
                    case PageKeyAction.Paste:
                        if (_grid.AllowReorder) { _grid.PasteClipboard(); return true; }
                        break;
                    case PageKeyAction.Undo:
                        if (_grid.AllowReorder) { UndoOrder(); return true; }
                        break;
                    case PageKeyAction.Redo:
                        if (_grid.AllowReorder) { RedoOrder(); return true; }
                        break;
                    case PageKeyAction.GoTo:
                        ShowGoToPage();
                        return true;
                    case PageKeyAction.RotateRight:
                        if (_grid.AllowRotate) { _grid.RotateSelected(90); return true; }
                        break;
                    case PageKeyAction.RotateLeft:
                        if (_grid.AllowRotate) { _grid.RotateSelected(-90); return true; }
                        break;
                    case PageKeyAction.CancelClipboard:
                        if (_grid.TryCancelCut()) return true; // иначе Esc идёт дальше (диалоги и т.п.)
                        break;
                    case PageKeyAction.Swallow:
                        if (_grid.AllowReorder) return true;
                        break;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        /// <summary>
        /// Подключить события контекстного меню сетки к обработчикам формы (единая
        /// обвязка, DRY): «Удалить» — как клавиша Delete, «Перейти…» — как Ctrl+G,
        /// «Переместить после страницы N…» — диалог формы. Вызывать в BuildUi
        /// наследника сразу после создания _grid.
        /// </summary>
        protected void WireGridMenu()
        {
            _grid.DeleteRequested += delegate { if (_grid.AllowReorder) RemoveSelectedPages(); };
            _grid.GoToRequested += delegate { ShowGoToPage(); };
            _grid.MoveAfterRequested += delegate { ShowMoveAfterPages(); };
        }

        /// <summary>
        /// Печать набора страниц — общая для всех PDF-инструментов (DRY): диалог принтера на
        /// UI-потоке, сама печать в фоне через общий воркер, прогресс и отмена как у остальных
        /// операций. Наследник лишь отдаёт страницы, которые надо напечатать.
        /// </summary>
        protected void PrintPages(System.Collections.Generic.IList<PdfPageRef> pages)
        {
            if (Working || pages == null || pages.Count == 0)
                return;
            System.Drawing.Printing.PrinterSettings settings;
            using (var dialog = new PrintDialog())
            {
                dialog.AllowSomePages = false; // какие страницы печатать, выбирают в сетке
                dialog.UseEXDialog = true;
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                settings = dialog.PrinterSettings;
            }
            BeginOperation(Loc.T("common.status.printing"), pages.Count, Loc.T("common.status.printingPage"));
            Action<int, int> onProgress = UiProgress();
            Func<bool> cancel = CancelToken();
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                try { PdfPrintService.Print(pages, settings, onProgress, cancel); }
                catch (Exception ex) { error = ex; }
                OnUi(delegate
                {
                    if (!FinishOperation(error, Loc.T("common.status.notDone"), Loc.T("common.err.printFailed")))
                        return;
                    SetStatus(SuccessStatus(string.Format(Loc.T("common.status.printed"), pages.Count)), Theme.OkGreen);
                });
            });
        }

        /// <summary>Диалог «Перейти к странице» (Ctrl+G и контекстное меню): выделяет и показывает её.</summary>
        protected void ShowGoToPage()
        {
            if (_grid == null || _grid.Count == 0)
                return;
            int page = NumberPromptDialog.Show(this, Loc.T("goto.title"),
                string.Format(Loc.T("goto.prompt"), _grid.Count), Loc.T("goto.ok"), 1, _grid.Count, 1);
            if (page > 0)
                _grid.SelectIndex(page - 1);
        }

        /// <summary>
        /// Диалог «Переместить после страницы N…»: выбранные страницы встают сразу за
        /// страницей N (0 — в самое начало). Перенос идёт тем же событием, что и вставка
        /// вырезанных, — форма правит модель.
        /// </summary>
        protected void ShowMoveAfterPages()
        {
            if (_grid == null || !_grid.AllowReorder || _grid.Locked || _grid.SelectedCount == 0)
                return;
            int n = NumberPromptDialog.Show(this, Loc.T("moveafter.title"),
                string.Format(Loc.T("moveafter.prompt"), _grid.Count), Loc.T("moveafter.ok"),
                0, _grid.Count, 0);
            if (n >= 0)
                _grid.MoveSelectedTo(n); // «после страницы N» = вставка перед позицией N
        }

        /// <summary>Шпаргалка горячих клавиш сетки: набор зависит от возможностей инструмента.</summary>
        protected void ShowShortcuts()
        {
            if (_grid == null)
                return;
            Dialogs.Info(this, Loc.T("menu.shortcuts"), Loc.T("shortcuts.title"),
                BuildShortcuts(_grid.AllowReorder, _grid.AllowRotate, Loc.T));
        }

        /// <summary>
        /// Тело шпаргалки: общие клавиши всегда, буфер/перестановка — при reorder,
        /// поворот — при rotate. Переводчик t инъектируется (чистая — под тест).
        /// </summary>
        internal static string BuildShortcuts(bool reorder, bool rotate, Func<string, string> t)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine(t("shortcuts.zoom"));
            sb.AppendLine(t("shortcuts.selectAll"));
            sb.AppendLine(t("shortcuts.goto"));
            if (reorder)
            {
                sb.AppendLine(t("shortcuts.move"));
                sb.AppendLine(t("shortcuts.cutcopy"));
                sb.AppendLine(t("shortcuts.paste"));
                sb.AppendLine(t("shortcuts.delete"));
                sb.AppendLine(t("shortcuts.undo"));
            }
            if (rotate)
                sb.AppendLine(t("shortcuts.rotate"));
            return sb.ToString().TrimEnd();
        }

        /// <summary>Откат последнего изменения порядка (Ctrl+Z). Переопределяют формы с моделью порядка.</summary>
        protected virtual void UndoOrder() { }

        /// <summary>Возврат отката (Ctrl+Y / Ctrl+Shift+Z). Переопределяют формы с моделью порядка.</summary>
        protected virtual void RedoOrder() { }

        /// <summary>Удалить выбранные страницы (Delete). Зовётся только при AllowReorder.</summary>
        protected virtual void RemoveSelectedPages() { }

        /// <summary>Сдвинуть выбранную страницу раньше/позже (Alt+←/→). Зовётся только при AllowReorder.</summary>
        protected virtual void MoveSelectedPage(bool later) { }

        /// <summary>
        /// Запомнить масштаб и уровень сжатия (load-modify-save сохраняет прочие настройки:
        /// другие окна держат устаревшую копию, поэтому пишем поверх свежей загрузки).
        /// </summary>
        private void SaveViewPrefs()
        {
            try
            {
                UserSettings s = UserSettings.Load();
                // Источник истины масштаба — регулятор: 60-мс троттлинг мог ещё не применить
                // его последнее движение к сетке (закрытие сразу после сдвига).
                int zoom = _zoom != null ? _zoom.Value : (_grid != null ? _grid.TileWidth : s.ZoomWidth);
                int compression = _compress != null ? (int)_compress.Level : s.CompressionLevel;
                s.SaveView(zoom, compression); // масштаб/сжатие — явно (общий Save их бы сохранил с диска)
            }
            catch { } // настройки — не критично: сбой сохранения не мешает закрытию
        }

        /// <summary>Сообщение при попытке закрыть окно во время фоновой операции.</summary>
        protected virtual string BusyMessage
        {
            get { return Loc.T("common.busy"); }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_busy)
            {
                SetStatus(BusyMessage, Theme.WarnOrange);
                e.Cancel = true; // фоновая операция занимает секунды, иначе остался бы зомби-процесс
                return;          // база не вызывается, поэтому границы окна тоже не сохранятся
            }
            SaveViewPrefs(); // запомнить масштаб и уровень сжатия между запусками
            // Фоновый рендер сетки останавливает её Dispose (stop + join): здесь звать
            // StopRendering нельзя — остановка необратима, а закрытие ещё может быть
            // ветировано (например, отмена завершения сеанса Windows после FormClosing).
            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_zoomTimer != null)
                    _zoomTimer.Dispose();
                // ToolTip — не дочерний контрол: авто-освобождение не срабатывает.
                if (_tips != null)
                    _tips.Dispose();
                _taskbar.Dispose(); // обёртка COM панели задач — детерминированно, как всё COM
            }
            base.Dispose(disposing); // _grid освобождается как дочерний контрол
        }
    }
}
