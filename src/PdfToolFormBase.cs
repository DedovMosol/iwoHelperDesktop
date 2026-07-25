using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Общая база инструментов работы с PDF («Объединение» и «Разделение»): держит
    /// сетку миниатюр, ползунок масштаба с троттлингом, выбор сжатия, строку статуса
    /// и подсказки; единообразно закрывается (дожидаясь фоновой операции) и
    /// детерминированно освобождает ресурсы. Наследник строит своё содержимое в
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
        protected Label _zoomPct; // подпись текущего масштаба в процентах
        protected System.Windows.Forms.Timer _zoomTimer;
        protected CompressionPicker _compress;
        protected Label _lblStatus;
        protected ToolTip _tips;
        protected bool _busy; // идёт фоновая операция (только UI-поток)
        // Отмена длинных операций: кнопка «Отмена» показывается поверх кнопки действия во
        // время операции от CancelPageThreshold единиц. Флаг ставит UI-поток, читает воркер.
        private Button _actionButton;
        private Button _btnCancel;
        private volatile bool _cancelRequested;
        private bool _cancelOffered; // на эту операцию показана кнопка отмены (восстановить в EndProgress)
        private const int CancelPageThreshold = 5; // от скольких единиц предлагать отмену

        /// <summary>Идёт операция: смена языка это окно не пересоздаёт (см. IBusyAware).</summary>
        public bool IsBusy
        {
            get { return _busy; }
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
            Font = new Font("Segoe UI", 9.75f);
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
        }

        /// <summary>
        /// Меню «Справка», брендовая шапка и кнопка «Главная» (единые, DRY). Вызывать
        /// после <see cref="InitShell"/>. Содержимое наследник кладёт ниже
        /// HelpMenu.Height + высоты шапки (76).
        /// </summary>
        protected void BuildHeaderWithHome(string title, string subtitle, Color colorTop, Color colorBottom, Action showHelp)
        {
            // Пункт «Горячие клавиши» — только у PDF-инструментов (у них сетка с клавишами).
            var shortcuts = new ToolStripMenuItem(Loc.T("menu.shortcuts"));
            shortcuts.Click += delegate { ShowShortcuts(); };
            MenuStrip menu = HelpMenu.Create(this, showHelp, shortcuts);
            MainMenuStrip = menu;
            Controls.Add(menu);

            var header = new HeaderBand(title, subtitle, colorTop, colorBottom);
            header.SetBounds(0, HelpMenu.Height, ClientSize.Width, 76);
            header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            header.TabIndex = 100; // «Главная» — в конце обхода Tab, а не в начале
            Controls.Add(header);
            Ui.HomeOnHeader(header, _showHub, _tips, 22);
        }

        /// <summary>Троттлинг пересборки плиток при перетаскивании ползунка масштаба.</summary>
        protected void ScheduleZoom()
        {
            _zoomTimer.Stop();
            _zoomTimer.Start();
        }

        /// <summary>Обновить подпись масштаба в процентах (от значения ползунка — мгновенно, до пересборки плиток).</summary>
        private void UpdateZoomPercent()
        {
            if (_zoomPct != null)
                _zoomPct.Text = ThumbZoom.Percent(_zoom.Value) + " %";
        }

        /// <summary>Ctrl+колесо над ползунком масштаба меняет масштаб шагом сетки (без Ctrl — родное поведение TrackBar).</summary>
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
        /// Текст «покоя» статуса: при выделении — «Выбрано N из M», иначе — idle. Чистая — под тест.
        /// </summary>
        internal static string RestingStatus(int selected, int total, string idle, string selectedFormat)
        {
            return selected > 0 ? string.Format(selectedFormat, selected, total) : idle;
        }

        /// <summary>Idle-текст статуса конкретного инструмента (без выделения). Переопределяют формы.</summary>
        protected virtual string IdleStatusText() { return ""; }

        /// <summary>
        /// Обновить статус «покоя»: счётчик выделения или idle-текст инструмента. Во время
        /// операции не трогаем статус (там прогресс). Зовут формы на смене выделения и мутациях.
        /// </summary>
        protected void RefreshRestingStatus()
        {
            if (_busy || _grid == null)
                return;
            SetStatus(RestingStatus(_grid.SelectedCount, _grid.Count, IdleStatusText(),
                Loc.T("common.status.selected")), Theme.TextMuted);
        }

        /// <summary>
        /// Общий нижний строй PDF-инструментов: подпись + ползунок масштаба (с
        /// троттлинг-таймером), выбор сжатия, полоса прогресса и строка статуса.
        /// Вызывать ПОСЛЕ создания сетки (_grid) — к ней привязан масштаб.
        /// right — правый край рабочей области; actionWidth — ширина кнопки действия
        /// в правом нижнем углу: статус обрезается до неё и не уходит под кнопку.
        /// </summary>
        protected void BuildBottomStrip(int right, string statusText, int actionWidth, bool withCompress = true)
        {
            int h = ClientSize.Height;
            // Восстановить масштаб прошлой сессии ДО создания ползунка (0 — не задан).
            UserSettings prefs = UserSettings.Load();
            if (prefs.ZoomWidth >= ThumbZoom.MinWidth)
                _grid.SetTileWidth(prefs.ZoomWidth);

            Ui.Label(this, Loc.T("common.zoom"), 20, h - 144, Font, Theme.TextMuted)
                .Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

            // Сжатие (если есть) создаём ПЕРВЫМ: от его левого края отмеряем подпись % и правый
            // край растяжимого ползунка. Подпись и сжатие оба на правом якоре → между ними
            // ПОСТОЯННЫЙ зазор на любой ширине окна, поэтому они не наезжают друг на друга даже
            // при минимальной ширине (замер: длинная метка сжатия иначе перекрывала бы «%»).
            const int pctW = 52; // ширина подписи «400 %»
            int pctRight;        // правый край подписи процента
            if (withCompress)
            {
                _compress = new CompressionPicker();
                _compress.SetLevel((CompressionLevel)prefs.CompressionLevel); // сжатие прошлой сессии
                _compress.Location = new Point(right - _compress.Width, h - 146);
                _compress.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
                Controls.Add(_compress);
                pctRight = _compress.Left - 12; // постоянный зазор до выбора сжатия
            }
            else
            {
                pctRight = right; // сжатия нет — подпись у правого края ряда
            }
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
            _zoom.ValueChanged += delegate { ScheduleZoom(); UpdateZoomPercent(); };
            _zoom.MouseWheel += OnZoomWheel; // Ctrl+колесо над ползунком — как в сетке
            _tips.SetToolTip(_zoom, Loc.T("common.tip.zoom"));
            Controls.Add(_zoom);
            _grid.ZoomChanged += delegate(int w) { _zoom.Value = w; }; // Ctrl+колесо в сетке двигает ползунок (→ %)

            // Текущий масштаб в процентах — справа от ползунка, правым якорём (левее сжатия).
            _zoomPct = Ui.Label(this, "", pctLeft, h - 138, Font, Theme.TextMuted);
            _zoomPct.AutoSize = false;
            _zoomPct.SetBounds(pctLeft, h - 138, pctW, 20);
            _zoomPct.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            UpdateZoomPercent();

            _zoomTimer = new System.Windows.Forms.Timer();
            _zoomTimer.Interval = 60; // троттлинг пересборки плиток при перетаскивании ползунка
            _zoomTimer.Tick += delegate { _zoomTimer.Stop(); _grid.SetTileWidth(_zoom.Value); };

            // Статус делит нижний ряд с кнопкой действия справа: фиксированная ширина до
            // кнопки + многоточие, длинный текст целиком покажет подсказка AutoEllipsis.
            _lblStatus = Ui.Label(this, statusText, 20, h - 50, Font, Theme.TextMuted);
            _lblStatus.AutoSize = false;
            _lblStatus.AutoEllipsis = true;
            _lblStatus.SetBounds(20, h - 50, right - actionWidth - 12 - 20, 20);
            _lblStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            // Полоса прогресса — отдельный ряд между «Масштаб/Сжатие» (низ h-103: ползунок
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
            _progress.Value = 0;
            _progressPct.Text = "0 %";
            _progress.Visible = true;
            _progressPct.Visible = true;
            _taskbar.SetState(Handle, TaskbarProgressState.Normal);
            _taskbar.SetValue(Handle, 0, 100);
            ShowCancelButton(ShouldOfferCancel(total));
        }

        /// <summary>Стоит ли предлагать отмену для операции такого размера (в единицах). Чистая — под тест.</summary>
        internal static bool ShouldOfferCancel(int total)
        {
            return total >= CancelPageThreshold;
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
            if (_actionButton == null)
                return;
            if (!show)
            {
                _cancelOffered = false;
                if (_btnCancel != null)
                    _btnCancel.Visible = false;
                _actionButton.Visible = true;
                return;
            }
            if (_btnCancel == null)
            {
                _btnCancel = new RoundedButton(false);
                _btnCancel.Bounds = _actionButton.Bounds;
                _btnCancel.Anchor = _actionButton.Anchor;
                _btnCancel.Click += delegate
                {
                    _cancelRequested = true;
                    _btnCancel.Enabled = false;
                    _btnCancel.Text = Loc.T("common.canceling");
                };
                Controls.Add(_btnCancel);
            }
            _btnCancel.Text = Loc.T("common.cancel");
            _btnCancel.Enabled = true;
            _btnCancel.Visible = true;
            _btnCancel.BringToFront();
            _actionButton.Visible = false;
            _cancelOffered = true;
        }

        /// <summary>Спрятать полосу по завершении/ошибке. Только UI-поток.</summary>
        protected void EndProgress()
        {
            _progress.Visible = false;
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
                try
                {
                    if (IsHandleCreated && !IsDisposed)
                        BeginInvoke((MethodInvoker)delegate { ApplyProgress(pct); });
                }
                catch (InvalidOperationException) { } // окно закрылось между проверкой и вызовом
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
                int zoom = _grid != null ? _grid.TileWidth : s.ZoomWidth;
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
                e.Cancel = true; // фоновая операция занимает секунды; иначе остался бы зомби-процесс
                return;
            }
            SaveViewPrefs(); // запомнить масштаб и уровень сжатия между запусками
            if (_grid != null)
                _grid.StopRendering(); // разбудить и остановить фоновый рендер
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
            }
            base.Dispose(disposing); // _grid освобождается как дочерний контрол
        }
    }
}
