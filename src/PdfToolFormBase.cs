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
        protected System.Windows.Forms.Timer _zoomTimer;
        protected CompressionPicker _compress;
        protected Label _lblStatus;
        protected ToolTip _tips;
        protected bool _busy; // идёт фоновая операция (только UI-поток)

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
            MenuStrip menu = HelpMenu.Create(this, showHelp);
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

        protected void SetStatus(string text, Color color)
        {
            _lblStatus.Text = text;
            _lblStatus.ForeColor = color;
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
            Ui.Label(this, Loc.T("common.zoom"), 20, h - 144, Font, Theme.TextMuted)
                .Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

            _zoom = new TrackBar();
            // Реальная высота TrackBar — 45 (AutoSize), заданную меньшую он игнорирует.
            _zoom.SetBounds(85, h - 148, 180, 45);
            _zoom.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            _zoom.Minimum = ThumbZoom.MinWidth;
            _zoom.Maximum = ThumbZoom.MaxWidth;
            _zoom.Value = _grid.TileWidth; // до подписки на ValueChanged — не триггерит ScheduleZoom
            _zoom.TickFrequency = 32;
            _zoom.SmallChange = 16;
            _zoom.LargeChange = 32;
            _zoom.ValueChanged += delegate { ScheduleZoom(); };
            _tips.SetToolTip(_zoom, Loc.T("common.tip.zoom"));
            Controls.Add(_zoom);
            _grid.ZoomChanged += delegate(int w) { _zoom.Value = w; }; // Ctrl+колесо в сетке двигает ползунок

            _zoomTimer = new System.Windows.Forms.Timer();
            _zoomTimer.Interval = 60; // троттлинг пересборки плиток при перетаскивании ползунка
            _zoomTimer.Tick += delegate { _zoomTimer.Stop(); _grid.SetTileWidth(_zoom.Value); };

            if (withCompress)
            {
                _compress = new CompressionPicker();
                _compress.Location = new Point(right - _compress.Width, h - 146);
                _compress.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
                Controls.Add(_compress);
            }

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

        /// <summary>Показать полосу в начале операции. Только UI-поток.</summary>
        protected void BeginProgress()
        {
            _lastPct = -1;
            _progress.Value = 0;
            _progressPct.Text = "0 %";
            _progress.Visible = true;
            _progressPct.Visible = true;
            _taskbar.SetState(Handle, TaskbarProgressState.Normal);
            _taskbar.SetValue(Handle, 0, 100);
        }

        /// <summary>Спрятать полосу по завершении/ошибке. Только UI-поток.</summary>
        protected void EndProgress()
        {
            _progress.Visible = false;
            _progressPct.Visible = false;
            _taskbar.SetState(Handle, TaskbarProgressState.None);
        }

        /// <summary>Применить процент к полосе и панели задач. Только UI-поток.</summary>
        private void ApplyProgress(int pct)
        {
            if (!_busy || pct == _lastPct)
                return; // поздний вызов после завершения или без изменения — пропускаем
            _lastPct = pct;
            SetBarInstant(pct);
            _progressPct.Text = pct + " %";
            _taskbar.SetValue(Handle, pct, 100);
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
            Cut, Copy, Paste, GoTo, RotateRight, RotateLeft, CancelClipboard
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
        /// обвязка, DRY): «Удалить» — как клавиша Delete, «Перейти…» — как Ctrl+G.
        /// Вызывать в BuildUi наследника сразу после создания _grid.
        /// </summary>
        protected void WireGridMenu()
        {
            _grid.DeleteRequested += delegate { if (_grid.AllowReorder) RemoveSelectedPages(); };
            _grid.GoToRequested += delegate { ShowGoToPage(); };
        }

        /// <summary>Диалог «Перейти к странице» (Ctrl+G и контекстное меню): выделяет и показывает её.</summary>
        protected void ShowGoToPage()
        {
            if (_grid == null || _grid.Count == 0)
                return;
            int page = GoToPageDialog.Show(this, _grid.Count);
            if (page > 0)
                _grid.SelectIndex(page - 1);
        }

        /// <summary>Удалить выбранные страницы (Delete). Зовётся только при AllowReorder.</summary>
        protected virtual void RemoveSelectedPages() { }

        /// <summary>Сдвинуть выбранную страницу раньше/позже (Alt+←/→). Зовётся только при AllowReorder.</summary>
        protected virtual void MoveSelectedPage(bool later) { }

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
