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
    public abstract class PdfToolFormBase : Form
    {
        protected readonly Action _showHub;
        protected PdfPageGrid _grid;
        protected TrackBar _zoom;
        protected System.Windows.Forms.Timer _zoomTimer;
        protected CompressionPicker _compress;
        protected Label _lblStatus;
        protected ToolTip _tips;
        protected bool _busy; // идёт фоновая операция (только UI-поток)

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
            if (_showHub != null)
            {
                Button home = Ui.HomeButton(_showHub);
                home.SetBounds(header.Width - 180, 22, 160, 30);
                home.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                _tips.SetToolTip(home, "Открыть экран выбора инструмента");
                header.Controls.Add(home);
            }
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
        /// Общий нижний строй обоих PDF-инструментов: подпись + ползунок масштаба (с
        /// троттлинг-таймером), выбор сжатия и строка статуса. Вызывать ПОСЛЕ создания
        /// сетки (_grid) — к ней привязан масштаб. right — правый край рабочей области.
        /// </summary>
        protected void BuildBottomStrip(int right, string statusText, bool withCompress = true)
        {
            int h = ClientSize.Height;
            Ui.Label(this, "Масштаб:", 20, h - 104, Font, Theme.TextMuted)
                .Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

            _zoom = new TrackBar();
            _zoom.SetBounds(85, h - 108, 180, 30);
            _zoom.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            _zoom.Minimum = ThumbZoom.MinWidth;
            _zoom.Maximum = ThumbZoom.MaxWidth;
            _zoom.Value = _grid.TileWidth; // до подписки на ValueChanged — не триггерит ScheduleZoom
            _zoom.TickFrequency = 32;
            _zoom.SmallChange = 16;
            _zoom.LargeChange = 32;
            _zoom.ValueChanged += delegate { ScheduleZoom(); };
            _tips.SetToolTip(_zoom, "Масштаб миниатюр (также Ctrl+колесо мыши)");
            Controls.Add(_zoom);
            _grid.ZoomChanged += delegate(int w) { _zoom.Value = w; }; // Ctrl+колесо в сетке двигает ползунок

            _zoomTimer = new System.Windows.Forms.Timer();
            _zoomTimer.Interval = 60; // троттлинг пересборки плиток при перетаскивании ползунка
            _zoomTimer.Tick += delegate { _zoomTimer.Stop(); _grid.SetTileWidth(_zoom.Value); };

            if (withCompress)
            {
                _compress = new CompressionPicker();
                _compress.Location = new Point(right - _compress.Width, h - 106);
                _compress.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
                Controls.Add(_compress);
            }

            _lblStatus = Ui.Label(this, statusText, 20, h - 50, Font, Theme.TextMuted);
            _lblStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

            // Полоса прогресса — в свободной зоне между рядом «Масштаб/Сжатие» (низ ~ h-78) и рядом
            // «статус/кнопка» (верх ~ h-58): целиком выше кнопки, поэтому ничего не перекрывает.
            _progress = new ProgressBar();
            _progress.SetBounds(20, h - 76, right - 20 - 52, 16);
            _progress.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _progress.Minimum = 0;
            _progress.Maximum = 100;
            _progress.Style = ProgressBarStyle.Continuous;
            _progress.Visible = false;
            Controls.Add(_progress);

            _progressPct = Ui.Label(this, "", right - 46, h - 76, Font, Theme.TextMuted);
            _progressPct.AutoSize = false; // фикс. ширина + выключка вправо: «0 %»/«100 %» стоят одинаково
            _progressPct.SetBounds(right - 46, h - 76, 46, 16);
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
            _progress.Value = pct;
            _progressPct.Text = pct + " %";
            _taskbar.SetValue(Handle, pct, 100);
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

        /// <summary>Сообщение при попытке закрыть окно во время фоновой операции.</summary>
        protected virtual string BusyMessage
        {
            get { return "Дождитесь завершения…"; }
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
