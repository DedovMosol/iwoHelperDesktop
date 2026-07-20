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

        protected PdfToolFormBase(Action showHub)
        {
            _showHub = showHub;
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
        protected void BuildBottomStrip(int right, string statusText)
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

            _compress = new CompressionPicker();
            _compress.Location = new Point(right - _compress.Width, h - 106);
            _compress.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Controls.Add(_compress);

            _lblStatus = Ui.Label(this, statusText, 20, h - 50, Font, Theme.TextMuted);
            _lblStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
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
