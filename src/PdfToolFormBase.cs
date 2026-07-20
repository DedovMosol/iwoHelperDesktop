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
