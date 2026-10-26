using System;
using System.Diagnostics;
using System.Threading;

namespace ExcelMerger
{
    /// <summary>Почему сторож остановил процесс Office (None — не останавливал).</summary>
    internal enum WatchdogStop
    {
        None,
        Cancel,  // человек отменил операцию, а текущий вызов не закончился за отведённое время
        Dialog,  // скрытый экземпляр показал модальное окно и ждёт ответа, которого не будет
        Timeout  // шаг идёт дольше предела (экземпляр занят без конца)
    }

    /// <summary>Процесс Office под присмотром сторожа. Реализация — <see cref="OwnedOfficeProcess"/>.</summary>
    internal interface IWatchedProcess
    {
        /// <summary>
        /// Заголовок модального окна, которое показал скрытый экземпляр ("" — окно без
        /// заголовка); null — такого окна нет или экземпляром пользуется человек.
        /// </summary>
        string ModalWindowTitle();

        /// <summary>Завершить процесс; false — нельзя (экземпляром пользуется человек) или он уже вышел.</summary>
        bool TryStop();
    }

    /// <summary>
    /// Отказ от операции после остановки сторожем: вызывающий узнаёт причину у сторожа
    /// (<see cref="ComWatchdog.Stop"/>) и переводит её в отмену или понятную ошибку.
    /// </summary>
    internal sealed class OfficeStoppedException : Exception
    {
        public OfficeStoppedException() : base("процесс Office остановлен сторожем") { }
    }

    /// <summary>
    /// Сторож зависшего COM-вызова. Вызов в скрытый Office блокирует STA-поток, пока сервер
    /// не ответит: модальное окно в невидимом экземпляре (на него некому ответить) вешало
    /// свод и запись .docx навсегда, отмена не работала, процесс оставался. Сторож живёт в
    /// своём потоке и опрашивает процесс; если шаг упёрся в модальное окно, идёт дольше
    /// предела или отменён и не закончился за отсрочку — завершает СВОЙ процесс (правила
    /// <see cref="OwnedOfficeProcess"/>: тот же PID и время старта, без окон человека).
    /// Зависший вызов после этого сразу возвращается с ошибкой RPC (проверено на живых
    /// Excel и Word 15.0: 0x800706BE за доли секунды).
    ///
    /// Работа размечается шагами Begin/End. Вне шага сторож ничего не делает, поэтому
    /// результат шага, закончившегося без остановки, не может испортить остановка «вдогонку»:
    /// решение и завершение процесса идут под той же блокировкой, что Begin/End. Все
    /// COM-вызовы вызывающий делает внутри шагов. Причина остановки не сбрасывается: процесс
    /// уже не вернуть.
    /// </summary>
    internal sealed class ComWatchdog : IDisposable
    {
        /// <summary>Пределы сторожа, мс. StepMs = 0 — без предела длительности шага.</summary>
        internal sealed class Limits
        {
            public int PollMs = 500;
            public int DialogMs = 10000;  // окно держится столько — ответа не будет
            public int CancelMs = 5000;   // отсрочка: текущий вызов успевает закончиться сам
            public int StepMs;
        }

        private readonly IWatchedProcess _target;
        private readonly Func<bool> _cancelRequested;
        private readonly Limits _limits;
        private readonly Func<long> _clock;
        private readonly object _gate = new object();
        private readonly ManualResetEvent _disposed;
        private readonly Thread _thread;

        private bool _armed;
        private bool _cancellable;
        private long _stepStart;
        private long _dialogSince = -1;
        private long _cancelSince = -1;
        private WatchdogStop _stop;
        private string _dialogTitle;
        private bool _isDisposed;

        /// <summary>
        /// target == null — процесс не определён: сторож ничего не делает (вызывающему не нужны
        /// проверки на null). cancelRequested читается из потока сторожа — должен быть
        /// потокобезопасным (volatile-флаг); null — отмены нет.
        /// </summary>
        public ComWatchdog(IWatchedProcess target, Func<bool> cancelRequested, Limits limits)
            : this(target, cancelRequested, limits, null, true)
        {
        }

        /// <summary>clock — монотонное время в мс (null — Stopwatch); run = false — без потока, Poll зовёт тест.</summary>
        internal ComWatchdog(IWatchedProcess target, Func<bool> cancelRequested, Limits limits,
            Func<long> clock, bool run)
        {
            if (limits == null)
                throw new ArgumentNullException("limits");
            _target = target;
            _cancelRequested = cancelRequested;
            _limits = limits;
            if (clock == null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                clock = delegate { return watch.ElapsedMilliseconds; };
            }
            _clock = clock;
            if (target == null || !run)
                return;
            _disposed = new ManualResetEvent(false);
            _thread = new Thread(Loop);
            _thread.IsBackground = true; // не держит процесс приложения при выходе
            _thread.Name = "Сторож COM-вызова";
            _thread.Start();
        }

        /// <summary>Начать шаг: отсчёт длительности, окна и отмены — заново. cancellable = false — шаг нельзя прерывать отменой (запись файла).</summary>
        public void Begin(bool cancellable)
        {
            lock (_gate)
            {
                _armed = true;
                _cancellable = cancellable;
                _stepStart = _clock();
                _dialogSince = -1;
                _cancelSince = -1;
            }
        }

        /// <summary>Закончить шаг; возвращает причину остановки (None — процесс работает).</summary>
        public WatchdogStop End()
        {
            lock (_gate)
            {
                _armed = false;
                return _stop;
            }
        }

        /// <summary>Закончить шаг; если процесс остановлен — <see cref="OfficeStoppedException"/>.</summary>
        public void EndStep()
        {
            if (End() != WatchdogStop.None)
                throw new OfficeStoppedException();
        }

        /// <summary>Причина остановки (None — не останавливал).</summary>
        public WatchdogStop Stop
        {
            get { lock (_gate) return _stop; }
        }

        /// <summary>Заголовок окна, из-за которого процесс остановлен ("" — без заголовка; null — причина другая).</summary>
        public string DialogTitle
        {
            get { lock (_gate) return _dialogTitle; }
        }

        /// <summary>Предел длительности шага, мс (для сообщения о таймауте).</summary>
        public int StepLimitMs
        {
            get { return _limits.StepMs; }
        }

        /// <summary>
        /// Что делать сейчас. dialogMs, cancelMs — сколько держится модальное окно и запрошенная
        /// отмена (-1 — нет). Отмена первой: человек уже решил, причина окна ему не важна.
        /// Чистая — под тест.
        /// </summary>
        internal static WatchdogStop Decide(long stepMs, long dialogMs, long cancelMs, Limits limits)
        {
            if (cancelMs >= 0 && cancelMs >= limits.CancelMs)
                return WatchdogStop.Cancel;
            if (dialogMs >= 0 && dialogMs >= limits.DialogMs)
                return WatchdogStop.Dialog;
            if (limits.StepMs > 0 && stepMs >= limits.StepMs)
                return WatchdogStop.Timeout;
            return WatchdogStop.None;
        }

        /// <summary>Один опрос (поток сторожа; тесты зовут напрямую с подставным временем).</summary>
        internal void Poll(long now)
        {
            lock (_gate)
            {
                if (!_armed || _stop != WatchdogStop.None || _target == null)
                    return;
                string title = _target.ModalWindowTitle();
                if (title == null)
                    _dialogSince = -1;
                else if (_dialogSince < 0)
                    _dialogSince = now;
                bool cancel = _cancellable && _cancelRequested != null && _cancelRequested();
                if (!cancel)
                    _cancelSince = -1;
                else if (_cancelSince < 0)
                    _cancelSince = now;
                WatchdogStop stop = Decide(now - _stepStart, Since(_dialogSince, now), Since(_cancelSince, now), _limits);
                // Человек завладел экземпляром — не останавливаем, смотрим дальше: его окно
                // закроется, а наш вызов тогда вернётся сам или будет отклонён фильтром сообщений.
                if (stop == WatchdogStop.None || !_target.TryStop())
                    return;
                _stop = stop;
                _dialogTitle = stop == WatchdogStop.Dialog ? title : null;
            }
        }

        private static long Since(long since, long now)
        {
            return since < 0 ? -1 : now - since;
        }

        private void Loop()
        {
            while (!_disposed.WaitOne(_limits.PollMs))
            {
                try
                {
                    Poll(_clock());
                }
                catch
                {
                    // Сбой опроса (перебор окон, доступ к процессу) не должен ронять приложение:
                    // необработанное исключение фонового потока завершило бы весь процесс.
                }
            }
        }

        /// <summary>
        /// Понятное сообщение об остановке процесса app («Excel», «Word») — по Dialog и Timeout;
        /// для None и Cancel — null (отмену вызывающий показывает как отмену).
        /// </summary>
        public string Describe(string app)
        {
            WatchdogStop stop = Stop;
            if (stop == WatchdogStop.Dialog)
            {
                string title = DialogTitle;
                return string.IsNullOrEmpty(title)
                    ? string.Format(Loc.T("err.office.dialogUntitled"), app)
                    : string.Format(Loc.T("err.office.dialog"), app, title);
            }
            if (stop == WatchdogStop.Timeout)
                return string.Format(Loc.T("err.office.timeout"), app, LimitMinutes(_limits.StepMs));
            return null;
        }

        /// <summary>Предел в минутах для текста: с округлением вверх, не меньше одной. Чистая — под тест.</summary>
        internal static int LimitMinutes(int ms)
        {
            return Math.Max(1, (int)((ms + 59999L) / 60000));
        }

        /// <summary>Остановить поток сторожа и дождаться его (опрос конечен: перебор окон и Kill с ожиданием 2 с).</summary>
        public void Dispose()
        {
            lock (_gate)
            {
                if (_isDisposed)
                    return;
                _isDisposed = true;
                _armed = false;
            }
            if (_thread == null)
                return;
            _disposed.Set();
            _thread.Join();
            _disposed.Close();
        }
    }
}
