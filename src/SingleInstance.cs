using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Один экземпляр приложения на сеанс пользователя. Повторный запуск ярлыка больше не
    /// поднимает второй процесс (в диспетчере задач он выглядел как ещё одно «приложение»),
    /// а просит уже работающий показать экран выбора инструмента и сразу завершается.
    ///
    /// Признак «я первый» — именованный мьютекс в сеансе (префикс Local\), поэтому у разных
    /// пользователей одной машины экземпляры независимы. Владение НЕ запрашиваем: важен сам
    /// факт существования имени, а не захват, и брошенного мьютекса после аварийного
    /// завершения не остаётся. Имя живо, пока открыт хотя бы один хэндл, поэтому объект
    /// держим в поле — иначе сборщик закрыл бы хэндл и следующий запуск счёл бы себя первым.
    ///
    /// Сигнал живому экземпляру — зарегистрированное оконное сообщение, разосланное окнам
    /// верхнего уровня. Принимает его скрытое окно <see cref="ActivationListener"/>, поэтому
    /// сигнал доходит и когда хаб закрыт, а открыт только инструмент. Message-only окно тут не
    /// годится: широковещательные сообщения до таких окон не доходят.
    ///
    /// Самостоятельность окон не страдает: инструменты и раньше жили в ОДНОМ процессе и
    /// переживали закрытие хаба (см. <see cref="ShellContext"/>), здесь это не меняется.
    /// </summary>
    internal static class SingleInstance
    {
        private const string MutexName = @"Local\iwoHelperDesktop.instance";
        private const string MessageName = "iwoHelperDesktop.ShowHub";
        private static readonly IntPtr Broadcast = new IntPtr(0xffff); // HWND_BROADCAST
        private const int ExToolWindow = 0x80; // WS_EX_TOOLWINDOW — не показывать в Alt+Tab

        private static Mutex _held; // ссылка на всю жизнь процесса: держит имя занятым

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern uint RegisterWindowMessage(string lpString);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool AllowSetForegroundWindow(int dwProcessId);

        private const int AsfwAny = -1; // ASFW_ANY — разрешить любому процессу

        /// <summary>
        /// Занять место единственного экземпляра. false — приложение уже работает, надо
        /// разбудить его через <see cref="SignalExisting"/> и выйти. Сбой проверки трактуем
        /// как «я первый»: не запустить приложение хуже, чем открыть второе окно.
        /// </summary>
        public static bool TryAcquire()
        {
            Mutex acquired;
            bool first = TryAcquire(MutexName, out acquired);
            if (first)
                _held = acquired;
            return first;
        }

        /// <summary>
        /// Занять имя: true и созданный мьютекс, если имя было свободно. Если имя уже занято,
        /// возвращает false и NULL, закрыв свой хэндл сразу — держать его незачем.
        /// Отдельная перегрузка с именем, чтобы механизм проверялся тестом на своём имени.
        /// </summary>
        internal static bool TryAcquire(string mutexName, out Mutex acquired)
        {
            acquired = null;
            try
            {
                bool createdNew;
                var mutex = new Mutex(false, mutexName, out createdNew);
                if (createdNew)
                    acquired = mutex;
                else
                    mutex.Close();
                return createdNew;
            }
            catch
            {
                return true; // нет прав на имя или иная беда — запускаемся как обычно
            }
        }

        /// <summary>
        /// Попросить работающий экземпляр показать экран выбора инструмента. Достаточно одной
        /// посылки: если тот ещё только стартует и слушателя не создал, он и так покажет хаб
        /// сам, а если работает давно — окно-слушатель на месте.
        /// </summary>
        public static void SignalExisting()
        {
            try
            {
                uint message = RegisterWindowMessage(MessageName);
                if (message == 0)
                    return;
                // Право выйти на передний план принадлежит НАМ (нас только что запустил
                // пользователь), а поднять окно должен работающий экземпляр. Без явной передачи
                // права Windows запретит ему смену переднего плана и лишь помигает кнопкой в
                // панели задач: окно развернётся, но останется позади чужих.
                AllowSetForegroundWindow(AsfwAny);
                PostMessage(Broadcast, message, IntPtr.Zero, IntPtr.Zero);
            }
            catch { } // разбудить не вышло — просто выходим, второго процесса всё равно не будет
        }

        /// <summary>Начать слушать сигнал повторного запуска. Вызывать с UI-потока, освобождать при выходе.</summary>
        public static IDisposable Listen(Action onActivate)
        {
            try { return new ActivationListener(onActivate); }
            catch { return null; } // без слушателя приложение работает, просто не всплывает по клику
        }

        /// <summary>
        /// Скрытое окно верхнего уровня, принимающее сигнал повторного запуска. Не видимо
        /// (без WS_VISIBLE) и помечено WS_EX_TOOLWINDOW, поэтому ни в панели задач, ни в
        /// Alt+Tab не появляется. Сообщения приходят на тот поток, который создал хэндл.
        /// </summary>
        private sealed class ActivationListener : NativeWindow, IDisposable
        {
            private readonly uint _message;
            private readonly Action _onActivate;

            public ActivationListener(Action onActivate)
            {
                _onActivate = onActivate;
                _message = RegisterWindowMessage(MessageName);
                var cp = new CreateParams();
                cp.Caption = MessageName;
                cp.X = 0;
                cp.Y = 0;
                cp.Width = 0;
                cp.Height = 0;
                cp.Style = 0; // без WS_VISIBLE — окно существует, но не показывается
                cp.ExStyle = ExToolWindow;
                CreateHandle(cp);
            }

            protected override void WndProc(ref Message m)
            {
                if (_message != 0 && unchecked((uint)m.Msg) == _message && _onActivate != null)
                {
                    // Обработчик поднимает окно, а ошибки внутри не должны ронять чужой
                    // широковещательный поток сообщений.
                    try { _onActivate(); }
                    catch { }
                }
                base.WndProc(ref m);
            }

            public void Dispose()
            {
                DestroyHandle();
            }
        }
    }
}
