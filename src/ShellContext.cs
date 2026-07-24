using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Владелец жизненного цикла приложения. Держит хаб выбора и открытые
    /// инструменты как независимые окна: закрытие хаба НЕ закрывает инструменты,
    /// а процесс завершается, только когда закрыто последнее окно. Кнопка
    /// «Главная» в инструменте снова показывает хаб (пересоздаёт, если был закрыт).
    /// Смена языка (<see cref="Loc.Changed"/>) пересобирает открытые окна на новом
    /// языке на тех же местах. Только UI-поток.
    /// </summary>
    internal sealed class ShellContext : ApplicationContext
    {
        private const string AppTitle = "iwo Helper Desktop";
        private readonly List<Form> _windows = new List<Form>();
        private readonly ToolRegistry _tools = new ToolRegistry();
        // Фабрики открытых инструментов (ключ → имя + как пересоздать) — для пересборки при смене языка.
        private readonly Dictionary<string, ToolEntry> _openTools = new Dictionary<string, ToolEntry>();
        private StartForm _hub;
        private bool _rebuilding; // идёт пересборка окон при смене языка — не завершать приложение

        internal sealed class ToolEntry
        {
            public string Name;
            public Func<Action, Form> Factory;
        }

        public ShellContext()
        {
            Loc.Changed += OnLanguageChanged;
            ShowHub();
        }

        /// <summary>Показать экран выбора инструмента; создать заново, если был закрыт.</summary>
        public void ShowHub()
        {
            if (_hub != null && !_hub.IsDisposed)
            {
                BringToFront(_hub);
                return;
            }
            _hub = new StartForm(this);
            Track(_hub);
            _hub.Show();
            BringToFront(_hub);
        }

        /// <summary>
        /// Открыть инструмент по ключу немодально. Повторное открытие того же —
        /// уведомление и фокус на уже открытое окно.
        /// </summary>
        public void OpenTool(string key, string name, Func<Action, Form> factory)
        {
            Form existing;
            if (_tools.TryGetOpen(key, out existing))
            {
                Dialogs.Info(_hub, AppTitle, Loc.T("shell.toolOpen.title"),
                    string.Format(Loc.T("shell.toolOpen.body"), name));
                BringToFront(existing);
                return;
            }
            SpawnTool(key, name, factory, null, FormWindowState.Normal);
        }

        /// <summary>Создать окно инструмента через фабрику, отследить и (опционально) поставить на место.</summary>
        private void SpawnTool(string key, string name, Func<Action, Form> factory, Point? location, FormWindowState state)
        {
            Form tool = factory(ShowHub); // кнопка «Главная» в инструменте показывает хаб
            _tools.Add(key, tool);
            _openTools[key] = new ToolEntry { Name = name, Factory = factory };
            Track(tool);
            tool.FormClosed += delegate { _tools.Remove(key); _openTools.Remove(key); };
            tool.Show(); // отдельное окно, без владельца — переживает закрытие хаба
            if (state == FormWindowState.Normal && location.HasValue)
                tool.Location = location.Value; // вернуть на прежнее место при пересборке
            else if (state != FormWindowState.Normal)
                tool.WindowState = state;
            BringToFront(tool);
        }

        /// <summary>
        /// Смена языка пришла из клика по пункту меню окна, которое мы будем закрывать.
        /// Пересоздавать это окно ПРЯМО в обработчике его же меню нельзя — пункт меню
        /// разрушится под собой (краш после возврата из обработчика). Поэтому откладываем
        /// пересборку на «после текущего события» через BeginInvoke. Если маршалить некуда
        /// (нет окон с хэндлом) — выполняем сразу.
        /// </summary>
        private void OnLanguageChanged()
        {
            Control ctx = MarshalTarget();
            if (ctx != null)
                ctx.BeginInvoke((MethodInvoker)RebuildOpenWindows);
            else
                RebuildOpenWindows();
        }

        /// <summary>Окно с созданным хэндлом для маршалинга отложенной пересборки (хаб или любой инструмент).</summary>
        private Control MarshalTarget()
        {
            if (_hub != null && !_hub.IsDisposed && _hub.IsHandleCreated)
                return _hub;
            foreach (Form f in _windows)
                if (f != null && !f.IsDisposed && f.IsHandleCreated)
                    return f;
            return null;
        }

        /// <summary>
        /// Пересобрать открытые окна на новом языке, сохранив их положение И активное окно.
        /// Порядок: неактивные инструменты → хаб → АКТИВНОЕ окно последним — последнее
        /// показанное остаётся сверху, поэтому пользователь после смены языка остаётся в том
        /// окне, из меню которого её выбрал (иначе хаб выпрыгивал поверх). Занятые инструменты
        /// (идёт операция) не закрываются — останутся в прежнем языке и обновятся при следующем
        /// открытии; занятое активное окно в конце поднимается наверх как есть. Флаг _rebuilding
        /// не даёт приложению завершиться, пока окна на миг закрыты между «закрыть» и «открыть».
        /// </summary>
        private void RebuildOpenWindows()
        {
            _rebuilding = true;
            try
            {
                Form active = Form.ActiveForm; // окно, из меню которого сменили язык

                // Снимок открытых инструментов ДО закрытия (закрытие чистит словари).
                var snap = new List<ToolSnapshot>();
                foreach (KeyValuePair<string, ToolEntry> kv in _openTools)
                {
                    Form f;
                    if (_tools.TryGetOpen(kv.Key, out f) && f != null && !f.IsDisposed)
                        snap.Add(new ToolSnapshot
                        {
                            Key = kv.Key,
                            Entry = kv.Value,
                            Location = f.Location,
                            State = f.WindowState,
                            WasActive = ReferenceEquals(f, active)
                        });
                }
                MoveActiveLast(snap);
                bool hubOpen = _hub != null && !_hub.IsDisposed;
                bool hubWasActive = hubOpen && ReferenceEquals(_hub, active);
                Point hubLoc = hubOpen ? _hub.Location : Point.Empty;

                // Хаб пересоздаётся ДО активного инструмента (и после неактивных); если активен
                // был сам хаб — он в самом конце, как раньше.
                bool hubRebuilt = false;
                Form busyActive = null; // активное занятое окно не пересоздаётся — поднимем его
                foreach (ToolSnapshot t in snap)
                {
                    if (hubOpen && !hubRebuilt && !hubWasActive && t.WasActive)
                    {
                        RebuildHub(hubLoc);
                        hubRebuilt = true;
                    }
                    Form old;
                    if (_tools.TryGetOpen(t.Key, out old) && old != null && !old.IsDisposed)
                    {
                        // Занятое окно не трогаем вовсе: его Close() показал бы модальный вопрос
                        // «Прервать операцию?» посреди смены языка (а «Да» отменил бы работу).
                        var busy = old as IBusyAware;
                        if (busy != null && busy.IsBusy)
                        {
                            if (t.WasActive)
                                busyActive = old; // остаётся на прежнем языке, но наверху
                            continue;
                        }
                        old.Close();
                        if (!old.IsDisposed)
                        {
                            if (t.WasActive)
                                busyActive = old; // закрытие отменено — оставляем и поднимаем
                            continue;
                        }
                    }
                    SpawnTool(t.Key, t.Entry.Name, t.Entry.Factory, t.Location, t.State);
                }

                if (hubOpen && !hubRebuilt)
                    RebuildHub(hubLoc);
                BringToFront(busyActive); // null-безопасно: активное непересозданное окно — наверх
            }
            finally
            {
                _rebuilding = false;
            }
            if (_windows.Count == 0)
                ExitThread(); // страховка: если пересобирать было нечего
        }

        /// <summary>Пересоздать хаб на прежнем месте (часть пересборки при смене языка).</summary>
        private void RebuildHub(Point hubLoc)
        {
            StartForm oldHub = _hub;
            _hub = null;
            if (oldHub != null && !oldHub.IsDisposed)
                oldHub.Close();
            ShowHub();
            if (_hub != null && !hubLoc.IsEmpty)
                _hub.Location = hubLoc;
        }

        /// <summary>
        /// Активный снапшот — в конец списка (последнее показанное окно остаётся сверху);
        /// порядок остальных сохраняется. Чистая — под тест.
        /// </summary>
        internal static void MoveActiveLast(List<ToolSnapshot> snap)
        {
            for (int i = 0; i < snap.Count; i++)
            {
                if (!snap[i].WasActive)
                    continue;
                ToolSnapshot active = snap[i];
                snap.RemoveAt(i);
                snap.Add(active);
                return; // активное окно одно
            }
        }

        internal sealed class ToolSnapshot
        {
            public string Key;
            public ToolEntry Entry;
            public Point Location;
            public FormWindowState State;
            public bool WasActive;
        }

        private void Track(Form form)
        {
            _windows.Add(form);
            form.FormClosed += OnWindowClosed;
        }

        private void OnWindowClosed(object sender, FormClosedEventArgs e)
        {
            var form = (Form)sender;
            form.FormClosed -= OnWindowClosed;
            _windows.Remove(form);
            if (form == _hub)
                _hub = null; // хаб закрыт — инструменты продолжают работать
            if (_windows.Count == 0 && !_rebuilding)
                ExitThread(); // закрыто последнее окно — завершаем приложение (не во время пересборки)
        }

        private static void BringToFront(Form form)
        {
            if (form == null || form.IsDisposed)
                return;
            if (form.WindowState == FormWindowState.Minimized)
                form.WindowState = FormWindowState.Normal;
            form.Activate();
            form.BringToFront();
        }
    }
}
