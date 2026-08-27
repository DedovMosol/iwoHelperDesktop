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
        private IDisposable _activation; // скрытое окно: повторный запуск ярлыка поднимает хаб

        internal sealed class ToolEntry
        {
            public string Name;
            public Func<Action, Form> Factory;
            /// <summary>Раздел хаба, откуда инструмент открыт: туда и вернёт его «Главная».</summary>
            public HubLevel Home;
        }

        public ShellContext()
        {
            // Слушателя создаём ДО первого окна: повторный запуск, случившийся сразу за нашим,
            // должен застать его на месте. ShowHub годится обработчиком как есть — он идемпотентен
            // и поднимает уже открытый хаб вместо создания второго.
            _activation = SingleInstance.Listen(ShowHub);
            Loc.Changed += OnLanguageChanged;
            ShowHub();
        }

        /// <summary>
        /// Запустить фоновую проверку обновлений и показать окно, только если версия новее.
        /// Зовётся ИЗ ПУТИ ЗАПУСКА, а не из конструктора: конструктор создают ещё и живые
        /// тесты окон, и проверка в нём означала бы обращение в сеть на каждом прогоне —
        /// медленно, ненадёжно и с упором в лимит запросов GitHub на сборочной машине.
        /// </summary>
        public void CheckForUpdatesOnStart()
        {
            if (_hub != null && !_hub.IsDisposed)
                UpdateUi.CheckOnStart(_hub);
        }

        public void ShowWhatsNewOnStart()
        {
            if (_hub != null && !_hub.IsDisposed)
                Ui.OnUi(_hub, delegate { WhatsNewUi.ShowIfNeeded(_hub); });
        }


        public void ShowHub()
        {
            ShowHub(null, null);
        }

        /// <summary>
        /// Показать хаб сразу нужным разделом — так «Главная» из инструмента возвращает туда,
        /// откуда его открыли, а не на верхний экран: между PDF-инструментами ходят чаще всего,
        /// и лишний клик на каждый переход был бы платой за новую структуру.
        /// </summary>
        public void ShowHub(HubLevel level)
        {
            ShowHub(null, level);
        }

        /// <summary>
        /// Показать хаб. place задан только при ПЕРЕСБОРКЕ (смена языка): окно встаёт ровно
        /// туда, где стояло, и НЕ поднимается на передний план — иначе свёрнутый хаб развернул
        /// бы сам себя, а окно, из меню которого сменили язык, потеряло бы фокус.
        /// level — раздел, который нужно показать (null — какой откроется, тот и откроется).
        /// </summary>
        private void ShowHub(WindowSnapshot? place, HubLevel? level)
        {
            if (_hub != null && !_hub.IsDisposed)
            {
                if (level.HasValue)
                    _hub.ShowLevel(level.Value);
                BringToFront(_hub);
                return;
            }
            _hub = new StartForm(this, level ?? HubLevel.Main);
            Track(_hub);
            Show(_hub, place);
            if (!place.HasValue)
                BringToFront(_hub);
        }

        /// <summary>
        /// Открыть инструмент по ключу немодально. Повторное открытие того же —
        /// уведомление и фокус на уже открытое окно.
        /// </summary>
        public void OpenTool(string key, string name, Func<Action, Form> factory, HubLevel home)
        {
            Form existing;
            if (_tools.TryGetOpen(key, out existing))
            {
                Dialogs.Info(_hub, AppTitle, Loc.T("shell.toolOpen.title"),
                    string.Format(Loc.T("shell.toolOpen.body"), name));
                BringToFront(existing);
                return;
            }
            SpawnTool(key, name, factory, home, null);
        }

        /// <summary>
        /// Открыть инструмент по ключу и передать ему файлы (дроп PDF на карточку хаба):
        /// уже открытое окно получает файлы без диалога «уже открыт», новое — сразу после
        /// показа. Инструмент грузит их сам (<see cref="IFileAcceptor"/>).
        /// </summary>
        public void OpenToolWithFiles(string key, string name, Func<Action, Form> factory, string[] files,
            HubLevel home)
        {
            Form tool;
            if (!_tools.TryGetOpen(key, out tool))
                tool = SpawnTool(key, name, factory, home, null);
            else
                BringToFront(tool);
            var acceptor = tool as IFileAcceptor;
            if (acceptor != null && files != null && files.Length > 0)
                acceptor.AcceptFiles(files);
        }

        /// <summary>
        /// Создать и показать окно инструмента. snap — размещение при ПЕРЕСБОРКЕ (смена языка),
        /// иначе окно встаёт само. activate — поднимать ли его на передний план: при пересборке
        /// нельзя, потому что BringToFront разворачивает свёрнутое окно и забирает фокус, и все
        /// окна вылезали поверх работы. Наверх поднимается только бывшее активным и после цикла.
        /// </summary>
        private Form SpawnTool(string key, string name, Func<Action, Form> factory, HubLevel home,
            WindowSnapshot? snap, bool activate = true)
        {
            HubLevel section = home;
            // «Главная» в инструменте показывает хаб СРАЗУ его разделом.
            Form tool = factory(delegate { ShowHub(section); });
            _tools.Add(key, tool);
            _openTools[key] = new ToolEntry { Name = name, Factory = factory, Home = home };
            Track(tool);
            tool.FormClosed += delegate { _tools.Remove(key); _openTools.Remove(key); };
            Show(tool, snap); // отдельное окно, без владельца — переживает закрытие хаба
            if (activate)
                BringToFront(tool);
            return tool;
        }

        /// <summary>Показать окно: на прежнем месте при пересборке, иначе как оно умеет само.</summary>
        private static void Show(Form form, WindowSnapshot? snap)
        {
            if (snap.HasValue)
                WindowPlacement.ShowAt(form, snap.Value);
            else
                form.Show();
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
                            Place = WindowPlacement.Snapshot(f),
                            WasActive = ReferenceEquals(f, active)
                        });
                }
                MoveActiveLast(snap);
                bool hubOpen = _hub != null && !_hub.IsDisposed;
                bool hubWasActive = hubOpen && ReferenceEquals(_hub, active);
                WindowSnapshot hubPlace = hubOpen ? WindowPlacement.Snapshot(_hub) : new WindowSnapshot();
                // Раздел хаба переносим на новое окно: иначе смена языка выбрасывала бы
                // человека из «Инструментов PDF» на верхний экран — та же по природе потеря
                // состояния, что и уехавший за экран свёрнутый хаб в 1.17.8.
                HubLevel hubLevel = hubOpen ? _hub.Level : HubLevel.Main;

                // Хаб пересоздаётся ДО активного инструмента (и после неактивных); если активен
                // был сам хаб — он в самом конце, как раньше.
                bool hubRebuilt = false;
                Form busyActive = null; // активное непересозданное окно — поднимем его
                int deferred = 0;       // dirty/busy окна останутся на прежнем языке
                foreach (ToolSnapshot t in snap)
                {
                    if (hubOpen && !hubRebuilt && !hubWasActive && t.WasActive)
                    {
                        RebuildHub(hubPlace, hubLevel);
                        hubRebuilt = true;
                    }
                    Form old;
                    if (_tools.TryGetOpen(t.Key, out old) && old != null && !old.IsDisposed)
                    {
                        // Занятое или dirty-окно не трогаем: пересборка без state-transfer
                        // потеряла бы выполняемую работу либо набранные страницы/файлы.
                        var busy = old as IBusyAware;
                        var dirty = old as IUnsavedStateAware;
                        if ((busy != null && busy.IsBusy) ||
                            (dirty != null && dirty.HasUncommittedState))
                        {
                            deferred++;
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
                    Form rebuilt = SpawnTool(t.Key, t.Entry.Name, t.Entry.Factory, t.Entry.Home,
                        t.Place, false);
                    if (t.WasActive)
                        busyActive = rebuilt; // поднимем в конце — только это окно и было активным
                }

                if (hubOpen && !hubRebuilt)
                    RebuildHub(hubPlace, hubLevel);
                BringToFront(busyActive); // null-безопасно: активное непересозданное окно — наверх
                if (deferred > 0)
                    Dialogs.Info(busyActive ?? _hub, AppTitle, Loc.T("lang.deferred.title"),
                        string.Format(Loc.T("lang.deferred.body"), deferred));
            }
            finally
            {
                _rebuilding = false;
            }
            if (_windows.Count == 0)
                ExitThread(); // страховка: если пересобирать было нечего
        }

        /// <summary>Пересоздать хаб на прежнем месте и в прежнем состоянии (часть пересборки при смене языка).</summary>
        private void RebuildHub(WindowSnapshot place, HubLevel level)
        {
            StartForm oldHub = _hub;
            _hub = null;
            if (oldHub != null && !oldHub.IsDisposed)
                oldHub.Close();
            ShowHub(place, level);
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
            public WindowSnapshot Place;
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

        protected override void Dispose(bool disposing)
        {
            // Скрытое окно-слушатель не дочернее ничему, поэтому убираем его сами.
            if (disposing && _activation != null)
            {
                _activation.Dispose();
                _activation = null;
            }
            base.Dispose(disposing);
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
