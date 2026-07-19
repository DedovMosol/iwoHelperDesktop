using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Владелец жизненного цикла приложения. Держит хаб выбора и открытые
    /// инструменты как независимые окна: закрытие хаба НЕ закрывает инструменты,
    /// а процесс завершается, только когда закрыто последнее окно. Кнопка
    /// «Главная» в инструменте снова показывает хаб (пересоздаёт, если был закрыт).
    /// Только UI-поток.
    /// </summary>
    internal sealed class ShellContext : ApplicationContext
    {
        private const string AppTitle = "iwo Helper Desktop";
        private readonly List<Form> _windows = new List<Form>();
        private readonly ToolRegistry _tools = new ToolRegistry();
        private StartForm _hub;

        public ShellContext()
        {
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
                Dialogs.Info(_hub, AppTitle, "Инструмент уже открыт",
                    "«" + name + "» уже запущен — открыто его окно.");
                BringToFront(existing);
                return;
            }

            Form tool = factory(ShowHub); // кнопка «Главная» в инструменте показывает хаб
            _tools.Add(key, tool);
            Track(tool);
            tool.FormClosed += delegate { _tools.Remove(key); };
            tool.Show(); // отдельное окно, без владельца — переживает закрытие хаба
            BringToFront(tool);
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
            if (_windows.Count == 0)
                ExitThread(); // закрыто последнее окно — завершаем приложение
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
