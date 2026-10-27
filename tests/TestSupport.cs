using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Общие помощники тестов окон (TS-13): ответ на модальные сообщения и текст окна. Раньше
    /// каждый файл тестов заводил свои копии.
    /// </summary>
    internal static partial class UnitTests
    {
        /// <summary>
        /// Отвечать result на модальные окна сообщений (MessageForm), пока их не наберётся max;
        /// seen получает весь текст окна. Возвращает таймер — его Dispose снимает ответчика.
        /// Без ответчика окно ошибки, открытое кодом под тестом, держало бы прогон до сторожа.
        /// </summary>
        private static Timer AnswerMessages(DialogResult result, int max, Action<string> seen)
        {
            int answered = 0;
            var timer = new Timer { Interval = 50 };
            timer.Tick += delegate
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (!(form is MessageForm) || !form.Visible || form.DialogResult != DialogResult.None)
                        continue;
                    if (seen != null)
                        seen(ControlText(form));
                    form.DialogResult = result;
                    if (++answered >= max)
                        timer.Stop();
                    return; // коллекция окон меняется при закрытии — следующее на следующем тике
                }
            };
            timer.Start();
            return timer;
        }

        /// <summary>Текст окна и всех его элементов — для проверок «что увидел человек».</summary>
        private static string ControlText(Control root)
        {
            var text = new StringBuilder();
            var stack = new Stack<Control>();
            stack.Push(root);
            while (stack.Count > 0)
            {
                Control c = stack.Pop();
                if (!string.IsNullOrEmpty(c.Text))
                    text.Append(c.Text).Append('\n');
                foreach (Control child in c.Controls)
                    stack.Push(child);
            }
            return text.ToString();
        }
    }
}
