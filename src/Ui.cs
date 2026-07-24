using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Общие фабрики элементов интерфейса (главное окно и «О программе»).</summary>
    internal static class Ui
    {
        /// <summary>
        /// Включить двойную буферизацию ListView (убирает мерцание при добавлении строк);
        /// свойство защищённое — только через reflection. Общее для всех списков (DRY).
        /// </summary>
        public static void EnableDoubleBuffer(ListView list)
        {
            PropertyInfo p = typeof(ListView).GetProperty("DoubleBuffered",
                BindingFlags.Instance | BindingFlags.NonPublic);
            if (p != null)
                p.SetValue(list, true, null);
        }

        /// <summary>
        /// Кнопка «Главная» в правом углу брендовой шапки (единая для всех инструментов, DRY).
        /// showHub == null (запуск вне хаба, напр. смоук-тест) — кнопки нет.
        /// </summary>
        public static void HomeOnHeader(HeaderBand header, Action showHub, ToolTip tips, int top)
        {
            if (showHub == null)
                return;
            Button home = HomeButton(showHub);
            home.SetBounds(header.Width - 180, top, 160, 30);
            home.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            if (tips != null)
                tips.SetToolTip(home, Loc.T("common.homeTip"));
            header.Controls.Add(home);
        }

        public static Label Label(Control parent, string text, int x, int y, Font font, Color color)
        {
            var l = new Label();
            l.Text = text;
            l.Location = new Point(x, y);
            l.AutoSize = true;
            l.Font = font;
            l.ForeColor = color;
            l.BackColor = Color.Transparent;
            parent.Controls.Add(l);
            return l;
        }

        public static LinkLabel Link(Control parent, string text, int x, int y)
        {
            var l = new LinkLabel();
            l.Text = text;
            l.Location = new Point(x, y);
            l.AutoSize = true;
            l.LinkColor = Theme.Accent;
            l.ActiveLinkColor = Theme.AccentPressed;
            parent.Controls.Add(l);
            return l;
        }

        /// <summary>Ссылка, открывающая URL в браузере по умолчанию.</summary>
        public static LinkLabel UrlLink(Control parent, string text, int x, int y, string url)
        {
            LinkLabel l = Link(parent, text, x, y);
            l.LinkClicked += delegate
            {
                try { Process.Start(url); }
                catch { } // нет браузера/ассоциации — молча, ссылку видно текстом
            };
            return l;
        }

        /// <summary>Стилизованная кнопка «Главная» — показывает экран выбора инструмента.</summary>
        public static Button HomeButton(Action showHub)
        {
            var b = new RoundedButton(false);
            b.Text = Loc.T("common.home");
            b.Click += delegate { showHub(); };
            return b;
        }

        /// <summary>Иконка приложения из exe (или null, если недоступна). Общая для всех окон.</summary>
        public static Icon AppIcon()
        {
            try { return Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
            catch { return null; } // без иконки — со стандартной системной
        }

        /// <summary>Акцентная полоса заданного цвета в верхней части окна.</summary>
        public static Panel AccentBar(Control parent, int y, Color color)
        {
            var bar = new Panel();
            bar.SetBounds(0, y, parent.ClientSize.Width, 3);
            bar.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            bar.BackColor = color;
            parent.Controls.Add(bar);
            return bar;
        }
    }
}
