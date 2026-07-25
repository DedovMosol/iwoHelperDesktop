using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Общие фабрики элементов интерфейса (главное окно и «О программе»).</summary>
    internal static class Ui
    {
        // Кэш общих шрифтов: WinForms НЕ владеет Control.Font, поэтому шрифт, создаваемый
        // на каждое окно/диалог, жил бы до финализатора — GDI-мусор при каждом диалоге и
        // каждой пересборке окон сменой языка. Кэш живёт весь процесс и не освобождается;
        // ключей конечное число (фиксированные литералы вызовов). Только UI-поток.
        private static readonly Dictionary<string, Font> _fonts = new Dictionary<string, Font>();

        /// <summary>Общий кэшированный шрифт (по умолчанию — Segoe UI). НЕ освобождать у получателя.</summary>
        public static Font Font(float size, FontStyle style = FontStyle.Regular, string family = "Segoe UI")
        {
            string key = family + "|" + size.ToString(CultureInfo.InvariantCulture) + "|" + (int)style;
            Font f;
            if (!_fonts.TryGetValue(key, out f))
            {
                f = new Font(family, size, style);
                _fonts[key] = f;
            }
            return f;
        }
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

        private static Icon _appIcon;
        private static bool _appIconTried;

        /// <summary>
        /// Иконка приложения из exe (или null, если недоступна). Общая для всех окон.
        /// Кэш на процесс: ExtractAssociatedIcon создаёт НОВЫЙ HICON на каждый вызов, а
        /// Form.Icon его не освобождает — без кэша каждый диалог/предпросмотр копил бы
        /// хэндлы до финализатора. Кэш не освобождается (живёт весь процесс). UI-поток.
        /// </summary>
        public static Icon AppIcon()
        {
            if (!_appIconTried)
            {
                _appIconTried = true;
                try { _appIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
                catch { _appIcon = null; } // без иконки — со стандартной системной
            }
            return _appIcon;
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
