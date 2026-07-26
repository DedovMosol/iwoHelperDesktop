using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Reflection;
using System.Threading;
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

        /// <summary>
        /// Подпись переменной длины (путь, текст ошибки, строка статуса) — на фиксированную
        /// ширину с многоточием вместо AutoSize. Иначе длинный текст уезжает за край окна,
        /// обрезанный посреди слова и без всякого признака, что он обрезан; полный текст
        /// показывает всплывающая подсказка, которую даёт AutoEllipsis. Единая обвязка (DRY).
        /// </summary>
        public static Label Ellipsize(Label label, int width, AnchorStyles anchor)
        {
            label.AutoSize = false;
            label.AutoEllipsis = true;
            label.SetBounds(label.Left, label.Top, width, 20);
            label.Anchor = anchor;
            return label;
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

        /// <summary>
        /// Открыть файл (или папку — asFolder) в проводнике/ShellExecute. Единая точка
        /// авто-открытия результата операции (DRY): PDF-инструменты больше не зовут Process.Start
        /// врозь с разным экранированием. Молча — нет ассоциации/проводника: файл всё равно создан.
        /// Явное «Открыть…» с диалогом об ошибке — отдельно (см. MainForm.OpenPath).
        /// </summary>
        public static void OpenPath(string path, bool asFolder = false)
        {
            if (string.IsNullOrEmpty(path))
                return;
            try
            {
                if (asFolder)
                    using (Process.Start("explorer.exe", "\"" + path + "\"")) { }
                else
                    using (Process.Start(path)) { }
            }
            catch { } // нет ассоциации/проводника — молча
        }

        /// <summary>Ссылка, открывающая URL в браузере по умолчанию.</summary>
        public static LinkLabel UrlLink(Control parent, string text, int x, int y, string url)
        {
            LinkLabel l = Link(parent, text, x, y);
            l.LinkClicked += delegate
            {
                try { using (Process.Start(url)) { } }
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

        /// <summary>
        /// Отправить брендовую шапку в КОНЕЦ обхода Tab, чтобы при открытии окна фокус
        /// достался рабочему контролу, а не кнопке «Главная» (иначе первый же Enter
        /// возвращает пользователя в меню, ничего не сделав).
        ///
        /// Задать шапке большой TabIndex ПРИ ДОБАВЛЕНИИ нельзя: ControlCollection.Add
        /// выдаёт каждому новому контролу следующий свободный индекс, поэтому всё
        /// добавленное после шапки получает ещё больший — и «большой» индекс, наоборот,
        /// ставит шапку ПЕРВОЙ. Поэтому пересчитываем, когда окно уже собрано (OnLoad).
        /// </summary>
        public static void HeaderLastInTabOrder(Form form)
        {
            if (form == null)
                return;
            Control header = null;
            int max = 0;
            foreach (Control c in form.Controls)
            {
                if (c is HeaderBand)
                    header = c;
                else if (c.TabIndex > max)
                    max = c.TabIndex;
            }
            if (header != null)
                header.TabIndex = max + 1;
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

        /// <summary>
        /// Безопасная доставка действия в UI-поток контрола из фонового: BeginInvoke под
        /// guard'ом IsHandleCreated/IsDisposed с перехватом гонки «окно разрушено между
        /// проверкой и вызовом». true — делегат принят очередью UI; false — не принят
        /// (вызывающий освобождает свои ресурсы сам). Единая обвязка всех воркеров (DRY):
        /// ручные копии guard'а уже дважды теряли catch.
        /// </summary>
        public static bool OnUi(Control target, Action action)
        {
            try
            {
                if (target != null && target.IsHandleCreated && !target.IsDisposed)
                {
                    target.BeginInvoke(action);
                    return true;
                }
            }
            catch (InvalidOperationException) { } // окно разрушено между проверкой и вызовом
            return false;
        }

        /// <summary>
        /// Запустить фоновый воркер (IsBackground; sta — для Office COM). Возвращает
        /// запущенный поток. Результаты доставлять через <see cref="OnUi"/>.
        /// </summary>
        public static Thread RunWorker(Action work, bool sta = false)
        {
            var thread = new Thread(delegate() { work(); });
            if (sta)
                thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
            return thread;
        }

        /// <summary>
        /// Общий каркас модального диалога приложения: заголовок, иконка, кэшированный шрифт,
        /// белый фон, неизменяемая рамка, центр по родителю и DPI-масштабирование. Четыре
        /// диалога («О программе», сообщения, статистика, ввод номера) держали эти девять
        /// строк своей копией — и три из них по-разному: где-то забывалась иконка окна.
        /// Звать ПЕРВЫМ делом в конструкторе: масштабирование должно быть задано до
        /// добавления контролов, а размер окна каждый диалог ставит себе сам.
        /// </summary>
        public static void InitDialog(Form form, string title)
        {
            form.Text = title;
            Icon icon = AppIcon();
            if (icon != null)
                form.Icon = icon;
            form.Font = Font(9.75f); // общий кэшированный шрифт (не освобождать)
            form.BackColor = Color.White;
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.MaximizeBox = false;
            form.MinimizeBox = false;
            form.ShowInTaskbar = false;
            form.StartPosition = FormStartPosition.CenterParent;
            form.AutoScaleDimensions = new SizeF(96f, 96f);
            form.AutoScaleMode = AutoScaleMode.Dpi;
        }

        /// <summary>
        /// Контур прямоугольника со скруглёнными углами — общая примитивная фигура всех
        /// рисуемых вручную контролов (кнопка, карточка, флажок): раньше жила в трёх
        /// одинаковых копиях, и защита от вырожденного радиуса была только в одной из них.
        /// Освобождает вызывающий (все места рисования — под using).
        /// </summary>
        public static GraphicsPath RoundedRect(RectangleF r, float radius)
        {
            var p = new GraphicsPath();
            float d = Math.Max(2f, radius * 2f); // нулевой диаметр AddArc не рисует
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        /// <summary>
        /// Выделить все строки списка (одним обновлением, без мерцания). Общая для списка
        /// файлов свода и сетки страниц: обе делали это одинаково своей копией.
        /// </summary>
        public static void SelectAllItems(ListView list)
        {
            if (list == null || list.Items.Count == 0)
                return;
            list.BeginUpdate();
            foreach (ListViewItem item in list.Items)
                item.Selected = true;
            list.EndUpdate();
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
