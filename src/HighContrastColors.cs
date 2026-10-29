using System;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Высокий контраст для обычных контролов окна (UI-7). Окна задают цвета явно (белый фон,
    /// тёмно-серый текст, цветные подписи состояния), а явный цвет в WinForms главнее системного:
    /// в чёрной теме высокого контраста окно оставалось белым. Пока режим включён, каждый
    /// несистемный цвет фона, текста и ссылки сбрасывается к системному по умолчанию — и тот,
    /// что задан при постройке, и тот, что окно поставит позже (подпись состояния, новая
    /// карточка). Выключили режим — возвращаются цвета, которые окно задало последними.
    /// Системный цвет, выбранный окном нарочно (сравнение PDF берёт Window для поля), не трогается.
    /// Рисованные контролы выбирают цвета сами (<see cref="Theme.HighContrast"/>).
    /// Только UI-поток.
    /// </summary>
    internal static class HighContrastColors
    {
        /// <summary>Цвета, заданные окном, — на время высокого контраста.</summary>
        private sealed class Intended
        {
            public Color Back = Color.Empty, Fore = Color.Empty, Link = Color.Empty, ActiveLink = Color.Empty;
        }

        // Слабая таблица: запись живёт, пока жив контрол, и не держит закрытые окна в памяти.
        private static readonly ConditionalWeakTable<Control, Intended> Saved =
            new ConditionalWeakTable<Control, Intended>();

        /// <summary>
        /// Подключить окно: режим применяется, когда создаётся хэндл (всё построенное в
        /// конструкторе уже на месте), и заново при каждой смене системных цветов — включение
        /// и выключение высокого контраста присылает окнам WM_SYSCOLORCHANGE. Подписка — на
        /// события самого окна, поэтому закрытое окно ничем не удерживается.
        /// </summary>
        internal static void Attach(Form form)
        {
            form.HandleCreated += delegate { if (Theme.HighContrast) Enforce(form); };
            form.SystemColorsChanged += delegate { Apply(form); };
        }

        /// <summary>Привести окно к текущему режиму: включить системные цвета или вернуть свои.</summary>
        internal static void Apply(Control root)
        {
            if (root == null)
                return;
            if (Theme.HighContrast)
                Enforce(root);
            else
                Restore(root);
        }

        private static void Enforce(Control c)
        {
            Intended intended;
            if (!Saved.TryGetValue(c, out intended))
            {
                intended = new Intended();
                Saved.Add(c, intended);
                c.BackColorChanged += OnColorChanged;
                c.ForeColorChanged += OnColorChanged;
                c.ControlAdded += OnControlAdded;
            }
            Reset(c, intended);
            foreach (Control child in c.Controls)
                Enforce(child);
        }

        private static void Restore(Control c)
        {
            Intended intended;
            if (Saved.TryGetValue(c, out intended))
            {
                c.BackColorChanged -= OnColorChanged;
                c.ForeColorChanged -= OnColorChanged;
                c.ControlAdded -= OnControlAdded;
                Saved.Remove(c);
                if (!intended.Back.IsEmpty) c.BackColor = intended.Back;
                if (!intended.Fore.IsEmpty) c.ForeColor = intended.Fore;
                var link = c as LinkLabel;
                if (link != null)
                {
                    if (!intended.Link.IsEmpty) link.LinkColor = intended.Link;
                    if (!intended.ActiveLink.IsEmpty) link.ActiveLinkColor = intended.ActiveLink;
                }
            }
            foreach (Control child in c.Controls)
                Restore(child);
        }

        /// <summary>
        /// Запомнить несистемный цвет и сбросить его к системному по умолчанию. Родитель
        /// сбрасывается раньше детей, поэтому унаследованный цвет ребёнка к этому времени уже
        /// системный и своим не считается.
        /// </summary>
        private static void Reset(Control c, Intended intended)
        {
            if (!c.BackColor.IsSystemColor)
            {
                intended.Back = c.BackColor;
                c.ResetBackColor();
            }
            if (!c.ForeColor.IsSystemColor)
            {
                intended.Fore = c.ForeColor;
                c.ResetForeColor();
            }
            var link = c as LinkLabel;
            if (link != null)
            {
                // HotTrack — цвет гиперссылки в теме высокого контраста. Не «пусто»: пустой цвет
                // ссылки WinForms берёт из настроек Internet Explorer (синий), а не из темы.
                if (!link.LinkColor.IsSystemColor)
                {
                    intended.Link = link.LinkColor;
                    link.LinkColor = SystemColors.HotTrack;
                }
                if (!link.ActiveLinkColor.IsSystemColor)
                {
                    intended.ActiveLink = link.ActiveLinkColor;
                    link.ActiveLinkColor = SystemColors.HotTrack;
                }
            }
        }



        private static void OnColorChanged(object sender, EventArgs e)
        {
            var c = (Control)sender;
            Intended intended;
            if (Theme.HighContrast && Saved.TryGetValue(c, out intended))
                Reset(c, intended); // системный цвет после сброса сюда же — и уже ничего не меняет
        }

        private static void OnControlAdded(object sender, ControlEventArgs e)
        {
            if (Theme.HighContrast)
                Enforce(e.Control);
        }
    }
}
