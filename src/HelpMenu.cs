using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Единое меню окон-инструментов «☰ Меню»: «Как пользоваться» (F1), «Статистика», выбор
    /// языка и произвольные дополнительные пункты. «О программе» вынесена на стартовый экран.
    /// По гайдлайнам Windows это последний (и единственный) пункт строки меню; многоточия у
    /// команд нет (окна не требуют дополнительного ввода). Тексты — из <see cref="Loc"/>.
    /// </summary>
    internal static class HelpMenu
    {
        public const int Height = 28;

        /// <summary>
        /// Меню рисуются шрифтом ПРИЛОЖЕНИЯ, а не системным шрифтом меню. Причина простая:
        /// полоса меню имеет фиксированную высоту (<see cref="Height"/>, её масштабирует DPI),
        /// а системный размер шрифта задаётся ОТДЕЛЬНОЙ настройкой Windows — «размер текста» в
        /// специальных возможностях. С крупным системным шрифтом подписи в полосе обрезались
        /// бы снизу, и это единственное место, где системный шрифт вообще доставал до нашей
        /// раскладки: все окна свой шрифт задают явно.
        /// </summary>
        private static Font AppFont { get { return Ui.Font(9.75f); } }

        public static MenuStrip Create(Form owner, Action showHowTo, params ToolStripMenuItem[] extras)
        {
            var menu = new MenuStrip();
            menu.Font = AppFont; // не системный шрифт меню — см. AppFont
            menu.AutoSize = false;
            menu.Height = Height;
            menu.Dock = DockStyle.Top;
            menu.BackColor = Color.White;
            menu.Padding = new Padding(12, 4, 0, 0);

            var root = new ToolStripMenuItem(Loc.T("menu.root"));

            var howTo = new ToolStripMenuItem(Loc.T("menu.howTo"));
            howTo.ShortcutKeys = Keys.F1;
            howTo.Click += delegate { showHowTo(); };
            root.DropDownItems.Add(howTo);

            // «Статистика» живёт внутри «Настроек» — отдельным пунктом её здесь больше нет:
            // два входа в одно окно приходится помнить в двух местах при каждой правке.
            var settings = new ToolStripMenuItem(Loc.T("settings.title"));
            settings.Click += delegate { using (var form = new SettingsForm()) form.ShowDialog(owner); };
            root.DropDownItems.Add(settings);

            root.DropDownItems.Add(new ToolStripSeparator());
            root.DropDownItems.Add(BuildLanguageMenu());

            if (extras != null && extras.Length > 0)
            {
                root.DropDownItems.Add(new ToolStripSeparator());
                foreach (ToolStripMenuItem item in extras)
                    root.DropDownItems.Add(item);
            }

            menu.Items.Add(root);
            return menu;
        }

        /// <summary>Подменю «Язык / Language» с отметкой текущего языка; смена зовёт <see cref="Loc.Set"/>.</summary>
        internal static ToolStripMenuItem BuildLanguageMenu()
        {
            var lang = new ToolStripMenuItem(Loc.T("menu.language"));
            lang.DropDownItems.Add(LangItem(Loc.T("menu.lang.ru"), Lang.Ru));
            lang.DropDownItems.Add(LangItem(Loc.T("menu.lang.en"), Lang.En));
            return lang;
        }

        /// <summary>Плоское контекстное меню выбора языка (для кнопки-глобуса на главной). DRY с подменю.</summary>
        internal static ContextMenuStrip LanguageContextMenu()
        {
            var menu = new ContextMenuStrip();
            menu.Font = AppFont; // как и полоса меню — своим шрифтом
            menu.Items.Add(LangItem(Loc.T("menu.lang.ru"), Lang.Ru));
            menu.Items.Add(LangItem(Loc.T("menu.lang.en"), Lang.En));
            return menu;
        }

        private static ToolStripMenuItem LangItem(string text, Lang lang)
        {
            var item = new ToolStripMenuItem(text);
            item.Image = Flags.For(lang);            // флаг страны перед кодом языка
            item.ImageScaling = ToolStripItemImageScaling.None; // показать флаг 24×16 без сжатия до 16
            item.Checked = Loc.Current == lang;
            item.Click += delegate { Loc.Set(lang); };
            return item;
        }
    }
}
