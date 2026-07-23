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

        public static MenuStrip Create(Form owner, Action showHowTo, params ToolStripMenuItem[] extras)
        {
            var menu = new MenuStrip();
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

            var stats = new ToolStripMenuItem(Loc.T("menu.stats"));
            stats.Click += delegate { using (var form = new StatsForm()) form.ShowDialog(owner); };
            root.DropDownItems.Add(stats);

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
