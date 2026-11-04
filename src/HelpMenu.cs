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
        /// <summary>Полосы меню нет: пункты на шапке. Ноль, чтобы раскладка окон не оставляла дыру.</summary>
        public const int Height = 0;

        /// <summary>
        /// Меню рисуются шрифтом ПРИЛОЖЕНИЯ, а не системным шрифтом меню: системный кегль —
        /// отдельная настройка Windows, и с ней подписи в меню обрезались бы. Все окна свой
        /// шрифт задают явно.
        /// </summary>
        private static Font AppFont { get { return Ui.Font(9.75f); } }

        public static ContextMenuStrip Create(Form owner, Action showHowTo, params ToolStripMenuItem[] extras)
        {
            var menu = new ContextMenuStrip();
            menu.Font = AppFont;
            var howTo = new ToolStripMenuItem(Loc.T("menu.howTo"));
            howTo.ShortcutKeyDisplayString = "F1";
            howTo.Click += delegate { showHowTo(); };
            menu.Items.Add(howTo);

            var settings = new ToolStripMenuItem(Loc.T("settings.title"));
            settings.Click += delegate { using (var form = new SettingsForm()) form.ShowDialog(owner); };
            menu.Items.Add(settings);

            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(BuildLanguageMenu());

            if (extras != null && extras.Length > 0)
            {
                menu.Items.Add(new ToolStripSeparator());
                foreach (ToolStripMenuItem item in extras)
                    menu.Items.Add(item);
            }

            BindF1(owner, showHowTo);
            if (owner != null)
                owner.Disposed += delegate { try { menu.Dispose(); } catch { } };
            return menu;
        }

        internal static void BindF1(Form form, Action showHowTo)
        {
            if (form == null || showHowTo == null)
                return;
            form.KeyPreview = true;
            form.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode != Keys.F1 || e.Alt || e.Control || e.Shift)
                    return;
                e.Handled = true;
                showHowTo();
            };
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
            item.Click += delegate
            {
                if (Loc.Set(lang))
                    return;
                ToolStrip parent = item.GetCurrentParent();
                Form owner = parent == null ? null : parent.FindForm();
                Dialogs.Error(owner, Loc.T("settings.title"),
                    Loc.T("settings.err.save.title"), Loc.T("settings.err.save.body"));
            };
            return item;
        }
    }
}
