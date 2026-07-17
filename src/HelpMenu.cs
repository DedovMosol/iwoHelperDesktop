using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Единое меню «Справка» для окон инструментов: «Как пользоваться» (F1),
    /// произвольные дополнительные пункты и «О программе». По гайдлайнам Windows
    /// «Справка» — последний пункт строки меню; многоточия у команд нет (окна
    /// не требуют дополнительного ввода).
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

            var help = new ToolStripMenuItem("Справка");

            var howTo = new ToolStripMenuItem("Как пользоваться");
            howTo.ShortcutKeys = Keys.F1;
            howTo.Click += delegate { showHowTo(); };
            help.DropDownItems.Add(howTo);

            if (extras != null && extras.Length > 0)
            {
                help.DropDownItems.Add(new ToolStripSeparator());
                foreach (ToolStripMenuItem item in extras)
                    help.DropDownItems.Add(item);
            }

            help.DropDownItems.Add(new ToolStripSeparator());
            var about = new ToolStripMenuItem("О программе");
            about.Click += delegate
            {
                using (var form = new AboutForm())
                    form.ShowDialog(owner);
            };
            help.DropDownItems.Add(about);

            menu.Items.Add(help);
            return menu;
        }
    }
}
