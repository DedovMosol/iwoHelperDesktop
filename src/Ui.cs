using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Общие фабрики элементов интерфейса (главное окно и «О программе»).</summary>
    internal static class Ui
    {
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

        /// <summary>Акцентная полоса в верхней части окна.</summary>
        public static Panel AccentBar(Control parent, int y)
        {
            var bar = new Panel();
            bar.SetBounds(0, y, parent.ClientSize.Width, 3);
            bar.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            bar.BackColor = Theme.Accent;
            parent.Controls.Add(bar);
            return bar;
        }
    }
}
