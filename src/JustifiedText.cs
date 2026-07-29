using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Абзац, выключенный ПО ШИРИНЕ. WinForms такого выравнивания не предлагает ни у TextBox,
    /// ни у RichTextBox (в перечислении только лево, право и центр), поэтому оно выставляется
    /// формату абзаца напрямую сообщением EM_SETPARAFORMAT — тем же способом, которым элемент
    /// делает все остальные выравнивания.
    ///
    /// Вынесено из окна «О программе», когда ровный правый край понадобился и в «Настройках»:
    /// возня с форматом абзаца через сообщения — ровно то, что не стоит держать в двух копиях.
    /// </summary>
    internal static class JustifiedText
    {
        /// <summary>
        /// Создать и добавить в parent абзац с выключкой по ширине. Текст только для чтения,
        /// но выделяемый и копируемый; внешне — обычная подпись на белом, без рамки.
        /// </summary>
        public static RichTextBox Paragraph(Control parent, string text, int x, int y, int width, Color color)
        {
            var rtb = new RichTextBox();
            rtb.Multiline = true;
            rtb.WordWrap = true;
            rtb.ScrollBars = RichTextBoxScrollBars.None;
            rtb.BorderStyle = BorderStyle.None;
            rtb.BackColor = Color.White;
            rtb.ForeColor = color;
            rtb.Font = parent.Font;
            rtb.TabStop = false;
            rtb.ReadOnly = true;
            rtb.Text = text;
            rtb.SetBounds(x, y, width, Height(text, parent.Font, width));
            // Формат абзаца хранится в самом окне элемента, а WinForms окно пересоздаёт (смена
            // родителя, пересборка интерфейса) — поэтому ставим выключку на КАЖДОЕ создание
            // хэндла, иначе разовая установка в конструкторе однажды потерялась бы.
            rtb.HandleCreated += delegate { Justify(rtb); };
            parent.Controls.Add(rtb);
            return rtb;
        }

        /// <summary>
        /// Высота абзаца под заданную ширину. Меряем чуть более узкой строкой и добавляем запас:
        /// поле переносит слова по своей внутренней ширине, которая на пару пикселей меньше
        /// заданной, и без запаса последняя строка обрезалась бы.
        /// </summary>
        public static int Height(string text, Font font, int width)
        {
            const int Inset = 4; // внутренние поля поля ввода с обеих сторон
            Size measured = TextRenderer.MeasureText(text, font,
                new Size(Math.Max(width - Inset, 1), int.MaxValue), TextFormatFlags.WordBreak);
            return measured.Height + Inset;
        }

        private const int WmUser = 0x0400;
        private const int EmGetParaFormat = WmUser + 61;
        private const int EmSetParaFormat = WmUser + 71;
        private const int PfmAlignment = 0x00000008;
        private const short PfaJustify = 4;  // выключка по ширине

        [StructLayout(LayoutKind.Sequential)]
        private struct ParaFormat2
        {
            public int cbSize, dwMask;
            public short wNumbering, wReserved;
            public int dxStartIndent, dxRightIndent, dxOffset;
            public short wAlignment, cTabCount;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)]
            public int[] rgxTabs;
            public int dySpaceBefore, dySpaceAfter, dyLineSpacing;
            public short sStyle;
            public byte bLineSpacingRule, bOutlineLevel;
            public short wShadingWeight, wShadingStyle;
            public short wNumberingStart, wNumberingStyle, wNumberingTab, wBorderSpace, wBorderWidth, wBorders;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref ParaFormat2 lParam);

        /// <summary>Выключить весь текст по ширине. Сбой не критичен — останется выключка влево.</summary>
        public static void Justify(RichTextBox rtb)
        {
            try
            {
                var fmt = new ParaFormat2();
                fmt.cbSize = Marshal.SizeOf(typeof(ParaFormat2));
                fmt.dwMask = PfmAlignment;
                fmt.wAlignment = PfaJustify;
                fmt.rgxTabs = new int[32];
                // Сообщение действует на абзацы ВЫДЕЛЕНИЯ (wParam обязан быть нулём), поэтому
                // выделяем всё, применяем и снимаем выделение — на экране это не мелькает.
                rtb.SelectAll();
                SendMessage(rtb.Handle, EmSetParaFormat, IntPtr.Zero, ref fmt);
                rtb.Select(0, 0);
            }
            catch { } // выключка — оформление, а не работа: не получилось, текст всё равно читается
        }

        /// <summary>Выключен ли текст по ширине (для проверок).</summary>
        public static bool IsJustified(RichTextBox rtb)
        {
            var fmt = new ParaFormat2();
            fmt.cbSize = Marshal.SizeOf(typeof(ParaFormat2));
            fmt.rgxTabs = new int[32];
            SendMessage(rtb.Handle, EmGetParaFormat, IntPtr.Zero, ref fmt);
            return (fmt.dwMask & PfmAlignment) != 0 && fmt.wAlignment == PfaJustify;
        }
    }
}
