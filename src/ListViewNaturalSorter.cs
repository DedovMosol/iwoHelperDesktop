using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Сортировка ListView кликом по заголовку: естественное сравнение текста
    /// колонки (как в Проводнике), повторный клик меняет направление,
    /// направление показывается системной стрелкой в заголовке.
    /// </summary>
    internal sealed class ListViewNaturalSorter : IComparer
    {
        private readonly ListView _list;
        private int _column = -1;
        private bool _descending;

        public ListViewNaturalSorter(ListView list)
        {
            _list = list;
        }

        /// <summary>Клик по заголовку: сортирует по колонке или меняет направление.</summary>
        public void SortBy(int column)
        {
            _descending = column == _column && !_descending;
            _column = column;
            _list.ListViewItemSorter = this;
            _list.Sort();
            ShowArrow();
        }

        /// <summary>Сброс перед новым прогоном: строки идут в порядке обработки.</summary>
        public void Reset()
        {
            _column = -1;
            _descending = false;
            _list.ListViewItemSorter = null;
            ShowArrow();
        }

        int IComparer.Compare(object x, object y)
        {
            if (_column < 0)
                return 0;
            string a = CellText((ListViewItem)x);
            string b = CellText((ListViewItem)y);
            int result = NaturalStringComparer.Instance.Compare(a, b);
            return _descending ? -result : result;
        }

        private string CellText(ListViewItem item)
        {
            return _column < item.SubItems.Count ? item.SubItems[_column].Text : "";
        }

        // ---------- системная стрелка направления в заголовке ----------

        private const int LVM_GETHEADER = 0x101F;
        private const int HDM_GETITEM = 0x120B;  // HDM_GETITEMW
        private const int HDM_SETITEM = 0x120C;  // HDM_SETITEMW
        private const int HDI_FORMAT = 0x0004;
        private const int HDF_SORTUP = 0x0400;
        private const int HDF_SORTDOWN = 0x0200;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct HDITEM
        {
            public int mask;
            public int cxy;
            public IntPtr pszText;
            public IntPtr hbm;
            public int cchTextMax;
            public int fmt;
            public IntPtr lParam;
            public int iImage;
            public int iOrder;
            public int type;
            public IntPtr pvFilter;
            public int state;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr wParam, ref HDITEM lParam);

        private void ShowArrow()
        {
            try
            {
                IntPtr header = SendMessage(_list.Handle, LVM_GETHEADER, IntPtr.Zero, IntPtr.Zero);
                if (header == IntPtr.Zero)
                    return;
                for (int i = 0; i < _list.Columns.Count; i++)
                {
                    var item = new HDITEM();
                    item.mask = HDI_FORMAT;
                    if (SendMessage(header, HDM_GETITEM, (IntPtr)i, ref item) == IntPtr.Zero)
                        continue;
                    item.fmt &= ~(HDF_SORTUP | HDF_SORTDOWN);
                    if (i == _column)
                        item.fmt |= _descending ? HDF_SORTDOWN : HDF_SORTUP;
                    SendMessage(header, HDM_SETITEM, (IntPtr)i, ref item);
                }
            }
            catch { } // стрелка — украшение, не причина падать
        }
    }
}
