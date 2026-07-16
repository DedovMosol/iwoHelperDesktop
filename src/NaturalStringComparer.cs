using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ExcelMerger
{
    /// <summary>
    /// Естественное сравнение строк — как сортирует Проводник Windows:
    /// «Отчет 2» меньше «Отчет 10», регистр не учитывается.
    /// Обёртка над StrCmpLogicalW (shlwapi.dll, есть в Windows начиная с XP).
    /// </summary>
    public sealed class NaturalStringComparer : IComparer<string>
    {
        public static readonly NaturalStringComparer Instance = new NaturalStringComparer();

        private NaturalStringComparer() { }

        public int Compare(string x, string y)
        {
            if (ReferenceEquals(x, y))
                return 0;
            if (x == null)
                return -1;
            if (y == null)
                return 1;
            return StrCmpLogicalW(x, y);
        }

        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        private static extern int StrCmpLogicalW(string psz1, string psz2);
    }
}
