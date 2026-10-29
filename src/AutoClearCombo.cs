using System;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Выбор давности автоочистки: выключена, ежедневно, 7 и 30 дней. Один на два окна —
    /// историю операций (Настройки) и статистику: раньше каждое держало свою копию дней и
    /// подписей, и связь «пункт списка — дни» жила в двух местах (UI-21).
    /// </summary>
    internal sealed class AutoClearCombo : ComboBox
    {
        private static readonly int[] DayChoices = { 0, 1, 7, 30 };

        public AutoClearCombo()
        {
            DropDownStyle = ComboBoxStyle.DropDownList;
            Items.AddRange(new object[]
            {
                Loc.T("stats.auto.off"), Loc.T("stats.auto.daily"),
                Loc.T("stats.auto.7days"), Loc.T("stats.auto.30days")
            });
        }

        /// <summary>
        /// Дни выбранного пункта. Давность, которой нет в списке (файл правили руками),
        /// показывается как «выключена» — так же, как прежде в обоих окнах.
        /// </summary>
        public int Days
        {
            get { return SelectedIndex < 0 ? 0 : DayChoices[SelectedIndex]; }
            set
            {
                int index = Array.IndexOf(DayChoices, value);
                SelectedIndex = index >= 0 ? index : 0;
            }
        }
    }
}
