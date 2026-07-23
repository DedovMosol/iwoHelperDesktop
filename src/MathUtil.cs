using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>Общая числовая мелочь конвейера «PDF → Word» (DRY: медиана жила в трёх копиях).</summary>
    internal static class MathUtil
    {
        /// <summary>Нижняя медиана (при чётном числе — меньший из средних). Пустой список — 0. Вход не меняется.</summary>
        public static double Median(List<double> values)
        {
            if (values == null || values.Count == 0)
                return 0;
            var copy = new List<double>(values);
            copy.Sort();
            return copy[(copy.Count - 1) / 2];
        }
    }
}
