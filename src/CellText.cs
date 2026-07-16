using System;

namespace ExcelMerger
{
    /// <summary>
    /// Безопасная запись текста в ячейки Excel. Проверено экспериментально:
    /// при присваивании Value2 строка «=x» превращается в формулу, а ведущий
    /// апостроф строки съедается как префикс (число- и дата-подобные строки
    /// Value2, в отличие от UI-ввода, не преобразует). Экранирование апострофом
    /// закрывает оба случая; для остальных строк оно прозрачно — ячейка получает
    /// буквально тот же текст.
    /// </summary>
    public static class CellText
    {
        /// <summary>Экранирует одну строку для ввода в ячейку.</summary>
        public static string EscapeForEntry(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;
            return "'" + value;
        }

        /// <summary>
        /// Экранирует строки внутри значения Value2 (скаляр или двумерный массив
        /// области). Числа, даты-серии, ошибки и null проходят без изменений.
        /// </summary>
        public static object EscapeValues(object value)
        {
            var s = value as string;
            if (s != null)
                return EscapeForEntry(s);

            var array = value as object[,];
            if (array != null)
            {
                int rowLo = array.GetLowerBound(0), rowHi = array.GetUpperBound(0);
                int colLo = array.GetLowerBound(1), colHi = array.GetUpperBound(1);
                for (int r = rowLo; r <= rowHi; r++)
                {
                    for (int c = colLo; c <= colHi; c++)
                    {
                        var cell = array[r, c] as string;
                        if (cell != null)
                            array[r, c] = EscapeForEntry(cell);
                    }
                }
                return array;
            }

            return value;
        }
    }
}
