using System;
using System.Runtime.InteropServices;

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
        /// Готовит значение в формате Value2 (скаляр или двумерный массив области) к
        /// обратной записи в ячейки: строки экранируются, коды ошибок заворачиваются в
        /// <see cref="ErrorWrapper"/>. Числа, даты-серии, булевы и null проходят без изменений.
        /// Контракт — именно Value2: числа в нём всегда Double, а Int32 означает ошибку
        /// (#Н/Д, #ДЕЛ/0!…). Проверено на живом Excel: записанный обратно Int32 становился
        /// числом -2146826246 вместо #Н/Д, а ErrorWrapper даёт настоящее значение-ошибку.
        /// Поэтому целые номера в собственных массивах вызывающие передают как double.
        /// </summary>
        public static object EscapeValues(object value)
        {
            object scalar;
            if (TryEscapeScalar(value, out scalar))
                return scalar;

            var array = value as object[,];
            if (array != null)
            {
                int rowLo = array.GetLowerBound(0), rowHi = array.GetUpperBound(0);
                int colLo = array.GetLowerBound(1), colHi = array.GetUpperBound(1);
                for (int r = rowLo; r <= rowHi; r++)
                {
                    for (int c = colLo; c <= colHi; c++)
                    {
                        object cell;
                        if (TryEscapeScalar(array[r, c], out cell))
                            array[r, c] = cell;
                    }
                }
                return array;
            }

            return value;
        }

        /// <summary>Строка или код ошибки — заменяемое значение; прочее остаётся как есть.</summary>
        private static bool TryEscapeScalar(object value, out object escaped)
        {
            var s = value as string;
            if (s != null)
            {
                escaped = EscapeForEntry(s);
                return true;
            }
            if (value is int)
            {
                escaped = new ErrorWrapper((int)value); // VT_ERROR: ячейка снова получает ошибку, а не число
                return true;
            }
            escaped = null;
            return false;
        }
    }
}
