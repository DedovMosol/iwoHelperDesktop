using System;

namespace ExcelMerger
{
    /// <summary>
    /// Кооперативная отмена длинных операций: сервисы проверяют предикат между единицами
    /// работы (страница/часть/файл) и бросают <see cref="OperationCanceledException"/>.
    /// Предикат обычно читает volatile-флаг, выставляемый кнопкой «Отмена» на UI-потоке;
    /// чтение bool из другого потока безопасно (без рвущихся чтений). null — без отмены.
    /// </summary>
    internal static class Cancellation
    {
        public static void ThrowIf(Func<bool> cancelled)
        {
            if (cancelled != null && cancelled())
                throw new OperationCanceledException();
        }
    }
}
