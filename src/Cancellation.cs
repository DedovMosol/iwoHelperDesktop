using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Кооперативная отмена длинных операций: сервисы проверяют предикат между единицами
    /// работы (страница/часть/файл) и бросают <see cref="OperationCanceledException"/>.
    /// Предикат обычно читает volatile-флаг, выставляемый кнопкой «Отмена» на UI-потоке;
    /// чтение bool из другого потока безопасно (без рвущихся чтений). null — без отмены.
    ///
    /// Вторая половина того же договора — инвариант «отмена не оставляет частичного
    /// результата» (<see cref="NoPartialOutput"/>): сервисы, пишущие ОДИН файл, сохраняют
    /// его в самом конце, а пишущие по файлу на единицу работы удаляют созданное.
    /// </summary>
    internal static class Cancellation
    {
        public static void ThrowIf(Func<bool> cancelled)
        {
            if (cancelled != null && cancelled())
                throw new OperationCanceledException();
        }

        /// <summary>
        /// Выполнить работу, создающую файлы ПО ОДНОМУ, под инвариантом «отмена не оставляет
        /// частичного результата»: созданное к моменту отмены удаляется, а сама отмена уходит
        /// наверх. Работа складывает путь в переданный список сразу после создания каждого
        /// файла — иначе удалять будет нечего. Возвращает этот список.
        ///
        /// Обёртка, а не catch по месту: блок повторялся дословно в трёх режимах разделения, а
        /// в четвёртом месте (сохранение страниц картинками) его просто забыли — отменённый
        /// экспорт оставлял в папке уже сохранённые картинки. ОШИБКИ поведения не меняют:
        /// файлы, записанные до сбоя, остаются, как и прежде.
        /// </summary>
        public static List<string> NoPartialOutput(Action<List<string>> produce)
        {
            var created = new List<string>();
            try
            {
                produce(created);
            }
            catch (OperationCanceledException)
            {
                DeleteQuietly(created);
                throw;
            }
            return created;
        }

        /// <summary>Удалить файлы, созданные до отмены (best-effort — сбой удаления не важен).</summary>
        private static void DeleteQuietly(IEnumerable<string> paths)
        {
            foreach (string path in paths)
                try { if (File.Exists(path)) File.Delete(path); } catch { }
        }
    }
}
