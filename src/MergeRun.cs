using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Один запуск «Свода Excel» целиком: из какой папки, с какими параметрами, когда, над
    /// какими файлами и с каким итогом. Повтор пропущенных и записка читают всё это из ОДНОГО
    /// объекта: раньше окно хранило папку, параметры и время отдельно от итога и переписывало
    /// их при каждом запуске, а итог — только при успехе. Повтор после неудачного или
    /// отменённого запуска B сливал пропуски запуска A с параметрами B (смена «Содержания»
    /// могла удалить настоящий лист «Содержание» свода A), а записка ставила дату и папку B
    /// под итогом A.
    /// </summary>
    internal sealed class MergeRun
    {
        public readonly string InputFolder;
        public readonly MergeOptions Options;
        public readonly DateTime StartedAt;
        /// <summary>Файлы этого запуска: весь набор или пропуски для повтора — их строки списка и сбрасываются.</summary>
        public readonly IList<string> Paths;
        public readonly MergeResult Result; // null — запуск идёт или итога не дал

        public MergeRun(string inputFolder, MergeOptions options, DateTime startedAt, IList<string> paths)
            : this(inputFolder, options, startedAt, paths, null)
        {
        }

        private MergeRun(string inputFolder, MergeOptions options, DateTime startedAt, IList<string> paths,
            MergeResult result)
        {
            InputFolder = inputFolder;
            Options = options;
            StartedAt = startedAt;
            Paths = paths ?? new List<string>();
            Result = result;
        }

        /// <summary>
        /// Повтор пропущенных этого запуска: та же папка и те же параметры (дослияние идёт в его
        /// свод), набор — его пропуски. Только у запуска с итогом. Чистая — под тест.
        /// </summary>
        public MergeRun ForRetry(DateTime startedAt)
        {
            if (Result == null)
                throw new InvalidOperationException("повторять можно только запуск с итогом");
            return new MergeRun(InputFolder, Options, startedAt, MergeRetry.Paths(Result));
        }

        /// <summary>
        /// Какой запуск помнить «последним» после завершения current. Только запуск с итогом
        /// (не упавший и не отменённый) заменяет прежний — вместе со своими папкой, параметрами
        /// и временем; иначе остаётся last целиком: такой прогон свод не менял. Чистая — под тест.
        /// </summary>
        public static MergeRun Settle(MergeRun last, MergeRun current, MergeResult result)
        {
            if (current == null || result == null || result.Cancelled)
                return last;
            return new MergeRun(current.InputFolder, current.Options, current.StartedAt, current.Paths, result);
        }
    }
}
