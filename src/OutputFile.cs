using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Выбор имени для создаваемого файла. Отделено от <see cref="NameTemplate"/> нарочно:
    /// там чистая сборка имени из шаблона (и потому всё под тестами), а здесь единственное,
    /// что требует обращения к диску — проверка занятости имени.
    /// </summary>
    public static class OutputFile
    {
        /// <summary>
        /// Путь в папке со СВОБОДНЫМ именем: «имя.pdf», занято — «имя_2.pdf», «имя_3.pdf»…
        /// Ни один инструмент не перезаписывает молча: результат прошлого запуска остаётся
        /// на месте, а новый ложится рядом. Имя уже должно быть очищено вызывающим.
        /// </summary>
        public static string Unique(string dir, string safeName, string extension)
        {
            string path = Path.Combine(dir, safeName + extension);
            int n = 2;
            while (File.Exists(path))
            {
                path = Path.Combine(dir, safeName + "_" + n + extension);
                n++;
            }
            return path;
        }
    }
}
