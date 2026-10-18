using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Извлечение путей из перетаскивания в окно — общее для всех PDF-инструментов
    /// (объединение берёт все, разделение — первый). DRY.
    /// </summary>
    internal static class PdfDrop
    {
        /// <summary>Набор по умолчанию: инструменты работают с PDF.</summary>
        public static readonly string[] PdfOnly = { ".pdf" };

        /// <summary>
        /// PDF и картинки — набор «Прочих операций»: они собирают документ и из снимков, а
        /// перетаскивание должно принимать ровно то же, что и кнопки того же окна.
        /// </summary>
        public static readonly string[] PdfAndImages = WithImages();

        public static string[] ExtractPaths(DragEventArgs e)
        {
            return ExtractPaths(e, PdfOnly);
        }

        /// <summary>Брошенные пути с одним из расширений (регистр не важен, файл должен существовать).</summary>
        public static string[] ExtractPaths(DragEventArgs e, string[] extensions)
        {
            if (e == null || !e.Data.GetDataPresent(DataFormats.FileDrop))
                return new string[0];
            return Filter(e.Data.GetData(DataFormats.FileDrop) as string[], extensions, File.Exists);
        }

        /// <summary>
        /// Сначала дешёвое расширение, потом существование: DragOver приходит на каждое
        /// движение мыши, а File.Exists по недоступному сетевому пути вешает UI-поток.
        /// Папки и чужие типы не трогают диск вовсе. Чистая при подставном exists — под тест.
        /// </summary>
        internal static string[] Filter(IEnumerable<string> items, string[] extensions,
            Func<string, bool> exists)
        {
            var picked = new List<string>();
            if (items == null || exists == null)
                return picked.ToArray();
            foreach (string item in items)
                if (Matches(item, extensions) && exists(item))
                    picked.Add(item);
            return picked.ToArray();
        }

        /// <summary>Расширение пути есть в наборе. Чистая — под тест.</summary>
        internal static bool Matches(string path, string[] extensions)
        {
            if (string.IsNullOrEmpty(path) || extensions == null)
                return false;
            string ext = Path.GetExtension(path);
            foreach (string known in extensions)
                if (string.Equals(ext, known, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        private static string[] WithImages()
        {
            var all = new List<string>(PdfOnly);
            all.AddRange(ImageToPdfService.Extensions);
            return all.ToArray();
        }
    }
}
