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
            var items = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (items == null)
                return new string[0];
            var picked = new List<string>();
            foreach (string item in items)
                if (File.Exists(item) && Matches(item, extensions))
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
