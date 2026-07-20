using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Извлечение путей PDF из перетаскивания в окно — общее для обоих PDF-инструментов
    /// (объединение берёт все, разделение — первый). DRY.
    /// </summary>
    internal static class PdfDrop
    {
        public static string[] ExtractPaths(DragEventArgs e)
        {
            if (e == null || !e.Data.GetDataPresent(DataFormats.FileDrop))
                return new string[0];
            var items = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (items == null)
                return new string[0];
            var pdfs = new List<string>();
            foreach (string item in items)
                if (File.Exists(item) &&
                    string.Equals(Path.GetExtension(item), ".pdf", StringComparison.OrdinalIgnoreCase))
                    pdfs.Add(item);
            return pdfs.ToArray();
        }
    }
}
