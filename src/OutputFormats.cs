using System;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Поддерживаемые форматы итогового файла. Формат определяется расширением
    /// пути — единственный источник истины для GUI, CLI и сервиса.
    /// </summary>
    public static class OutputFormats
    {
        /// <summary>Расширения в порядке отображения в интерфейсе.</summary>
        public static readonly string[] Extensions = { ".xlsx", ".xlsm", ".xlsb", ".xls" };

        /// <summary>
        /// Код XlFileFormat для SaveAs по расширению пути; 0 — формат не поддержан.
        /// </summary>
        public static int FileFormatFor(string path)
        {
            string ext = Path.GetExtension(path);
            if (string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase))
                return 51; // xlOpenXMLWorkbook
            if (string.Equals(ext, ".xlsm", StringComparison.OrdinalIgnoreCase))
                return 52; // xlOpenXMLWorkbookMacroEnabled
            if (string.Equals(ext, ".xlsb", StringComparison.OrdinalIgnoreCase))
                return 50; // xlExcel12
            if (string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase))
                return 56; // xlExcel8 (Excel 97–2003)
            return 0;
        }

        /// <summary>Срезает известное расширение, если пользователь ввёл его в имя файла.</summary>
        public static string StripKnownExtension(string fileName)
        {
            foreach (string ext in Extensions)
            {
                if (fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                    return fileName.Substring(0, fileName.Length - ext.Length).TrimEnd();
            }
            return fileName;
        }
    }
}
