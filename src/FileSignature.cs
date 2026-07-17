using System;
using System.IO;

namespace ExcelMerger
{
    /// <summary>Тип контейнера файла по сигнатуре (magic bytes).</summary>
    public enum ExcelContainer
    {
        NotExcel,    // ни ZIP, ни OLE2 — текст, пустой, мусор
        Zip,         // ZIP — незашифрованный OOXML (xlsx/xlsm/xlsb)
        Ole2,        // OLE2/CFB — xls или зашифрованная (парольная) книга OOXML
        Unreadable   // не удалось прочитать — решение оставляем Excel
    }

    /// <summary>
    /// Распознаёт контейнер книги Excel по первым байтам, не открывая её в Excel.
    /// Нужно, чтобы отсеять до Excel то, что его подвешивает: битый файл
    /// (например, текст с расширением .xlsx) и зашифрованную парольную книгу —
    /// их <c>Workbooks.Open</c> заклинивает так, что и следующие файлы перестают
    /// открываться. Незашифрованный OOXML всегда ZIP; если у .xlsx/.xlsm/.xlsb
    /// контейнер OLE2 — книга зашифрована. Чистая функция — покрыта тестами.
    /// </summary>
    public static class FileSignature
    {
        // OLE2/CFB (xls, зашифрованные книги): D0 CF 11 E0 A1 B1 1A E1.
        private static readonly byte[] Cfb = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        public static ExcelContainer Detect(string path)
        {
            byte[] head = ReadHead(path, 8);
            if (head == null)
                return ExcelContainer.Unreadable;
            if (head.Length >= 2 && head[0] == 0x50 && head[1] == 0x4B)
                return ExcelContainer.Zip;  // "PK"
            if (head.Length >= 8 && StartsWith(head, Cfb))
                return ExcelContainer.Ole2;
            return ExcelContainer.NotExcel;
        }

        private static byte[] ReadHead(string path, int count)
        {
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var buf = new byte[count];
                    int read = fs.Read(buf, 0, count);
                    if (read == count)
                        return buf;
                    var trimmed = new byte[read];
                    Array.Copy(buf, trimmed, read);
                    return trimmed;
                }
            }
            catch
            {
                return null; // недоступен/занят — не отсеиваем
            }
        }

        private static bool StartsWith(byte[] data, byte[] prefix)
        {
            for (int i = 0; i < prefix.Length; i++)
                if (data[i] != prefix[i])
                    return false;
            return true;
        }
    }
}
