using System;
using System.IO;
using System.Text;

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
    /// контейнер OLE2 — это либо шифрованная книга (поток EncryptedPackage), либо
    /// чужой OLE2-файл под этим расширением (обычно переименованный .xls).
    /// Чистая функция — покрыта тестами.
    /// </summary>
    public static class FileSignature
    {
        // OLE2/CFB (xls, зашифрованные книги): D0 CF 11 E0 A1 B1 1A E1.
        private static readonly byte[] Cfb = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        // Поток, в котором Office хранит шифрованный пакет OOXML (MS-OFFCRYPTO).
        private const string EncryptedPackageName = "EncryptedPackage";
        private const int HeaderSize = 512;
        private const int DirEntrySize = 128;
        private const int HeaderDifatCount = 109;

        /// <summary>
        /// Потолок цепочки каталога. Пакету MS-OFFCRYPTO хватает 2–4 секторов (около дюжины
        /// записей: EncryptionInfo, EncryptedPackage, хранилище DataSpaces с потомками); 64 — с запасом.
        /// Без потолка враждебный файл с зацикленной цепочкой разбирался квадратичное время
        /// (64 МБ — 8 с, 256 МБ — почти 2 мин), и свод стоял без отмены. Сверх потолка — «не
        /// знаем», то есть false: формулировка останется нейтральной.
        /// </summary>
        internal const int MaxDirectorySectors = 64;

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

        /// <summary>
        /// Есть ли в каталоге OLE2 поток EncryptedPackage — признак шифрованной паролем
        /// книги OOXML. Без него OLE2 под расширением .xlsx — не «пароль», а другой формат
        /// (чаще всего переименованный .xls), и сообщать о пароле было бы неправдой.
        /// false — потока нет, файл не OLE2 или каталог не разобрать.
        /// </summary>
        public static bool HasEncryptedPackage(string path)
        {
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    return HasEncryptedPackage(fs);
            }
            catch
            {
                return false; // недоступен — формулировка останется нейтральной
            }
        }

        /// <summary>Разбор каталога CFB по цепочке FAT (MS-CFB): ищем запись-поток EncryptedPackage.</summary>
        internal static bool HasEncryptedPackage(Stream s)
        {
            var header = new byte[HeaderSize];
            if (ReadAt(s, 0, header) < HeaderSize || !StartsWith(header, Cfb))
                return false;
            int shift = BitConverter.ToUInt16(header, 0x1E);
            if (shift != 9 && shift != 12)
                return false; // версии 3 и 4 — других секторов спецификация не знает
            int sectorSize = 1 << shift;
            var sector = new byte[sectorSize];
            // Цепочка DIFAT битого файла тоже может быть зацикленной: секторов не больше, чем в файле.
            long maxSectors = s.Length / sectorSize + 1;
            // Последний найденный сектор FAT: соседние сектора каталога описаны в одном секторе
            // FAT, и DIFAT не проходится заново на каждый (на самоцикле — ни разу повторно).
            var fat = new FatSectorCache { Index = -1 };

            int sid = BitConverter.ToInt32(header, 0x30); // первый сектор каталога
            for (int n = 0; sid >= 0 && n < MaxDirectorySectors; n++)
            {
                if (ReadAt(s, SectorOffset(sid, shift), sector) < sectorSize)
                    return false;
                for (int off = 0; off + DirEntrySize <= sectorSize; off += DirEntrySize)
                {
                    if (IsEncryptedPackageEntry(sector, off))
                        return true;
                }
                sid = NextSector(s, header, shift, sid, maxSectors, ref fat);
            }
            return false;
        }

        /// <summary>Номер сектора FAT (Sector) для его индекса в DIFAT (Index); -1 — ещё не искали.</summary>
        private struct FatSectorCache
        {
            public int Index;
            public int Sector;
        }

        private static bool IsEncryptedPackageEntry(byte[] sector, int off)
        {
            int nameBytes = BitConverter.ToUInt16(sector, off + 0x40); // длина с завершающим нулём
            byte type = sector[off + 0x42];                              // 2 — поток
            if (type != 2 || nameBytes != (EncryptedPackageName.Length + 1) * 2)
                return false;
            string name = Encoding.Unicode.GetString(sector, off, EncryptedPackageName.Length * 2);
            return string.Equals(name, EncryptedPackageName, StringComparison.Ordinal);
        }

        /// <summary>Следующий сектор цепочки по FAT; отрицательное — конец цепочки или сбой.</summary>
        private static int NextSector(Stream s, byte[] header, int shift, int sid, long maxSectors, ref FatSectorCache fat)
        {
            int perSector = (1 << shift) / 4;
            int index = sid / perSector;
            if (fat.Index != index)
            {
                fat.Sector = DifatEntry(s, header, shift, index, maxSectors);
                fat.Index = index;
            }
            int fatSector = fat.Sector;
            if (fatSector < 0)
                return -1;
            var buf = new byte[4];
            if (ReadAt(s, SectorOffset(fatSector, shift) + (long)(sid % perSector) * 4, buf) < 4)
                return -1;
            return BitConverter.ToInt32(buf, 0);
        }

        /// <summary>Номер i-го сектора FAT: первые 109 — в заголовке, дальше — цепочка DIFAT.</summary>
        private static int DifatEntry(Stream s, byte[] header, int shift, int index, long maxSectors)
        {
            if (index < HeaderDifatCount)
                return BitConverter.ToInt32(header, 0x4C + index * 4);
            int perDifat = (1 << shift) / 4 - 1; // последний int сектора — ссылка на следующий
            int rest = index - HeaderDifatCount;
            int difat = BitConverter.ToInt32(header, 0x44);
            var buf = new byte[4];
            for (long guard = 0; difat >= 0 && guard < maxSectors; guard++)
            {
                long pos = SectorOffset(difat, shift);
                if (rest < perDifat)
                {
                    if (ReadAt(s, pos + (long)rest * 4, buf) < 4)
                        return -1;
                    return BitConverter.ToInt32(buf, 0);
                }
                rest -= perDifat;
                if (ReadAt(s, pos + (long)perDifat * 4, buf) < 4)
                    return -1;
                difat = BitConverter.ToInt32(buf, 0);
            }
            return -1;
        }

        // Заголовок занимает первый сектор: сектор №sid начинается со смещения (sid+1)·размер.
        private static long SectorOffset(int sid, int shift)
        {
            return ((long)sid + 1) << shift;
        }

        private static int ReadAt(Stream s, long position, byte[] buffer)
        {
            if (position < 0 || position >= s.Length)
                return 0;
            s.Position = position;
            int total = 0;
            while (total < buffer.Length)
            {
                int read = s.Read(buffer, total, buffer.Length - total);
                if (read <= 0)
                    break;
                total += read;
            }
            return total;
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
