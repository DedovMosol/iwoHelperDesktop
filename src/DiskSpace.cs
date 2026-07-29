using System;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// «На диске нет места» — самая обидная из ошибок записи: сообщение системы говорит, что
    /// места нет, но НЕ говорит, где именно, а это половина ответа. Результат пишется в одну
    /// папку, временные файлы Ghostscript и Word — в другую, на другом диске, и переполниться
    /// может любой из них. Здесь ошибка опознаётся по коду и дополняется буквой диска и тем,
    /// сколько на нём осталось, — чтобы человек знал, что чистить.
    /// </summary>
    internal static class DiskSpace
    {
        // Коды Windows: 0x27 — «диск переполнен» при записи через дескриптор, 0x70 — при
        // операции с файлом. В HRESULT они приходят с приставкой 0x8007 (FACILITY_WIN32).
        private const int DiskFull = unchecked((int)0x80070070);
        private const int HandleDiskFull = unchecked((int)0x80070027);

        /// <summary>
        /// Ниже этого остатка (МБ) диск считаем переполненным, даже если ошибка о месте молчит.
        /// Молчат многие: GDI+ на любую неудачу записи отвечает «в GDI+ произошла общая ошибка»,
        /// COM Word — своим кодом. Сказать «свободно 3 МБ» рядом с такой ошибкой полезнее, чем
        /// не сказать ничего: даже если причина окажется другой, названное число проверяемо.
        /// </summary>
        private const long LowSpaceMb = 16;

        /// <summary>Эта ошибка — про нехватку места? Чистая — под тест.</summary>
        public static bool IsFull(Exception ex)
        {
            while (ex != null)
            {
                if (ex is IOException && (ex.HResult == DiskFull || ex.HResult == HandleDiskFull))
                    return true;
                ex = ex.InnerException;
            }
            return false;
        }

        /// <summary>
        /// Свободное место на диске пути, в мегабайтах; -1 — узнать не удалось (сетевой путь,
        /// нет прав). Отдельно от текста сообщения — чтобы величину можно было проверить.
        /// </summary>
        public static long FreeMegabytes(string path)
        {
            try
            {
                string root = Path.GetPathRoot(Path.GetFullPath(path));
                if (string.IsNullOrEmpty(root))
                    return -1;
                return new DriveInfo(root).AvailableFreeSpace / (1024 * 1024);
            }
            catch { return -1; }
        }

        /// <summary>
        /// Сообщение об ошибке записи, дополненное диском. path — куда писали.
        ///
        /// Два разных утверждения, и смешивать их нельзя. Если о нехватке места сказала САМА
        /// система — заменяем её сообщение своим, с диском и остатком: диск она не называет, а
        /// это половина ответа. Если система сказала что-то другое, но на диске почти ничего не
        /// осталось, — исходное сообщение сохраняем и лишь добавляем к нему остаток: причина
        /// может быть и другой, выдавать догадку за факт нельзя.
        /// </summary>
        public static string Describe(Exception ex, string path)
        {
            string root;
            try { root = Path.GetPathRoot(Path.GetFullPath(path)); }
            catch { root = null; }
            if (string.IsNullOrEmpty(root))
                return ex.Message;
            long free = FreeMegabytes(path);
            if (IsFull(ex))
                return free >= 0
                    ? string.Format(Loc.T("err.disk.fullFree"), root, free)
                    : string.Format(Loc.T("err.disk.full"), root);
            if (free >= 0 && free < LowSpaceMb)
                return ex.Message + string.Format(Loc.T("err.disk.lowFree"), root, free);
            return ex.Message;
        }
    }
}
