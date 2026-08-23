using System.Runtime.CompilerServices;

namespace ExcelMerger
{
    /// <summary>
    /// Сколько страниц в PDF — и открывается ли он вообще. Единственная проба такого рода:
    /// ею конвейер Ghostscript проверяет, что после прогона документ остался документом.
    ///
    /// Читаем PdfPig, а НЕ PdfSharp, и это не вкусовщина: PdfSharp 1.50 не понимает потоков
    /// объектов (PDF 1.5+) и падает на них с ArgumentOutOfRangeException — а такие файлы
    /// пишут по умолчанию и Word, и Acrobat. Проверка, которая не открывает половину
    /// принесённых файлов, превратилась бы в отказ на ровном месте: сжатие сообщало бы
    /// «не удалось» там, где движок отработал безупречно. Замер на выходе Ghostscript
    /// с -dCompatibilityLevel=1.5: PdfSharp — ArgumentOutOfRangeException, PdfPig — 1 страница.
    /// </summary>
    internal static class PdfPageProbe
    {
        /// <summary>Файл не открылся (битый, зашифрованный, не PDF). Отличается от «ноль страниц».</summary>
        public const int Unreadable = -1;

        /// <summary>
        /// Число страниц или <see cref="Unreadable"/>. Не бросает: нечитаемый файл — это
        /// ответ, а не авария, и решает по нему вызывающий.
        ///
        /// Цена вопроса измерена, а не предположена: PDF на 58 МБ и 900 страниц —
        /// 0,45 с и 92 МБ рабочего набора (файл читается в память целиком). Против
        /// десятков секунд, которые тот же файл проводит в Ghostscript, это незаметно, а
        /// профиль памяти не нов: кэш миниатюр держит до шести полных файлов сразу.
        /// </summary>
        public static int PageCount(string path)
        {
            EmbeddedAssemblies.Ensure();
            return PageCountCore(path);
        }

        /// <summary>
        /// Уцелел ли документ после прогона движком: вывод открывается и содержит СТОЛЬКО ЖЕ
        /// страниц, сколько исходник. Если исходник не читается нами самими (починка — там он
        /// битый по условию задачи; экзотика, которую понимает Ghostscript, но не PdfPig),
        /// сравнивать не с чем — тогда достаточно, чтобы вывод открывался и был непуст.
        /// Ровно поэтому правило одно на все три режима: строгое там, где есть с чем сверять,
        /// и не дающее ложных отказов там, где сверять не с чем. Чистая — под тест.
        /// </summary>
        public static bool PagesKept(int before, int after)
        {
            if (after <= 0)
                return false;                  // вывод не открылся или пуст — отказ всегда
            return before <= 0 || after == before;
        }

        /// <summary>
        /// Открыть документ PdfPig — единственное место, где это делается, и потому
        /// единственное, где помнят про пароль. Защищённый файл без пароля PdfPig не
        /// открывает вовсе, так что забыть подставить его здесь значило бы получить
        /// «файл не читается» на каждом шаге разбора.
        ///
        /// <see cref="EmbeddedAssemblies.Ensure"/> зовётся ЗДЕСЬ, а не только в публичных
        /// обёртках: иначе прямой вызов (например, из сравнения версий, если оно первое
        /// коснулось PdfPig в сеансе) падал «сборка не найдена» — вшита она ресурсом и без
        /// перехвата разрешения не загружается.
        /// </summary>
        [MethodImpl(MethodImplOptions.NoInlining)]
        internal static UglyToad.PdfPig.PdfDocument OpenPig(string path)
        {
            EmbeddedAssemblies.Ensure();
            string password = PdfPasswords.For(path);
            if (string.IsNullOrEmpty(password))
                return UglyToad.PdfPig.PdfDocument.Open(path);
            var options = new UglyToad.PdfPig.ParsingOptions();
            options.Password = password;
            return UglyToad.PdfPig.PdfDocument.Open(path, options);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static int PageCountCore(string path)
        {
            try
            {
                using (UglyToad.PdfPig.PdfDocument doc = OpenPig(path))
                    return doc.NumberOfPages;
            }
            catch { return Unreadable; }
        }
    }
}
