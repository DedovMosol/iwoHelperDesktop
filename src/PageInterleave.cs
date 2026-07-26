using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>Документ в текущем порядке страниц: непрерывный отрезок страниц одного файла.</summary>
    public sealed class PageRun
    {
        public string SourcePath;
        public int Start;  // индекс первой страницы отрезка в общем порядке
        public int Count;  // сколько страниц подряд
        public bool Reverse; // выдавать этот документ с конца (обратная сторона двусторонней пачки)

        /// <summary>Имя файла для показа в диалоге (без пути).</summary>
        public string FileName
        {
            get { return SourcePath == null ? "" : System.IO.Path.GetFileName(SourcePath); }
        }
    }

    /// <summary>
    /// Чередование страниц нескольких документов — «склейка» двух пачек, полученных
    /// односторонним сканером: лицевые стороны в одном файле, оборотные в другом и, как
    /// правило, в обратном порядке. Результат — 1-я лицевая, 1-я оборотная, 2-я лицевая…
    ///
    /// Чистые функции без UI и без PDFsharp: разбор текущего порядка на документы и
    /// само перемешивание. Порядок страниц внутри документа сохраняется (или
    /// разворачивается), сами ссылки не копируются — переставляются те же объекты,
    /// поэтому назначенные пользователем повороты едут вместе со страницами.
    /// </summary>
    public static class PageInterleave
    {
        /// <summary>
        /// Разбить текущий порядок на документы: каждый максимальный непрерывный отрезок
        /// страниц ОДНОГО файла. Один файл, добавленный дважды, даёт два отрезка — так и
        /// нужно: пользователь видит их в сетке как две пачки. Чистая — под тест.
        /// </summary>
        public static List<PageRun> SplitIntoRuns(IList<PdfPageRef> pages)
        {
            var runs = new List<PageRun>();
            if (pages == null)
                return runs;
            for (int i = 0; i < pages.Count; i++)
            {
                PageRun last = runs.Count > 0 ? runs[runs.Count - 1] : null;
                bool sameAsPrevious = last != null && SamePath(last.SourcePath, pages[i].SourcePath) &&
                                      last.Start + last.Count == i;
                if (sameAsPrevious)
                    last.Count++;
                else
                    runs.Add(new PageRun { SourcePath = pages[i].SourcePath, Start = i, Count = 1 });
            }
            return runs;
        }

        /// <summary>
        /// Перемешать документы, беря из каждого по pace страниц по кругу, пока страницы не
        /// кончатся. Документ с Reverse отдаёт свои страницы с конца. Исчерпавшийся документ
        /// просто выбывает — хвост более длинного дописывается в конец (пачки со сканера
        /// часто отличаются на одну страницу). pace &lt; 1 считается за 1.
        ///
        /// Возвращает НОВЫЙ порядок из тех же самых ссылок: ничего не теряется и не
        /// дублируется — это перестановка. Чистая — под тест.
        /// </summary>
        public static List<PdfPageRef> Interleave(IList<PdfPageRef> pages, IList<PageRun> runs, int pace)
        {
            var result = new List<PdfPageRef>();
            if (pages == null || runs == null || runs.Count == 0)
                return result;
            if (pace < 1)
                pace = 1;

            // Сколько страниц уже отдал каждый документ.
            var taken = new int[runs.Count];
            bool any = true;
            while (any)
            {
                any = false;
                for (int r = 0; r < runs.Count; r++)
                {
                    PageRun run = runs[r];
                    for (int k = 0; k < pace && taken[r] < run.Count; k++)
                    {
                        int offset = run.Reverse ? run.Count - 1 - taken[r] : taken[r];
                        result.Add(pages[run.Start + offset]);
                        taken[r]++;
                        any = true;
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Осмысленно ли предлагать чередование: нужны минимум два документа. Один документ
        /// перемешивать не с чем — кнопка гасится. Чистая — под тест.
        /// </summary>
        public static bool CanInterleave(IList<PageRun> runs)
        {
            return runs != null && runs.Count >= 2;
        }

        /// <summary>
        /// Один ли документ у двух соседних страниц. Пустые страницы (без файла) считаются
        /// одинаковыми между собой: подряд идущие пустые — это одна вставка, и разбивать их
        /// на отдельные документы незачем.
        /// </summary>
        private static bool SamePath(string a, string b)
        {
            if (a == null || b == null)
                return a == null && b == null;
            return string.Equals(a, b, System.StringComparison.OrdinalIgnoreCase);
        }
    }
}
