using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace ExcelMerger
{
    /// <summary>
    /// Добивка документов до чётного числа страниц — для двусторонней печати: без неё первая
    /// страница следующего документа печатается на обороте последней страницы предыдущего.
    ///
    /// Считаем ЧИСТОЙ функцией и применяем при записи, а не заводим пустые страницы в модели:
    /// иначе их пришлось бы рисовать в сетке, тащить через буфер и отмену, запрещать в
    /// «PDF → Word» — цена несопоставима с пользой. Документом считается каждый непрерывный
    /// отрезок страниц одного файла (та же разметка, что у чередования, — см.
    /// <see cref="PageInterleave.SplitIntoRuns"/>), поэтому перестановка страниц вручную
    /// разметку не ломает: пользователь видит в сетке ровно те же пачки.
    /// </summary>
    public static class BlankPages
    {
        /// <summary>
        /// Позиции, ПЕРЕД которыми нужно вставить пустую страницу, в порядке возрастания.
        /// Позиция равна числу исходных страниц слева от неё, поэтому вставлять их надо
        /// с конца — иначе каждая следующая позиция съезжает.
        ///
        /// Пустая страница добавляется после документа с НЕЧЁТНЫМ числом страниц, кроме
        /// последнего: добивать хвост документа нечем — за ним ничего не печатается.
        /// Чистая — под тест.
        /// </summary>
        public static List<int> InsertPositions(IList<PdfPageRef> pages)
        {
            var result = new List<int>();
            List<PageRun> runs = PageInterleave.SplitIntoRuns(pages);
            for (int i = 0; i < runs.Count - 1; i++) // последний документ пропускаем намеренно
                if (runs[i].Count % 2 != 0)
                    result.Add(runs[i].Start + runs[i].Count);
            return result;
        }

        /// <summary>Нужна ли вообще добивка для этого набора страниц. Чистая — под тест.</summary>
        public static bool Needed(IList<PdfPageRef> pages)
        {
            return InsertPositions(pages).Count > 0;
        }

        /// <summary>Лист A4 в пунктах — размер по умолчанию, когда взять его не у чего.</summary>
        public const double A4WidthPt = 595;
        public const double A4HeightPt = 842;

        /// <summary>
        /// Размер вставляемого пустого листа: как у страницы, рядом с которой он встаёт, иначе
        /// A4. Пустой лист чужого формата посреди документа выглядит ошибкой, а не намерением.
        /// Чистая — под тест.
        /// </summary>
        public static void SheetSize(PdfPageInfo neighbour, out double widthPt, out double heightPt)
        {
            if (neighbour != null && neighbour.WidthPt > 0 && neighbour.HeightPt > 0)
            {
                widthPt = neighbour.WidthPt;
                heightPt = neighbour.HeightPt;
                return;
            }
            widthPt = A4WidthPt;
            heightPt = A4HeightPt;
        }

        /// <summary>
        /// Записать одностраничный пустой PDF заданного размера. В сетку он попадает тем же
        /// путём, что и картинка: файл-обёртка во временной папке окна, а дальше — обычная
        /// страница, которую двигают, поворачивают, печатают и сохраняют.
        /// </summary>
        public static void WriteSheet(string path, double widthPt, double heightPt)
        {
            EmbeddedAssemblies.Ensure();
            WriteSheetCore(path, widthPt, heightPt);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void WriteSheetCore(string path, double widthPt, double heightPt)
        {
            using (var doc = new PdfSharp.Pdf.PdfDocument())
            {
                PdfSharp.Pdf.PdfPage page = doc.AddPage();
                page.Width = widthPt;
                page.Height = heightPt;
                doc.Save(path);
            }
        }
    }
}
