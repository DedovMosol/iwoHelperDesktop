using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace ExcelMerger
{
    /// <summary>
    /// Остался ли в документе цвет — проверка результата перевода в оттенки серого.
    /// Смотрим не в структуру файла, а в то, что видит человек: страницы рендерятся тем же
    /// системным движком, что и миниатюры, и в растре ищется насыщенный пиксель. Структуру
    /// проверять бессмысленно — цвет может прийти из ICC-профиля, палитры или группы
    /// прозрачности, и перечислить все ходы заранее нельзя, а растр их закрывает все разом.
    ///
    /// Проба умеет только ОПРОВЕРГАТЬ: положительный ответ означает «цвет доказан».
    /// Если страница не отрендерилась (старая Windows, недоступный WinRT, необычный файл),
    /// это НЕ «остался цвет» — иначе на такой машине перевод в серое сломался бы там, где
    /// до появления проверки работал. Молчание проверки не должно отнимать функцию.
    /// </summary>
    internal static class PdfColorProbe
    {
        /// <summary>
        /// Насколько разъехались R/G/B, чтобы считать пиксель цветным. Честно серый документ
        /// даёт строго ноль (DeviceGray рендерится в R=G=B, и сглаживание равенства не рушит),
        /// так что порог — не «чувствительность», а запас против чужой цветокоррекции.
        /// </summary>
        internal const int SaturationThreshold = 24;

        /// <summary>
        /// Доля насыщенных пикселей, начиная с которой страница считается цветной. У честно
        /// серой страницы она равна нулю, поэтому любое разумное малое значение годится;
        /// смысл порога — не дать одинокому артефакту масштабирования отменить операцию,
        /// которая на деле удалась.
        /// </summary>
        internal const double MaxColorShare = 0.005;

        /// <summary>
        /// Сколько страниц смотрим. Цвет в документе почти никогда не живёт на одной странице,
        /// а рендер — единственная дорогая часть проверки, поэтому берём выборку по всей
        /// длине, а не каждую страницу: на 500-страничном документе разница между «мгновенно»
        /// и «полминуты».
        /// </summary>
        internal const int MaxSampledPages = 8;

        /// <summary>Ширина рендера. Цвет виден и на маленькой картинке, а памяти и времени просит мало.</summary>
        private const int RenderWidth = 140;

        /// <summary>
        /// Потолок высоты растра. PDF разрешает лист до 14400 пунктов при ширине в единицы, и
        /// при обычной ширине такая страница даёт растр в сотни тысяч пикселей высотой.
        /// Надеяться, что движок сам откажется, нельзя — и это не догадка: на машине
        /// разработчика он за лист 3×14400 не взялся (пусто, ноль лишней памяти), а на
        /// сборочной честно нарисовал его целиком и забрал 400 МБ, чем и уронил проверку.
        /// С потолком тот же лист стоит меньше мегабайта.
        /// </summary>
        private const int MaxRenderHeight = 2000;

        /// <summary>
        /// Индексы страниц для выборки: все подряд, пока их не больше max, иначе max штук,
        /// равномерно по всей длине документа (первая всегда попадает). Чистая — под тест.
        /// </summary>
        internal static int[] SamplePages(int pageCount, int max)
        {
            if (pageCount <= 0 || max <= 0)
                return new int[0];
            if (pageCount <= max)
            {
                var all = new int[pageCount];
                for (int i = 0; i < pageCount; i++)
                    all[i] = i;
                return all;
            }
            var picked = new int[max];
            for (int i = 0; i < max; i++)
                picked[i] = (int)((long)i * pageCount / max);
            return picked;
        }

        /// <summary>
        /// Есть ли на растре насыщенные пиксели в заметной доле. Через LockBits, а не GetPixel:
        /// на выборке из восьми страниц это разница в два порядка по времени. Чистая — под тест.
        /// </summary>
        internal static bool HasColor(Bitmap bmp)
        {
            if (bmp == null || bmp.Width <= 0 || bmp.Height <= 0)
                return false;
            // Формат просим явно: WinRT отдаёт то 24, то 32 бита на пиксель, а GDI+ по запросу
            // приводит к одному виду — иначе разбор строки зависел бы от источника.
            BitmapData data = bmp.LockBits(new Rectangle(0, 0, bmp.Width, bmp.Height),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppRgb);
            try
            {
                int lineBytes = bmp.Width * 4;      // без хвоста выравнивания строки
                var row = new byte[lineBytes];
                long colored = 0;
                long total = (long)bmp.Width * bmp.Height;
                long limit = (long)(total * MaxColorShare);
                for (int y = 0; y < bmp.Height; y++)
                {
                    Marshal.Copy(IntPtr.Add(data.Scan0, y * data.Stride), row, 0, lineBytes);
                    for (int x = 0; x < lineBytes; x += 4)
                    {
                        int b = row[x], g = row[x + 1], r = row[x + 2];
                        int max = r > g ? (r > b ? r : b) : (g > b ? g : b);
                        int min = r < g ? (r < b ? r : b) : (g < b ? g : b);
                        if (max - min > SaturationThreshold && ++colored > limit)
                            return true;       // порог взят — дальше считать нечего
                    }
                }
                return false;
            }
            finally { bmp.UnlockBits(data); }
        }

        /// <summary>
        /// Доказан ли цвет в файле. false — либо документ серый, либо посмотреть не удалось
        /// (см. описание класса: проба только опровергает). Только с фонового потока —
        /// рендерер однопоточный и создаётся здесь свой, чтобы не делить кэш с сеткой
        /// миниатюр: файл вот-вот подменят, и чужой кэш остался бы со старой версией.
        /// </summary>
        public static bool HasColorPages(string path)
        {
            int pages = PdfPageProbe.PageCount(path);
            if (pages <= 0)
                return false;
            using (var renderer = new PdfThumbnailRenderer())
            {
                foreach (int index in SamplePages(pages, MaxSampledPages))
                {
                    using (Bitmap bmp = renderer.Render(path, index, RenderWidth, MaxRenderHeight))
                        if (HasColor(bmp))
                            return true;
                }
            }
            return false;
        }
    }
}
