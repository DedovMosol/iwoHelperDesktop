using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Подложки слайдов: страница PDF, отрисованная БЕЗ текста. Так на слайд попадает всё, чего
    /// в текстовой модели нет и быть не может — фон, рамки, диаграммы, векторные логотипы, — а
    /// текст ложится поверх редактируемыми надписями. Без этого слоя презентация превращается в
    /// набор подписей на белом листе: у типового слайда почти вся картинка нарисована вектором.
    ///
    /// Слой опционален во всех смыслах: нет Ghostscript — молча возвращаем пустой результат и
    /// пишем презентацию из одного текста. Ошибка отрисовки одной страницы не трогает остальные.
    /// </summary>
    internal static class PageBackgrounds
    {
        /// <summary>Разрешение подложек по числу страниц: колода на сотню страниц не должна весить сотни мегабайт.</summary>
        internal static int Dpi(int pageCount)
        {
            if (pageCount <= 50) return 150;
            if (pageCount <= 150) return 120;
            return 96;
        }

        /// <summary>
        /// Общий бюджет подложек. Исчерпан — остальные страницы идут без подложки: пусть файл
        /// будет беднее, но останется файлом, который открывается и правится.
        /// </summary>
        private const long MediaBudgetBytes = 64L * 1024 * 1024;

        /// <summary>
        /// Отрисовать подложки для страниц заказа. Результат — массив длиной с заказ: элемент
        /// null означает «подложки нет» (пустая страница, нет Ghostscript, исчерпан бюджет).
        /// progress — «сделано/всего» по страницам, cancelled — отмена между страницами.
        /// </summary>
        public static Background[] Render(IList<PdfPageRef> order, Action<int, int> progress = null,
            Func<bool> cancelled = null)
        {
            if (order == null || order.Count == 0)
                return new Background[0];
            var result = new Background[order.Count];
            if (!Ghostscript.Available)
                return result; // молчаливый откат: презентация будет из одного текста

            int dpi = Dpi(order.Count);
            long spent = 0;
            string root = Path.Combine(Path.GetTempPath(), "iwo_bg_" + Guid.NewGuid().ToString("N"));
            int done = 0;
            int sourceNo = 0;
            try
            {
                foreach (KeyValuePair<string, List<int>> source in GroupBySource(order))
                {
                    sourceNo++;
                    Cancellation.ThrowIf(cancelled);
                    List<int> slots = source.Value;
                    int firstPage = int.MaxValue, lastPage = 0;
                    foreach (int slot in slots)
                    {
                        int page = order[slot].PageIndex + 1; // Ghostscript нумерует страницы с единицы
                        if (page < firstPage) firstPage = page;
                        if (page > lastPage) lastPage = page;
                    }

                    // Один запуск движка на файл: процесс дорог, а страницы он рисует пачкой.
                    // Своя папка на каждый источник — имена файлов у движка одинаковые.
                    string dir = Path.Combine(root, "src" + sourceNo.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    List<string> files = PageRasterizer.RenderPagesWithoutText(source.Key, firstPage, lastPage, dpi, dir);
                    foreach (int slot in slots)
                    {
                        Cancellation.ThrowIf(cancelled);
                        done++;
                        if (progress != null)
                            progress(done, order.Count);
                        int index = order[slot].PageIndex + 1 - firstPage;
                        if (index < 0 || index >= files.Count)
                            continue; // страницу отрисовать не удалось — слайд останется без подложки
                        if (spent >= MediaBudgetBytes)
                            continue;
                        Background bg = Encode(files[index], order[slot].Rotation);
                        if (bg == null)
                            continue;
                        spent += bg.Data.Length;
                        result[slot] = bg;
                    }
                }
            }
            finally
            {
                try { if (Directory.Exists(root)) Directory.Delete(root, true); } catch { }
            }
            return result;
        }

        /// <summary>Страницы заказа по исходным файлам: значение — позиции в заказе.</summary>
        internal static Dictionary<string, List<int>> GroupBySource(IList<PdfPageRef> order)
        {
            var map = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < order.Count; i++)
            {
                PdfPageRef r = order[i];
                if (r == null || string.IsNullOrEmpty(r.SourcePath) || r.PageIndex < 0)
                    continue;
                List<int> slots;
                if (!map.TryGetValue(r.SourcePath, out slots))
                    map[r.SourcePath] = slots = new List<int>();
                slots.Add(i);
            }
            return map;
        }

        /// <summary>
        /// Прочитать отрисованную страницу, довернуть её так же, как пользователь довернул
        /// страницу в сетке, и выбрать кодек. Пустая (одноцветная) страница подложки не
        /// заслуживает — это самый частый случай в текстовых документах, и он же экономит вес.
        /// </summary>
        private static Background Encode(string file, int rotation)
        {
            try
            {
                using (var loaded = new Bitmap(file))
                using (Bitmap bmp = Rotate(loaded, rotation))
                {
                    if (RasterUtil.IsSolidColor(bmp))
                        return null;
                    byte[] png = SavePng(bmp);
                    if (png == null)
                        return null;
                    bool jpeg;
                    byte[] data = RasterUtil.PreferSmaller(png, out jpeg);
                    return new Background { Data = data, IsJpeg = jpeg };
                }
            }
            catch { return null; } // одна подложка не удалась — остальные не при чём
        }

        /// <summary>Копия растра, довёрнутая на угол пользователя (0 — та же картинка).</summary>
        private static Bitmap Rotate(Bitmap source, int rotation)
        {
            var copy = new Bitmap(source); // независимая копия: исходник закрыт вместе с файлом
            try
            {
                if (rotation != 0)
                    copy.RotateFlip(PageRotation.FlipFor(rotation));
                return copy;
            }
            catch
            {
                copy.Dispose(); // поворот не удался — копия наружу не уходит и не течёт
                throw;
            }
        }

        private static byte[] SavePng(Bitmap bmp)
        {
            try
            {
                using (var ms = new MemoryStream())
                {
                    bmp.Save(ms, ImageFormat.Png);
                    return ms.ToArray();
                }
            }
            catch { return null; }
        }
    }

    /// <summary>Подложка страницы: готовые байты картинки и её формат.</summary>
    internal sealed class Background
    {
        public byte[] Data;
        public bool IsJpeg;
    }
}
