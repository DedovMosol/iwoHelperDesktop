using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Пересчёт геометрии PDF в геометрию слайда. Всё чистое и покрыто тестами: ошибка здесь
    /// не выпадает исключением, а тихо сдвигает текст, и увидеть её можно только глазами.
    ///
    /// Две системы координат. В PDF единица — пункт, ось Y направлена ВВЕРХ, начало в левом
    /// нижнем углу. В OOXML единица — EMU (English Metric Unit, 914400 на дюйм, то есть 12700
    /// на пункт), ось Y направлена ВНИЗ, начало в левом верхнем углу. Перевод — не только
    /// умножение: вертикаль обязана переворачиваться от высоты страницы.
    ///
    /// Размер слайда в презентации ОДИН на все слайды — это ограничение формата. Когда в
    /// набор попали страницы разных размеров (альбомная колода и книжный документ), слайд
    /// берётся по самому частому размеру, а остальные страницы вписываются в него целиком
    /// с сохранением пропорций (масштаб применяется и к координатам, и к кеглю) и ставятся
    /// по центру.
    /// </summary>
    internal static class PptxGeometry
    {
        public const long EmuPerPoint = 12700;
        public const long EmuPerInch = 914400;

        /// <summary>Разумные пределы размера слайда: от дюйма до 56 дюймов (предел PowerPoint).</summary>
        public const double MinSlidePt = 72;
        public const double MaxSlidePt = 4032;

        /// <summary>
        /// Пункты → EMU для КООРДИНАТЫ. Знак сохраняется: рамка, чуть вылезшая за край
        /// страницы, так и останется чуть за краем, а не прилипнет к нулю. Чистая — под тест.
        /// </summary>
        public static long Emu(double pt)
        {
            if (double.IsNaN(pt) || double.IsInfinity(pt))
                return 0;
            return (long)Math.Round(pt * EmuPerPoint, MidpointRounding.AwayFromZero);
        }

        /// <summary>Пункты → EMU для РАЗМЕРА: отрицательной ширины не бывает. Чистая — под тест.</summary>
        public static long EmuSize(double pt)
        {
            long emu = Emu(pt);
            return emu > 0 ? emu : 0;
        }

        /// <summary>
        /// Кегль в сотых пункта, как того требует разметка (12 pt → 1200), с клампом по
        /// пределам самой разметки (1…4000 pt). Отсев МУСОРНОГО кегля — не здесь, а в
        /// <see cref="FontResolver.ClampSizePt"/>: там это про исходник, тут про единицы,
        /// и смешивать два клампа нельзя (уменьшенный масштабом кегль иначе «отскакивал»
        /// бы к умолчанию вместо того, чтобы стать меньше). Чистая — под тест.
        /// </summary>
        public static int Hundredths(double pt)
        {
            if (double.IsNaN(pt))
                return (int)(FontResolver.DefaultFontSize * 100);
            double v = Math.Round(pt * 100, MidpointRounding.AwayFromZero);
            if (v < 100) return 100;          // 1 pt — минимум разметки
            if (v > 400000) return 400000;    // 4000 pt — максимум разметки
            return (int)v;
        }

        /// <summary>Размер в допустимых пределах слайда.</summary>
        public static double ClampSlidePt(double pt)
        {
            if (double.IsNaN(pt) || pt < MinSlidePt)
                return MinSlidePt;
            return pt > MaxSlidePt ? MaxSlidePt : pt;
        }

        /// <summary>
        /// Размер слайда для колоды: самый частый размер страницы (при равенстве — тот, что
        /// встретился раньше), с клампом. Пустой набор — лист A4 книжной ориентации.
        /// Сравнение размеров огрублено до пункта: PDF нередко приносит 720.00001.
        /// Чистая — под тест.
        /// </summary>
        public static void SlideSizePt(IList<PdfPageText> pages, out double widthPt, out double heightPt)
        {
            widthPt = 595;
            heightPt = 842;
            if (pages == null || pages.Count == 0)
                return;

            var counts = new Dictionary<string, int>();
            var firstSeen = new Dictionary<string, int>();
            var sizes = new Dictionary<string, double[]>();
            for (int i = 0; i < pages.Count; i++)
            {
                PdfPageText page = pages[i];
                if (page == null || page.WidthPt <= 0 || page.HeightPt <= 0)
                    continue;
                double w = ClampSlidePt(page.WidthPt), h = ClampSlidePt(page.HeightPt);
                string key = Math.Round(w) + "x" + Math.Round(h);
                if (!counts.ContainsKey(key))
                {
                    counts[key] = 0;
                    firstSeen[key] = i;
                    sizes[key] = new[] { w, h };
                }
                counts[key]++;
            }
            if (counts.Count == 0)
                return;

            string best = null;
            foreach (KeyValuePair<string, int> kv in counts)
                if (best == null || kv.Value > counts[best]
                    || (kv.Value == counts[best] && firstSeen[kv.Key] < firstSeen[best]))
                    best = kv.Key;
            widthPt = sizes[best][0];
            heightPt = sizes[best][1];
        }

        /// <summary>
        /// Как лечь странице на слайд: масштаб и сдвиг центрирования. Совпал размер — масштаб 1
        /// и нулевые сдвиги (самый частый случай, никаких пересчётов). Не совпал — вписываем
        /// целиком по меньшей стороне и центрируем. Чистая — под тест.
        /// </summary>
        public static SlideFit Fit(double pageWidthPt, double pageHeightPt, double slideWidthPt, double slideHeightPt)
        {
            // Вырожденная страница (битый источник): переворот оси считаем от высоты СЛАЙДА —
            // от нулевой высоты он дал бы зеркальные координаты.
            if (pageWidthPt <= 0 || pageHeightPt <= 0 || slideWidthPt <= 0 || slideHeightPt <= 0)
                return new SlideFit(1, 0, 0, pageHeightPt > 0 ? pageHeightPt : slideHeightPt);
            // Огрубление до пункта: 720.00001 против 720 — это тот же размер, а не повод масштабировать.
            if (Math.Abs(pageWidthPt - slideWidthPt) < 0.5 && Math.Abs(pageHeightPt - slideHeightPt) < 0.5)
                return new SlideFit(1, 0, 0, pageHeightPt);

            double scale = Math.Min(slideWidthPt / pageWidthPt, slideHeightPt / pageHeightPt);
            double offsetX = (slideWidthPt - pageWidthPt * scale) / 2;
            double offsetY = (slideHeightPt - pageHeightPt * scale) / 2;
            return new SlideFit(scale, offsetX, offsetY, pageHeightPt);
        }
    }

    /// <summary>
    /// Перевод координат ОДНОЙ страницы в координаты слайда: масштаб, сдвиг центрирования и
    /// переворот вертикальной оси. Значимый тип (структура) — переводов столько же, сколько
    /// страниц, и каждый живёт ровно в пределах своей.
    /// </summary>
    internal struct SlideFit
    {
        private readonly double _scale;
        private readonly double _offsetXPt;
        private readonly double _offsetYPt;
        private readonly double _pageHeightPt;

        public SlideFit(double scale, double offsetXPt, double offsetYPt, double pageHeightPt)
        {
            _scale = scale > 0 ? scale : 1;
            _offsetXPt = offsetXPt;
            _offsetYPt = offsetYPt;
            _pageHeightPt = pageHeightPt;
        }

        /// <summary>Масштаб страницы (1 — вписывать не понадобилось).</summary>
        public double Scale { get { return _scale; } }

        /// <summary>Левая граница (pt, ось X) → EMU слайда.</summary>
        public long X(double leftPt)
        {
            return PptxGeometry.Emu(leftPt * _scale + _offsetXPt);
        }

        /// <summary>Верхняя граница (pt, ось Y ВВЕРХ) → EMU слайда (ось вниз).</summary>
        public long Y(double topPt)
        {
            return PptxGeometry.Emu((_pageHeightPt - topPt) * _scale + _offsetYPt);
        }

        /// <summary>Длина (pt) → EMU слайда с учётом масштаба (не отрицательная).</summary>
        public long Size(double pt)
        {
            return PptxGeometry.EmuSize(pt * _scale);
        }

        /// <summary>
        /// Кегль (pt) → сотые пункта: сперва отсев мусорного значения ИСХОДНИКА, затем масштаб
        /// вписывания. Порядок важен — см. <see cref="PptxGeometry.Hundredths"/>.
        /// </summary>
        public int FontHundredths(double pt)
        {
            return PptxGeometry.Hundredths(FontResolver.ClampSizePt(pt) * _scale);
        }
    }
}
