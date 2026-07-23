using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>Найденная печать ЭП: рамка (PDF, ось Y вверх) и слова, которые она накрывает.</summary>
    internal sealed class StampRegion
    {
        public double Left;
        public double Right;
        public double Bottom;
        public double Top;
        public List<PdfWord> Words = new List<PdfWord>(); // слова печати — их снимаем из текста (печать станет картинкой)

        public double Width { get { return Right - Left; } }
        public double Height { get { return Top - Bottom; } }
    }

    /// <summary>
    /// Распознавание штампа электронной подписи, вписанного в PDF ТЕКСТОМ (а не картинкой).
    /// Такой штамп при конвертации иначе растекается обычными строками и теряет вид печати —
    /// поэтому область распознаётся и переносится изображением (рендер-кроп в PdfTextExtract).
    ///
    /// Детекция НАРОЧНО строгая: нужны все опорные слова (см. <see cref="Anchors"/>) в компактной
    /// вертикальной полосе. Отсюда ключевое свойство: ложные срабатывания практически исключены
    /// (в обычном тексте эти слова не стоят кучно в рамке), а если штамп не распознан — поведение
    /// прежнее (текст остаётся текстом), то есть регрессии нет. Растровые печати (нарисованные
    /// картинкой) сюда не попадают — у них нет извлекаемого текста, их и так переносит
    /// ExtractImages. Чистая логика без типов PdfPig — под юнит-тест.
    /// </summary>
    internal static class StampDetector
    {
        // Опорные слова штампа (в нижнем регистре, без окаймляющей пунктуации). Нужны ВСЕ.
        private static readonly string[] Anchors = { "подписан", "сертификат", "владелец", "действителен" };

        // Компактность опорных слов: если ключевые слова разбросаны шире — это не печать, а проза.
        private const double AnchorMaxWidthFrac = 0.5;
        private const double AnchorMaxHeightFrac = 0.2;
        // Компактность итоговой рамки (с содержимым строк — номер сертификата, ФИО, даты).
        private const double RegionMaxWidthFrac = 0.8;
        private const double RegionMaxHeightFrac = 0.22;
        private const double BandPadEmFactor = 0.4; // полоса строк штампа расширяется на долю кегля вверх/вниз
        private const int MinStampWords = 6;         // печать — минимум несколько слов (заголовок + поля)

        /// <summary>
        /// Найти текстовый штамп ЭП среди слов страницы. Возвращает область и её слова, либо null.
        /// pageW/pageH — размеры страницы (pt) для проверки компактности. Чистая — под тест.
        /// </summary>
        public static StampRegion Detect(IList<PdfWord> words, double pageW, double pageH)
        {
            if (words == null || words.Count < MinStampWords || pageW <= 0 || pageH <= 0)
                return null;

            // 1) Опорные слова: по одному представителю каждой категории достаточно, но берём все
            //    совпадения (у категории может быть несколько вхождений) — для устойчивой рамки.
            var anchorWords = new List<PdfWord>();
            var seen = new bool[Anchors.Length];
            foreach (PdfWord w in words)
            {
                int cat = AnchorCategory(w.Text);
                if (cat < 0)
                    continue;
                seen[cat] = true;
                anchorWords.Add(w);
            }
            for (int i = 0; i < seen.Length; i++)
                if (!seen[i])
                    return null; // нет всех четырёх опорных слов — не штамп

            // 2) Компактность опорных слов (иначе это разрозненные слова в тексте, не печать).
            Box a = Bounds(anchorWords);
            if (a.Width > AnchorMaxWidthFrac * pageW || a.Height > AnchorMaxHeightFrac * pageH)
                return null;

            // 3) Полоса строк штампа = вертикальный размах опорных слов с запасом на кегль; в неё
            //    попадают и не-опорные слова строк (номер сертификата, ФИО, даты).
            double em = MedianHeight(anchorWords);
            if (em <= 0) em = 1;
            double bandBottom = a.Bottom - BandPadEmFactor * em;
            double bandTop = a.Top + BandPadEmFactor * em;

            var region = new StampRegion();
            foreach (PdfWord w in words)
            {
                double cy = (w.Top + w.Bottom) / 2;
                if (cy >= bandBottom && cy <= bandTop)
                    region.Words.Add(w);
            }
            if (region.Words.Count < MinStampWords)
                return null;

            // 4) Итоговая рамка и её компактность (если полоса захватила ползоны текста — отклоняем).
            Box r = Bounds(region.Words);
            if (r.Width > RegionMaxWidthFrac * pageW || r.Height > RegionMaxHeightFrac * pageH)
                return null;

            region.Left = r.Left;
            region.Right = r.Right;
            region.Bottom = r.Bottom;
            region.Top = r.Top;
            return region;
        }

        /// <summary>Категория опорного слова (индекс в Anchors) или -1. Регистр и пунктуация по краям игнорируются.</summary>
        private static int AnchorCategory(string text)
        {
            string norm = NormalizeWord(text);
            if (norm.Length == 0)
                return -1;
            for (int i = 0; i < Anchors.Length; i++)
                if (norm == Anchors[i])
                    return i;
            return -1;
        }

        /// <summary>Слово в нижнем регистре без окаймляющих не-буквенных символов («Сертификат:» → «сертификат»).</summary>
        private static string NormalizeWord(string text)
        {
            if (string.IsNullOrEmpty(text))
                return "";
            int start = 0, end = text.Length - 1;
            while (start <= end && !char.IsLetter(text[start])) start++;
            while (end >= start && !char.IsLetter(text[end])) end--;
            if (end < start)
                return "";
            return text.Substring(start, end - start + 1).ToLowerInvariant();
        }

        private struct Box
        {
            public double Left, Right, Bottom, Top;
            public double Width { get { return Right - Left; } }
            public double Height { get { return Top - Bottom; } }
        }

        private static Box Bounds(List<PdfWord> ws)
        {
            var b = new Box { Left = double.MaxValue, Right = double.MinValue, Bottom = double.MaxValue, Top = double.MinValue };
            foreach (PdfWord w in ws)
            {
                if (w.Left < b.Left) b.Left = w.Left;
                if (w.Right > b.Right) b.Right = w.Right;
                if (w.Bottom < b.Bottom) b.Bottom = w.Bottom;
                if (w.Top > b.Top) b.Top = w.Top;
            }
            return b;
        }

        private static double MedianHeight(List<PdfWord> ws)
        {
            var hs = new List<double>(ws.Count);
            foreach (PdfWord w in ws) hs.Add(w.Height);
            hs.Sort();
            return hs.Count == 0 ? 0 : hs[(hs.Count - 1) / 2];
        }
    }
}
