using System;
using System.IO;
using System.Runtime.CompilerServices;

namespace ExcelMerger
{
    /// <summary>
    /// Из чего состоит документ — ровно настолько, чтобы объяснить человеку, почему сжатие
    /// ничего не дало.
    ///
    /// Повод не выдуманный. Скан на четыре страницы весом 1,6 МБ: текста в нём нет вовсе,
    /// а четыре изображения занимают 99,8% файла. Уровень «Очень хорошо» изображения не
    /// трогает по своему определению, поэтому пересобранный документ вышел на 0,1% БОЛЬШЕ
    /// исходного, был отвергнут — и человек увидел прежний размер и слова «файл уже
    /// оптимизирован». Файл при этом отлично сжимался: «Хорошо» давало −64%, «Нормально» −88%.
    /// Сообщение было формально верным и по сути ложным: оптимизирован не файл, а просто
    /// выбранный уровень не про эти байты.
    ///
    /// Профиль считается ТОЛЬКО когда сжатие не дало выигрыша — то есть редко, — и потому
    /// может позволить себе прочитать документ целиком.
    /// </summary>
    internal static class PdfContentProfile
    {
        /// <summary>Долю посчитать не удалось (файл не открылся или пуст).</summary>
        public const double Unknown = -1.0;

        /// <summary>
        /// Какую часть файла занимают изображения (0..1) или <see cref="Unknown"/>.
        /// Берутся сжатые байты картинок, как они лежат в файле, — именно их и уменьшает
        /// пересчёт, поэтому сравнивать с размером файла честно.
        /// </summary>
        public static double ImageShare(string path)
        {
            EmbeddedAssemblies.Ensure();
            return ImageShareCore(path);
        }

        /// <summary>
        /// Стоит ли посоветовать уровень, уменьшающий изображения: сжатие не помогло, файл
        /// в основном из картинок, а выбранный уровень их не пересчитывает. Половина файла —
        /// порог с запасом: у скана доля под единицу, у обычного документа — доли процента.
        /// Чистая — под тест.
        /// </summary>
        public static bool ShouldSuggestDownsampling(CompressionLevel used, double imageShare)
        {
            return imageShare >= 0.5 && !PdfCompression.Downsamples(used);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static double ImageShareCore(string path)
        {
            try
            {
                long total = new FileInfo(path).Length;
                if (total <= 0)
                    return Unknown;
                long images = 0;
                using (UglyToad.PdfPig.PdfDocument doc = PdfPageProbe.OpenPig(path))
                {
                    foreach (UglyToad.PdfPig.Content.Page page in doc.GetPages())
                    {
                        // Страница с испорченными ресурсами не должна отменять весь подсчёт:
                        // пропускаем её и считаем по остальным.
                        try
                        {
                            foreach (UglyToad.PdfPig.Content.IPdfImage img in page.GetImages())
                                images += img.RawBytes.Length; // длина без копирования байтов
                        }
                        catch { }
                    }
                }
                return (double)images / total;
            }
            catch { return Unknown; }
        }
    }
}
