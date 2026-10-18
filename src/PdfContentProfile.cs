using System;
using System.Collections.Generic;
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
                // Байты, уже учтённые: картинка, общая для многих страниц (логотип в шапке),
                // лежит в файле ОДИН раз, и считать её надо один раз. Иначе логотип на сотне
                // страниц тянул долю выше единицы, и текстовому документу советовали
                // пересчитать изображения, которых в нём почти нет. Номера объекта PdfPig не
                // отдаёт, а байты разбирает заново на каждой странице (буферы разные — проверено),
                // поэтому объект узнаётся по отпечатку байтов. Две отдельно вложенные одинаковые
                // картинки сольются в одну — недооценка в безопасную сторону: совет не возникнет
                // зря, а дубликаты картинок пересчёт разрешения всё равно не уменьшает.
                var seen = new HashSet<Tuple<int, int, int, long>>();
                using (UglyToad.PdfPig.PdfDocument doc = PdfPageProbe.OpenPig(path))
                {
                    foreach (UglyToad.PdfPig.Content.Page page in doc.GetPages())
                    {
                        // Страница с испорченными ресурсами не должна отменять весь подсчёт:
                        // пропускаем её и считаем по остальным.
                        try
                        {
                            foreach (UglyToad.PdfPig.Content.IPdfImage img in page.GetImages())
                            {
                                if (!img.IsInlineImage && !seen.Add(Fingerprint(img)))
                                    continue; // этот объект уже посчитан на другой странице
                                images += img.RawMemory.Length; // длина без копирования байтов
                            }
                        }
                        catch { }
                    }
                }
                // Страховка на случай, если библиотека отдаст общий объект разными буферами:
                // доля по смыслу не больше единицы, и совет не должен держаться на переборе.
                return Math.Min(1.0, (double)images / total);
            }
            catch { return Unknown; }
        }

        /// <summary>
        /// Отпечаток картинки: размеры, длина и выборка из 64 равномерно взятых байтов.
        /// Дешёвый — не читает картинку целиком (многомегабайтный скан на каждой странице
        /// хэшировать незачем), а у разных картинок совпасть всем этим сразу практически
        /// невозможно.
        /// </summary>
        private static Tuple<int, int, int, long> Fingerprint(UglyToad.PdfPig.Content.IPdfImage img)
        {
            ReadOnlySpan<byte> bytes = img.RawMemory.Span;
            const int samples = 64;
            long hash = unchecked((long)14695981039346656037UL); // FNV-1a
            for (int i = 0; i < samples && bytes.Length > 0; i++)
            {
                hash ^= bytes[(int)((long)i * bytes.Length / samples)];
                hash = unchecked(hash * 1099511628211L);
            }
            return Tuple.Create(bytes.Length, img.WidthInSamples, img.HeightInSamples, hash);
        }
    }
}
