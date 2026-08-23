using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace ExcelMerger
{
    /// <summary>Формат картинок при экспорте страниц.</summary>
    public enum ImageExportFormat
    {
        Png = 0, // без потерь, крупнее — под дальнейшую обработку
        Jpeg = 1 // с потерями, компактнее — под пересылку и вставку в документы
    }

    /// <summary>
    /// Экспорт страниц PDF наружу: картинками и простым текстом.
    ///
    /// Новых зависимостей не заводит — обе половины переиспользуют то, что уже работает:
    /// картинки рисует тот же <see cref="PdfThumbnailRenderer"/>, что и миниатюры сетки,
    /// размеры листов даёт тот же <see cref="PdfMergeService.LoadPages"/>, что и объединение,
    /// текст — тот же <see cref="PdfTextExtract"/>, что и «PDF → Word».
    ///
    /// UI здесь нет: прогресс и отмена приходят колбэками, как у остальных сервисов.
    /// Вызывать с фонового потока.
    /// </summary>
    public static class PdfExportService
    {
        /// <summary>Разрешения на выбор: экран, печать, высокое, очень высокое.</summary>
        public static readonly int[] DpiChoices = { 96, 150, 300, 600 };

        /// <summary>
        /// Качество JPEG. GDI+ по умолчанию берёт 75, и на сканах это видно кольцами вокруг
        /// букв, поэтому задаём явно: 90 — общепринятая граница, за которой файл растёт
        /// быстрее, чем различимое качество.
        /// </summary>
        private const long JpegQuality = 90L;

        /// <summary>Ширина листа, когда настоящую узнать не удалось: A4 в пунктах.</summary>
        private const double DefaultWidthPt = 595.28;

        /// <summary>
        /// Ширина страницы в пикселях для заданного разрешения: ширина листа приходит в
        /// пунктах (1/72 дюйма). Меньше пикселя страниц не бывает. Чистая — под тест.
        /// </summary>
        internal static int PixelWidth(double widthPt, int dpi)
        {
            if (widthPt <= 0 || dpi <= 0)
                return 1;
            int px = (int)Math.Round(widthPt / 72.0 * dpi, MidpointRounding.AwayFromZero);
            return px < 1 ? 1 : px;
        }

        /// <summary>Расширение файла для формата. Чистая — под тест.</summary>
        internal static string Extension(ImageExportFormat format)
        {
            return format == ImageExportFormat.Jpeg ? ".jpg" : ".png";
        }

        /// <summary>
        /// Сохранить указанные страницы картинками в папку, вернуть список созданных файлов.
        /// Имя каждого собирает <see cref="NameTemplate"/>.
        ///
        /// Ширина считается от РЕАЛЬНОГО размера каждой страницы: в одном документе бывают
        /// листы разного формата, и общая ширина растянула бы альбомную вставку.
        ///
        /// Рендерер создаётся и освобождается ВНУТРИ вызова, на одном потоке: он держит
        /// открытые документы WinRT, и трогать их с другого потока нельзя.
        ///
        /// Здесь файл создаётся на КАЖДУЮ страницу, поэтому отмена — не только исключение:
        /// уже сохранённые картинки удаляются (<see cref="Cancellation.NoPartialOutput"/>),
        /// иначе «Отменено» оставляло бы в папке половину экспорта.
        /// </summary>
        public static List<string> ToImages(string sourcePath, IList<int> pageIndexes, string outDir,
            string template, ImageExportFormat format, int dpi,
            Action<int, int> progress = null, Func<bool> cancelled = null)
        {
            if (pageIndexes == null || pageIndexes.Count == 0)
                throw new MergeException(Loc.T("err.export.noPages"));
            Directory.CreateDirectory(outDir);

            // Ширины листов в пунктах. Отдельный разбор файла — вспомогательный: движок
            // отрисовки открывает и то, что не открывает разбор (файл после чужой программы,
            // защищённый от изменения). Ронять из-за этого сохранение картинок нельзя —
            // при неудаче считаем лист обычным A4, соотношение сторон всё равно даст рендер.
            List<PdfPageInfo> sizes;
            try { sizes = PdfMergeService.LoadPages(sourcePath); }
            catch { sizes = new List<PdfPageInfo>(); }
            string baseName = Path.GetFileNameWithoutExtension(sourcePath);
            DateTime startedAt = DateTime.Now;

            using (var renderer = new PdfThumbnailRenderer())
            {
                return Cancellation.NoPartialOutput(delegate(List<string> written)
                {
                    for (int i = 0; i < pageIndexes.Count; i++)
                    {
                        Cancellation.ThrowIf(cancelled);
                        int pageIndex = pageIndexes[i];
                        var values = new NameValues
                        {
                            BaseName = baseName,
                            FileNumber = i + 1,
                            TotalFiles = pageIndexes.Count,
                            CurrentPage = pageIndex + 1,
                            Timestamp = startedAt
                        };
                        string path = OutputFile.Unique(outDir,
                            NameTemplate.Apply(template, values), Extension(format));

                        double widthPt = pageIndex >= 0 && pageIndex < sizes.Count && sizes[pageIndex].WidthPt > 0
                            ? sizes[pageIndex].WidthPt : DefaultWidthPt;
                        using (Bitmap bmp = renderer.Render(sourcePath, pageIndex, PixelWidth(widthPt, dpi)))
                        {
                            if (bmp == null)
                                throw new MergeException(string.Format(Loc.T("err.export.pageFailed"), pageIndex + 1));
                            try
                            {
                                Save(bmp, path, format);
                            }
                            catch (Exception ex) when (MergeException.ShouldWrap(ex))
                            {
                                // GDI+ на любую неудачу записи отвечает «произошла общая ошибка»,
                                // поэтому диск здесь особенно нужен: см. DiskSpace.Describe.
                                throw new MergeException(string.Format(Loc.T("err.export.saveFailed"),
                                    Path.GetFileName(path), DiskSpace.Describe(ex, path)));
                            }
                        }
                        written.Add(path);
                        if (progress != null)
                            progress(i + 1, pageIndexes.Count);
                    }
                });
            }
        }

        /// <summary>
        /// Сохранить текстовый слой документа в .txt, вернуть путь к файлу.
        ///
        /// Кодировка — UTF-8 С меткой порядка байтов: без неё «Блокнот» на части систем
        /// читает кириллицу как набор посторонних знаков. Переводы строк приводим к принятым
        /// в Windows, иначе весь текст слипается в одну строку в том же «Блокноте».
        /// </summary>
        public static string ToText(string sourcePath, string outPath,
            Action<int, int> progress = null, Func<bool> cancelled = null)
        {
            if (OutputFile.IsSameFile(sourcePath, outPath))
                throw new MergeException(Loc.T("err.output.sameSource"));
            List<PdfPageText> pages = PdfTextExtract.Extract(sourcePath, progress, null, cancelled);
            Cancellation.ThrowIf(cancelled);
            string text = PlainText.Document(pages).Replace("\n", Environment.NewLine);
            try
            {
                using (var output = new AtomicOutput(outPath))
                {
                    File.WriteAllText(output.TempPath, text, System.Text.Encoding.UTF8);
                    Cancellation.ThrowIf(cancelled);
                    output.Commit();
                }
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                throw new MergeException(string.Format(Loc.T("err.export.saveFailed"),
                    Path.GetFileName(outPath), DiskSpace.Describe(ex, outPath)));
            }
            return outPath;
        }

        /// <summary>
        /// Записать картинку. JPEG идёт через кодировщик с явно заданным качеством, PNG —
        /// напрямую (у него качества нет, сжатие без потерь).
        /// </summary>
        private static void Save(Bitmap bmp, string path, ImageExportFormat format)
        {
            if (format != ImageExportFormat.Jpeg)
            {
                bmp.Save(path, ImageFormat.Png);
                return;
            }
            using (var ps = new EncoderParameters(1))
            {
                ps.Param[0] = new EncoderParameter(Encoder.Quality, JpegQuality);
                bmp.Save(path, JpegCodec, ps);
            }
        }

        // Кодировщик JPEG входит в GDI+ и есть всегда; отдельной ветки «его нет» не держим —
        // сбой любого сохранения и так поднимется наружу и покажется пользователем ошибкой.
        private static readonly ImageCodecInfo JpegCodec = FindJpegCodec();

        private static ImageCodecInfo FindJpegCodec()
        {
            foreach (ImageCodecInfo c in ImageCodecInfo.GetImageEncoders())
                if (c.FormatID == ImageFormat.Jpeg.Guid)
                    return c;
            return null;
        }
    }
}
