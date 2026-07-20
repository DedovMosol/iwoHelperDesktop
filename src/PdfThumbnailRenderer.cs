using System;
using System.Drawing;
using System.IO;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;

namespace ExcelMerger
{
    /// <summary>
    /// Рендер миниатюр страниц PDF системным движком Windows.Data.Pdf (WinRT).
    /// Вся зависимость от WinRT изолирована здесь; любой сбой (старая Windows,
    /// защищённый/битый файл, недоступный API) даёт null — вызывающий код
    /// показывает страницу без картинки, приложение не падает.
    /// Экземпляр не потокобезопасен: вызывать из одного (фонового) потока;
    /// открытые документы держатся в ограниченном LRU-кэше (снимает рост памяти и
    /// файловых хэндлов при переборе файлов) и освобождаются при вытеснении и в Dispose.
    /// </summary>
    public sealed class PdfThumbnailRenderer : IDisposable
    {
        // WinRT PdfDocument держит файл и нативные буферы. Инструменты открывают
        // документы последовательно (в «Разделении» — по одному; в «Объединении»
        // видимое окно охватывает обычно 1–2 файла), поэтому небольшой LRU не вызывает
        // перерендера, но ограничивает память и число открытых файлов.
        private const int MaxCachedDocuments = 6;

        private readonly LruCache<PdfDocument> _docs =
            new LruCache<PdfDocument>(MaxCachedDocuments, ComSafe.Release);
        private bool _disposed;

        /// <summary>
        /// Миниатюра страницы шириной targetWidth пикселей (высота пропорциональна).
        /// null — отрендерить не удалось. Bitmap принадлежит вызывающему.
        /// </summary>
        public Bitmap Render(string path, int pageIndex, int targetWidth)
        {
            if (_disposed)
                return null;
            try
            {
                PdfDocument doc = GetDocument(path);
                if (doc == null || pageIndex < 0 || pageIndex >= (int)doc.PageCount)
                    return null;
                using (PdfPage page = doc.GetPage((uint)pageIndex))
                using (var ras = new InMemoryRandomAccessStream())
                {
                    var opts = new PdfPageRenderOptions();
                    opts.DestinationWidth = (uint)targetWidth;
                    page.RenderToStreamAsync(ras, opts).AsTask().GetAwaiter().GetResult();
                    using (Stream managed = ras.AsStreamForRead())
                    using (var decoded = new Bitmap(managed))
                        return new Bitmap(decoded); // копия, независимая от потока
                }
            }
            catch
            {
                return null; // страница без миниатюры — не причина падать
            }
        }

        private PdfDocument GetDocument(string path)
        {
            string key = Path.GetFullPath(path);
            PdfDocument doc;
            if (_docs.TryGet(key, out doc))
                return doc;
            StorageFile file = StorageFile.GetFileFromPathAsync(key).AsTask().GetAwaiter().GetResult();
            doc = PdfDocument.LoadFromFileAsync(file).AsTask().GetAwaiter().GetResult();
            _docs.Add(key, doc); // при переполнении самый несвежий документ будет освобождён (ComSafe.Release)
            return doc;
        }

        public void Dispose()
        {
            _disposed = true;
            // Детерминированно освобождаем COM-обёртки WinRT, не полагаясь на финализатор
            // (иначе возможен сбой при выгрузке процесса). Тот же поток/апартамент, где
            // документы создавались (см. PdfPageGrid.ThumbWorker).
            _docs.Clear();
        }
    }
}
