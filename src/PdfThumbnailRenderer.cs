using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices.WindowsRuntime; // AsBuffer: byte[] -> IBuffer
using Windows.Data.Pdf;
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
    ///
    /// ВАЖНО: документы грузятся ИЗ КОПИИ В ПАМЯТИ, а не напрямую из файла
    /// (<see cref="GetDocument"/>). Иначе WinRT держит файл отображённым в память
    /// (user-mapped section), и тогда объединение/разделение не может сохранить итог
    /// ПОД ИМЕНЕМ файла, показанного в сетке миниатюр (ERROR_USER_MAPPED_FILE).
    /// </summary>
    public sealed class PdfThumbnailRenderer : IDisposable
    {
        // WinRT PdfDocument держит нативные буферы (и поток-источник байтов). Инструменты
        // открывают документы последовательно (в «Разделении» — по одному; в «Объединении»
        // видимое окно охватывает обычно 1–2 файла), поэтому небольшой LRU не вызывает
        // перерендера, но ограничивает память и число открытых документов.
        private const int MaxCachedDocuments = 6;

        private readonly LruCache<CachedDoc> _docs =
            new LruCache<CachedDoc>(MaxCachedDocuments, ReleaseCached);
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
            CachedDoc cached;
            if (_docs.TryGet(key, out cached))
                return cached.Doc;
            // Читаем файл в память и грузим WinRT ИЗ ПОТОКА, а не из файла. LoadFromFileAsync держит
            // файл отображённым в память (user-mapped section) на всё время, пока документ в кэше, и
            // тогда сохранить итог объединения/разделения ПОД ИМЕНЕМ показанного в сетке файла
            // невозможно (ERROR_USER_MAPPED_FILE). Копия в памяти снимает блокировку файла на диске;
            // поток держим живым, пока жив документ (WinRT читает страницы из него лениво).
            byte[] bytes = File.ReadAllBytes(key);
            var ras = new InMemoryRandomAccessStream();
            try
            {
                ras.WriteAsync(bytes.AsBuffer()).AsTask().GetAwaiter().GetResult();
                ras.Seek(0);
                PdfDocument doc = PdfDocument.LoadFromStreamAsync(ras).AsTask().GetAwaiter().GetResult();
                _docs.Add(key, new CachedDoc { Doc = doc, Stream = ras }); // вытеснение освободит оба
                return doc;
            }
            catch
            {
                ras.Dispose(); // документ не создан — поток не попал в кэш, освобождаем сразу
                throw;
            }
        }

        public void Dispose()
        {
            _disposed = true;
            // Детерминированно освобождаем COM-обёртки WinRT и потоки, не полагаясь на финализатор
            // (иначе возможен сбой при выгрузке процесса). Тот же поток/апартамент, где
            // документы создавались (см. PdfPageGrid.ThumbWorker).
            _docs.Clear();
        }

        /// <summary>Документ WinRT и поток-источник его байтов (файл на диске не мапится): освобождаем вместе.</summary>
        private sealed class CachedDoc
        {
            public PdfDocument Doc;
            public InMemoryRandomAccessStream Stream;
        }

        /// <summary>Освободить COM-обёртку документа и его поток (при вытеснении из LRU или в Dispose).</summary>
        private static void ReleaseCached(CachedDoc c)
        {
            if (c == null)
                return;
            ComSafe.Release(c.Doc);
            try { if (c.Stream != null) c.Stream.Dispose(); }
            catch { }
        }
    }
}
