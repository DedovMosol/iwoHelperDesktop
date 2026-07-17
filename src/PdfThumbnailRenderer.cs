using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
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
    /// открытые документы кэшируются и освобождаются в Dispose.
    /// </summary>
    public sealed class PdfThumbnailRenderer : IDisposable
    {
        private readonly Dictionary<string, PdfDocument> _docs =
            new Dictionary<string, PdfDocument>(StringComparer.OrdinalIgnoreCase);
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
            PdfDocument doc;
            string key = Path.GetFullPath(path);
            if (_docs.TryGetValue(key, out doc))
                return doc;
            StorageFile file = StorageFile.GetFileFromPathAsync(key).AsTask().GetAwaiter().GetResult();
            doc = PdfDocument.LoadFromFileAsync(file).AsTask().GetAwaiter().GetResult();
            _docs[key] = doc;
            return doc;
        }

        public void Dispose()
        {
            _disposed = true;
            // Детерминированно освобождаем COM-обёртки WinRT, не полагаясь на
            // финализатор (иначе возможен сбой при выгрузке процесса).
            foreach (PdfDocument doc in _docs.Values)
            {
                try
                {
                    if (doc != null && Marshal.IsComObject(doc))
                        Marshal.FinalReleaseComObject(doc);
                }
                catch { }
            }
            _docs.Clear();
        }
    }
}
