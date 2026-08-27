using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Data.Pdf;
using Windows.Storage.Streams;

namespace ExcelMerger
{
    /// <summary>
    /// Рендер PDF системным Windows.Data.Pdf. Документы грузятся только из памяти, чтобы
    /// WinRT не оставлял source memory-mapped. Экземпляр однопоточен; документы — bounded LRU,
    /// а каждый raster допускается общим process-wide бюджетом ДО native/decode allocations.
    /// </summary>
    public sealed class PdfThumbnailRenderer : IDisposable
    {
        private static readonly int MaxCachedDocuments = IntPtr.Size == 8 ? 6 : 3;
        private readonly LruCache<CachedDoc> _docs =
            new LruCache<CachedDoc>(MaxCachedDocuments, ReleaseCached);
        private bool _disposed;

        public Bitmap Render(string path, int pageIndex, int targetWidth)
        {
            return Render(path, pageIndex, targetWidth, 0, RasterBudget.DefaultRenderPixels);
        }

        public Bitmap Render(string path, int pageIndex, int targetWidth, int maxHeight)
        {
            return Render(path, pageIndex, targetWidth, maxHeight,
                RasterBudget.DefaultRenderPixels);
        }

        public Bitmap Render(string path, int pageIndex, int targetWidth,
            int maxHeight, long maxPixels)
        {
            BudgetedBitmap owned = RenderOwned(path, pageIndex, targetWidth,
                maxHeight, maxPixels);
            return owned == null ? null : owned.DetachUnbudgeted();
        }

        /// <summary>Внутренний retaining-контракт: один owner для bitmap и memory lease.</summary>
        internal BudgetedBitmap RenderOwned(string path, int pageIndex, int targetWidth,
            int maxHeight, long maxPixels)
        {
            if (_disposed || string.IsNullOrWhiteSpace(path) || targetWidth <= 0 ||
                maxPixels <= 0)
                return null;
            PdfMemoryLease working = null;
            Bitmap result = null;
            try
            {
                PdfDocument doc = GetDocument(path);
                if (doc == null || pageIndex < 0 || pageIndex >= (int)doc.PageCount)
                    return null;
                using (PdfPage page = doc.GetPage((uint)pageIndex))
                {
                    Windows.Foundation.Size size = page.Size;
                    int width = RasterBudget.FitWidth(targetWidth, size.Width,
                        size.Height, maxPixels, maxHeight);
                    if (width <= 0)
                        return null;
                    int expectedHeight = RasterBudget.ExpectedHeight(width,
                        size.Height / size.Width);
                    long peak = RasterBudget.BitmapWorkingSetBytes(width, expectedHeight, 3);
                    if (!PdfMemoryBudget.TryAcquire(peak, out working))
                        return null;

                    using (var ras = new InMemoryRandomAccessStream())
                    {
                        var options = new PdfPageRenderOptions
                        {
                            DestinationWidth = (uint)width
                        };
                        page.RenderToStreamAsync(ras, options).AsTask().GetAwaiter().GetResult();
                        using (Stream managed = ras.AsStreamForRead())
                        using (var decoded = new Bitmap(managed))
                        {
                            if (!RasterBudget.IsWithin(decoded.Width, decoded.Height, maxPixels))
                                return null;
                            result = new Bitmap(decoded);
                        }
                    }
                }

                long finalBytes = PdfMemoryBudget.EstimateBitmapBytes(result.Width, result.Height);
                if (finalBytes > working.Bytes && !working.TryGrow(finalBytes - working.Bytes))
                {
                    result.Dispose();
                    result = null;
                    return null;
                }
                working.ReduceTo(finalBytes);
                var owned = new BudgetedBitmap(result, working);
                result = null;
                working = null;
                return owned;
            }
            catch (OutOfMemoryException)
            {
                if (result != null) result.Dispose();
                throw;
            }
            catch
            {
                if (result != null) result.Dispose();
                return null;
            }
            finally
            {
                if (working != null)
                    working.Dispose();
            }
        }

        internal static int FitWidth(int targetWidth, double pageWidth, double pageHeight,
            int maxHeight)
        {
            if (targetWidth <= 0)
                return targetWidth;
            if (pageWidth <= 0 || pageHeight <= 0 ||
                double.IsNaN(pageWidth) || double.IsNaN(pageHeight) ||
                double.IsInfinity(pageWidth) || double.IsInfinity(pageHeight))
                return Math.Min(targetWidth, RasterBudget.MaxRenderDimension);
            return RasterBudget.FitWidth(targetWidth, pageWidth, pageHeight,
                long.MaxValue, maxHeight);
        }

        private PdfDocument GetDocument(string path)
        {
            string key = Path.GetFullPath(path);
            CachedDoc cached;
            if (_docs.TryGet(key, out cached))
                return cached.Doc;

            long reserved = PdfMemoryBudget.EstimateDocumentBytes(new FileInfo(key).Length);
            PdfMemoryLease lease;
            while (!PdfMemoryBudget.TryAcquire(reserved, out lease))
                if (!_docs.TryEvictOldest())
                    return null;

            InMemoryRandomAccessStream stream = null;
            PdfDocument document = null;
            try
            {
                stream = new InMemoryRandomAccessStream();
                byte[] bytes = File.ReadAllBytes(key);
                long actual = PdfMemoryBudget.EstimateDocumentBytes(bytes.LongLength);
                if (actual > lease.Bytes)
                {
                    long extra = actual - lease.Bytes;
                    while (!lease.TryGrow(extra))
                        if (!_docs.TryEvictOldest())
                            return null;
                }
                else
                    lease.ReduceTo(actual);

                stream.WriteAsync(bytes.AsBuffer()).AsTask().GetAwaiter().GetResult();
                stream.Seek(0);
                string password = PdfPasswords.For(key);
                document = string.IsNullOrEmpty(password)
                    ? PdfDocument.LoadFromStreamAsync(stream).AsTask().GetAwaiter().GetResult()
                    : PdfDocument.LoadFromStreamAsync(stream, password).AsTask().GetAwaiter().GetResult();
                var entry = new CachedDoc { Doc = document, Stream = stream, Lease = lease };
                _docs.Add(key, entry);
                document = null;
                stream = null;
                lease = null;
                return entry.Doc;
            }
            finally
            {
                if (document != null) ComSafe.Release(document);
                try { if (stream != null) stream.Dispose(); } catch { }
                if (lease != null) lease.Dispose();
            }
        }

        public void Dispose()
        {
            if (_disposed)
                return;
            _disposed = true;
            _docs.Clear();
        }

        private sealed class CachedDoc
        {
            public PdfDocument Doc;
            public InMemoryRandomAccessStream Stream;
            public PdfMemoryLease Lease;
        }

        private static void ReleaseCached(CachedDoc cached)
        {
            if (cached == null)
                return;
            ComSafe.Release(cached.Doc);
            try { if (cached.Stream != null) cached.Stream.Dispose(); } catch { }
            if (cached.Lease != null) cached.Lease.Dispose();
            cached.Lease = null;
        }
    }
}
