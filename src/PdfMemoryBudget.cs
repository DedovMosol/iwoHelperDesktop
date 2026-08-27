using System;
using System.Threading;

namespace ExcelMerger
{
    /// <summary>
    /// Process-wide admission budget for PDF byte copies, WinRT source streams and page
    /// rasters. Lease ownership makes every reservation single-release and keeps transient
    /// allocations accounted before they happen, not after an OOM-prone render.
    /// </summary>
    internal static class PdfMemoryBudget
    {
        internal static readonly long LimitBytes = IntPtr.Size == 8 ? 768L << 20 : 128L << 20;
        private static readonly object Gate = new object();
        private static long _used;
        private static int _waiters;
        internal static event Action MemoryReleased;

        internal static long Used { get { lock (Gate) return _used; } }

        internal static long EstimateBitmapBytes(int width, int height)
        {
            if (width <= 0 || height <= 0)
                return 0;
            long pixels = (long)width * height;
            return pixels > long.MaxValue / 4 ? long.MaxValue : pixels * 4;
        }

        internal static long EstimateDocumentBytes(long fileBytes)
        {
            if (fileBytes <= 0)
                return 0;
            return fileBytes > long.MaxValue / 2 ? long.MaxValue : fileBytes * 2;
        }

        internal static bool TryAcquire(long bytes, out PdfMemoryLease lease)
        {
            lease = null;
            if (bytes <= 0 || bytes > LimitBytes)
                return false;
            lock (Gate)
            {
                if (bytes > LimitBytes - _used)
                {
                    Interlocked.Exchange(ref _waiters, 1);
                    return false;
                }
                _used += bytes;
                lease = new PdfMemoryLease(bytes);
                return true;
            }
        }

        internal static bool TryGrow(PdfMemoryLease lease, long additional)
        {
            if (lease == null || additional <= 0)
                return additional == 0;
            lock (Gate)
            {
                if (lease.Disposed)
                    return false;
                if (additional > LimitBytes - _used)
                {
                    Interlocked.Exchange(ref _waiters, 1);
                    return false;
                }
                _used += additional;
                lease.BytesCore += additional;
                return true;
            }
        }

        internal static void Reduce(PdfMemoryLease lease, long keepBytes)
        {
            if (lease == null)
                return;
            bool released = false;
            lock (Gate)
            {
                if (lease.Disposed)
                    return;
                long keep = Math.Max(0, Math.Min(keepBytes, lease.BytesCore));
                long difference = lease.BytesCore - keep;
                if (difference > 0)
                {
                    lease.BytesCore = keep;
                    _used -= difference;
                    released = true;
                }
            }
            if (released)
                RaiseReleased();
        }

        internal static void DisposeLease(PdfMemoryLease lease)
        {
            if (lease == null)
                return;
            bool released = false;
            lock (Gate)
            {
                if (lease.Disposed)
                    return;
                lease.Disposed = true;
                if (lease.BytesCore > 0)
                {
                    _used -= lease.BytesCore;
                    lease.BytesCore = 0;
                    released = true;
                }
            }
            if (released)
                RaiseReleased();
        }

        private static void RaiseReleased()
        {
            if (Interlocked.Exchange(ref _waiters, 0) == 0)
                return;
            Action handler = MemoryReleased;
            if (handler != null)
                try { handler(); } catch { }
        }
    }

    internal sealed class PdfMemoryLease : IDisposable
    {
        internal long BytesCore;
        internal bool Disposed;

        internal PdfMemoryLease(long bytes) { BytesCore = bytes; }

        internal long Bytes { get { return BytesCore; } }

        internal bool TryGrow(long additional)
        {
            return PdfMemoryBudget.TryGrow(this, additional);
        }

        internal void ReduceTo(long bytes)
        {
            PdfMemoryBudget.Reduce(this, bytes);
        }

        ~PdfMemoryLease()
        {
            Dispose();
        }

        public void Dispose()
        {
            PdfMemoryBudget.DisposeLease(this);
            GC.SuppressFinalize(this);
        }
    }
}
