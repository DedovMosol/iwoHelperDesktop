using System;
using System.Drawing;

namespace ExcelMerger
{
    /// <summary>Один владелец bitmap и его process-memory lease.</summary>
    internal sealed class BudgetedBitmap : IDisposable
    {
        internal Bitmap Bitmap { get; private set; }
        private PdfMemoryLease _lease;

        internal BudgetedBitmap(Bitmap bitmap, PdfMemoryLease lease)
        {
            if (bitmap == null || lease == null)
                throw new ArgumentNullException(bitmap == null ? "bitmap" : "lease");
            Bitmap = bitmap;
            _lease = lease;
        }

        /// <summary>Legacy short-lived caller получает bitmap, а accounting ownership снимается.</summary>
        internal Bitmap DetachUnbudgeted()
        {
            Bitmap bitmap = Bitmap;
            Bitmap = null;
            if (_lease != null) _lease.Dispose();
            _lease = null;
            return bitmap;
        }

        internal bool TryGrow(long bytes)
        {
            return _lease != null && _lease.TryGrow(bytes);
        }

        internal void ReplaceAfterTransform(Bitmap bitmap)
        {
            Bitmap = bitmap;
            ReduceToBitmap();
        }

        internal void ReduceToBitmap()
        {
            if (_lease != null && Bitmap != null)
                _lease.ReduceTo(PdfMemoryBudget.EstimateBitmapBytes(
                    Bitmap.Width, Bitmap.Height));
        }

        public void Dispose()
        {
            if (Bitmap != null) Bitmap.Dispose();
            Bitmap = null;
            if (_lease != null) _lease.Dispose();
            _lease = null;
        }
    }
}
