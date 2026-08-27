using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// Single-document redline compositor: the later/right page is the only raster. Additions
    /// stay green; deletions are red geometry/strike markers so reflowed source glyphs can never
    /// stack over the later text. Exact earlier content remains available side by side.
    internal static class PdfReviewUnifiedRenderer
    {
        internal static void DrawDeletedMarkers(Bitmap target, PdfReviewHighlight deleted)
        {
            if (target == null || deleted == null || deleted.ViewWidthPt <= 0 ||
                deleted.ViewHeightPt <= 0)
                return;
            Color color, edge;
            PdfReviewPageView.ResolveHighlightColors(deleted,
                SystemInformation.HighContrast, out color, out edge);
            var mapped = new List<RectangleF>();
            using (Graphics graphics = Graphics.FromImage(target))
            using (var outline = new Pen(edge, 2f))
            using (var strike = new Pen(color, 1.5f))
            {
                foreach (PdfReviewBox box in deleted.Boxes)
                {
                    RectangleF rect = PdfReviewGeometry.ToPixelRect(box,
                        deleted.ViewWidthPt, deleted.ViewHeightPt,
                        target.Width, target.Height);
                    rect = Clamp(rect, target.Size);
                    if (rect.Width < 1 || rect.Height < 1)
                        continue;
                    // Deletion geometry remains visible, but old glyphs are never painted over
                    // the later page: reflow and indentation make those coordinate systems diverge.
                    graphics.DrawRectangle(outline, rect.X, rect.Y,
                        rect.Width, rect.Height);
                    float y = rect.Top + rect.Height * 0.56f;
                    graphics.DrawLine(strike, rect.Left, y, rect.Right, y);
                    mapped.Add(rect);
                }
                DrawWhitespace(graphics, target.Size, target.Size, deleted);
                DrawRails(graphics, target.Size, mapped, edge);
            }
        }

        private static RectangleF Clamp(RectangleF rect, Size bounds)
        {
            float left = Math.Max(0, rect.Left), top = Math.Max(0, rect.Top);
            float right = Math.Min(bounds.Width, rect.Right);
            float bottom = Math.Min(bounds.Height, rect.Bottom);
            return right > left && bottom > top
                ? RectangleF.FromLTRB(left, top, right, bottom) : RectangleF.Empty;
        }

        private static RectangleF Scale(RectangleF source, Size from, Size to)
        {
            if (from.Width <= 0 || from.Height <= 0)
                return RectangleF.Empty;
            return new RectangleF(
                source.X * to.Width / from.Width,
                source.Y * to.Height / from.Height,
                source.Width * to.Width / from.Width,
                source.Height * to.Height / from.Height);
        }

        private static bool Usable(RectangleF rect, Size bounds)
        {
            return rect.Width > 0 && rect.Height > 0 && rect.Right > 0 && rect.Bottom > 0 &&
                rect.Left < bounds.Width && rect.Top < bounds.Height;
        }

        private static void DrawWhitespace(Graphics graphics, Size target, Size earlier,
            PdfReviewHighlight deleted)
        {
            if (deleted.WhitespaceMarkers == null)
                return;
            using (var font = new Font(FontFamily.GenericSansSerif,
                Math.Max(9f, Math.Min(15f, target.Width / 80f)), FontStyle.Bold,
                GraphicsUnit.Pixel))
                foreach (PdfReviewWhitespaceMarker marker in deleted.WhitespaceMarkers)
                {
                    if (marker == null || string.IsNullOrEmpty(marker.Text))
                        continue;
                    RectangleF source = PdfReviewGeometry.ToPixelRect(marker.Box,
                        deleted.ViewWidthPt, deleted.ViewHeightPt,
                        earlier.Width, earlier.Height);
                    RectangleF destination = Scale(source, earlier, target);
                    if (!Usable(destination, target))
                        continue;
                    Color color = marker.Style == PdfReviewHighlightStyle.Removed
                        ? Theme.ReviewDeleteMarker : Theme.ReviewInsertMarker;
                    if (SystemInformation.HighContrast)
                    {
                        Color highContrastColor, highContrastEdge;
                        PdfReviewPageView.ResolveHighlightColors(deleted, true,
                            out highContrastColor, out highContrastEdge);
                        color = highContrastEdge;
                    }
                    using (var brush = new SolidBrush(color))
                        graphics.DrawString(marker.Text, font, brush,
                            destination.Left, destination.Top);
                }
        }

        private static void DrawRails(Graphics graphics, Size target,
            IList<RectangleF> rectangles, Color color)
        {
            if (rectangles == null || rectangles.Count == 0)
                return;
            float x = Math.Max(4f, Math.Min(10f, target.Width * 0.008f));
            using (var pen = new Pen(color, 3f))
            using (var font = new Font(FontFamily.GenericSansSerif, 13f,
                FontStyle.Bold, GraphicsUnit.Pixel))
            using (var brush = new SolidBrush(color))
                foreach (RectangleF rect in rectangles)
                {
                    float top = Math.Max(3f, rect.Top - 2f);
                    float bottom = Math.Min(target.Height - 3f, rect.Bottom + 2f);
                    graphics.DrawLine(pen, x, top, x, bottom);
                    graphics.DrawString("−", font, brush, x + 4f,
                        Math.Max(0f, (top + bottom) / 2f - 7f));
                }
        }
    }
}
