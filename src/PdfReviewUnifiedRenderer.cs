using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Single-document redline compositor: later/right page is the base, additions stay
    /// green, and only deleted fragments from earlier/left are projected in red. Unchanged
    /// pixels from the earlier page never enter the result.
    /// </summary>
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
            using (var fill = new SolidBrush(Color.FromArgb(
                SystemInformation.HighContrast ? 36 : 86, color)))
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
                    graphics.FillRectangle(fill, rect);
                    graphics.DrawRectangle(outline, rect.X, rect.Y,
                        rect.Width, rect.Height);
                    float y = rect.Top + rect.Height * 0.56f;
                    graphics.DrawLine(strike, rect.Left, y, rect.Right, y);
                    mapped.Add(rect);
                }
                DrawWhitespace(graphics, target.Size, target.Size, deleted, edge);
                DrawRails(graphics, target.Size, mapped, edge);
            }
        }

        internal static void OverlayDeletedFragments(Bitmap target, Bitmap earlier,
            PdfReviewHighlight deleted)
        {
            if (target == null || earlier == null || deleted == null ||
                deleted.ViewWidthPt <= 0 || deleted.ViewHeightPt <= 0)
                return;
            Color color, edge;
            PdfReviewPageView.ResolveHighlightColors(deleted,
                SystemInformation.HighContrast, out color, out edge);
            var mapped = new List<RectangleF>();
            using (Graphics graphics = Graphics.FromImage(target))
            using (var attributes = OverlayAttributes(SystemInformation.HighContrast ? 1f : 0.86f))
            using (var outline = new Pen(Color.FromArgb(255, edge), 2f))
            using (var strike = new Pen(Color.FromArgb(255, color), 1.5f))
            {
                graphics.CompositingMode = CompositingMode.SourceOver;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                foreach (PdfReviewBox box in deleted.Boxes)
                {
                    RectangleF source = PdfReviewGeometry.ToPixelRect(box,
                        deleted.ViewWidthPt, deleted.ViewHeightPt,
                        earlier.Width, earlier.Height);
                    source = Clamp(source, earlier.Size);
                    RectangleF destination = Scale(source, earlier.Size, target.Size);
                    if (!Usable(source, earlier.Size) || !Usable(destination, target.Size))
                        continue;
                    Rectangle sourceInt = Rectangle.Round(source);
                    RectangleF targetRect = Clamp(destination, target.Size);
                    if (sourceInt.Width < 1 || sourceInt.Height < 1 ||
                        targetRect.Width < 1 || targetRect.Height < 1)
                        continue;
                    graphics.DrawImage(earlier, Rectangle.Round(targetRect), sourceInt.X,
                        sourceInt.Y, sourceInt.Width, sourceInt.Height,
                        GraphicsUnit.Pixel, attributes);
                    graphics.DrawRectangle(outline, targetRect.X, targetRect.Y,
                        targetRect.Width, targetRect.Height);
                    float y = targetRect.Top + targetRect.Height * 0.56f;
                    graphics.DrawLine(strike, targetRect.Left, y, targetRect.Right, y);
                    mapped.Add(targetRect);
                }
                DrawWhitespace(graphics, target.Size, earlier.Size, deleted, edge);
                DrawRails(graphics, target.Size, mapped, edge);
            }
        }

        private static ImageAttributes OverlayAttributes(float alpha)
        {
            var attributes = new ImageAttributes();
            var matrix = new ColorMatrix { Matrix33 = alpha };
            attributes.SetColorMatrix(matrix, ColorMatrixFlag.Default,
                ColorAdjustType.Bitmap);
            return attributes;
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

        private static RectangleF Clamp(RectangleF rect, Size bounds)
        {
            float left = Math.Max(0, rect.Left), top = Math.Max(0, rect.Top);
            float right = Math.Min(bounds.Width, rect.Right);
            float bottom = Math.Min(bounds.Height, rect.Bottom);
            return right > left && bottom > top
                ? RectangleF.FromLTRB(left, top, right, bottom) : RectangleF.Empty;
        }

        private static void DrawWhitespace(Graphics graphics, Size target, Size earlier,
            PdfReviewHighlight deleted, Color color)
        {
            if (deleted.WhitespaceMarkers == null)
                return;
            using (var font = new Font(FontFamily.GenericSansSerif,
                Math.Max(9f, Math.Min(15f, target.Width / 80f)), FontStyle.Bold,
                GraphicsUnit.Pixel))
            using (var brush = new SolidBrush(color))
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
