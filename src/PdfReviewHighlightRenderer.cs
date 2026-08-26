using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace ExcelMerger
{
    /// <summary>
    /// Word-подобная заливка authoritative word-box. Меняет только нейтральную бумагу и
    /// светлые серые края glyph; тёмный/chromatic ink и alpha остаются исходными.
    /// </summary>
    internal static class PdfReviewHighlightRenderer
    {
        internal const int DarkInkMaximum = 150;
        internal const int PaperMinimum = 238;
        internal const int NeutralChromaTolerance = 18;

        private struct PixelInterval
        {
            public int Left;
            public int Right;

            public PixelInterval(int left, int right)
            {
                Left = left;
                Right = right;
            }
        }

        internal static Bitmap Create32BppCopy(Bitmap source)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            var copy = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
            try
            {
                if (source.HorizontalResolution > 0 && source.VerticalResolution > 0)
                    copy.SetResolution(source.HorizontalResolution, source.VerticalResolution);
                using (Graphics graphics = Graphics.FromImage(copy))
                {
                    graphics.CompositingMode =
                        System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                    graphics.DrawImageUnscaled(source, 0, 0);
                }
                return copy;
            }
            catch
            {
                copy.Dispose();
                throw;
            }
        }

        internal static void ApplyWordFills(Bitmap bitmap,
            IList<RectangleF> wordRectangles, Color fill)
        {
            if (bitmap == null)
                throw new ArgumentNullException("bitmap");
            if (wordRectangles == null || wordRectangles.Count == 0 ||
                bitmap.Width <= 0 || bitmap.Height <= 0)
                return;

            List<PixelInterval>[] rows = BuildIntervals(wordRectangles,
                bitmap.Width, bitmap.Height);
            BitmapData data = null;
            try
            {
                data = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height),
                    ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
                int bytes = checked(bitmap.Width * 4);
                var row = new byte[bytes];
                for (int y = 0; y < rows.Length; y++)
                {
                    List<PixelInterval> intervals = rows[y];
                    if (intervals == null || intervals.Count == 0)
                        continue;
                    MergeIntervals(intervals);
                    IntPtr address = IntPtr.Add(data.Scan0, checked(y * data.Stride));
                    Marshal.Copy(address, row, 0, bytes);
                    foreach (PixelInterval interval in intervals)
                    {
                        for (int x = interval.Left; x < interval.Right; x++)
                        {
                            int offset = x * 4;
                            ApplyPixel(row, offset, fill);
                        }
                    }
                    Marshal.Copy(row, 0, address, bytes);
                }
            }
            finally
            {
                if (data != null)
                    bitmap.UnlockBits(data);
            }
        }

        internal static Color ComposePixel(Color source, Color fill)
        {
            byte[] pixel = { source.B, source.G, source.R, source.A };
            ApplyPixel(pixel, 0, fill);
            return Color.FromArgb(pixel[3], pixel[2], pixel[1], pixel[0]);
        }

        private static void ApplyPixel(byte[] row, int offset, Color fill)
        {
            int blue = row[offset];
            int green = row[offset + 1];
            int red = row[offset + 2];
            int alpha = row[offset + 3];
            if (alpha != 255)
                return;

            int maximum = Math.Max(red, Math.Max(green, blue));
            int minimum = Math.Min(red, Math.Min(green, blue));
            if (maximum <= DarkInkMaximum || maximum - minimum > NeutralChromaTolerance)
                return;

            if (minimum >= PaperMinimum)
            {
                row[offset] = fill.B;
                row[offset + 1] = fill.G;
                row[offset + 2] = fill.R;
                return;
            }

            // Нейтральный edge рассматривается как black ink с fractional coverage над
            // бумагой: тот же coverage накладывается на новый фон, без белого ореола.
            int luminance = (red * 54 + green * 183 + blue * 19 + 128) >> 8;
            row[offset] = Multiply(fill.B, luminance);
            row[offset + 1] = Multiply(fill.G, luminance);
            row[offset + 2] = Multiply(fill.R, luminance);
        }

        private static byte Multiply(int channel, int factor)
        {
            return (byte)((channel * factor + 127) / 255);
        }

        private static List<PixelInterval>[] BuildIntervals(IList<RectangleF> rectangles,
            int width, int height)
        {
            var rows = new List<PixelInterval>[height];
            foreach (RectangleF value in rectangles)
            {
                if (!IsFinite(value.Left) || !IsFinite(value.Top) ||
                    !IsFinite(value.Right) || !IsFinite(value.Bottom) ||
                    value.Width <= 0 || value.Height <= 0)
                    continue;
                int left = PixelFloor(value.Left, width);
                int top = PixelFloor(value.Top, height);
                int right = PixelCeiling(value.Right, width);
                int bottom = PixelCeiling(value.Bottom, height);
                if (right <= left || bottom <= top)
                    continue;
                for (int y = top; y < bottom; y++)
                {
                    if (rows[y] == null)
                        rows[y] = new List<PixelInterval>();
                    rows[y].Add(new PixelInterval(left, right));
                }
            }
            return rows;
        }

        private static void MergeIntervals(List<PixelInterval> intervals)
        {
            if (intervals.Count <= 1)
                return;
            intervals.Sort(delegate(PixelInterval left, PixelInterval right)
            {
                int byLeft = left.Left.CompareTo(right.Left);
                return byLeft != 0 ? byLeft : left.Right.CompareTo(right.Right);
            });
            int output = 0;
            for (int i = 1; i < intervals.Count; i++)
            {
                PixelInterval current = intervals[output];
                PixelInterval next = intervals[i];
                if (next.Left < current.Right)
                {
                    if (next.Right > current.Right)
                        current.Right = next.Right;
                    intervals[output] = current;
                }
                else
                {
                    output++;
                    intervals[output] = next;
                }
            }
            if (output + 1 < intervals.Count)
                intervals.RemoveRange(output + 1, intervals.Count - output - 1);
        }

        private static int PixelFloor(float value, int limit)
        {
            if (value <= 0f)
                return 0;
            if (value >= limit)
                return limit;
            return (int)Math.Floor(value);
        }

        private static int PixelCeiling(float value, int limit)
        {
            if (value <= 0f)
                return 0;
            if (value >= limit)
                return limit;
            return (int)Math.Ceiling(value);
        }

        private static bool IsFinite(float value)
        {
            return !float.IsNaN(value) && !float.IsInfinity(value);
        }
    }
}
