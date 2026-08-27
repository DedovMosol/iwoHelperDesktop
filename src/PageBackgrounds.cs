using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace ExcelMerger
{
    internal sealed class BackgroundRenderReport
    {
        public int Added;
        public int Failed;
        public bool EngineMissing;
        public bool BudgetExhausted;
    }

    internal sealed class BackgroundRenderResult : IDisposable
    {
        internal readonly Background[] Items;
        internal readonly BackgroundRenderReport Report;

        internal BackgroundRenderResult(Background[] items, BackgroundRenderReport report)
        {
            Items = items;
            Report = report;
        }

        public void Dispose()
        {
            PageBackgrounds.Release(Items);
        }
    }

    /// <summary>
    /// Подложки слайдов: PDF без текста. Рендер ограничен по диапазонам, пикселям,
    /// временному месту и process memory; duplicate page/rotation кодируется один раз.
    /// Ошибка отдельной страницы оставляет текстовый слайд, отмена останавливает GS сразу.
    /// </summary>
    internal static class PageBackgrounds
    {
        internal static int Dpi(int pageCount)
        {
            if (pageCount <= 50) return 150;
            if (pageCount <= 150) return 120;
            return 96;
        }

        private const long MediaBudgetBytes = 64L * 1024 * 1024;
        private const int MaxRunsPerSource = 64;

        internal static BackgroundRenderResult Render(IList<PdfPageRef> order,
            Action<int, int> progress, Func<bool> cancelled)
        {
            var report = new BackgroundRenderReport();
            if (order == null || order.Count == 0)
                return new BackgroundRenderResult(new Background[0], report);
            var result = new Background[order.Count];
            if (!Ghostscript.Available)
            {
                report.EngineMissing = true;
                return new BackgroundRenderResult(result, report);
            }

            int requestedDpi = Dpi(order.Count);
            long spent = 0;
            string root = Path.Combine(Path.GetTempPath(),
                "iwo_bg_" + Guid.NewGuid().ToString("N"));
            int done = 0;
            int sourceNo = 0;
            bool completed = false;
            try
            {
                foreach (KeyValuePair<string, List<int>> source in GroupBySource(order))
                {
                    Cancellation.ThrowIf(cancelled);
                    if (spent >= MediaBudgetBytes)
                    {
                        report.BudgetExhausted = true;
                        done += source.Value.Count;
                        if (progress != null) progress(done, order.Count);
                        continue;
                    }
                    sourceNo++;
                    List<PdfPageInfo> sizes;
                    try
                    {
                        sizes = PdfMergeService.LoadPages(source.Key);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch
                    {
                        report.Failed += source.Value.Count;
                        done += source.Value.Count;
                        if (progress != null) progress(done, order.Count);
                        continue;
                    }

                    List<int> pageNumbers = new List<int>();
                    foreach (int slot in source.Value)
                    {
                        int index = order[slot].PageIndex;
                        if (index >= 0 && index < sizes.Count)
                            pageNumbers.Add(index + 1);
                    }
                    List<Tuple<int, int>> runs = RenderRuns(pageNumbers);
                    var encoded = new Dictionary<string, Background>(
                        StringComparer.OrdinalIgnoreCase);
                    var pendingSlots = new HashSet<int>(source.Value);
                    for (int runIndex = 0; runIndex < runs.Count; runIndex++)
                    {
                        Cancellation.ThrowIf(cancelled);
                        if (runIndex >= MaxRunsPerSource || spent >= MediaBudgetBytes)
                        {
                            report.BudgetExhausted = spent >= MediaBudgetBytes;
                            break;
                        }
                        Tuple<int, int> run = runs[runIndex];
                        int dpi = SafeDpi(sizes, run.Item1, run.Item2, requestedDpi);
                        if (dpi <= 0)
                            continue;
                        string dir = Path.Combine(root,
                            "src" + sourceNo.ToString(System.Globalization.CultureInfo.InvariantCulture),
                            "run" + runIndex.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        List<string> files = PageRasterizer.RenderPagesWithoutText(source.Key,
                            run.Item1, run.Item2, dpi, dir, cancelled);
                        Cancellation.ThrowIf(cancelled);
                        var fileByPage = new Dictionary<int, string>();
                        for (int i = 0; i < files.Count; i++)
                            fileByPage[run.Item1 + i] = files[i];

                        foreach (int slot in source.Value)
                        {
                            if (!pendingSlots.Contains(slot))
                                continue;
                            int pageNumber = order[slot].PageIndex + 1;
                            if (pageNumber < run.Item1 || pageNumber > run.Item2)
                                continue;
                            pendingSlots.Remove(slot);
                            ProcessSlot(order, slot, sizes, fileByPage, encoded, result,
                                ref spent, report, ref done, progress, cancelled);
                        }
                        try { if (Directory.Exists(dir)) Directory.Delete(dir, true); } catch { }
                        if (report.BudgetExhausted)
                            break;
                    }

                    foreach (int slot in pendingSlots)
                    {
                        Cancellation.ThrowIf(cancelled);
                        done++;
                        if (!report.BudgetExhausted)
                            report.Failed++;
                        if (progress != null) progress(done, order.Count);
                    }
                }
                completed = true;
            }
            finally
            {
                if (!completed)
                    Release(result);
                try { if (Directory.Exists(root)) Directory.Delete(root, true); } catch { }
            }
            return new BackgroundRenderResult(result, report);
        }

        private static void ProcessSlot(IList<PdfPageRef> order, int slot,
            IList<PdfPageInfo> sizes, IDictionary<int, string> files,
            IDictionary<string, Background> encoded, Background[] result,
            ref long spent, BackgroundRenderReport report, ref int done,
            Action<int, int> progress, Func<bool> cancelled)
        {
            Cancellation.ThrowIf(cancelled);
            done++;
            if (progress != null) progress(done, order.Count);
            PdfPageRef page = order[slot];
            if (page.PageIndex < 0 || page.PageIndex >= sizes.Count)
            {
                report.Failed++;
                return;
            }
            string cacheKey = page.PageIndex + "|" + page.Rotation;
            Background cached;
            if (encoded.TryGetValue(cacheKey, out cached))
            {
                result[slot] = cached;
                if (cached != null) report.Added++;
                return;
            }

            string raster;
            if (!files.TryGetValue(page.PageIndex + 1, out raster))
            {
                report.Failed++;
                encoded[cacheKey] = null;
                return;
            }
            if (spent >= MediaBudgetBytes)
            {
                report.BudgetExhausted = true;
                encoded[cacheKey] = null;
                return;
            }
            bool failed, memoryDenied;
            Background background = Encode(raster, page.Rotation,
                out failed, out memoryDenied);
            if (memoryDenied)
                report.BudgetExhausted = true;
            if (failed)
            {
                report.Failed++;
                encoded[cacheKey] = null;
                return;
            }
            if (background == null)
            {
                encoded[cacheKey] = null;
                return;
            }
            if (background.Data == null ||
                background.Data.Length > MediaBudgetBytes - spent)
            {
                background.Dispose();
                report.BudgetExhausted = true;
                encoded[cacheKey] = null;
                return;
            }
            spent += background.Data.Length;
            encoded[cacheKey] = background;
            result[slot] = background;
            report.Added++;
        }


        internal static void Release(Background[] backgrounds)
        {
            if (backgrounds == null)
                return;
            foreach (Background background in backgrounds)
                if (background != null)
                    background.Dispose();
        }

        internal static List<Tuple<int, int>> RenderRuns(IList<int> pageNumbers)
        {
            int selected;
            List<Tuple<int, int>> source = ContinuousRuns(pageNumbers, out selected);
            if (source.Count == 0)
                return source;
            long span = (long)source[source.Count - 1].Item2 - source[0].Item1 + 1L;
            if (source.Count > 1 && span <= PageRasterizer.MaxRangePages && selected > 0 &&
                span <= (long)selected * 3L)
                source = new List<Tuple<int, int>>
                    { Tuple.Create(source[0].Item1, source[source.Count - 1].Item2) };

            var chunks = new List<Tuple<int, int>>();
            foreach (Tuple<int, int> run in source)
            {
                int first = run.Item1;
                while (first <= run.Item2)
                {
                    int last = (int)Math.Min((long)run.Item2,
                        (long)first + PageRasterizer.MaxRangePages - 1L);
                    chunks.Add(Tuple.Create(first, last));
                    if (last == int.MaxValue) break;
                    first = last + 1;
                }
            }
            return chunks;
        }

        internal static List<Tuple<int, int>> ContinuousRuns(IEnumerable<int> pageNumbers)
        {
            int ignored;
            return ContinuousRuns(pageNumbers, out ignored);
        }

        private static List<Tuple<int, int>> ContinuousRuns(IEnumerable<int> pageNumbers,
            out int selected)
        {
            var sorted = new SortedSet<int>();
            if (pageNumbers != null)
                foreach (int page in pageNumbers)
                    if (page > 0)
                        sorted.Add(page);
            selected = sorted.Count;
            var result = new List<Tuple<int, int>>();
            int first = -1, last = -1;
            foreach (int page in sorted)
            {
                if (first < 0)
                {
                    first = last = page;
                    continue;
                }
                if ((long)page == (long)last + 1L)
                {
                    last = page;
                    continue;
                }
                result.Add(Tuple.Create(first, last));
                first = last = page;
            }
            if (first > 0)
                result.Add(Tuple.Create(first, last));
            return result;
        }

        internal static Dictionary<string, List<int>> GroupBySource(IList<PdfPageRef> order)
        {
            var map = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
            if (order == null)
                return map;
            for (int i = 0; i < order.Count; i++)
            {
                PdfPageRef page = order[i];
                if (page == null || string.IsNullOrEmpty(page.SourcePath) || page.PageIndex < 0)
                    continue;
                List<int> slots;
                if (!map.TryGetValue(page.SourcePath, out slots))
                    map[page.SourcePath] = slots = new List<int>();
                slots.Add(i);
            }
            return map;
        }

        private static int SafeDpi(List<PdfPageInfo> sizes, int first, int last,
            int requested)
        {
            int dpi = requested;
            if (sizes == null || first < 1 || last > sizes.Count || last < first)
                return 0;
            for (int page = first; page <= last; page++)
            {
                PdfPageInfo info = sizes[page - 1];
                int safe = RasterBudget.SafeDpi(info.WidthPt, info.HeightPt, requested,
                    RasterBudget.BackgroundPixels);
                if (safe <= 0)
                    return 0;
                dpi = Math.Min(dpi, safe);
            }
            return dpi;
        }

        private static Background Encode(string file, int rotation, out bool failed,
            out bool memoryDenied)
        {
            failed = false;
            memoryDenied = false;
            int width, height;
            if (!TryPngDimensions(file, out width, out height) ||
                !RasterBudget.IsWithin(width, height, RasterBudget.BackgroundPixels))
            {
                failed = true;
                return null;
            }
            PdfMemoryLease working;
            int copies = rotation == 0 ? 2 : 3;
            if (!PdfMemoryBudget.TryAcquire(
                RasterBudget.BitmapWorkingSetBytes(width, height, copies), out working))
            {
                memoryDenied = true;
                failed = true;
                return null;
            }
            try
            {
                using (var loaded = new Bitmap(file))
                {
                    Bitmap rotated = null;
                    Bitmap bitmap = loaded;
                    try
                    {
                        if (rotation != 0)
                        {
                            rotated = new Bitmap(loaded);
                            rotated.RotateFlip(PageRotation.FlipFor(rotation));
                            bitmap = rotated;
                        }
                        if (RasterUtil.IsSolidColor(bitmap))
                            return null;
                        byte[] png = SavePng(bitmap);
                        if (png == null)
                        {
                            failed = true;
                            return null;
                        }
                        bool jpeg;
                        byte[] data = RasterUtil.PreferSmaller(png, out jpeg);
                        PdfMemoryLease retained;
                        if (data == null || !PdfMemoryBudget.TryAcquire(
                            Math.Max(1L, data.LongLength), out retained))
                        {
                            memoryDenied = true;
                            failed = true;
                            return null;
                        }
                        return new Background { Data = data, IsJpeg = jpeg, Lease = retained };
                    }
                    finally { if (rotated != null) rotated.Dispose(); }
                }
            }
            catch (OutOfMemoryException) { throw; }
            catch
            {
                failed = true;
                return null;
            }
            finally { working.Dispose(); }
        }

        private static bool TryPngDimensions(string path, out int width, out int height)
        {
            width = height = 0;
            try
            {
                using (var stream = File.OpenRead(path))
                {
                    var header = new byte[24];
                    if (stream.Read(header, 0, header.Length) != header.Length ||
                        header[0] != 137 || header[1] != 80 || header[2] != 78 ||
                        header[3] != 71)
                        return false;
                    width = ReadBigEndian(header, 16);
                    height = ReadBigEndian(header, 20);
                    return width > 0 && height > 0;
                }
            }
            catch { return false; }
        }

        private static int ReadBigEndian(byte[] bytes, int offset)
        {
            uint value = ((uint)bytes[offset] << 24) | ((uint)bytes[offset + 1] << 16) |
                ((uint)bytes[offset + 2] << 8) | bytes[offset + 3];
            return value > int.MaxValue ? 0 : (int)value;
        }

        private static byte[] SavePng(Bitmap bitmap)
        {
            try
            {
                using (var stream = new MemoryStream())
                {
                    bitmap.Save(stream, ImageFormat.Png);
                    return stream.ToArray();
                }
            }
            catch { return null; }
        }
    }

    internal sealed class Background : IDisposable
    {
        public byte[] Data;
        public bool IsJpeg;
        internal PdfMemoryLease Lease;

        public void Dispose()
        {
            if (Lease != null)
                Lease.Dispose();
            Lease = null;
        }
    }
}
