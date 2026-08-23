using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

namespace ExcelMerger
{
    /// <summary>
    /// Локальная оркестрация review двух явно выбранных PDF: проверки, извлечение,
    /// лёгкая проекция страниц и pure diff. Успешный текст кэшируется только в памяти.
    /// </summary>
    internal static class PdfReviewService
    {
        private const string AlgorithmVersion = "review-1";
        private static readonly object CacheGate = new object();
        private static readonly LruCache<PdfReviewDocument> Cache =
            new LruCache<PdfReviewDocument>(4, null);

        public static PdfReviewResult Compare(string leftPath, string rightPath,
            Action<int, int> progress = null, Func<bool> cancelled = null,
            PdfReviewLimits limits = null)
        {
            limits = limits ?? PdfReviewLimits.Default();
            string left = Canonical(leftPath), right = Canonical(rightPath);
            if (left == null || right == null)
                throw new PdfReviewException(PdfReviewFailure.Unreadable, null,
                    Loc.T("review.err.pickBoth"));
            if (OutputFile.IsSameFile(left, right))
                throw new PdfReviewException(PdfReviewFailure.Unreadable, left,
                    Loc.T("review.err.sameFile"));

            Action<int, int> leftProgress = progress == null ? null :
                (Action<int, int>)delegate(int done, int total)
                {
                    progress(total > 0 ? done : 0, Math.Max(1, total) * 2);
                };
            PdfReviewDocument leftDoc = Load(left, leftProgress, cancelled, limits);
            Cancellation.ThrowIf(cancelled);
            Action<int, int> rightProgress = progress == null ? null :
                (Action<int, int>)delegate(int done, int total)
                {
                    int t = Math.Max(1, total);
                    progress(t + done, t * 2);
                };
            PdfReviewDocument rightDoc = Load(right, rightProgress, cancelled, limits);
            Cancellation.ThrowIf(cancelled);
            PdfReviewResult result = PdfReviewDiff.Compare(leftDoc, rightDoc, limits);
            if (progress != null)
                progress(2, 2);
            return result;
        }

        public static PdfReviewDocument Load(string path, Action<int, int> progress,
            Func<bool> cancelled, PdfReviewLimits limits)
        {
            // PdfPig вшит в exe ресурсом и без перехвата разрешения не грузится: резолвер
            // обязан быть зарегистрирован ДО JIT-компиляции ядра, в теле которого есть типы
            // UglyToad (строка с «PdfDocument probe» одна роняла первое же сравнение
            // в однофайловой сборке: «Не удалось загрузить файл или сборку…»).
            EmbeddedAssemblies.Ensure();
            return LoadCore(path, progress, cancelled, limits);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static PdfReviewDocument LoadCore(string path, Action<int, int> progress,
            Func<bool> cancelled, PdfReviewLimits limits)
        {
            string full = Canonical(path);
            if (full == null || !File.Exists(full))
                throw Failure(PdfReviewFailure.Unreadable, path, Loc.T("review.err.unreadable"));
            FileStamp before = Stamp(full);
            if (before.Length > limits.MaxFileBytes)
                throw Failure(PdfReviewFailure.TooLarge, full, Loc.T("review.err.tooLarge"));
            string key = CacheKey(full, before);
            lock (CacheGate)
            {
                PdfReviewDocument cached;
                if (Cache.TryGet(key, out cached))
                    return cached;
            }

            // Открываем документ САМИ, а не через PdfPageProbe: тот глотает любое исключение
            // в «-1», и причина отказа терялась ещё до того, как её можно было понять. Здесь
            // исключение нужно живьём: по нему отличаем защищённый паролем файл (пользователю
            // будет предложен пароль) от битого/недоступного (сообщаем истинную причину).
            UglyToad.PdfPig.PdfDocument probe;
            try
            {
                probe = PdfPageProbe.OpenPig(full);
            }
            catch (Exception ex)
            {
                throw ReadFailure(full, ex);
            }
            int count;
            using (probe)
                count = probe.NumberOfPages;
            if (count <= 0)
                throw Failure(PdfReviewFailure.Unreadable, full, Loc.T("review.err.unreadable"));
            if (count > limits.MaxPages)
                throw Failure(PdfReviewFailure.TooLarge, full, Loc.T("review.err.tooLarge"));
            Cancellation.ThrowIf(cancelled);

            try
            {
                if (!PdfTextExtract.AnyPageHasText(full))
                    throw Failure(PdfReviewFailure.NoText, full, Loc.T("review.err.noText"));
                List<PdfPageText> extracted = PdfTextExtract.Extract(full, progress, null, cancelled);
                var doc = new PdfReviewDocument { Path = full };
                foreach (PdfPageText page in extracted)
                {
                    Cancellation.ThrowIf(cancelled);
                    string text = PlainText.Page(page);
                    string normalized = PdfReviewDiff.Normalize(text);
                    doc.CharacterCount += PdfReviewDiff.CodePoints(normalized);
                    if (doc.CharacterCount > limits.MaxCharacters)
                        throw Failure(PdfReviewFailure.TooLarge, full, Loc.T("review.err.tooLarge"));
                    var reviewPage = new PdfReviewPage
                    {
                        PageIndex = page.PageIndex,
                        Text = text,
                        NormalizedText = normalized,
                        WidthPt = page.WidthPt,
                        HeightPt = page.HeightPt,
                        Fingerprint = PdfReviewDiff.Fingerprint(normalized, page.WidthPt, page.HeightPt)
                    };
                    double viewW, viewH;
                    PdfReviewGeometry.ViewSize(page.WidthPt, page.HeightPt, page.NativeRotation,
                        out viewW, out viewH);
                    reviewPage.ViewWidthPt = viewW;
                    reviewPage.ViewHeightPt = viewH;
                    BuildWords(page, reviewPage);
                    doc.WordCount += reviewPage.Words.Count;
                    doc.Pages.Add(reviewPage);
                }
                bool any = false;
                foreach (PdfReviewPage page in doc.Pages)
                    if (!string.IsNullOrWhiteSpace(page.NormalizedText)) { any = true; break; }
                if (!any)
                    throw Failure(PdfReviewFailure.NoText, full, Loc.T("review.err.noText"));
                FileStamp after = Stamp(full);
                if (before.Length != after.Length || before.WriteTicks != after.WriteTicks)
                    throw Failure(PdfReviewFailure.ChangedDuringRead, full,
                        Loc.T("review.err.changedDuringRead"));
                lock (CacheGate)
                    Cache.Add(key, doc);
                return doc;
            }
            catch (OperationCanceledException) { throw; }
            catch (PdfReviewException) { throw; }
            catch (Exception ex)
            {
                PdfReviewFailure reason = PdfPasswords.LooksPasswordProtected(ex)
                    ? PdfReviewFailure.PasswordRequired : PdfReviewFailure.Unreadable;
                string message = reason == PdfReviewFailure.PasswordRequired
                    ? Loc.T("review.err.password")
                    : string.Format(Loc.T("review.err.unreadableFile"), Path.GetFileName(full), ex.Message);
                throw Failure(reason, full, message);
            }
        }

        private struct FileStamp
        {
            public long Length;
            public long WriteTicks;
        }

        /// <summary>
        /// Слова страницы в ПОРЯДКЕ ЧТЕНИЯ с рамками в пространстве отображения. Тот же
        /// порядок, что и у обычного текста (строки сверху вниз, слова слева направо):
        /// ворд-дифф сравнивает именно эту последовательность, а рамки из неё же ложатся на
        /// отрендеренную страницу — один источник для диффа и подсветки.
        /// </summary>
        private static void BuildWords(PdfPageText page, PdfReviewPage reviewPage)
        {
            if (page.Words == null || page.Words.Count == 0)
                return;
            foreach (OcrLayout.Line line in OcrLayout.ToLines(page.Words))
            {
                foreach (PdfWord word in line.Words)
                {
                    string text = word.Text == null ? "" : word.Text.Trim();
                    if (text.Length == 0)
                        continue; // чистый пробел — не слово
                    reviewPage.Words.Add(new PdfReviewWord
                    {
                        Text = word.Text,
                        Key = text,
                        Box = PdfReviewGeometry.RawToView(word.Left, word.Bottom, word.Right, word.Top,
                            page.NativeRotation, page.WidthPt, page.HeightPt)
                    });
                }
            }
        }

        private static FileStamp Stamp(string path)
        {
            try
            {
                var file = new FileInfo(path);
                return new FileStamp { Length = file.Length, WriteTicks = file.LastWriteTimeUtc.Ticks };
            }
            catch
            {
                throw Failure(PdfReviewFailure.Unreadable, path, Loc.T("review.err.unreadable"));
            }
        }

        private static string CacheKey(string path, FileStamp stamp)
        {
            return path.ToLowerInvariant() + "|" + stamp.Length + "|" + stamp.WriteTicks + "|" + AlgorithmVersion;
        }

        private static string Canonical(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return null;
            try { return Path.GetFullPath(path); }
            catch { return null; } // недопустимые символы пути → вызывающий выдаст «не читается»
        }

        private static PdfReviewException Failure(PdfReviewFailure reason, string path, string message)
        {
            return new PdfReviewException(reason, path, message);
        }

        /// <summary>
        /// Разобрать исключение ОТКРЫТИЯ документа: защищённый паролем файл — отдельный исход
        /// (форма предложит ввести пароль), всё остальное — «не читается» с истинной причиной.
        /// Раньше сюда не доходило вовсе: подсчёт страниц глотал исключение в «-1», и человек
        /// получал безликое «не удалось прочитать» даже на паролем защищённом документе.
        /// </summary>
        private static PdfReviewException ReadFailure(string path, Exception ex)
        {
            PdfReviewFailure reason = PdfPasswords.LooksPasswordProtected(ex)
                ? PdfReviewFailure.PasswordRequired : PdfReviewFailure.Unreadable;
            string message = reason == PdfReviewFailure.PasswordRequired
                ? Loc.T("review.err.password")
                : string.Format(Loc.T("review.err.unreadableFile"), Path.GetFileName(path), ex.Message);
            return Failure(reason, path, message);
        }

    }
}
