using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Одна страница одного исходного PDF (индекс с нуля). Rotation — дополнительный
    /// поворот по часовой (0/90/180/270), назначенный пользователем в сетке миниатюр;
    /// при записи складывается с собственным /Rotate исходной страницы.
    /// Каждая позиция списка страниц — ОТДЕЛЬНЫЙ экземпляр (инвариант моделей):
    /// копирование страниц идёт через <see cref="Clone"/>, а не общие ссылки.
    /// </summary>
    public class PdfPageRef
    {
        public string SourcePath;
        public int PageIndex;
        public int Rotation;

        public string FileName
        {
            get { return Path.GetFileName(SourcePath); }
        }

        public PdfPageRef Clone()
        {
            return new PdfPageRef { SourcePath = SourcePath, PageIndex = PageIndex, Rotation = Rotation };
        }

        /// <summary>Сумма поворотов, нормализованная в {0, 90, 180, 270} (отрицательные допустимы). Чистая — под тест.</summary>
        internal static int ComposeRotation(int a, int b)
        {
            int sum = ((a + b) % 360 + 360) % 360;
            return sum - sum % 90; // не кратный 90 мусор из чужого PDF срезаем вниз до кратного
        }
    }

    /// <summary>Размеры страницы в пунктах — для проверок и подписей.</summary>
    public class PdfPageInfo
    {
        public int PageIndex;
        public double WidthPt;
        public double HeightPt;
    }

    /// <summary>
    /// Слияние PDF-документов с произвольным порядком страниц (PDFsharp, MIT).
    /// Страницы копируются как есть, без переконвертации — сканы, печати
    /// и подписи не искажаются. Публичные методы не содержат типов PdfSharp
    /// в телах: сначала EmbeddedAssemblies.Ensure(), затем [NoInlining]-ядро.
    /// </summary>
    public static class PdfMergeService
    {
        /// <summary>Страницы документа с размерами. Битый/зашифрованный файл — MergeException.</summary>
        public static List<PdfPageInfo> LoadPages(string path)
        {
            EmbeddedAssemblies.Ensure();
            return LoadPagesCore(path);
        }

        /// <summary>
        /// Собирает страницы в порядке order в новый PDF. Пустой порядок — ошибка.
        /// cancelled — кооперативная отмена (проверяется между страницами; при отмене файл
        /// не создаётся — сохранение идёт лишь в самом конце): OperationCanceledException.
        /// </summary>
        public static void Merge(IList<PdfPageRef> order, string outputPath, Action<int, int> progress = null,
            Func<bool> cancelled = null, bool padToEven = false)
        {
            if (order == null || order.Count == 0)
                throw new MergeException(Loc.T("err.pdf.noPages"));
            foreach (PdfPageRef page in order)
                if (page != null && OutputFile.IsSameFile(page.SourcePath, outputPath))
                    throw new MergeException(Loc.T("err.output.sameSource"));
            string writeError = OutputFile.PdfWriteProblem(outputPath);
            if (writeError != null)
                throw new MergeException(writeError);
            EmbeddedAssemblies.Ensure();
            using (var output = new AtomicOutput(outputPath))
            {
                MergeCore(order, output.TempPath, progress, cancelled, padToEven, null);
                output.Commit();
            }
        }

        /// <summary>
        /// Записать в путь, уже принадлежащий внешней транзакции. Никакой второй
        /// AtomicOutput/journal здесь не создаётся; публикует только вызывающий.
        /// </summary>
        internal static void WriteUnpublished(IList<PdfPageRef> order, string path,
            Action<int, int> progress = null, Func<bool> cancelled = null,
            bool padToEven = false)
        {
            if (order == null || order.Count == 0)
                throw new MergeException(Loc.T("err.pdf.noPages"));
            EmbeddedAssemblies.Ensure();
            MergeCore(order, path, progress, cancelled, padToEven, null);
        }

        /// <summary>
        /// То же, что <see cref="WriteUnpublished(IList{PdfPageRef}, string, Action{int, int}, Func{bool}, bool)"/>,
        /// но источники берутся из общего набора, открытого вызывающим на всю серию записей
        /// (части разделения). Набор вызывающий и закрывает. Только после
        /// <see cref="EmbeddedAssemblies.Ensure"/> — набор держит типы PdfSharp.
        /// </summary>
        internal static void WriteUnpublished(IList<PdfPageRef> order, string path,
            PdfSourceSet sources, Func<bool> cancelled)
        {
            if (order == null || order.Count == 0)
                throw new MergeException(Loc.T("err.pdf.noPages"));
            MergeCore(order, path, null, cancelled, false, sources);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static List<PdfPageInfo> LoadPagesCore(string path)
        {
            try
            {
                using (PdfDocument doc = PdfReader.Open(path, PdfPasswords.For(path), PdfDocumentOpenMode.Import))
                {
                    var pages = new List<PdfPageInfo>();
                    for (int i = 0; i < doc.PageCount; i++)
                    {
                        PdfPage page = doc.Pages[i];
                        var info = new PdfPageInfo();
                        info.PageIndex = i;
                        info.WidthPt = page.Width.Point;
                        info.HeightPt = page.Height.Point;
                        pages.Add(info);
                    }
                    if (pages.Count == 0)
                        throw new MergeException(string.Format(Loc.T("err.pdf.noPagesIn"), Path.GetFileName(path)));
                    return pages;
                }
            }
            catch (MergeException)
            {
                throw;
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                throw new MergeException(string.Format(Loc.T("err.pdf.cantOpen"), Path.GetFileName(path), ex.Message));
            }
        }

        /// <param name="shared">Набор источников вызывающего (серия записей) или null — тогда
        /// источники открываются здесь и здесь же закрываются.</param>
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void MergeCore(IList<PdfPageRef> order, string outputPath, Action<int, int> progress,
            Func<bool> cancelled, bool padToEven, PdfSourceSet shared)
        {
            // Позиции добивочных пустых страниц считаем ЗАРАНЕЕ и по исходной нумерации:
            // вставка по ходу сдвигала бы все последующие позиции.
            var padAfter = new HashSet<int>();
            if (padToEven)
                foreach (int at in BlankPages.InsertPositions(order))
                    padAfter.Add(at);
            // Каждый источник открывается один раз, сколько бы страниц из него ни брали.
            PdfSourceSet sources = shared ?? new PdfSourceSet();
            // Куда переехала каждая взятая страница — отдельно по каждому источнику, плюс
            // порядок первого появления источников: по этому и переносится оглавление.
            // Страница, взятая дважды, ведёт на ПЕРВОЕ вхождение — закладка должна вести
            // в начало раздела, а не в его повтор.
            var movedBySource = new Dictionary<string, Dictionary<int, int>>(StringComparer.OrdinalIgnoreCase);
            var sourceOrder = new List<string>();
            PdfDocument output = null;
            try
            {
                output = new PdfDocument();
                int added = 0;
                for (int index = 0; index < order.Count; index++)
                {
                    PdfPageRef page = order[index];
                    Cancellation.ThrowIf(cancelled); // между страницами; файл ещё не создан
                    string key = Path.GetFullPath(page.SourcePath);
                    PdfDocument source = sources.Get(key, page.FileName);
                    Dictionary<int, int> moved;
                    if (!movedBySource.TryGetValue(key, out moved))
                    {
                        moved = new Dictionary<int, int>();
                        movedBySource.Add(key, moved);
                        sourceOrder.Add(key);
                    }
                    if (page.PageIndex < 0 || page.PageIndex >= source.PageCount)
                        throw new MergeException(string.Format(Loc.T("err.pdf.pageGone"), page.FileName, page.PageIndex + 1));
                    PdfPage copied = output.AddPage(source.Pages[page.PageIndex]);
                    if (!moved.ContainsKey(page.PageIndex))          // только первое вхождение
                        moved[page.PageIndex] = output.PageCount - 1;
                    if (page.Rotation != 0) // поворот пользователя поверх собственного /Rotate страницы
                        copied.Rotate = PdfPageRef.ComposeRotation(copied.Rotate, page.Rotation);
                    // Пустая страница ПОСЛЕ документа с нечётным числом страниц: размер берём
                    // у только что добавленной, иначе добивка выпала бы из формата пачки.
                    if (padAfter.Contains(index + 1))
                    {
                        PdfPage blank = output.AddPage();
                        blank.Width = copied.Width;
                        blank.Height = copied.Height;
                    }
                    added++;
                    if (progress != null)
                        progress(added, order.Count);
                }

                // Оглавление переносим ПОСЛЕ всех страниц: закладке нужна страница результата,
                // а она существует только когда собран весь документ. Источники идут в порядке
                // первого появления — «содержание первого файла, затем второго».
                var bookmarks = new List<PdfBookmark>();
                foreach (string key in sourceOrder)
                    bookmarks.AddRange(PdfBookmarks.Remap(sources.Bookmarks(key), movedBySource[key]));
                PdfBookmarks.Write(output, bookmarks);

                try
                {
                    output.Save(outputPath);
                }
                catch (Exception ex) when (MergeException.ShouldWrap(ex))
                {
                    throw new MergeException(string.Format(Loc.T("err.pdf.saveFailed"), DiskSpace.Describe(ex, outputPath)));
                }
            }
            finally
            {
                if (output != null)
                    try { output.Dispose(); } catch { }
                if (shared == null)
                    sources.Dispose();
            }
        }
    }

    /// <summary>
    /// Открытые исходники на серию записей. Разделение пишет много частей из ОДНОГО файла, и
    /// раньше каждая часть заново открывала и целиком разбирала исходник (и заново читала его
    /// оглавление): на 300-страничном документе, разбитом постранично, — триста полных
    /// разборов одного и того же файла. Импорт страниц исходник не меняет, поэтому один
    /// открытый документ безопасно отдаёт страницы в любое число новых.
    ///
    /// Держит типы PdfSharp — создаётся только после <see cref="EmbeddedAssemblies.Ensure"/>.
    /// Не потокобезопасен: одна серия — один поток.
    /// </summary>
    internal sealed class PdfSourceSet : IDisposable
    {
        private readonly Dictionary<string, PdfDocument> _documents =
            new Dictionary<string, PdfDocument>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, List<PdfBookmark>> _bookmarks =
            new Dictionary<string, List<PdfBookmark>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Сколько раз этот поток разобрал исходник с диска — только для теста «исходник
        /// открывается один раз». На поток, а не общий: чужие фоновые операции не сбивают счёт.
        /// </summary>
        [ThreadStatic] internal static int OpensForTests;

        /// <summary>
        /// Документ по пути; с диска — только при первом обращении. errorKey — формат
        /// сообщения об ошибке открытия ({0} — имя файла, {1} — причина): объединение и
        /// разделение исторически говорят об этом по-разному.
        /// </summary>
        internal PdfDocument Get(string path, string displayName, string errorKey = "err.pdf.cantOpenShort")
        {
            string key = Path.GetFullPath(path);
            PdfDocument document;
            if (_documents.TryGetValue(key, out document))
                return document;
            try
            {
                document = PdfReader.Open(key, PdfPasswords.For(key), PdfDocumentOpenMode.Import);
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                throw new MergeException(string.Format(Loc.T(errorKey), displayName, ex.Message));
            }
            OpensForTests++;
            _documents.Add(key, document);
            return document;
        }

        /// <summary>Оглавление источника; читается один раз на серию (Remap его не меняет).</summary>
        internal List<PdfBookmark> Bookmarks(string path)
        {
            string key = Path.GetFullPath(path);
            List<PdfBookmark> marks;
            if (!_bookmarks.TryGetValue(key, out marks))
            {
                marks = PdfBookmarks.Read(Get(key, Path.GetFileName(key)));
                _bookmarks.Add(key, marks);
            }
            return marks;
        }

        public void Dispose()
        {
            foreach (PdfDocument document in _documents.Values)
                try { document.Dispose(); } catch { }
            _documents.Clear();
            _bookmarks.Clear();
        }
    }
}
