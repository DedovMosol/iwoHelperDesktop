using System;
using System.IO;
using System.Runtime.CompilerServices;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Чтение и запись свойств документа (заголовок, автор, тема, ключевые слова).
    ///
    /// Как и остальные сервисы вокруг PdfSharp: публичные методы не содержат его типов в
    /// телах — сначала <see cref="EmbeddedAssemblies.Ensure"/>, затем [NoInlining]-ядро,
    /// иначе JIT увидит тип вшитой сборки раньше, чем зарегистрирован её резолвер.
    ///
    /// Запись идёт в ДРУГОЙ файл: исходники приложение не меняет никогда.
    /// </summary>
    public static class PdfMetadataService
    {
        /// <summary>Текущие свойства документа. Нечитаемый файл — пустые значения, а не ошибка.</summary>
        public static PdfMetadata Read(string path)
        {
            EmbeddedAssemblies.Ensure();
            return ReadCore(path);
        }

        /// <summary>
        /// Записать свойства в новый файл outputPath (исходник не меняется). Пустое значение
        /// удаляет свойство — это осознанный способ убрать из файла имя автора перед отправкой.
        /// </summary>
        public static void Write(string sourcePath, string outputPath, PdfMetadata meta)
        {
            if (meta == null)
                throw new ArgumentNullException("meta");
            if (OutputFile.IsSameFile(sourcePath, outputPath))
                throw new MergeException(Loc.T("err.output.sameSource"));
            string writeError = OutputFile.PdfWriteProblem(outputPath);
            if (writeError != null)
                throw new MergeException(writeError);
            EmbeddedAssemblies.Ensure();
            using (var output = new AtomicOutput(outputPath))
            {
                WriteCore(sourcePath, output.TempPath, meta);
                output.Commit();
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static PdfMetadata ReadCore(string path)
        {
            try
            {
                using (PdfDocument doc = PdfReader.Open(path, PdfPasswords.For(path), PdfDocumentOpenMode.InformationOnly))
                    return new PdfMetadata
                    {
                        Title = doc.Info.Title ?? "",
                        Author = doc.Info.Author ?? "",
                        Subject = doc.Info.Subject ?? "",
                        Keywords = doc.Info.Keywords ?? ""
                    };
            }
            catch
            {
                return new PdfMetadata(); // свойства не прочитались — предложим заполнить с нуля
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void WriteCore(string sourcePath, string outputPath, PdfMetadata meta)
        {
            try
            {
                using (PdfDocument doc = PdfReader.Open(sourcePath, PdfPasswords.For(sourcePath), PdfDocumentOpenMode.Modify))
                {
                    SetOrRemove(doc.Info, "/Title", meta.Title);
                    SetOrRemove(doc.Info, "/Author", meta.Author);
                    SetOrRemove(doc.Info, "/Subject", meta.Subject);
                    SetOrRemove(doc.Info, "/Keywords", meta.Keywords);
                    // Свойства документа живут в ДВУХ местах: в словаре Info и в потоке XMP
                    // (/Metadata каталога, dc:creator/dc:title). Вьюеры предпочитают XMP, поэтому
                    // правка одного Info была видимостью: «убрать автора перед отправкой» молча не
                    // срабатывало — Acrobat продолжал показывать его из XMP. Синхронно переписать
                    // XMP без библиотеки XMP — значит рисковать битым пакетом; удаляем его, и
                    // единственным источником правды остаётся Info, который мы только что задали.
                    doc.Internals.Catalog.Elements.Remove("/Metadata");
                    doc.Save(outputPath);
                }
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                throw new MergeException(string.Format(Loc.T("err.pdf.saveFailed"), DiskSpace.Describe(ex, outputPath)));
            }
        }

        /// <summary>
        /// Пустое значение — ключ удаляется, а не пишется пустой строкой «()»: пустое свойство
        /// всё равно остаётся в файле следом, и часть программ показывает его как «есть, но пусто».
        /// </summary>
        private static void SetOrRemove(PdfDocumentInformation info, string key, string value)
        {
            if (string.IsNullOrEmpty(value))
                info.Elements.Remove(key);
            else
                info.Elements.SetString(key, value);
        }
    }
}
