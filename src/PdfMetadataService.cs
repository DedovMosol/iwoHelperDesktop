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
        /// очищает свойство — это осознанный способ убрать из файла имя автора перед отправкой.
        /// </summary>
        public static void Write(string sourcePath, string outputPath, PdfMetadata meta)
        {
            if (meta == null)
                throw new ArgumentNullException("meta");
            string lockError = MergeService.CheckOutputWritable(outputPath);
            if (lockError != null)
                throw new MergeException(Loc.T("err.pdf.fileBusy"));
            EmbeddedAssemblies.Ensure();
            WriteCore(sourcePath, outputPath, meta);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static PdfMetadata ReadCore(string path)
        {
            try
            {
                using (PdfDocument doc = PdfReader.Open(path, PdfDocumentOpenMode.InformationOnly))
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
                using (PdfDocument doc = PdfReader.Open(sourcePath, PdfDocumentOpenMode.Modify))
                {
                    doc.Info.Title = meta.Title ?? "";
                    doc.Info.Author = meta.Author ?? "";
                    doc.Info.Subject = meta.Subject ?? "";
                    doc.Info.Keywords = meta.Keywords ?? "";
                    doc.Save(outputPath);
                }
            }
            catch (Exception ex) when (MergeException.ShouldWrap(ex))
            {
                throw new MergeException(string.Format(Loc.T("err.pdf.saveFailed"), ex.Message));
            }
        }
    }
}
