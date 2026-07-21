using System;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Общий каркас записи .docx через COM установленного Word: открыть скрытый Word,
    /// создать документ, дать вызывающему его наполнить, сохранить и детерминированно
    /// закрыть. Действуют те же правила COM, что в MergeService: после Close/Quit —
    /// никаких динамических операций (только статические ссылки, см. ComSafe.Release).
    /// Вызывать в STA-потоке (требование Word COM). DRY: используется и запиской, и
    /// извлечением «PDF → Word».
    /// </summary>
    internal static class WordCom
    {
        private const int WdFormatXmlDocument = 12; // .docx
        private const int WdDoNotSaveChanges = 0;

        /// <summary>
        /// Наполняет и сохраняет .docx. build получает (word, document) как object —
        /// приводить к dynamic внутри; вызывается до сохранения и закрытия (документ
        /// открыт). lockLabel — как называть файл в сообщении о занятости.
        /// </summary>
        public static void WriteDocx(string path, string lockLabel, Action<object, object> build)
        {
            if (build == null)
                throw new ArgumentNullException("build");

            Type wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
                throw new MergeException("Microsoft Word не установлен: COM-компонент Word.Application не найден.");

            string lockError = MergeService.CheckOutputWritable(path);
            if (lockError != null)
                throw new MergeException(lockError.Replace("Итоговый файл", lockLabel));

            dynamic word = null;
            dynamic doc = null;
            ComMessageFilter.Register();
            try
            {
                word = Activator.CreateInstance(wordType);
                word.Visible = false;
                word.DisplayAlerts = 0; // wdAlertsNone
                doc = word.Documents.Add();

                build((object)word, (object)doc);

                try
                {
                    doc.SaveAs2(path, WdFormatXmlDocument);
                }
                catch (Exception ex)
                {
                    throw new MergeException("Не удалось сохранить «" + Path.GetFileName(path) + "»: " + ex.Message);
                }
            }
            finally
            {
                // Статические ссылки: после Close/Quit — никаких динамических операций.
                object docObj = doc;
                if (docObj != null)
                {
                    try { doc.Close(WdDoNotSaveChanges); } catch { } // файл уже сохранён
                    ComSafe.Release(docObj);
                }
                object wordObj = word;
                if (wordObj != null)
                {
                    try { word.Quit(WdDoNotSaveChanges); } catch { }
                    ComSafe.Release(wordObj);
                }
                ComSafe.Collect();
                ComMessageFilter.Revoke();
            }
        }
    }
}
