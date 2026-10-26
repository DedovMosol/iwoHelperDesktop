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
        /// Пределы сторожа зависшего вызова Word: без предела длительности — наполнение
        /// растёт с числом страниц; модальное окно и отмена — как у Excel. Тесты подставляют короткие.
        /// </summary>
        internal static ComWatchdog.Limits WatchdogLimits = new ComWatchdog.Limits();

        /// <summary>
        /// Наполняет и сохраняет .docx. build получает (word, document) как object —
        /// приводить к dynamic внутри; вызывается до сохранения и закрытия (документ
        /// открыт). lockLabel — как называть файл в сообщении о занятости. cancelled —
        /// отмена (потокобезопасный флаг; null — нет): наполнение, которое Word не закончил за
        /// отсрочку сторожа, прерывается остановкой Word — OperationCanceledException.
        /// Модальное окно в скрытом Word (ответить некому) — остановка и MergeException с
        /// причиной. Запись файла отменой не прерывается.
        /// </summary>
        public static void WriteDocx(string path, string lockLabel, Action<object, object> build,
            Func<bool> cancelled = null)
        {
            if (build == null)
                throw new ArgumentNullException("build");

            Type wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
                throw new MergeException(Loc.T("err.word.notInstalled"));

            EnsureWritable(path, lockLabel);

            dynamic word = null;
            dynamic doc = null;
            ComWatchdog watchdog = null;
            ComMessageFilter.Register();
            try
            {
                word = Activator.CreateInstance(wordType);
                word.Visible = false;
                word.DisplayAlerts = 0; // wdAlertsNone
                doc = word.Documents.Add();
                // Окно Word (а с ним и процесс) появляется только с документом; запуск до этого — без сторожа.
                watchdog = new ComWatchdog(OwnWordProcess((object)word), cancelled, WatchdogLimits);

                watchdog.Begin(true);
                build((object)word, (object)doc);
                watchdog.EndStep();

                watchdog.Begin(false); // запись не прерываем отменой
                try
                {
                    SaveDocx(doc, path);
                }
                catch (Exception ex)
                {
                    throw new MergeException(string.Format(Loc.T("err.word.saveFailed"), Path.GetFileName(path), DiskSpace.Describe(ex, path)));
                }
                watchdog.End(); // запись вернулась без ошибки — файл записан, успех не перечёркиваем
            }
            catch (Exception) when (watchdog != null && watchdog.Stop != WatchdogStop.None)
            {
                // Любая ошибка после остановки — её следствие; человеку — причина.
                if (watchdog.Stop == WatchdogStop.Cancel)
                    throw new OperationCanceledException();
                throw new MergeException(watchdog.Describe("Word"));
            }
            finally
            {
                if (watchdog != null)
                    watchdog.Begin(true); // закрытие тоже под сторожем
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
                    QuitOrHandOver(wordObj);
                    ComSafe.Release(wordObj);
                }
                if (watchdog != null)
                    watchdog.Dispose();
                ComSafe.Collect();
                ComMessageFilter.Revoke();
            }
        }

        /// <summary>
        /// Процесс Word по окну документа (Window.Hwnd, Word 2013+). null — определить не удалось
        /// (старый Word, сбой): запись идёт без сторожа, как раньше.
        /// </summary>
        internal static OwnedOfficeProcess OwnWordProcess(object wordObj)
        {
            dynamic word = wordObj;
            try
            {
                return OwnedOfficeProcess.FromWindow((long)word.ActiveWindow.Hwnd);
            }
            catch
            {
                return null;
            }
        }

        private const int WdAlertsAll = -1;

        /// <summary>
        /// Закрыть свой Word — или отдать его человеку, если в нём открыты чужие документы:
        /// Windows открывает документ, по которому щёлкнули дважды, в уже запущенном Word, в том
        /// числе в нашем скрытом, а Quit(wdDoNotSaveChanges) отбрасывает изменения ВСЕХ
        /// документов экземпляра. Свой документ к этому моменту закрыт, поэтому всё открытое —
        /// чужое (то же правило у Excel, см. MergeService.QuitOrHandOver). Возвращает число
        /// чужих документов (0 — Word закрыт). После Quit динамических вызовов нет.
        /// </summary>
        internal static int QuitOrHandOver(object wordObj)
        {
            dynamic word = wordObj;
            int open = 0;
            try { open = (int)word.Documents.Count; }
            catch { open = 0; } // Word не отвечает — закрываем, как раньше
            if (open == 0)
            {
                try { word.Quit(WdDoNotSaveChanges); } catch { }
                return 0;
            }
            try { word.DisplayAlerts = WdAlertsAll; } catch { }
            try { word.Visible = true; } catch { }
            return open;
        }

        /// <summary>
        /// Занятый или недоступный для записи файл — понятная ошибка сразу, до запуска Word и
        /// всей работы. lockLabel — как называть файл в сообщении.
        /// </summary>
        public static void EnsureWritable(string path, string lockLabel)
        {
            if (MergeService.CheckOutputWritable(path) != null)
                throw new MergeException(string.Format(Loc.T("err.word.fileBusy"), lockLabel));
        }

        /// <summary>
        /// Сохранить .docx: сначала SaveAs2 (Word 2010+), с фолбэком на классический SaveAs.
        /// На некоторых сборках/состояниях Word позднее связывание не резолвит SaveAs2, и
        /// dynamic-вызов падает с RuntimeBinderException («__ComObject не содержит определения
        /// для SaveAs2») — тогда SaveAs, присутствующий во ВСЕХ версиях, отрабатывает штатно.
        /// Ошибка самой записи (занятый файл и т.п.) приходит COMException-ом и пробрасывается
        /// наружу как есть — её обернёт вызывающий; фолбэк ловит ТОЛЬКО «метод не найден».
        /// </summary>
        private static void SaveDocx(dynamic doc, string path)
        {
            try
            {
                doc.SaveAs2(path, WdFormatXmlDocument);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                doc.SaveAs(path, WdFormatXmlDocument); // старые/нестандартные сборки Word без SaveAs2
            }
        }
    }
}
