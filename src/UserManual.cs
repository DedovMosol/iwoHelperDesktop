using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Руководство пользователя (.docx). Лежит РЕСУРСОМ внутри exe и распаковывается рядом с
    /// настройками при первом открытии: так оно есть и у установленной версии, и у portable,
    /// и работает без интернета — приложение офлайновое, ссылка на сеть тут была бы обманом.
    ///
    /// Имя ресурса — ASCII (кириллица в именах ресурсов лишний повод для сюрпризов), имя файла
    /// на диске — человеческое: пользователь увидит его в заголовке Word.
    /// </summary>
    internal static class UserManual
    {
        private const string ResourceName = "manual.docx";
        private const string FileName = "Инструкция пользователя.docx";

        /// <summary>Куда распаковывается документ (рядом с настройками приложения).</summary>
        internal static string FilePath
        {
            get { return Path.Combine(AppPaths.Root, FileName); }
        }

        /// <summary>
        /// Распаковать при необходимости и открыть в программе, назначенной для .docx.
        /// Ошибку показываем диалогом: это явное действие пользователя, молчать в ответ на
        /// нажатие нельзя.
        /// </summary>
        public static void Open(IWin32Window owner, string title)
        {
            string path = FilePath;
            try
            {
                Extract(path);
            }
            catch (Exception ex)
            {
                // Не распаковалось (нет прав, нет места, копия открыта в Word). Если прошлая
                // копия на месте — она ничем не хуже, открываем её и молчим.
                if (!File.Exists(path))
                {
                    Dialogs.Error(owner, title, Loc.T("common.err.openFailed"), ex.Message);
                    return;
                }
            }
            Ui.OpenPathOrWarn(owner, title, path);
        }

        /// <summary>
        /// Положить документ на диск. Перезаписываем, только если файла нет или он другого
        /// размера (обновилась версия программы): лишняя запись упала бы на ровном месте,
        /// если копия сейчас открыта в Word.
        /// </summary>
        private static void Extract(string path)
        {
            using (Stream source = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName))
            {
                if (source == null)
                    throw new FileNotFoundException(ResourceName);
                var existing = new FileInfo(path);
                if (existing.Exists && existing.Length == source.Length)
                    return;
                Directory.CreateDirectory(AppPaths.Root);
                using (var target = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                    source.CopyTo(target);
            }
        }
    }
}
