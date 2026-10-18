using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Документ, вшитый ресурсом в exe (руководство, лицензии), распаковывается рядом с
    /// настройками и открывается программой по умолчанию. Одна реализация на всех: две копии
    /// уже разошлись бы при первой правке, а ошибка в них была общей.
    /// </summary>
    internal static class EmbeddedFile
    {
        /// <summary>
        /// Распаковать при необходимости и открыть в программе, назначенной для типа файла.
        /// Ошибку показываем диалогом: это явное действие пользователя, молчать в ответ на
        /// нажатие нельзя.
        /// </summary>
        internal static void Open(IWin32Window owner, string title, string resourceName, string path)
        {
            try
            {
                Extract(resourceName, path);
            }
            catch (Exception ex)
            {
                // Не распаковалось (нет прав, нет места, копия открыта в Word). Публикация
                // атомарная, поэтому прошлая копия на месте целой — открываем её и молчим.
                if (!File.Exists(path))
                {
                    Dialogs.Error(owner, title, Loc.T("common.err.openFailed"), ex.Message);
                    return;
                }
            }
            Ui.OpenPathOrWarn(owner, title, path);
        }

        private static void Extract(string resourceName, string path)
        {
            byte[] content;
            using (Stream source = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (source == null)
                    throw new FileNotFoundException(resourceName);
                using (var buffer = new MemoryStream())
                {
                    source.CopyTo(buffer);
                    content = buffer.ToArray();
                }
            }
            WriteIfChanged(path, content);
        }

        /// <summary>
        /// Положить content в path, если там лежит что-то другое. true — файл переписан.
        ///
        /// Сравниваем СОДЕРЖИМОЕ, а не длину: правка в ресурсе того же размера (опечатка,
        /// дата) иначе не доезжала бы до диска никогда. Файлы маленькие, побайтное сравнение
        /// дёшево. Совпавшее не переписываем: лишняя запись упала бы на ровном месте, если
        /// копия сейчас открыта в Word.
        ///
        /// Пишем через временный файл и атомарную замену: прямая перезапись сначала обнуляла
        /// файл, и сбой посреди копирования оставлял битый документ, который тут же и открывался.
        /// </summary>
        internal static bool WriteIfChanged(string path, byte[] content)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            if (SameContent(path, content))
                return false;
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
            using (var output = new AtomicOutput(path))
            {
                File.WriteAllBytes(output.TempPath, content);
                output.Commit();
            }
            return true;
        }

        private static bool SameContent(string path, byte[] content)
        {
            try
            {
                var info = new FileInfo(path);
                if (!info.Exists || info.Length != content.Length)
                    return false;
                byte[] existing = File.ReadAllBytes(path);
                for (int i = 0; i < existing.Length; i++)
                    if (existing[i] != content[i])
                        return false;
                return true;
            }
            catch
            {
                // Не прочитали (копия занята) — считаем другой: запись либо пройдёт, либо
                // упадёт и оставит прежнюю копию, которую Open всё равно откроет.
                return false;
            }
        }
    }
}
