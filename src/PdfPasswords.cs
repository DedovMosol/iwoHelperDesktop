using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>
    /// Пароли к защищённым PDF, введённые в этом сеансе. До 1.18.3 такой файл просто не
    /// открывался: инструмент отказывал ровно тогда, когда был нужен.
    ///
    /// Реестр, а не параметр в каждом методе, — решение осознанное. Пароль нужен ВЕЗДЕ, где
    /// файл открывается: страницы читает PdfSharp, миниатюры рисует WinRT, текст разбирает
    /// PdfPig, сжатие запускает Ghostscript. Протаскивать его через все эти слои значило бы
    /// поменять с десяток сигнатур ради данных, которые по смыслу принадлежат не операции,
    /// а файлу. Здесь он живёт один раз и достаётся по пути, а слои остаются прежними.
    ///
    /// НА ДИСК НЕ ПОПАДАЕТ НИКОГДА. Ни в настройки, ни в отчёты, ни в журнал сбоев: пароль
    /// живёт в памяти процесса и уходит вместе с ним. Это не забывчивость, а условие —
    /// программа, которая сохранила бы чужой пароль в открытом виде, хуже, чем программа,
    /// которая спросит его ещё раз.
    /// </summary>
    internal static class PdfPasswords
    {
        private static readonly Dictionary<string, string> Known =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // Читают из фоновых потоков (разбор, рендер миниатюр), пишет UI-поток после диалога.
        private static readonly object Gate = new object();

        /// <summary>Пароль к файлу или null. Путь приводится к полному — тот же файл под другим написанием пути должен считаться тем же.</summary>
        public static string For(string path)
        {
            string key = Key(path);
            if (key == null)
                return null;
            lock (Gate)
            {
                string password;
                return Known.TryGetValue(key, out password) ? password : null;
            }
        }

        /// <summary>Запомнить пароль на время сеанса. Пустой стирает запись.</summary>
        public static void Remember(string path, string password)
        {
            string key = Key(path);
            if (key == null)
                return;
            lock (Gate)
            {
                if (string.IsNullOrEmpty(password))
                    Known.Remove(key);
                else
                    Known[key] = password;
            }
        }

        /// <summary>Забыть всё (нужно тестам: реестр общий на процесс).</summary>
        public static void Clear()
        {
            lock (Gate)
                Known.Clear();
        }

        /// <summary>
        /// Похоже ли, что открыть помешал именно пароль. PdfSharp говорит об этом двумя
        /// разными фразами — «нужен пароль» и «пароль неверен», — и обе означают одно:
        /// спрашивать пароль. Ошибиться здесь не страшно: лишний вопрос человек закроет,
        /// а вот молчаливый отказ на защищённом файле выглядит поломкой программы.
        /// Чистая — под тест.
        /// </summary>
        public static bool LooksPasswordProtected(Exception error)
        {
            for (Exception e = error; e != null; e = e.InnerException)
            {
                string message = e.Message;
                if (!string.IsNullOrEmpty(message) &&
                    (message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     message.IndexOf("парол", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     e.GetType().Name.IndexOf("Encrypted", StringComparison.OrdinalIgnoreCase) >= 0))
                    return true;
            }
            return false;
        }

        private static string Key(string path)
        {
            if (string.IsNullOrEmpty(path))
                return null;
            try { return Path.GetFullPath(path); }
            catch { return path; }
        }
    }
}
