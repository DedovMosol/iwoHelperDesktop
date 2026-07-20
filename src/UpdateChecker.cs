using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Проверка обновлений через GitHub Releases — только чтение последней версии.
    /// Ничего не скачивает и не заменяет: для портативного самоподписанного
    /// приложения это лучшая практика (самозаменяющиеся exe ловят антивирусы;
    /// «без сети» — козырь приложения). Страницу релиза открывает браузер по клику.
    /// </summary>
    internal static class UpdateChecker
    {
        private const string LatestApi = "https://api.github.com/repos/DedovMosol/iwoHelperDesktop/releases/latest";
        public const string ReleasesPage = "https://github.com/DedovMosol/iwoHelperDesktop/releases/latest";

        /// <summary>Тег «v1.11.2» / «1.11.2» → Version (null, если не разобрать). Чистая — под тест.</summary>
        public static Version ParseTag(string tag)
        {
            if (string.IsNullOrEmpty(tag))
                return null;
            string s = tag.Trim();
            if (s.Length > 0 && (s[0] == 'v' || s[0] == 'V'))
                s = s.Substring(1);
            Version v;
            return Version.TryParse(s, out v) ? v : null;
        }

        /// <summary>latest строго новее current. Чистая — под тест.</summary>
        public static bool IsNewer(Version latest, Version current)
        {
            return latest != null && current != null && latest > current;
        }

        /// <summary>Запрос последнего тега с GitHub (сеть). Бросает при ошибке/недоступности.</summary>
        public static string FetchLatestTag()
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var request = (HttpWebRequest)WebRequest.Create(LatestApi);
            request.UserAgent = "iwoHelperDesktop"; // GitHub требует User-Agent
            request.Accept = "application/vnd.github+json";
            request.Timeout = 10000;
            using (var response = (HttpWebResponse)request.GetResponse())
            using (var reader = new StreamReader(response.GetResponseStream()))
            {
                string json = reader.ReadToEnd();
                Match m = Regex.Match(json, "\"tag_name\"\\s*:\\s*\"([^\"]+)\"");
                if (!m.Success)
                    throw new Exception("не удалось прочитать версию из ответа GitHub");
                return m.Groups[1].Value;
            }
        }
    }

    /// <summary>Интерактивная проверка обновлений: сеть в фоне, результат — в UI-потоке.</summary>
    internal static class UpdateUi
    {
        private const string Title = "iwo Helper Desktop";

        public static void Check(Form owner)
        {
            var thread = new Thread(delegate()
            {
                string tag = null;
                Exception error = null;
                try { tag = UpdateChecker.FetchLatestTag(); }
                catch (Exception ex) { error = ex; }
                try
                {
                    if (owner != null && owner.IsHandleCreated && !owner.IsDisposed)
                        owner.BeginInvoke((MethodInvoker)delegate { ShowResult(owner, tag, error); });
                }
                catch (InvalidOperationException) { } // окно уже закрыто
            });
            thread.IsBackground = true;
            thread.Start();
        }

        private static void ShowResult(Form owner, string tag, Exception error)
        {
            if (error != null)
            {
                Dialogs.Error(owner, Title, "Не удалось проверить обновления",
                    "Проверьте подключение к интернету. (" + error.Message + ")");
                return;
            }
            Version latest = UpdateChecker.ParseTag(tag);
            Version current = Assembly.GetExecutingAssembly().GetName().Version;
            if (latest == null)
            {
                Dialogs.Error(owner, Title, "Не удалось проверить обновления", "Непонятный ответ сервера.");
                return;
            }
            if (UpdateChecker.IsNewer(latest, current))
            {
                if (Dialogs.ConfirmWarning(owner, Title, "Доступна новая версия " + latest.ToString(3),
                        "У вас " + current.ToString(3) + ". Открыть страницу загрузки в браузере?"))
                {
                    try { Process.Start(UpdateChecker.ReleasesPage); }
                    catch { } // нет браузера — ссылку видно в диалоге
                }
            }
            else
            {
                Dialogs.Info(owner, Title, "Обновлений нет", "У вас последняя версия (" + current.ToString(3) + ").");
            }
        }
    }
}
