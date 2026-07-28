using System;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text.RegularExpressions;
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

        /// <summary>
        /// Показывать ли уведомление о новой версии при запуске. Чистая — под тест.
        ///
        /// Пропущенная версия сравнивается по «не старше», а не по равенству: человек просил
        /// не напоминать про 1.18.0, вышла 1.18.1 — и промолчать значило бы, что один флажок
        /// отключил уведомления навсегда. Наоборот, если пропущенная версия ВЫШЕ найденной
        /// (откатились с бета-версии), молчим — новостью это не является.
        ///
        /// Непонятная пропущенная версия (мусор в настройках) не должна глушить проверку,
        /// поэтому неразобранное значение считается отсутствующим.
        /// </summary>
        public static bool ShouldNotify(Version latest, Version current, string skippedVersion)
        {
            if (!IsNewer(latest, current))
                return false;
            Version skipped = ParseTag(skippedVersion);
            return skipped == null || latest > skipped;
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
                    throw new Exception(Loc.T("update.err.parseVersion"));
                return m.Groups[1].Value;
            }
        }
    }

    /// <summary>Интерактивная проверка обновлений: сеть в фоне, результат — в UI-потоке.</summary>
    internal static class UpdateUi
    {
        private const string Title = "iwo Helper Desktop";

        /// <summary>
        /// Проверить обновления: запрос в фоне, результат — в UI-потоке. done вызывается
        /// перед показом результата и возвращает в строй кнопку, которую вызывающий погасил
        /// на время запроса (сеть с таймаутом 10 с). Если окно успели закрыть, не выполнится
        /// ни done, ни показ — возвращать в строй уже нечего.
        /// </summary>
        public static void Check(Form owner, Action done)
        {
            Ui.RunWorker(delegate()
            {
                string tag = null;
                Exception error = null;
                try { tag = UpdateChecker.FetchLatestTag(); }
                catch (Exception ex) { error = ex; }
                Ui.OnUi(owner, delegate // общий guard: своя копия уже дважды теряла catch
                {
                    if (done != null)
                        done();
                    ShowResult(owner, tag, error);
                });
            });
        }

        private static void ShowResult(Form owner, string tag, Exception error)
        {
            if (error != null)
            {
                Dialogs.Error(owner, Title, Loc.T("update.err.title"),
                    string.Format(Loc.T("update.err.network"), error.Message));
                return;
            }
            Version latest = UpdateChecker.ParseTag(tag);
            Version current = Assembly.GetExecutingAssembly().GetName().Version;
            if (latest == null)
            {
                Dialogs.Error(owner, Title, Loc.T("update.err.title"), Loc.T("update.err.badResponse"));
                return;
            }
            if (UpdateChecker.IsNewer(latest, current))
                OfferUpdate(owner, latest, current, false);
            else
                Dialogs.Info(owner, Title, Loc.T("update.none.title"), string.Format(Loc.T("update.none.body"), current.ToString(3)));
        }

        /// <summary>
        /// Проверка ПРИ ЗАПУСКЕ: молчит обо всём, кроме найденной новой версии. Ни ошибка
        /// сети, ни «у вас последняя» здесь не показываются — человек об этой проверке не
        /// просил, и отвечать ему диалогом на действие, которого он не совершал, значит
        /// мешать работать. По кнопке (<see cref="Check"/>) отвечаем на всё: там спросили.
        /// </summary>
        public static void CheckOnStart(Form owner)
        {
            if (!UserSettings.Load().UpdateCheckOnStart)
                return;
            Ui.RunWorker(delegate()
            {
                string tag = null;
                try { tag = UpdateChecker.FetchLatestTag(); }
                catch { return; } // нет сети, нет ответа, мусор в ответе — молча
                Ui.OnUi(owner, delegate
                {
                    Version latest = UpdateChecker.ParseTag(tag);
                    Version current = Assembly.GetExecutingAssembly().GetName().Version;
                    // Настройки перечитываем здесь, а не при запуске воркера: за десять
                    // секунд ожидания сети пользователь мог отказаться от напоминаний.
                    if (UpdateChecker.ShouldNotify(latest, current, UserSettings.Load().SkippedVersion))
                        OfferUpdate(owner, latest, current, true);
                });
            });
        }

        /// <summary>
        /// Одно на оба пути окно «доступна новая версия»: два одинаковых сообщения в двух
        /// местах разъехались бы при первой же правке текста. Флажок «больше не напоминать»
        /// нужен только проверке при запуске — по кнопке человек спросил сам, и предлагать
        /// ему отписаться от собственного действия незачем.
        /// </summary>
        private static void OfferUpdate(Form owner, Version latest, Version current, bool withSkipOption)
        {
            string header = string.Format(Loc.T("update.available.title"), latest.ToString(3));
            string body = string.Format(Loc.T("update.available.body"), current.ToString(3));
            // Ветки «с флажком» и «без» различаются ТОЛЬКО подписью флажка (пустая — флажка
            // нет): развилка из двух вызовов дала бы два разных значка у одного и того же
            // сообщения — синий при запуске и оранжевый по кнопке.
            bool skip;
            bool open = MessageForm.ShowConfirm(owner, Title, header, body,
                withSkipOption ? Loc.T("update.skip") : null, out skip);
            if (skip)
                UserSettings.SaveSkippedVersion(latest.ToString(3));
            if (open)
                Ui.OpenUrlOrShow(owner, Title, UpdateChecker.ReleasesPage);
        }
    }
}
