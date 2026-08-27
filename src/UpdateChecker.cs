using System;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Проверка обновлений через GitHub: чтение последней версии и, только если она
    /// новее, короткой локализованной сводки. Ничего не скачивает и не заменяет:
    /// самозаменяющиеся exe ловят антивирусы, а документам сеть вообще не нужна.
    /// Страницу релиза открывает браузер только по клику.
    /// </summary>
    internal static class UpdateChecker
    {
        private const string LatestApi = "https://api.github.com/repos/DedovMosol/iwoHelperDesktop/releases/latest";
        public const string ReleasesPage = "https://github.com/DedovMosol/iwoHelperDesktop/releases/latest";

        // Короткая сводка «что нового» лежит ОТДЕЛЬНЫМ файлом в репозитории, а не в теле
        // релиза: тело релиза читают люди на странице загрузки, и служебная разметка для
        // диалога там лишняя. Локальной строкой обойтись нельзя — описывать надо версию,
        // которой у пользователя ещё нет.
        private const string WhatsNewUrl = "https://raw.githubusercontent.com/DedovMosol/iwoHelperDesktop/main/docs/whatsnew.json";

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

        /// <summary>
        /// Версия для показа человеку: до трёх чисел, но не больше, чем в ней есть.
        ///
        /// Прямой <c>ToString(3)</c> БРОСАЕТ <see cref="ArgumentException"/> на версии из двух
        /// чисел, а тег вида «v1.18» на GitHub поставить никто не мешает — версия приходит
        /// СНАРУЖИ. На пути проверки при запуске это означало бы отчёт о сбое сразу после
        /// открытия программы, причём у всех пользователей разом и без их участия.
        /// Чистая — под тест.
        /// </summary>
        public static string Display(Version version)
        {
            if (version == null)
                return "";
            return version.ToString(version.Build >= 0 ? 3 : 2);
        }

        /// <summary>Запрос последнего тега с GitHub (сеть). Бросает при ошибке/недоступности.</summary>
        public static string FetchLatestTag()
        {
            string json = Download(LatestApi, "application/vnd.github+json");
            Match m = Regex.Match(json, "\"tag_name\"\\s*:\\s*\"([^\"]+)\"");
            if (!m.Success)
                throw new Exception(Loc.T("update.err.parseVersion"));
            return m.Groups[1].Value;
        }

        /// <summary>
        /// Короткая сводка «что нового» для найденной версии (сеть). НЕ бросает: сводка —
        /// дополнение к сообщению, а не его смысл, и из-за недоступного файла человек не
        /// должен вместо новости о новой версии получить отчёт о сбое. Пусто — значит
        /// покажем обычный текст.
        /// </summary>
        public static string FetchWhatsNew(string version, string lang)
        {
            try { return ExtractNotes(Download(WhatsNewUrl, "application/json"), version, lang); }
            catch { return ""; }
        }

        private static string Download(string url, string accept)
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.UserAgent = "iwoHelperDesktop"; // GitHub требует User-Agent
            request.Accept = accept;
            request.Timeout = 10000;
            using (var response = (HttpWebResponse)request.GetResponse())
            using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                return reader.ReadToEnd();
        }

        private static string UnescapeJson(string s)
        {
            if (string.IsNullOrEmpty(s))
                return s;
            var sb = new StringBuilder(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == '\\' && i + 1 < s.Length)
                {
                    char next = s[i + 1];
                    if (next == 'n') { sb.Append('\n'); i++; }
                    else if (next == 'r') { sb.Append('\r'); i++; }
                    else if (next == 't') { sb.Append('\t'); i++; }
                    else if (next == '"') { sb.Append('"'); i++; }
                    else if (next == '\\') { sb.Append('\\'); i++; }
                    else if (next == 'u' && i + 5 < s.Length)
                    {
                        string hex = s.Substring(i + 2, 4);
                        int code;
                        if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out code))
                        {
                            sb.Append((char)code);
                            i += 5;
                        }
                        else
                        {
                            sb.Append(s[i]);
                        }
                    }
                    else
                    {
                        sb.Append(s[i]);
                    }
                }
                else
                {
                    sb.Append(s[i]);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Сводка для версии на нужном языке из whatsnew.json. Пусто, если версии в файле
        /// нет, нет нужного языка или файл битый — сводка необязательна, см. FetchWhatsNew.
        ///
        /// Разбор регуляркой, а не парсером JSON: в .NET Framework без сторонних сборок
        /// своего JSON нет, а формат тут наш и плоский — версия, внутри две строки.
        /// Ключ ищется как «"версия": {», поэтому упоминание номера версии в тексте
        /// (в том числе в поясняющем _comment) за ключ не сойдёт. Чистая — под тест.
        /// </summary>
        public static string ExtractNotes(string json, string version, string lang)
        {
            if (string.IsNullOrEmpty(json) || string.IsNullOrEmpty(version) || string.IsNullOrEmpty(lang))
                return "";
            Match key = Regex.Match(json, "\"" + Regex.Escape(version) + "\"\\s*:\\s*\\{");
            if (!key.Success)
                return "";
            int start = key.Index + key.Length;
            int end = FindObjectEnd(json, start);
            if (end < 0)
                return "";
            Match m = Regex.Match(json.Substring(start, end - start),
                "\"" + Regex.Escape(lang) + "\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"",
                RegexOptions.IgnoreCase);
            return m.Success ? LimitNotes(UnescapeJson(m.Groups[1].Value).Trim()) : "";
        }

        /// <summary>
        /// Закрывающая скобка объекта, начатого на позиции start. Поиск по первому «}» врал бы
        /// на скобке ВНУТРИ текста сводки: обрезанный кусок мог оборваться до нужного языка,
        /// и английский текст пропал бы из-за скобки в русском. Строки пропускаем целиком,
        /// уважая экранирование. −1, если объект не закрыт (файл битый).
        /// </summary>
        private static int FindObjectEnd(string json, int start)
        {
            bool inString = false;
            for (int i = start; i < json.Length; i++)
            {
                char c = json[i];
                if (inString)
                {
                    if (c == '\\') i++;            // экранированная кавычка строку не закрывает
                    else if (c == '"') inString = false;
                }
                else if (c == '"') inString = true;
                else if (c == '}') return i;
                else if (c == '{') return -1;      // вложенный объект — формат не наш, не угадываем
            }
            return -1;
        }

        /// <summary>Сколько строк сводки влезает в диалог, не выпихивая кнопки за экран.</summary>
        private const int MaxNoteLines = 6;

        /// <summary>
        /// Обрезка сводки до <see cref="MaxNoteLines"/> строк. Окно сообщения — фиксированной
        /// ширины и растёт вниз без ограничения, поэтому длинная сводка уехала бы кнопками за
        /// нижний край экрана. Ограничиваем ЗДЕСЬ, а не в окне: текст приходит из сети, и
        /// править его после выпуска версии нельзя, а окно — общее для всех сообщений.
        /// Отброшенные строки заменяются на «…»: молча потерянный текст выглядел бы как
        /// оборванная на середине новость. Чистая — под тест.
        /// </summary>
        public static string LimitNotes(string notes)
        {
            if (string.IsNullOrEmpty(notes))
                return "";
            string[] lines = notes.Replace("\r\n", "\n").Split('\n');
            if (lines.Length <= MaxNoteLines)
                return notes;
            var kept = new StringBuilder();
            for (int i = 0; i < MaxNoteLines; i++)
            {
                if (i > 0)
                    kept.Append("\r\n");
                kept.Append(lines[i].TrimEnd());
            }
            kept.Append("\r\n…");
            return kept.ToString();
        }
    }

    /// <summary>Интерактивная проверка обновлений: сеть в фоне, результат — в UI-потоке.</summary>
    internal static class UpdateUi
    {
        private const string Title = "iwo Helper Desktop";

        // Окно об обновлении в любой момент времени ровно одно. Проверка при запуске и
        // проверка по кнопке — два НЕЗАВИСИМЫХ воркера: нажатие кнопки в первые десять
        // секунд после запуска давало бы два одинаковых сообщения, одно поверх другого.
        // Флаг трогает только UI-поток (оба показа приходят через Ui.OnUi), поэтому
        // блокировка не нужна, а вложенный показ отсекается сам: пока модальное окно
        // держит поток, флаг поднят.
        private static bool _windowOpen;

        internal static void ShowOnce(Action show)
        {
            if (_windowOpen)
                return;
            _windowOpen = true;
            try { show(); }
            finally { _windowOpen = false; }
        }

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
                string notes = null;
                Exception error = null;
                try
                {
                    tag = UpdateChecker.FetchLatestTag();
                    notes = FetchNotesIfNewer(tag);
                }
                catch (Exception ex) { error = ex; }
                Ui.OnUi(owner, delegate // общий guard: своя копия уже дважды теряла catch
                {
                    // Кнопку возвращаем в строй ВСЕГДА, даже если показ отменён чужим
                    // окном: иначе она осталась бы погашенной навсегда.
                    if (done != null)
                        done();
                    ShowOnce(delegate { ShowResult(owner, tag, notes, error); });
                });
            });
        }

        /// <summary>
        /// Сводка «что нового» для тега — вызывается из ФОНОВОГО потока сразу за запросом
        /// тега, чтобы окно открывалось с готовым текстом, а не досылало его потом.
        /// Второй запрос делаем только когда версия новее: у большинства запусков она
        /// не новее, и ходить за сводкой, которую никто не увидит, незачем.
        /// </summary>
        private static string FetchNotesIfNewer(string tag)
        {
            Version latest = UpdateChecker.ParseTag(tag);
            Version current = Assembly.GetExecutingAssembly().GetName().Version;
            if (!UpdateChecker.IsNewer(latest, current))
                return "";
            return UpdateChecker.FetchWhatsNew(UpdateChecker.Display(latest), Loc.Code(Loc.Current));
        }

        private static void ShowResult(Form owner, string tag, string notes, Exception error)
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
                OfferUpdate(owner, latest, current, notes, false);
            else
                Dialogs.Info(owner, Title, Loc.T("update.none.title"), string.Format(Loc.T("update.none.body"), UpdateChecker.Display(current)));
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
                string notes = null;
                try
                {
                    tag = UpdateChecker.FetchLatestTag();
                    notes = FetchNotesIfNewer(tag);
                }
                catch { return; } // нет сети, нет ответа, мусор в ответе — молча
                Ui.OnUi(owner, delegate
                {
                    Version latest = UpdateChecker.ParseTag(tag);
                    Version current = Assembly.GetExecutingAssembly().GetName().Version;
                    // Настройки перечитываем здесь, а не при запуске воркера: за десять
                    // секунд ожидания сети пользователь мог отказаться от напоминаний.
                    if (UpdateChecker.ShouldNotify(latest, current, UserSettings.Load().SkippedVersion))
                        ShowOnce(delegate { OfferUpdate(owner, latest, current, notes, true); });
                });
            });
        }

        /// <summary>
        /// Одно на оба пути окно «доступна новая версия»: два одинаковых сообщения в двух
        /// местах разъехались бы при первой же правке текста. Флажок «больше не напоминать»
        /// нужен только проверке при запуске — по кнопке человек спросил сам, и предлагать
        /// ему отписаться от собственного действия незачем.
        /// </summary>
        private static void OfferUpdate(Form owner, Version latest, Version current, string notes, bool withSkipOption)
        {
            string header = string.Format(Loc.T("update.available.title"), UpdateChecker.Display(latest));
            // Сводка ДОПОЛНЯЕТ вопрос, а не заменяет его: окно с двумя кнопками обязано
            // сказать, что случится по нажатию, иначе «Открыть» — это прыжок в неизвестность.
            string question = string.Format(Loc.T("update.available.body"), UpdateChecker.Display(current));
            string body = string.IsNullOrEmpty(notes) ? question : notes + "\r\n\r\n" + question;
            // Ветки «с флажком» и «без» различаются ТОЛЬКО подписью флажка (пустая — флажка
            // нет): развилка из двух вызовов дала бы два разных значка у одного и того же
            // сообщения — синий при запуске и оранжевый по кнопке.
            bool skip;
            bool open = MessageForm.ShowConfirm(owner, Title, header, body,
                withSkipOption ? Loc.T("update.skip") : null, out skip);
            if (skip && !UserSettings.SaveSkippedVersion(UpdateChecker.Display(latest)))
                Dialogs.Error(owner, Title, Loc.T("settings.err.save.title"),
                    Loc.T("settings.err.save.body"));
            if (open)
                Ui.OpenUrlOrShow(owner, Title, UpdateChecker.ReleasesPage);
        }
    }
}
