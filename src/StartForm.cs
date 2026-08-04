using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Экран стартового окна: выбор раздела и сами разделы.</summary>
    public enum HubLevel
    {
        /// <summary>Главный: два раздела — PDF и всё остальное.</summary>
        Main,
        /// <summary>Инструменты PDF: объединение, разделение, PDF → Word, прочие операции.</summary>
        Pdf,
        /// <summary>Иной функционал: пока один инструмент — объединение Excel.</summary>
        Other
    }

    /// <summary>
    /// Стартовый экран — хаб выбора инструмента. С 1.17.9 он двухуровневый: сначала раздел
    /// (PDF или иной функционал), внутри — сами инструменты. Уровни это ПАНЕЛИ одного окна, а
    /// не разные окна: <see cref="ShellContext"/> держит единственный хаб, к нему привязаны
    /// идемпотентный показ, ответ на повторный запуск ярлыка и пересборка при смене языка —
    /// размножать окна значило бы переписывать всё это. Размер окна на всех уровнях один,
    /// иначе оно прыгало бы при каждом переходе.
    ///
    /// Только представление: открытие инструментов, дедупликацию и жизненный цикл окон ведёт
    /// <see cref="ShellContext"/>. Закрытие хаба не закрывает уже открытые инструменты, кнопка
    /// «Главная» в инструменте снова показывает этот экран — и сразу нужным разделом.
    /// </summary>
    public class StartForm : Form
    {
        private const string AppTitle = "iwo Helper Desktop";
        // Раздел PDF — пять инструментов сеткой 3×2. Третий РЯД не поместился бы: окно и так
        // почти во всю высоту экрана 1366×768, а третья КОЛОНКА помещается свободно —
        // экраны шире, чем выше. Карточки при этом остаются полноразмерными.
        private const int CardW = 240, WideW = 756, CardH = 250, Row1 = 96, Row2 = 364;
        private const int Col2 = 282, Col3 = 540, Pad = 24;
        private const int HeaderH = 78, BottomRowY = 632, BottomRowH = 36;

        private readonly ShellContext _context;
        private ToolTip _langTip;           // подсказка кнопки-глобуса (компонент — освобождаем вручную)
        private ToolTip _bottomTip;         // подсказка значка настроек в нижнем ряду
        private ContextMenuStrip _langMenu; // меню выбора языка (одно на окно; окно пересоздаётся при смене языка)
        private HeaderBand _header;
        private Button _back;
        private Panel _levelMain, _levelPdf, _levelOther;
        private ChoiceCard _firstMain, _firstPdf, _firstOther;
        private HubLevel _level = HubLevel.Main;
        // Файлы, бро́шенные на карточку «PDF» главного уровня: инструмент ещё не выбран, поэтому
        // держим их до первого клика по карточке раздела. Набор ОДНОРАЗОВЫЙ — см. ClearPending.
        private string[] _pending;

        public StartForm() : this(null) { } // для смоук-теста; открытие инструментов недоступно

        internal StartForm(ShellContext context) : this(context, HubLevel.Main) { }

        internal StartForm(ShellContext context, HubLevel level)
        {
            _context = context;

            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            Text = AppTitle + " " + version.ToString(3);
            Icon startIcon = Ui.AppIcon();
            if (startIcon != null)
                Icon = startIcon;
            Font = Ui.Font(9.75f); // общий кэшированный шрифт (не освобождать)
            BackColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;
            // Окно разворачивается на весь экран, как и окна инструментов: на большом мониторе
            // карточки в маленьком неподвижном прямоугольнике выглядят потерянными, а часть
            // работы (перетащить сюда файлы) удобнее делать в развёрнутом окне.
            // Раскладка карточек задана координатами, поэтому при изменении размера они
            // ЦЕНТРИРУЮТСЯ целой группой (CenterLevels) — растягивать их незачем, а
            // прижиматься к левому верхнему углу они не должны.
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimumSize = new Size(820, 730);
            KeyPreview = true; // Esc — назад из раздела
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(804, 692);
            WindowChrome.Enable(this, Theme.HubBlue); // синий заголовок на Windows 11

            BuildHeader();
            BuildLevels();
            BuildBottomRow();

            _designSize = ClientSize;   // от него считается центрирование карточек
            AcceptButton = null; // Enter активирует карточку в фокусе
            GoTo(level, false);  // при пересборке (смена языка) возвращаемся в тот же раздел
        }

        /// <summary>Показанный раздел — его переносит на новое окно пересборка при смене языка.</summary>
        internal HubLevel Level { get { return _level; } }

        /// <summary>Показать раздел на уже открытом хабе («Главная» из инструмента).</summary>
        internal void ShowLevel(HubLevel level)
        {
            GoTo(level, true);
        }

        private void BuildHeader()
        {
            _header = new HeaderBand(AppTitle, Loc.T("hub.subtitle"), Theme.HubBlue, Theme.HubBlueDark);
            _header.Centered = true; // на стартовом экране заголовок и подпись по центру
            _header.SetBounds(0, 0, ClientSize.Width, HeaderH);
            _header.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(_header);

            // «Назад» — слева, зеркально глобусу справа. Показывается только внутри раздела.
            _back = new RoundedButton(false);
            _back.Text = Loc.T("hub.back");
            _back.SetBounds(12, 0, 104, 30);
            _back.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            _back.Visible = false;
            _back.Click += delegate { GoTo(HubLevel.Main, true); };
            _header.Controls.Add(_back);
            _header.AlignToText(_back);

            // Выбор языка — белый глиф-глобус в правом верхнем углу шапки (на синем, без рамки).
            _langMenu = HelpMenu.LanguageContextMenu(); // одно меню на жизнь окна
            // Запасной рисунок — флаг текущего языка: если шрифта глифов нет, кнопка всё равно
            // должна говорить, что она про язык (общий кэш на процесс — не освобождать).
            var globe = new GlyphButton("", 15f, "Segoe MDL2 Assets",
                delegate { return Flags.For(Loc.Current); }); // U+E774 — «глобус»
            globe.ForeColor = Color.White;
            globe.HoverFill = Color.FromArgb(46, 255, 255, 255); // на синей шапке отзывается высветлением
            globe.AccessibleName = Loc.T("lang.tooltip"); // с клавиатуры и для экранного диктора
            // Y задаёт сама шапка (AlignToText): глобус встаёт по центру пары «заголовок +
            // подпись» при любом масштабе экрана, здесь только колонка.
            globe.SetBounds(_header.Width - 42, 0, 30, 30);
            globe.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            globe.Click += delegate { _langMenu.Show(globe, new Point(globe.Width, globe.Height), ToolStripDropDownDirection.BelowLeft); };
            _langTip = new ToolTip();
            _langTip.SetToolTip(globe, Loc.T("lang.tooltip"));
            _header.Controls.Add(globe);
            _header.AlignToText(globe);
        }

        // ---------- уровни ----------

        private void BuildLevels()
        {
            _levelMain = AddLevelPanel();
            _levelPdf = AddLevelPanel();
            _levelOther = AddLevelPanel();

            // Главный уровень: два раздела во всю ширину. Описание перечисляет содержимое —
            // так видно, куда ведёт кнопка, ещё до нажатия.
            _firstMain = AddCard(_levelMain, CardGlyph.Pdf, Loc.T("hub.section.pdf.name"),
                Loc.T("hub.section.pdf.desc"), Pad, Row1, WideW);
            _firstMain.Click += delegate { GoTo(HubLevel.Pdf, true); };
            _firstMain.AcceptFiles(".pdf");
            // Дроп на раздел: инструмент ещё не выбран, поэтому переходим внутрь и придерживаем
            // файлы до первого клика по карточке инструмента.
            _firstMain.FilesDropped += delegate(string[] files) { GoTo(HubLevel.Pdf, true); SetPending(files); };

            ChoiceCard other = AddCard(_levelMain, CardGlyph.Other, Loc.T("hub.section.other.name"),
                Loc.T("hub.section.other.desc"), Pad, Row2, WideW);
            other.Click += delegate { GoTo(HubLevel.Other, true); };

            // Раздел PDF: четыре инструмента сеткой 2×2.
            Func<Action, Form> ocrFactory = delegate(Action back) { return new OcrForm(back); };
            Func<Action, Form> pptxFactory = delegate(Action back) { return new PptxForm(back); };
            Func<Action, Form> opsFactory = delegate(Action back) { return new PdfOpsForm(back); };
            // Мост в «Прочие операции»: им пользуются и «Разделение» (открытый документ), и
            // «Объединение» (собранный файл). Фабрику окна знает стартовый экран — он же и
            // композиционный корень инструментов, поэтому они не ссылаются друг на друга.
            Action<string> openOps = delegate(string path)
            {
                if (_context == null)
                    return;
                // Пустой путь — «просто открой окно»: инструмент есть инструмент, и нажатие
                // никогда не должно упираться в объяснение вместо действия. Тот же вызов и
                // открывает окно, и поднимает уже открытое, и отдаёт ему документ.
                string[] files = string.IsNullOrEmpty(path) ? new string[0] : new[] { path };
                _context.OpenToolWithFiles("ops", Loc.T("hub.ops.name"), opsFactory, files, HubLevel.Pdf);
            };
            Func<Action, Form> mergeFactory = delegate(Action back) { return new PdfMergeForm(back, openOps); };
            Func<Action, Form> splitFactory = delegate(Action back) { return new PdfSplitForm(back, openOps); };

            // Верхний ряд — операции НАД страницами, нижний — переводы в другой формат.
            _firstPdf = AddTool(_levelPdf, CardGlyph.Pdf, "pdf", "hub.pdf.name", "hub.pdf.desc",
                mergeFactory, Pad, Row1, CardW);
            AddTool(_levelPdf, CardGlyph.PdfSplit, "split", "hub.split.name", "hub.split.desc",
                splitFactory, Col2, Row1, CardW);
            AddTool(_levelPdf, CardGlyph.Tools, "ops", "hub.ops.name", "hub.ops.desc",
                opsFactory, Col3, Row1, CardW);
            AddTool(_levelPdf, CardGlyph.Ocr, "ocr", "hub.ocr.name", "hub.ocr.desc",
                ocrFactory, Pad, Row2, CardW);
            AddTool(_levelPdf, CardGlyph.Pptx, "pptx", "hub.pptx.name", "hub.pptx.desc",
                pptxFactory, Col2, Row2, CardW);

            // Иной функционал: пока один инструмент, место под следующие размечено той же
            // сеткой, что и в разделе PDF.
            _firstOther = AddTool(_levelOther, CardGlyph.Excel, "excel", "hub.excel.name", "hub.excel.desc",
                delegate(Action back) { return new MainForm(back); }, Pad, Row1, WideW);
        }

        /// <summary>Исходные места карточек: от них считается центрирование при изменении размера.</summary>
        private readonly Dictionary<Control, Rectangle> _cardHome = new Dictionary<Control, Rectangle>();
        private Size _designSize;

        /// <summary>
        /// Сдвинуть карточки каждого уровня так, чтобы группа стояла по центру окна. Растягивать
        /// их было бы хуже: у карточки есть свой размер, при котором значок, название и описание
        /// читаются, и растянутая на полэкрана карточка перестаёт быть карточкой.
        /// </summary>
        private void CenterLevels()
        {
            if (_designSize.Width <= 0)
                return;
            int dx = (ClientSize.Width - _designSize.Width) / 2;
            int dy = (ClientSize.Height - _designSize.Height) / 2;
            foreach (KeyValuePair<Control, Rectangle> home in _cardHome)
            {
                Rectangle r = home.Value;
                home.Key.Location = new Point(r.X + (dx > 0 ? dx : 0), r.Y + (dy > 0 ? dy : 0));
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            CenterLevels();
        }

        private Panel AddLevelPanel()
        {
            var panel = new Panel();
            // От шапки до нижнего ряда, а НЕ до низа окна: панель непрозрачна и, растянутая на
            // всё окно, закрыла бы собой «Проверить обновления» и «О программе» (они добавлены
            // позже, а значит ниже по z-порядку).
            panel.SetBounds(0, HeaderH, ClientSize.Width, BottomRowY - HeaderH);
            panel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            panel.BackColor = Color.White;
            panel.Visible = false;
            Controls.Add(panel);
            return panel;
        }

        /// <summary>Карточка раздела или инструмента. Координаты — в клиентских координатах окна.</summary>
        private ChoiceCard AddCard(Panel level, CardGlyph glyph, string title, string desc, int x, int y, int width)
        {
            var card = new ChoiceCard(glyph, title, desc);
            card.SetBounds(x, y - level.Top, width, CardH); // панель начинается под шапкой
            level.Controls.Add(card);
            _cardHome[card] = card.Bounds;   // отсюда считается центрирование при изменении размера
            return card;
        }

        /// <summary>Карточка инструмента: клик открывает его, дроп PDF — открывает с файлами.</summary>
        private ChoiceCard AddTool(Panel level, CardGlyph glyph, string key, string nameKey, string descKey,
            Func<Action, Form> factory, int x, int y, int width)
        {
            ChoiceCard card = AddCard(level, glyph, Loc.T(nameKey), Loc.T(descKey), x, y, width);
            HubLevel home = level == _levelOther ? HubLevel.Other : HubLevel.Pdf;
            card.Click += delegate { OpenTool(key, nameKey, factory, home); };
            if (key != "excel") // свод Excel работает с ПАПКОЙ, не с файлами — дроп ему не нужен
            {
                // «Прочие операции» собирают документ и из картинок, поэтому их карточка
                // принимает то же, что и само окно (иначе дроп снимка на неё молча отвергался бы).
                card.AcceptFiles(key == "ops" ? PdfDrop.PdfAndImages : PdfDrop.PdfOnly);
                card.FilesDropped += delegate(string[] files)
                {
                    if (_context != null)
                        _context.OpenToolWithFiles(key, Loc.T(nameKey), factory, files, home);
                    ClearPending();
                };
            }
            return card;
        }

        /// <summary>
        /// Открыть инструмент. Если на раздел бросали файлы и их ещё не разобрали — отдаём их
        /// выбранному инструменту: человек уже показал, с чем хочет работать.
        /// </summary>
        private void OpenTool(string key, string nameKey, Func<Action, Form> factory, HubLevel home)
        {
            if (_context == null)
                return;
            string[] files = _pending;
            ClearPending();
            if (files != null && files.Length > 0)
                _context.OpenToolWithFiles(key, Loc.T(nameKey), factory, files, home);
            else
                _context.OpenTool(key, Loc.T(nameKey), factory, home);
        }

        /// <summary>Показать раздел. focus — переводить ли фокус на первую карточку (переход руками).</summary>
        private void GoTo(HubLevel level, bool focus)
        {
            if (level != HubLevel.Pdf)
                ClearPending(); // набор живёт только внутри раздела PDF
            _level = level;
            _levelMain.Visible = level == HubLevel.Main;
            _levelPdf.Visible = level == HubLevel.Pdf;
            _levelOther.Visible = level == HubLevel.Other;
            _back.Visible = level != HubLevel.Main;
            _header.Subtitle = SubtitleFor(level);
            if (!focus)
                return;
            // Фокус обязателен: он остаётся на спрятанной карточке, и с клавиатуры экран
            // становится неуправляемым — Tab начинает обход с непонятного места.
            ChoiceCard first = level == HubLevel.Pdf ? _firstPdf
                : level == HubLevel.Other ? _firstOther : _firstMain;
            if (first != null && first.Visible)
                first.Focus();
        }

        /// <summary>Подпись шапки под текущий раздел (и под ожидающие файлы, если их бросили).</summary>
        private string SubtitleFor(HubLevel level)
        {
            if (level == HubLevel.Pdf)
                return _pending != null
                    ? string.Format(Loc.T("hub.pending"), _pending.Length)
                    : Loc.T("hub.subtitle.pdf");
            return level == HubLevel.Other ? Loc.T("hub.subtitle.other") : Loc.T("hub.subtitle");
        }

        private void SetPending(string[] files)
        {
            _pending = files != null && files.Length > 0 ? files : null;
            _header.Subtitle = SubtitleFor(_level);
        }

        /// <summary>
        /// Забыть придержанные файлы. Зовётся отовсюду, где набор перестаёт быть актуальным:
        /// инструмент открыт, ушли из раздела, бросили новые файлы. Иначе набор «прилипнет» и
        /// следующий клик откроет инструмент с чужими файлами.
        /// </summary>
        private void ClearPending()
        {
            if (_pending == null)
                return;
            _pending = null;
            _header.Subtitle = SubtitleFor(_level);
        }

        private void BuildBottomRow()
        {
            // Нижний ряд виден на всех уровнях: «Настройки» слева, «О программе» справа.
            // Проверка обновлений уехала ВНУТРЬ настроек: с 1.18.0 она идёт при запуске сама,
            // и ручная кнопка из первого ряда превратилась в редкое действие. Заодно рядом с
            // ней встал её собственный выключатель — раньше его негде было показать.
            // «Настройки» — значком-шестернёй, а не подписью: значок понятен без чтения и не
            // занимает четверть нижнего ряда. Подпись остаётся в подсказке и в имени для
            // экранного диктора — узнаваемость значка не должна быть единственной опорой.
            var settings = new GlyphButton("", 17f, "Segoe MDL2 Assets"); // U+E713 — «шестерня»
            // Свой цвет, а не общий тёмно-серый: две одинаково серые закорючки по краям белого
            // ряда сливались и между собой, и с фоном — по значку нельзя было понять, куда
            // ведёт нажатие, не наводя на него курсор.
            settings.ForeColor = Theme.SettingsBlue;
            settings.HoverFill = Color.FromArgb(30, Theme.SettingsBlue);
            settings.AccessibleName = Loc.T("settings.title");
            settings.SetBounds(Pad, BottomRowY, BottomRowH, BottomRowH);
            settings.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            settings.Click += delegate { using (var f = new SettingsForm()) f.ShowDialog(this); };
            Controls.Add(settings);
            _bottomTip = new ToolTip();
            _bottomTip.SetToolTip(settings, Loc.T("settings.title"));

            // «О программе» — знаком вопроса, парно к шестерне настроек: оба служебных действия
            // нижнего ряда одного размера и не спорят с карточками инструментов, но цвет у
            // каждого свой — форма значка на белом читается хуже, чем цветовое пятно.
            var about = new GlyphButton("", 17f, "Segoe MDL2 Assets"); // U+E9CE — «вопрос в круге»
            about.ForeColor = Theme.HelpTeal;
            about.HoverFill = Color.FromArgb(30, Theme.HelpTeal);
            about.AccessibleName = Loc.T("hub.about");
            about.SetBounds(ClientSize.Width - Pad - BottomRowH, BottomRowY, BottomRowH, BottomRowH);
            about.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            about.Click += delegate { using (var f = new AboutForm()) f.ShowDialog(this); };
            Controls.Add(about);
            _bottomTip.SetToolTip(about, Loc.T("hub.about"));

            BuildRecentFiles();
        }

        /// <summary>Сколько последних файлов показывать карточками.</summary>
        private const int RecentCount = 3;
        /// <summary>Полоса карточек недавних — над нижним рядом со значками.</summary>
        private const int RecentCardH = 52;
        private const int RecentRowY = BottomRowY - RecentCardH - 10;
        private readonly List<RecentCard> _recentCards = new List<RecentCard>();

        /// <summary>
        /// Карточки последних сделанных файлов — прямо на стартовом экране, значком того
        /// инструмента, который их сделал. Обновляются сами: подписка на историю снимает
        /// нужду перезапускать программу, чтобы увидеть только что полученный файл.
        ///
        /// История читается с диска, а существование файлов проверяется обращением к ним —
        /// и это НЕ работа для потока интерфейса: путь на отключённой сетевой шаре отвечает
        /// около секунды на файл (замерено), то есть три записи задержали бы появление окна
        /// на три секунды. Поэтому в фоне, а карточки строятся по готовому ответу.
        /// </summary>
        private void BuildRecentFiles()
        {
            OperationHistory.Changed += OnHistoryChanged;   // отписка — в Dispose
            RefreshRecentFiles();
        }

        private void OnHistoryChanged()
        {
            // Событие приходит с того потока, где закончилась операция, — возвращаемся к себе.
            Ui.OnUi(this, RefreshRecentFiles);
        }

        private void RefreshRecentFiles()
        {
            if (IsDisposed || Disposing)
                return;
            Ui.RunWorker(delegate()
            {
                List<HistoryEntry> found;
                try { found = OperationHistory.RecentFiles(OperationHistory.Load(), RecentCount, System.IO.File.Exists); }
                catch { return; } // история недоступна — стартовому экрану это безразлично
                Ui.OnUi(this, delegate { ShowRecentFiles(found); });
            });
        }

        /// <summary>Разложить карточки недавних файлов в нижней части экрана. Только UI-поток.</summary>
        private void ShowRecentFiles(List<HistoryEntry> recent)
        {
            if (IsDisposed || Disposing)
                return;
            // Прежние карточки убираем целиком: список пересобирается после каждой операции,
            // и подчищать его по одной записи было бы больше кода без единой выгоды.
            foreach (RecentCard old in _recentCards)
            {
                Controls.Remove(old);
                old.Dispose();
            }
            _recentCards.Clear();
            if (recent == null || recent.Count == 0)
                return;   // возвращаться не к чему — не занимаем место пустым обещанием

            int gap = 12;
            int width = (ClientSize.Width - 2 * Pad - (RecentCount - 1) * gap) / RecentCount;
            int x = Pad;
            foreach (HistoryEntry entry in recent)
            {
                var card = new RecentCard(entry);
                card.SetBounds(x, RecentRowY, width, RecentCardH);
                card.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
                Controls.Add(card);
                card.BringToFront();
                _recentCards.Add(card);
                x += width + gap;
            }
        }

        /// <summary>Esc возвращает из раздела на главный экран (как «Назад»).</summary>
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape && _level != HubLevel.Main)
            {
                GoTo(HubLevel.Main, true);
                e.Handled = true;
                return;
            }
            base.OnKeyDown(e);
        }

        /// <summary>
        /// Окно собрано: шапку с глобусом — в конец обхода Tab. Глобус доступен с клавиатуры,
        /// а шапка добавляется первой, поэтому иначе фокус при открытии доставался бы выбору
        /// языка вместо карточки инструмента. См. <see cref="Ui.HeaderLastInTabOrder"/>.
        /// </summary>
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            Ui.HeaderLastInTabOrder(this);
        }

        protected override void Dispose(bool disposing)
        {
            // ToolTip и ContextMenuStrip — компоненты (не дочерние контролы): освобождаем вручную.
            if (disposing)
            {
                // Отписка обязательна: статическое событие держало бы закрытое окно живым, а
                // при смене языка окно пересоздаётся — подписки копились бы за сеанс.
                OperationHistory.Changed -= OnHistoryChanged;
                if (_langTip != null) _langTip.Dispose();
                if (_bottomTip != null) _bottomTip.Dispose();
                if (_langMenu != null) _langMenu.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Иконка-глиф без рамки на прозрачном фоне (для белого глобуса на синей шапке).
        /// SupportsTransparentBackColor + BackColor=Transparent показывает фон родителя
        /// (градиент шапки); глиф рисуется по центру. Кликается как кнопка и доступна
        /// с клавиатуры (TabStop + Enter/Пробел): это единственный переключатель языка
        /// на стартовом экране, у которого нет меню «☰».
        ///
        /// Шрифта может не быть: «Segoe MDL2 Assets» появился только в Windows 10, а
        /// минимум приложения — Windows 8.1. GDI молча подставит другой шрифт, и глиф из
        /// области частного использования нарисовался бы пустотой — переключатель языка
        /// стал бы невидимым пятном. Подстановку видно по Font.Name, и тогда рисуем флаг
        /// текущего языка: он рисуется через GDI+ (Flags) и есть везде, а в самом меню
        /// выбора языка флаги и так стоят.
        /// </summary>
        private sealed class GlyphButton : Control
        {
            private readonly bool _glyphFontPresent;

            private readonly Func<Image> _fallback;

            private bool _hover, _pressed;

            /// <summary>
            /// Подложка под значком при наведении и нажатии. Значок без рамки ничем не отвечает
            /// на курсор, а «нажимается или нет» — это ровно то, что мышью выясняют наведением.
            /// Цвет задаёт вызывающий: на белом ряду нужна светло-серая, на синей шапке — белая
            /// полупрозрачная, и обратное в каждом случае было бы невидимо.
            /// </summary>
            public Color HoverFill = Color.Empty;

            public GlyphButton(string glyph, float size, string family, Func<Image> fallback = null)
            {
                _fallback = fallback;
                SetStyle(ControlStyles.SupportsTransparentBackColor | ControlStyles.UserPaint |
                         ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
                BackColor = Color.Transparent;
                Font = Ui.Font(size, FontStyle.Regular, family); // кэш: глобус пересоздаётся с окном
                _glyphFontPresent = string.Equals(Font.Name, family, StringComparison.OrdinalIgnoreCase);
                Text = glyph;
                Cursor = Cursors.Hand;
                TabStop = true;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                if (HoverFill != Color.Empty && (_hover || _pressed))
                {
                    // Круг, а не прямоугольник: значок квадратный и мелкий, прямоугольная плашка
                    // на его фоне выглядит крупнее самого значка.
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    int alpha = _pressed ? Math.Min(255, HoverFill.A * 2) : HoverFill.A;
                    using (var b = new SolidBrush(Color.FromArgb(alpha, HoverFill)))
                        e.Graphics.FillEllipse(b, 0, 0, Width - 1, Height - 1);
                    e.Graphics.SmoothingMode = SmoothingMode.Default;
                }
                Image fallback = _glyphFontPresent || _fallback == null ? null : _fallback();
                if (fallback != null)
                {
                    e.Graphics.DrawImage(fallback, (Width - fallback.Width) / 2, (Height - fallback.Height) / 2);
                }
                else
                {
                    // Рисуем СЕРЫМ сглаживанием, а не субпиксельным: у мелкого глифа на светлом
                    // фоне цветная кайма видна как грязь по краям — значок выглядит цветным,
                    // хотя он одноцветный.
                    TextRenderingHint was = e.Graphics.TextRenderingHint;
                    e.Graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
                    using (var brush = new SolidBrush(ForeColor))
                    using (var format = new StringFormat())
                    {
                        format.Alignment = StringAlignment.Center;
                        format.LineAlignment = StringAlignment.Center;
                        e.Graphics.DrawString(Text, Font, brush, ClientRectangle, format);
                    }
                    e.Graphics.TextRenderingHint = was;
                }
                if (Focused) // видимый фокус: иначе с клавиатуры непонятно, где ты находишься
                    ControlPaint.DrawFocusRectangle(e.Graphics, ClientRectangle);
            }

            /// <summary>Enter и Пробел работают как клик — иначе с клавиатуры кнопка бесполезна.</summary>
            protected override bool IsInputKey(Keys keyData)
            {
                return keyData == Keys.Enter || keyData == Keys.Space || base.IsInputKey(keyData);
            }

            protected override void OnKeyDown(KeyEventArgs e)
            {
                base.OnKeyDown(e);
                if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space)
                {
                    e.Handled = true;
                    InvokeOnClick(this, EventArgs.Empty);
                }
            }

            protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
            protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }

            protected override void OnMouseEnter(EventArgs e) { base.OnMouseEnter(e); _hover = true; Invalidate(); }
            protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); _hover = false; _pressed = false; Invalidate(); }
            protected override void OnMouseDown(MouseEventArgs e)
            {
                base.OnMouseDown(e);
                if (e.Button == MouseButtons.Left) { _pressed = true; Invalidate(); }
            }
            protected override void OnMouseUp(MouseEventArgs e) { base.OnMouseUp(e); _pressed = false; Invalidate(); }
        }
    }
}
