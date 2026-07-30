using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Чем один конвертер «PDF → документ» отличается от другого. Всё остальное у них общее,
    /// поэтому различия собраны в одном описании, а не размазаны по двум почти одинаковым
    /// окнам: раньше такой копией были «Объединение» и «PDF → Word», и любая правка требовала
    /// делать её дважды (см. <see cref="PdfOrderedToolFormBase"/>).
    ///
    /// Ключи каталога строятся от <see cref="Prefix"/> по единой схеме («ocr.btn.convert»,
    /// «pptx.btn.convert»), поэтому добавление третьего вывода не потребует ни одного нового
    /// поля здесь — только свою группу строк в каталоге.
    /// </summary>
    internal sealed class ConvertToolSpec
    {
        public string Prefix;        // «ocr» / «pptx» — префикс группы строк в каталоге
        public string NameKey;       // ключ названия инструмента (он же подпись карточки хаба)
        public Color Theme;          // цвет шапки окна
        public Color ThemeDark;
        public string Extension;     // «.docx» / «.pptx» — расширение результата
        public bool RequiresSta;     // нужен ли STA-поток (Word COM — да, свой писатель — нет)
        public string HistoryKey;    // подпись операции в истории
        public Action RecordUsage;   // счётчик в статистике
        public Func<IList<PdfPageRef>, string, Action<int, int>, Func<bool>, Action, ConvertResult> Convert;

        public string Key(string suffix) { return Prefix + "." + suffix; }
        public string Text(string suffix) { return Loc.T(Key(suffix)); }
    }

    /// <summary>
    /// Окно конвертера «PDF → документ»: сетка страниц с изменением порядка, кнопки правки
    /// набора и одно действие — конвертация. Различия инструментов вынесены в
    /// <see cref="ConvertToolSpec"/>, здесь только общее поведение: выбор файлов, порядок,
    /// печать, прогресс с отменой, разбор ошибок и запись в историю со статистикой.
    /// </summary>
    public class PdfConvertFormBase : PdfOrderedToolFormBase
    {
        private readonly ConvertToolSpec _spec;
        private Label _beta;
        private Button _btnOpen;
        private Button _btnUp;
        private Button _btnDown;
        private Button _btnRemove;
        private Button _btnPrint;
        private Button _btnConvert;

        internal PdfConvertFormBase(Action showHub, ConvertToolSpec spec) : base(showHub)
        {
            _spec = spec;
            BuildUi();
            SyncControls();
        }

        private string Title { get { return Loc.T(_spec.NameKey); } }

        protected override string ToolTitle { get { return Title; } }

        /// <summary>Idle-статус: пусто — подсказка добавления, иначе — число страниц к переводу.</summary>
        protected override string IdleStatusText()
        {
            return _order.Count == 0
                ? _spec.Text("status.addPdf")
                : string.Format(_spec.Text("status.pageCount"), _order.Count);
        }

        private void BuildUi()
        {
            InitShell(Title, new Size(800, 660), new Size(700, 560), _spec.Theme);
            WireFileDropAppend(); // дроп PDF на окно — добавить в конец (общая обвязка базы)
            BuildHeaderWithHome(Title, _spec.Text("header.subtitle"), _spec.Theme, _spec.ThemeDark, ShowHelp);

            int right = ClientSize.Width - 20;

            // Полоска «бета» — между шапкой и сеткой, где её нельзя не увидеть и где она
            // никому не мешает. Оба перевода берут ТЕКСТОВЫЙ слой PDF, а отсканированную
            // страницу читать нечем: без этой строки человек видит пустой слайд и считает
            // сломанной программу, а не понимает, что в источнике не было текста.
            // Высоту и место всей этой части окна считает LayoutBody на каждой раскладке:
            // строка длинная, и на узком окне ей нужно больше строк, чем на широком.
            _beta = new Label();
            _beta.Text = Loc.T("convert.beta");
            _beta.ForeColor = Theme.WarnOrange;
            _beta.AutoSize = false;
            Controls.Add(_beta);

            _grid = new PdfPageGrid();
            _grid.AllowReorder = true; // перестановка страниц перетаскиванием
            _grid.ShowPositionNumbers = true; // под плиткой — позиция в итоговом документе
            _grid.AllowRotate = true; // страница выправляется ДО анализа макета (боковой текст станет строками)
            _grid.EmptyHint = _spec.Text("grid.empty");
            _grid.DropHint = Loc.T("grid.dropHint");
            WireOrderGrid(); // события порядка + контекстное меню (общая обвязка базы)
            Controls.Add(_grid);

            _btnOpen = new RoundedButton(false);
            _btnOpen.Text = _spec.Text("btn.open");
            _btnOpen.Click += delegate { PickAndAddFiles(); };
            _tips.SetToolTip(_btnOpen, _spec.Text("tip.open"));
            Controls.Add(_btnOpen);

            _btnUp = AddPanelButton(Loc.T("common.earlier"), Loc.T("common.tip.earlier"));
            _btnUp.Click += delegate { MoveSelected(false); };
            _btnDown = AddPanelButton(Loc.T("common.later"), Loc.T("common.tip.later"));
            _btnDown.Click += delegate { MoveSelected(true); };
            _btnRemove = AddPanelButton(Loc.T("common.remove"), Loc.T("common.tip.remove"));
            _btnPrint = AddPanelButton(Loc.T("common.btn.print"), Loc.T("common.tip.print"));
            _btnPrint.Click += delegate { PrintPages(SelectedOrAllPages()); };
            _btnRemove.Click += delegate { RemoveSelected(); };
            LayoutBody();

            // Ширина действия — по его подписи: «Конвертировать в PowerPoint…» длиннее, чем
            // «Конвертировать в Word…», а в английском обе длиннее русских. Фиксированная
            // ширина обрезала бы надпись многоточием на ровном месте.
            string convertText = _spec.Text("btn.convert");
            int actionWidth = ActionWidth(convertText);
            BuildBottomStrip(right, _spec.Text("status.addPdf"), actionWidth, false);

            // Действие — в правом нижнем углу (как «Сохранить PDF» в «Объединении»).
            _btnConvert = new RoundedButton(true);
            _btnConvert.Text = convertText;
            _btnConvert.SetBounds(right - actionWidth, ClientSize.Height - 58, actionWidth, 38);
            _btnConvert.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _btnConvert.Click += OnConvertClick;
            _tips.SetToolTip(_btnConvert, _spec.Text("tip.convert"));
            Controls.Add(_btnConvert);
            AcceptButton = _btnConvert;
            RegisterActionButton(_btnConvert); // база подменит её кнопкой «Отмена» во время конвертации
        }

        // Раскладка тела окна. Все числа — отступы от низа меню и от верха тела; они же стояли
        // в SetBounds, когда полоска «бета» была ростом ровно в две строки.
        private const int NoteTop = 78;      // от низа меню до полоски «бета»
        private const int NoteGap = 12;      // просвет между полоской и телом
        private const int PanelWidth = 210;  // колонка кнопок справа от сетки
        private const int MinGridHeight = 80;

        /// <summary>
        /// Ставит полоску «бета», сетку и колонку кнопок от ФАКТИЧЕСКОЙ высоты полоски.
        /// Раньше высота была задана числом (две строки), и на узком окне хвост предупреждения
        /// просто исчезал под сеткой: человек читал обрезанную фразу и не узнавал главного —
        /// что скан переводить нечем. Теперь строка переносится, а всё, что ниже, съезжает
        /// ровно на столько, сколько она заняла.
        /// </summary>
        private void LayoutBody()
        {
            // Раскладка приходит и посреди сборки окна, когда половины элементов ещё нет.
            if (_beta == null || _grid == null || _btnPrint == null)
                return;

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;
            int width = right - 20;
            int noteH = NoteHeight(_beta.Text, _beta.Font, width);
            SetBoundsIfChanged(_beta, 20, m + NoteTop, width, noteH);

            int top = ContentTop(m, noteH);
            int gridH = Math.Max(MinGridHeight, ClientSize.Height - 152 - top);
            SetBoundsIfChanged(_grid, 20, top, width - PanelWidth, gridH);

            int px = right - PanelWidth + 10, pw = PanelWidth - 10;
            SetBoundsIfChanged(_btnOpen, px, top, pw, 32);
            SetBoundsIfChanged(_btnUp, px, top + 44, pw, 30);
            SetBoundsIfChanged(_btnDown, px, top + 80, pw, 30);
            SetBoundsIfChanged(_btnRemove, px, top + 124, pw, 30);
            SetBoundsIfChanged(_btnPrint, px, top + 160, pw, 30);
        }

        /// <summary>
        /// Верх тела окна: под полоской «бета» с просветом. Чистая — под тест.
        /// </summary>
        internal static int ContentTop(int menuHeight, int noteHeight)
        {
            return menuHeight + NoteTop + noteHeight + NoteGap;
        }

        /// <summary>
        /// Высота многострочного примечания при заданной ширине — по ПЕРЕНОСУ ПО СЛОВАМ, тем же
        /// шрифтом, каким подпись рисует. Меряем по чуть меньшей ширине: у надписи есть свои
        /// поля, и замер впритык давал на строку меньше, чем выходило на экране.
        /// </summary>
        internal static int NoteHeight(string text, Font font, int width)
        {
            if (string.IsNullOrEmpty(text) || font == null || width <= 0)
                return font != null ? font.Height : 0;
            Size measured = TextRenderer.MeasureText(text, font, new Size(Math.Max(1, width - 6), int.MaxValue),
                TextFormatFlags.WordBreak);
            return Math.Max(measured.Height, font.Height);
        }

        /// <summary>
        /// Ставит элемент на место, только если оно изменилось: присваивание внутри раскладки
        /// запускает её заново, и без этой проверки вышла бы бесконечная рекурсия.
        /// </summary>
        private static void SetBoundsIfChanged(Control c, int x, int y, int w, int h)
        {
            var want = new Rectangle(x, y, w, h);
            if (c.Bounds != want)
                c.Bounds = want;
        }

        /// <summary>
        /// Тело окна расставляем в OnLayout, а не в OnResize: раскладка идёт и при смене
        /// шрифта, и при появлении полос меню, и только к этому моменту известны настоящие
        /// размеры клиентской области.
        /// </summary>
        protected override void OnLayout(LayoutEventArgs levent)
        {
            base.OnLayout(levent);
            LayoutBody();
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, Title, Loc.T("menu.howTo"), _spec.Text("help.body"));
        }

        // ---------- конвертация ----------

        private void OnConvertClick(object sender, EventArgs e)
        {
            if (Working || _order.Count == 0)
                return;
            // Порядок/подмножество страниц из сетки (источник + индекс; страницы могут идти из разных файлов).
            List<PdfPageRef> order = _order.ToList();
            string outPath;
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = _spec.Text("filter");
                dialog.FileName = DefaultOutputName(order);
                dialog.InitialDirectory = Path.GetDirectoryName(order[0].SourcePath);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                outPath = dialog.FileName;
            }

            BeginOperation(_spec.Text("status.converting"), order.Count, _spec.Text("status.convertingPage"));
            Action<int, int> onProgress = UiProgress();
            Func<bool> cancel = CancelToken();
            // Точка невозврата: писатель зовёт это перед сохранением (документ уже наполнен,
            // отменить без следа нельзя) — снимаем предложение отмены, чтобы кнопка не «зависала».
            Action pointOfNoReturn = delegate { OnUi(delegate { StopOfferingCancel(); }); };

            Ui.RunWorker(delegate()
            {
                Exception error = null;
                ConvertResult result = null;
                try
                {
                    result = _spec.Convert(order, outPath, onProgress, cancel, pointOfNoReturn);
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                OnUi(delegate { OnConvertFinished(error, result, outPath); });
            }, sta: _spec.RequiresSta);
        }

        private void OnConvertFinished(Exception error, ConvertResult result, string outPath)
        {
            if (!FinishOperation(error, _spec.Text("status.failed"), _spec.Text("err.convertFailed")))
                return; // отмена или ошибка — статус и диалог уже показаны базой
            if (_spec.RecordUsage != null)
                _spec.RecordUsage();
            OperationHistory.Record(_spec.HistoryKey, outPath);
            SetStatus(DoneStatus(result), Theme.OkGreen);
            Ui.OpenPath(outPath); // авто-открытие результата; молча, если нет ассоциации
        }

        /// <summary>
        /// Строка результата. Если часть страниц пришла БЕЗ текста — говорим об этом прямо:
        /// такая страница в презентации есть, но пустая, и без объяснения это выглядит как
        /// поломка перевода, а не как отсутствие текста в источнике. Скан целиком мы отклоняем
        /// заранее, а смешанный документ (несколько слайдов «напечатаны» картинкой) через этот
        /// заслон проходит — и до сих пор проходил молча. Чистая — под тест.
        /// </summary>
        internal static string DoneStatus(ConvertResult result, string doneFormat)
        {
            string status = string.Format(doneFormat, result.Pages);
            int without = result.Pages - result.PagesWithText;
            if (without > 0)
                status += string.Format(Loc.T("convert.status.noTextPages"), without);
            return status;
        }

        private string DoneStatus(ConvertResult result)
        {
            return DoneStatus(result, _spec.Text("status.done"));
        }

        /// <summary>Доступность кнопок и блокировка сетки по текущему состоянию (операция/загрузка/выделение).</summary>
        protected override void SyncControls()
        {
            bool one = !Working && _grid.SelectedCount == 1;
            _grid.Locked = Working; // правки сетки (буфер, дроп) — только вне работы
            _btnOpen.Enabled = !Working;
            _btnUp.Enabled = one;
            _btnDown.Enabled = one;
            _btnRemove.Enabled = !Working && _grid.SelectedCount > 0;
            _btnPrint.Enabled = !Working && _order.Count > 0;
            _btnConvert.Enabled = !Working && _order.Count > 0;
        }

        /// <summary>Имя результата по умолчанию: из одного файла — его имя; из нескольких — «Объединённый».</summary>
        private string DefaultOutputName(List<PdfPageRef> order)
        {
            var distinct = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PdfPageRef r in order)
                distinct.Add(r.SourcePath);
            return distinct.Count == 1
                ? Path.GetFileNameWithoutExtension(order[0].SourcePath) + _spec.Extension
                : _spec.Text("defaultMerged");
        }

        /// <summary>
        /// Ширина кнопки действия под её подпись, но не меньше привычных 230 px (чтобы кнопка
        /// не «прыгала» между инструментами и языками) и не больше половины окна.
        /// </summary>
        private int ActionWidth(string caption)
        {
            // Мерить надо ТЕМ ЖЕ шрифтом, каким кнопка рисует: у главного действия он
            // полужирный и крупнее обычного, и подпись, померенная шрифтом окна, влезала
            // «на бумаге», а на экране обрезалась.
            // Шрифт КЭШИРОВАННЫЙ — не освобождать: его держат все кнопки приложения.
            Font font = Ui.Font(10.5f, FontStyle.Bold);
            // Поля 36 px подобраны так, чтобы «Конвертировать в Word…» (191 px) осталось при
            // прежних 230 px — иначе кнопка отняла бы место у строки состояния и та начала бы
            // обрезаться на минимальной ширине окна. «Конвертировать в PowerPoint…» (231 px)
            // получает свои 267 px с теми же полями по краям.
            int text = TextRenderer.MeasureText(caption, font).Width + 36;
            int min = 230, max = ClientSize.Width / 2;
            if (text < min) text = min;
            return text > max ? max : text;
        }

        /// <summary>Кнопка правой колонки: место ей найдёт <see cref="LayoutBody"/>.</summary>
        private Button AddPanelButton(string text, string tip)
        {
            var b = new RoundedButton(false);
            b.Text = text;
            _tips.SetToolTip(b, tip);
            Controls.Add(b);
            return b;
        }
    }
}
