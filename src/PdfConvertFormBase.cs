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

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;
            int panelW = 210;
            int gridBottom = ClientSize.Height - 152;

            // Полоска «бета» — между шапкой и сеткой, где её нельзя не увидеть и где она
            // никому не мешает. Оба перевода берут ТЕКСТОВЫЙ слой PDF, а отсканированную
            // страницу читать нечем: без этой строки человек видит пустой слайд и считает
            // сломанной программу, а не понимает, что в источнике не было текста.
            var beta = new Label();
            beta.Text = Loc.T("convert.beta");
            beta.ForeColor = Theme.WarnOrange;
            beta.AutoSize = false;
            beta.SetBounds(20, m + 78, right - 20, 32);
            beta.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(beta);
            int noteH = beta.Height + 6;

            _grid = new PdfPageGrid();
            _grid.AllowReorder = true; // перестановка страниц перетаскиванием
            _grid.ShowPositionNumbers = true; // под плиткой — позиция в итоговом документе
            _grid.AllowRotate = true; // страница выправляется ДО анализа макета (боковой текст станет строками)
            _grid.EmptyHint = _spec.Text("grid.empty");
            _grid.DropHint = Loc.T("grid.dropHint");
            _grid.SetBounds(20, m + 84 + noteH, right - 20 - panelW, gridBottom - (m + 84 + noteH));
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            WireOrderGrid(); // события порядка + контекстное меню (общая обвязка базы)
            Controls.Add(_grid);

            int px = right - panelW + 10;
            int pw = panelW - 10;
            _btnOpen = new RoundedButton(false);
            _btnOpen.Text = _spec.Text("btn.open");
            _btnOpen.SetBounds(px, m + 84 + noteH, pw, 32);
            _btnOpen.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnOpen.Click += delegate { PickAndAddFiles(); };
            _tips.SetToolTip(_btnOpen, _spec.Text("tip.open"));
            Controls.Add(_btnOpen);

            _btnUp = AddPanelButton(Loc.T("common.earlier"), px, m + 128 + noteH, pw, Loc.T("common.tip.earlier"));
            _btnUp.Click += delegate { MoveSelected(false); };
            _btnDown = AddPanelButton(Loc.T("common.later"), px, m + 164 + noteH, pw, Loc.T("common.tip.later"));
            _btnDown.Click += delegate { MoveSelected(true); };
            _btnRemove = AddPanelButton(Loc.T("common.remove"), px, m + 208 + noteH, pw, Loc.T("common.tip.remove"));
            _btnPrint = AddPanelButton(Loc.T("common.btn.print"), px, m + 244 + noteH, pw, Loc.T("common.tip.print"));
            _btnPrint.Click += delegate { PrintPages(SelectedOrAllPages()); };
            _btnRemove.Click += delegate { RemoveSelected(); };

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

        private Button AddPanelButton(string text, int x, int y, int w, string tip)
        {
            var b = new RoundedButton(false);
            b.Text = text;
            b.SetBounds(x, y, w, 30);
            b.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _tips.SetToolTip(b, tip);
            Controls.Add(b);
            return b;
        }
    }
}
