using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Инструмент «PDF → Word»: извлекает текстовый слой одного или НЕСКОЛЬКИХ цифровых PDF
    /// (сохранённых из Word, «Microsoft Print to PDF» и т.п.) в один редактируемый .docx.
    /// Страницы всех добавленных файлов показаны единой сеткой (<see cref="PdfPageGrid"/>) и
    /// собираются в выбранном порядке. Отсканированные документы (без текстового слоя) в
    /// настоящее время недоступны — при попытке будет понятное сообщение (файл цел).
    /// Конвертация обёрнута try/catch (<see cref="OnConvertClick"/>): любая ошибка — диалог,
    /// не краш. Модель порядка и её слой — в общей базе <see cref="PdfOrderedToolFormBase"/> (DRY).
    /// </summary>
    public class OcrForm : PdfOrderedToolFormBase
    {
        private static string Title { get { return Loc.T("hub.ocr.name"); } }

        private Button _btnOpen;
        private Button _btnUp;
        private Button _btnDown;
        private Button _btnRemove;
        private Button _btnConvert;

        public OcrForm() : this(null) { }

        public OcrForm(Action showHub) : base(showHub)
        {
            BuildUi();
            SyncControls();
        }

        protected override string ToolTitle { get { return Title; } }

        /// <summary>Idle-статус: пусто — подсказка добавления, иначе — число страниц к переводу.</summary>
        protected override string IdleStatusText()
        {
            return _order.Count == 0
                ? Loc.T("ocr.status.addPdf")
                : string.Format(Loc.T("ocr.status.pageCount"), _order.Count);
        }

        private void BuildUi()
        {
            InitShell(Title, new Size(800, 660), new Size(700, 560), Theme.WordViolet);
            WireFileDropAppend(); // дроп PDF на окно — добавить в конец (общая обвязка базы)
            BuildHeaderWithHome(Title,
                Loc.T("ocr.header.subtitle"),
                Theme.WordViolet, Theme.WordVioletDark, ShowHelp);

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;
            int panelW = 210;
            int gridBottom = ClientSize.Height - 152;

            _grid = new PdfPageGrid();
            _grid.AllowReorder = true; // перестановка страниц перетаскиванием
            _grid.ShowPositionNumbers = true; // под плиткой — позиция в итоговом документе
            _grid.AllowRotate = true; // страница выправляется ДО анализа макета (боковой текст станет строками)
            _grid.EmptyHint = Loc.T("ocr.grid.empty");
            _grid.DropHint = Loc.T("grid.dropHint");
            _grid.SetBounds(20, m + 84, right - 20 - panelW, gridBottom - (m + 84));
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            WireOrderGrid(); // события порядка + контекстное меню (общая обвязка базы)
            Controls.Add(_grid);

            int px = right - panelW + 10;
            int pw = panelW - 10;
            _btnOpen = new RoundedButton(false);
            _btnOpen.Text = Loc.T("ocr.btn.open");
            _btnOpen.SetBounds(px, m + 84, pw, 32);
            _btnOpen.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnOpen.Click += OnOpenClick;
            _tips.SetToolTip(_btnOpen, Loc.T("ocr.tip.open"));
            Controls.Add(_btnOpen);

            _btnUp = AddPanelButton(Loc.T("common.earlier"), px, m + 128, pw, Loc.T("common.tip.earlier"));
            _btnUp.Click += delegate { MoveSelected(false); };
            _btnDown = AddPanelButton(Loc.T("common.later"), px, m + 164, pw, Loc.T("common.tip.later"));
            _btnDown.Click += delegate { MoveSelected(true); };
            _btnRemove = AddPanelButton(Loc.T("common.remove"), px, m + 208, pw, Loc.T("common.tip.remove"));
            _btnRemove.Click += delegate { RemoveSelected(); };

            BuildBottomStrip(right, Loc.T("ocr.status.addPdf"), 230, false);

            // Действие — в правом нижнем углу (как «Сохранить PDF» в «Объединении»).
            _btnConvert = new RoundedButton(true);
            _btnConvert.Text = Loc.T("ocr.btn.convert");
            _btnConvert.SetBounds(right - 230, ClientSize.Height - 58, 230, 38);
            _btnConvert.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _btnConvert.Click += OnConvertClick;
            _tips.SetToolTip(_btnConvert, Loc.T("ocr.tip.convert"));
            Controls.Add(_btnConvert);
            AcceptButton = _btnConvert;
            RegisterActionButton(_btnConvert); // база подменит её кнопкой «Отмена» во время конвертации
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, Title, Loc.T("menu.howTo"), Loc.T("ocr.help.body"));
        }

        private void OnOpenClick(object sender, EventArgs e)
        {
            PickAndAddFiles();
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
                dialog.Filter = Loc.T("ocr.docxFilter");
                dialog.FileName = DefaultOutputName(order);
                dialog.InitialDirectory = Path.GetDirectoryName(order[0].SourcePath);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                outPath = dialog.FileName;
            }

            BeginOperation(Loc.T("ocr.status.converting"), order.Count, Loc.T("ocr.status.convertingPage"));
            Action<int, int> onProgress = UiProgress();
            Func<bool> cancel = CancelToken();
            // Точка невозврата: писатель зовёт это перед SaveDocx (Word уже наполнен, отменить
            // без следа нельзя) — снимаем предложение отмены, чтобы кнопка не «зависала».
            Action pointOfNoReturn = delegate { OnUi(delegate { StopOfferingCancel(); }); };

            Ui.RunWorker(delegate()
            {
                Exception error = null;
                ConvertResult result = null;
                try
                {
                    result = PdfToWordService.Convert(order, outPath, onProgress, cancel, pointOfNoReturn);
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                OnUi(delegate { OnConvertFinished(error, result, outPath); });
            }, sta: true); // требование Word COM
        }

        private void OnConvertFinished(Exception error, ConvertResult result, string outPath)
        {
            if (!FinishOperation(error, Loc.T("ocr.status.failed"), Loc.T("ocr.err.convertFailed")))
                return; // отмена или ошибка — статус и диалог уже показаны базой
            UsageStats.RecordPdfToWord();
            SetStatus(string.Format(Loc.T("ocr.status.done"), result.Pages), Theme.OkGreen);
            Ui.OpenPath(outPath); // авто-открытие результата; молча, если нет ассоциации .docx
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
            _btnConvert.Enabled = !Working && _order.Count > 0;
        }

        /// <summary>Имя .docx по умолчанию: из одного файла — его имя; из нескольких — «Объединённый».</summary>
        private static string DefaultOutputName(List<PdfPageRef> order)
        {
            var distinct = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PdfPageRef r in order)
                distinct.Add(r.SourcePath);
            return distinct.Count == 1
                ? Path.GetFileNameWithoutExtension(order[0].SourcePath) + ".docx"
                : Loc.T("ocr.defaultMerged");
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
