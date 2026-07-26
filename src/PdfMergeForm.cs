using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Инструмент «Объединение PDF»: сетка миниатюр (<see cref="PdfPageGrid"/>)
    /// страниц выбранных документов, масштаб, перестановка кнопками и
    /// перетаскиванием, удаление, сохранение в один PDF. Страницы копируются без
    /// переконвертации (PDFsharp). Модель порядка и её слой (добавление/перестановка/
    /// буфер/Ctrl+Z) — в общей базе <see cref="PdfOrderedToolFormBase"/> (DRY).
    /// </summary>
    public class PdfMergeForm : PdfOrderedToolFormBase
    {
        private static string Title { get { return Loc.T("hub.pdf.name"); } }

        // Сетка, зум, сжатие, статус, подсказки, флаг _busy и модель порядка — в базах.
        private Button _btnAdd;
        private Button _btnUp;
        private Button _btnDown;
        private Button _btnRemove;
        private Button _btnSave;

        public PdfMergeForm() : this(null) { }

        public PdfMergeForm(Action showHub) : base(showHub)
        {
            BuildUi();
            SyncControls();
        }

        protected override string ToolTitle { get { return Title; } }

        /// <summary>Во время сохранения окно не закрывается — иначе остался бы зомби-процесс.</summary>
        protected override string BusyMessage
        {
            get { return Loc.T("common.busySaving"); }
        }

        /// <summary>Idle-статус: пусто — подсказка добавления, иначе — число страниц.</summary>
        protected override string IdleStatusText()
        {
            return _order.Count == 0
                ? Loc.T("pdf.status.addPdf")
                : string.Format(Loc.T("common.status.pageCountList"), _order.Count);
        }

        private void BuildUi()
        {
            InitShell(Title, new Size(780, 660), new Size(660, 540), Theme.PdfRed);
            WireFileDropAppend(); // дроп PDF на окно — добавить в конец (общая обвязка базы)
            BuildHeaderWithHome(Title,
                Loc.T("pdf.header.subtitle"),
                Theme.PdfRed, Theme.PdfRedDark, ShowHelp);

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;

            _grid = new PdfPageGrid();
            _grid.AllowReorder = true;
            _grid.AllowRotate = true;          // поворот пишется в итоговый PDF (/Rotate)
            _grid.ShowPositionNumbers = true;  // под плиткой — позиция в итоговом наборе
            _grid.EmptyHint = Loc.T("pdf.grid.empty");
            _grid.DropHint = Loc.T("grid.dropHint");
            _grid.SetBounds(20, m + 80, right - 20 - 150, ClientSize.Height - (m + 80) - 152);
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            WireOrderGrid(); // события порядка + контекстное меню (общая обвязка базы)
            Controls.Add(_grid);

            int col = right - 130;
            _btnAdd = AddButton(Loc.T("common.addPdf"), col, m + 80, 130, 32);
            _btnAdd.Click += OnAddClick;
            _tips.SetToolTip(_btnAdd, Loc.T("common.tip.addPdf"));
            _btnUp = AddButton(Loc.T("common.earlier"), col, m + 124, 130, 30);
            _btnUp.Click += delegate { MoveSelected(false); };
            _tips.SetToolTip(_btnUp, Loc.T("common.tip.earlier"));
            _btnDown = AddButton(Loc.T("common.later"), col, m + 160, 130, 30);
            _btnDown.Click += delegate { MoveSelected(true); };
            _tips.SetToolTip(_btnDown, Loc.T("common.tip.later"));
            _btnRemove = AddButton(Loc.T("common.remove"), col, m + 204, 130, 30);
            _btnRemove.Click += delegate { RemoveSelected(); };
            _tips.SetToolTip(_btnRemove, Loc.T("common.tip.removePages"));

            BuildBottomStrip(right, Loc.T("pdf.status.addPdf"), 190);

            var save = new RoundedButton(true);
            save.Text = Loc.T("pdf.btn.save");
            save.SetBounds(right - 190, ClientSize.Height - 58, 190, 38);
            save.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            save.Click += OnSaveClick;
            Controls.Add(save);
            _btnSave = save;
            AcceptButton = save;
            RegisterActionButton(save); // база подменит её кнопкой «Отмена» во время сохранения
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, Title, Loc.T("menu.howTo"), Loc.T("pdf.help.body"));
        }

        private Button AddButton(string text, int x, int y, int w, int h)
        {
            var b = new RoundedButton(false);
            b.Text = text;
            b.SetBounds(x, y, w, h);
            b.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(b);
            return b;
        }

        private void OnAddClick(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = Loc.T("common.pdfFilter");
                dialog.Multiselect = true;
                dialog.Title = Loc.T("common.pickPdf");
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    AddFiles(dialog.FileNames);
            }
        }

        // ---------- сохранение ----------

        private void OnSaveClick(object sender, EventArgs e)
        {
            if (Working || _order.Count == 0)
                return;
            string outputPath;
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = Loc.T("common.pdfSaveFilter");
                dialog.FileName = Loc.T("pdf.defaultName");
                dialog.InitialDirectory = Path.GetDirectoryName(_order[0].SourcePath);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                outputPath = dialog.FileName;
            }

            var pages = _order.ToList();
            CompressionLevel level = _compress.Level; // читаем с UI-потока до старта воркера
            BeginOperation(Loc.T("common.status.saving"), pages.Count, Loc.T("pdf.status.savingPage"));
            Action<int, int> onProgress = UiProgress();
            Func<bool> cancel = CancelToken();

            Ui.RunWorker(delegate()
            {
                Exception error = null;
                bool compressed = false;
                try
                {
                    PdfMergeService.Merge(pages, outputPath, onProgress, cancel);
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                // Сжатие — на этом же воркере и ДО открытия файла (иначе замену
                // заблокирует вьюер). Ошибки сжатия не срывают сохранение.
                if (error == null)
                {
                    // Файл записан — точка невозврата: сжатие (Ghostscript) не прерываем, поэтому
                    // снимаем предложение отмены, чтобы кнопка не «зависала» на «Отмена…».
                    OnUi(delegate { StopOfferingCancel(); });
                    compressed = PdfCompression.Compress(outputPath, level);
                }
                bool didCompress = compressed;
                OnUi(delegate { OnSaveFinished(error, outputPath, pages.Count, didCompress, level); });
            });
        }

        private void OnSaveFinished(Exception error, string outputPath, int pageCount, bool compressed, CompressionLevel level)
        {
            EndOperation();
            if (error is OperationCanceledException)
            {
                SetStatus(Loc.T("common.status.canceled"), Theme.WarnOrange); // файл не создан
                return;
            }
            if (error != null)
            {
                SetStatus(Loc.T("pdf.status.saveFailed"), Theme.ErrRed);
                Dialogs.Error(this, Title, Loc.T("pdf.err.saveFailed"), error.Message);
                return;
            }
            UsageStats.RecordPdfMerge();
            if (compressed)
                UsageStats.RecordPdfCompress();
            SetStatus(SuccessStatus(
                string.Format(Loc.T("pdf.status.pagesSaved"), pageCount),
                CompressedPart(compressed, level)), Theme.OkGreen);
            Ui.OpenPath(outputPath); // авто-открытие результата; молча, если нет ассоциации PDF
        }

        /// <summary>Доступность кнопок и блокировка сетки по текущему состоянию (операция/загрузка/выделение).</summary>
        protected override void SyncControls()
        {
            bool one = !Working && _grid.SelectedCount == 1;
            _grid.Locked = Working; // правки сетки (буфер, поворот, дроп) — только вне работы
            _compress.Enabled = !Working;
            _btnAdd.Enabled = !Working;
            _btnUp.Enabled = one;
            _btnDown.Enabled = one;
            _btnRemove.Enabled = !Working && _grid.SelectedCount > 0;
            _btnSave.Enabled = !Working && _order.Count > 0;
        }
    }
}
