using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
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
    /// не краш. На базе <see cref="PdfToolFormBase"/> (сетка/зум/статус/закрытие/освобождение — DRY).
    /// </summary>
    public class OcrForm : PdfToolFormBase
    {
        private static string Title { get { return Loc.T("hub.ocr.name"); } }

        private readonly PdfPageOrder _order = new PdfPageOrder();
        private Button _btnOpen;
        private Button _btnUp;
        private Button _btnDown;
        private Button _btnRemove;
        private Button _btnConvert;

        public OcrForm() : this(null) { }

        public OcrForm(Action showHub) : base(showHub)
        {
            BuildUi();
            UpdateControls();
        }

        private void BuildUi()
        {
            InitShell(Title, new Size(800, 660), new Size(700, 560), Theme.WordViolet);
            DragEnter += OnFileDragEnter;
            DragDrop += OnFileDragDrop;
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
            // AllowRotate НЕ включаем: текстовый конвейер PDF → Word повороты не потребляет.
            _grid.SetBounds(20, m + 84, right - 20 - panelW, gridBottom - (m + 84));
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _grid.ReorderRequested += OnReorder;
            _grid.MoveRangeRequested += OnMoveRange;
            _grid.InsertPagesRequested += OnInsertPages;
            _grid.FilesDropped += delegate(string[] paths, int insertAt) { AddFiles(paths, insertAt); };
            _grid.SelectionChanged += delegate { UpdateControls(); };
            WireGridMenu();
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
            _btnRemove.Click += OnRemoveClick;

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
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, Title, Loc.T("menu.howTo"), Loc.T("ocr.help.body"));
        }

        // ---------- открытие ----------

        private void OnOpenClick(object sender, EventArgs e)
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

        private void OnFileDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = !_busy && PdfDrop.ExtractPaths(e).Length > 0 ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void OnFileDragDrop(object sender, DragEventArgs e)
        {
            if (!_busy)
                AddFiles(PdfDrop.ExtractPaths(e));
        }

        /// <summary>
        /// Добавить PDF-файлы в конец (кнопка, дроп на окно) или ПЕРЕД позицией insertAt
        /// (дроп на сетку). Ошибка файла — диалог, остальные добавляются.
        /// </summary>
        private void AddFiles(string[] paths, int insertAt = -1)
        {
            if (_busy || paths == null)
                return;
            int added = 0;
            int firstAdded = -1;
            int at = insertAt < 0 || insertAt > _order.Count ? _order.Count : insertAt;
            Cursor = Cursors.WaitCursor;
            try
            {
                foreach (string path in paths)
                {
                    try
                    {
                        int pages = PdfMergeService.LoadPages(path).Count;
                        int landed = _order.InsertDocument(at, path, pages);
                        if (firstAdded < 0)
                            firstAdded = landed;
                        at = landed + pages; // следующий файл — сразу за вставленным
                        added += pages;
                    }
                    catch (MergeException ex)
                    {
                        Dialogs.Error(this, Title, Loc.T("common.fileNotAdded"), ex.Message);
                    }
                }
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            if (added > 0)
            {
                RefreshGrid();
                if (insertAt >= 0)
                    _grid.SelectRange(firstAdded, added); // показать место вставки дропа
                SetStatus(string.Format(Loc.T("ocr.status.pageCount"), _order.Count), Theme.TextMuted);
            }
            UpdateControls();
        }

        // ---------- конвертация ----------

        private void OnConvertClick(object sender, EventArgs e)
        {
            if (_busy || _order.Count == 0)
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

            _busy = true;
            UpdateControls();
            SetStatus(Loc.T("ocr.status.converting"), Theme.TextMuted);
            BeginProgress();
            Action<int, int> onProgress = UiProgress();

            var thread = new Thread(delegate()
            {
                Exception error = null;
                ConvertResult result = null;
                try
                {
                    result = PdfToWordService.Convert(order, outPath, onProgress);
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                try
                {
                    if (IsHandleCreated && !IsDisposed)
                        BeginInvoke((MethodInvoker)delegate { OnConvertFinished(error, result, outPath); });
                }
                catch (InvalidOperationException) { }
            });
            thread.SetApartmentState(ApartmentState.STA); // требование Word COM
            thread.IsBackground = true;
            thread.Start();
        }

        private void OnConvertFinished(Exception error, ConvertResult result, string outPath)
        {
            _busy = false;
            EndProgress();
            UpdateControls();
            if (error != null)
            {
                SetStatus(Loc.T("ocr.status.failed"), Theme.ErrRed);
                Dialogs.Error(this, Title, Loc.T("ocr.err.convertFailed"), error.Message);
                return;
            }
            UsageStats.RecordPdfToWord();
            SetStatus(string.Format(Loc.T("ocr.status.done"), result.Pages), Theme.OkGreen);
            try { Process.Start(outPath); }
            catch { } // нет ассоциации .docx — файл всё равно создан
        }

        private void UpdateControls()
        {
            bool one = !_busy && _grid.SelectedCount == 1;
            _grid.Locked = _busy; // правки сетки (буфер, дроп) — только вне операции
            _btnOpen.Enabled = !_busy;
            _btnUp.Enabled = one;
            _btnDown.Enabled = one;
            _btnRemove.Enabled = !_busy && _grid.SelectedCount > 0;
            _btnConvert.Enabled = !_busy && _order.Count > 0;
        }

        // ---------- перестановка и удаление страниц ----------

        private void RefreshGrid()
        {
            _grid.SetPages(_order.ToList());
        }

        private void OnReorder(int from, int to)
        {
            if (_busy)
                return;
            _order.Move(from, to);
            RefreshGrid();
            _grid.SelectIndex(to > from ? to - 1 : to); // выделить страницу на новом месте
        }

        /// <summary>Вставка вырезанных страниц (Ctrl+X → Ctrl+V) — перенос набора внутри порядка.</summary>
        private void OnMoveRange(int[] indices, int insertAt)
        {
            if (_busy)
                return;
            int landed = _order.MoveRange(indices, insertAt);
            if (landed < 0)
                return;
            RefreshGrid();
            _grid.SelectRange(landed, indices.Length);
        }

        /// <summary>Вставка скопированных страниц (Ctrl+C → Ctrl+V) — новые экземпляры в позиции.</summary>
        private void OnInsertPages(PdfPageRef[] pages, int insertAt)
        {
            if (_busy || pages == null || pages.Length == 0)
                return;
            int landed = _order.InsertAt(insertAt, pages);
            RefreshGrid();
            _grid.SelectRange(landed, pages.Length);
            SetStatus(string.Format(Loc.T("ocr.status.pageCount"), _order.Count), Theme.TextMuted);
            UpdateControls();
        }

        private void MoveSelected(bool later)
        {
            if (_busy || _grid.SelectedCount != 1)
                return;
            int index = _grid.GetSelectedIndices()[0];
            int moved = later ? _order.MoveDown(index) : _order.MoveUp(index);
            if (moved == index)
                return; // уже с краю
            RefreshGrid();
            _grid.SelectIndex(moved);
        }

        private void OnRemoveClick(object sender, EventArgs e)
        {
            if (_busy || _grid.SelectedCount == 0)
                return;
            _order.RemoveAt(_grid.GetSelectedIndices());
            RefreshGrid();
            SetStatus(string.Format(Loc.T("ocr.status.pageCount"), _order.Count), Theme.TextMuted);
            UpdateControls();
        }

        // Горячие клавиши сетки (Delete, Alt+←/→, Ctrl+A, Enter) — в базе PdfToolFormBase.
        protected override void RemoveSelectedPages() { OnRemoveClick(this, EventArgs.Empty); }
        protected override void MoveSelectedPage(bool later) { MoveSelected(later); }

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
