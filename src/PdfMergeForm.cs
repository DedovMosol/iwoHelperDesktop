using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Инструмент «Объединение PDF»: сетка миниатюр (<see cref="PdfPageGrid"/>)
    /// страниц выбранных документов, масштаб, перестановка кнопками и
    /// перетаскиванием, удаление, сохранение в один PDF. Страницы копируются без
    /// переконвертации (PDFsharp).
    /// </summary>
    public class PdfMergeForm : PdfToolFormBase
    {
        private static string Title { get { return Loc.T("hub.pdf.name"); } }

        private readonly PdfPageOrder _order = new PdfPageOrder();

        // Сетка, зум, сжатие, статус, подсказки и флаг _busy — в базе PdfToolFormBase.
        private Button _btnAdd;
        private Button _btnUp;
        private Button _btnDown;
        private Button _btnRemove;
        private Button _btnSave;

        public PdfMergeForm() : this(null) { }

        public PdfMergeForm(Action showHub) : base(showHub)
        {
            BuildUi();
            UpdateButtons();
        }

        /// <summary>Во время сохранения окно не закрывается — иначе остался бы зомби-процесс.</summary>
        protected override string BusyMessage
        {
            get { return Loc.T("common.busySaving"); }
        }

        private void BuildUi()
        {
            InitShell(Title, new Size(780, 660), new Size(660, 540), Theme.PdfRed);
            DragEnter += OnFileDragEnter;
            DragDrop += OnFileDragDrop;
            BuildHeaderWithHome(Title,
                Loc.T("pdf.header.subtitle"),
                Theme.PdfRed, Theme.PdfRedDark, ShowHelp);

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;

            _grid = new PdfPageGrid();
            _grid.AllowReorder = true;
            _grid.AllowRotate = true;          // поворот пишется в итоговый PDF (/Rotate)
            _grid.ShowPositionNumbers = true;  // под плиткой — позиция в итоговом наборе
            _grid.SetBounds(20, m + 80, right - 20 - 150, ClientSize.Height - (m + 80) - 152);
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _grid.SelectionChanged += delegate { UpdateButtons(); };
            _grid.ReorderRequested += OnReorder;
            _grid.MoveRangeRequested += OnMoveRange;
            _grid.InsertPagesRequested += OnInsertPages;
            _grid.FilesDropped += delegate(string[] paths, int insertAt) { AddFiles(paths, insertAt); };
            WireGridMenu();
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
            _btnRemove.Click += OnRemoveClick;
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

        // ---------- добавление файлов ----------

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

        private void OnFileDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = !_busy && PdfDrop.ExtractPaths(e).Length > 0
                ? DragDropEffects.Copy
                : DragDropEffects.None;
        }

        private void OnFileDragDrop(object sender, DragEventArgs e)
        {
            if (!_busy)
                AddFiles(PdfDrop.ExtractPaths(e));
        }

        /// <summary>Добавить файлы в конец (кнопка, дроп на окно) или ПЕРЕД позицией insertAt (дроп на сетку).</summary>
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
                SetStatus(string.Format(Loc.T("common.status.pageCountList"), _order.Count), Theme.TextMuted);
            }
            UpdateButtons();
        }

        private void RefreshGrid()
        {
            _grid.SetPages(_order.ToList());
        }

        // ---------- перестановка и удаление ----------

        private void OnReorder(int from, int to)
        {
            _order.Move(from, to);
            RefreshGrid();
            int landed = to > from ? to - 1 : to;
            _grid.SelectIndex(landed);
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
            SetStatus(string.Format(Loc.T("common.status.pageCountList"), _order.Count), Theme.TextMuted);
            UpdateButtons();
        }

        private void MoveSelected(bool later)
        {
            if (_busy || _grid.SelectedCount != 1)
                return;
            int index = _grid.GetSelectedIndices()[0];
            int moved = later ? _order.MoveDown(index) : _order.MoveUp(index);
            if (moved == index)
                return;
            RefreshGrid();
            _grid.SelectIndex(moved);
        }

        private void OnRemoveClick(object sender, EventArgs e)
        {
            if (_busy || _grid.SelectedCount == 0)
                return;
            _order.RemoveAt(_grid.GetSelectedIndices());
            RefreshGrid();
            SetStatus(string.Format(Loc.T("common.status.pageCountList"), _order.Count), Theme.TextMuted);
            UpdateButtons();
        }

        // Горячие клавиши сетки (Delete, Alt+←/→, Ctrl+A, Enter) — в базе PdfToolFormBase.
        protected override void RemoveSelectedPages() { OnRemoveClick(this, EventArgs.Empty); }
        protected override void MoveSelectedPage(bool later) { MoveSelected(later); }

        // ---------- сохранение ----------

        private void OnSaveClick(object sender, EventArgs e)
        {
            if (_busy || _order.Count == 0)
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
            _busy = true;
            UpdateButtons();
            SetStatus(Loc.T("common.status.saving"), Theme.TextMuted);
            BeginProgress();
            Action<int, int> onProgress = UiProgress();

            var thread = new Thread(delegate()
            {
                Exception error = null;
                bool compressed = false;
                try
                {
                    PdfMergeService.Merge(pages, outputPath, onProgress);
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                // Сжатие — на этом же воркере и ДО открытия файла (иначе замену
                // заблокирует вьюер). Ошибки сжатия не срывают сохранение.
                if (error == null)
                    compressed = PdfCompression.Compress(outputPath, level);
                try
                {
                    if (IsHandleCreated && !IsDisposed)
                        BeginInvoke((MethodInvoker)delegate { OnSaveFinished(error, outputPath, pages.Count, compressed); });
                }
                catch (InvalidOperationException) { }
            });
            thread.IsBackground = true;
            thread.Start();
        }

        private void OnSaveFinished(Exception error, string outputPath, int pageCount, bool compressed)
        {
            _busy = false;
            EndProgress();
            UpdateButtons();
            if (error != null)
            {
                SetStatus(Loc.T("pdf.status.saveFailed"), Theme.ErrRed);
                Dialogs.Error(this, Title, Loc.T("pdf.err.saveFailed"), error.Message);
                return;
            }
            UsageStats.RecordPdfMerge();
            if (compressed)
                UsageStats.RecordPdfCompress();
            SetStatus(string.Format(Loc.T(compressed ? "pdf.status.savedCompressed" : "pdf.status.saved"), pageCount), Theme.OkGreen);
            try { Process.Start(outputPath); }
            catch { } // нет ассоциации PDF — файл всё равно сохранён
        }

        private void UpdateButtons()
        {
            bool one = !_busy && _grid.SelectedCount == 1;
            _grid.Locked = _busy; // правки сетки (буфер, поворот, дроп) — только вне операции
            _compress.Enabled = !_busy;
            _btnAdd.Enabled = !_busy;
            _btnUp.Enabled = one;
            _btnDown.Enabled = one;
            _btnRemove.Enabled = !_busy && _grid.SelectedCount > 0;
            _btnSave.Enabled = !_busy && _order.Count > 0;
        }

    }
}
