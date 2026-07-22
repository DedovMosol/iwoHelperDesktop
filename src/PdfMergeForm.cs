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
        private const string Title = "Объединение PDF";

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
            get { return "Дождитесь завершения сохранения…"; }
        }

        private void BuildUi()
        {
            InitShell(Title, new Size(780, 660), new Size(660, 540), Theme.PdfRed);
            DragEnter += OnFileDragEnter;
            DragDrop += OnFileDragDrop;
            BuildHeaderWithHome(Title,
                "Объединение документов формата *.pdf с возможностью изменения порядка страниц и сжатием.",
                Theme.PdfRed, Theme.PdfRedDark, ShowHelp);

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;

            _grid = new PdfPageGrid();
            _grid.AllowReorder = true;
            _grid.SetBounds(20, m + 80, right - 20 - 150, ClientSize.Height - (m + 80) - 152);
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _grid.SelectionChanged += delegate { UpdateButtons(); };
            _grid.ReorderRequested += OnReorder;
            Controls.Add(_grid);

            int col = right - 130;
            _btnAdd = AddButton("Добавить PDF…", col, m + 80, 130, 32);
            _btnAdd.Click += OnAddClick;
            _tips.SetToolTip(_btnAdd, "Файлы также можно перетащить в окно");
            _btnUp = AddButton("◀ Раньше", col, m + 124, 130, 30);
            _btnUp.Click += delegate { MoveSelected(false); };
            _tips.SetToolTip(_btnUp, "Переместить страницу раньше (Alt+←)");
            _btnDown = AddButton("Позже ▶", col, m + 160, 130, 30);
            _btnDown.Click += delegate { MoveSelected(true); };
            _tips.SetToolTip(_btnDown, "Переместить страницу позже (Alt+→)");
            _btnRemove = AddButton("Удалить", col, m + 204, 130, 30);
            _btnRemove.Click += OnRemoveClick;
            _tips.SetToolTip(_btnRemove, "Удалить выбранные страницы (Delete)");

            BuildBottomStrip(right, "Добавьте PDF-файлы — кнопкой или перетащив их в окно.", 190);

            var save = new RoundedButton(true);
            save.Text = "Сохранить PDF…";
            save.SetBounds(right - 190, ClientSize.Height - 58, 190, 38);
            save.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            save.Click += OnSaveClick;
            Controls.Add(save);
            _btnSave = save;
            AcceptButton = save;
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, Title, "Как пользоваться",
                "1. Добавьте PDF-файлы — кнопкой «Добавить PDF…» или перетащив их в окно.\n" +
                "2. Появится сетка миниатюр страниц. Масштаб — ползунком внизу или Ctrl+колесо мыши.\n" +
                "3. Задайте порядок: перетаскивайте миниатюры или используйте «◀ Раньше» / «Позже ▶».\n" +
                "   Лишние страницы удаляйте кнопкой «Удалить».\n" +
                "4. При необходимости выберите «Сжатие» (по умолчанию «Отлично» — без сжатия). " +
                "«Хорошо»/«Нормально» уменьшают размер за счёт понижения разрешения изображений " +
                "(как в Acrobat); текст сохраняется. Требуется Ghostscript.\n" +
                "5. «Сохранить PDF…» соберёт один документ в выбранном порядке.\n\n" +
                "Горячие клавиши: Delete — удалить выбранные, Alt+←/→ — порядок, " +
                "Ctrl+A — выделить всё, Ctrl+колесо — масштаб.\n" +
                "Страницы копируются как есть, без переконвертации — сканы, печати и подписи " +
                "не искажаются. Битые и защищённые паролем файлы пропускаются с причиной.\n" +
                "Сжатие меняет содержимое файла, поэтому у подписанных PDF подпись станет " +
                "недействительной (как и при сжатии в Acrobat) — сжимайте до подписания.");
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
                dialog.Filter = "Документы PDF (*.pdf)|*.pdf";
                dialog.Multiselect = true;
                dialog.Title = "Выберите PDF-файлы";
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

        private void AddFiles(string[] paths)
        {
            int added = 0;
            Cursor = Cursors.WaitCursor;
            try
            {
                foreach (string path in paths)
                {
                    try
                    {
                        int pages = PdfMergeService.LoadPages(path).Count;
                        _order.AddDocument(path, pages);
                        added += pages;
                    }
                    catch (MergeException ex)
                    {
                        Dialogs.Error(this, Title, "Файл не добавлен", ex.Message);
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
                SetStatus("Страниц в списке: " + _order.Count + ".", Theme.TextMuted);
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
            SetStatus("Страниц в списке: " + _order.Count + ".", Theme.TextMuted);
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
                dialog.Filter = "Документ PDF (*.pdf)|*.pdf";
                dialog.FileName = "Объединённый.pdf";
                dialog.InitialDirectory = Path.GetDirectoryName(_order[0].SourcePath);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                outputPath = dialog.FileName;
            }

            var pages = _order.ToList();
            CompressionLevel level = _compress.Level; // читаем с UI-потока до старта воркера
            _busy = true;
            UpdateButtons();
            SetStatus("Сохранение…", Theme.TextMuted);
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
                SetStatus("PDF не сохранён.", Theme.ErrRed);
                Dialogs.Error(this, Title, "PDF не сохранён", error.Message);
                return;
            }
            UsageStats.RecordPdfMerge();
            if (compressed)
                UsageStats.RecordPdfCompress();
            SetStatus("✓ Сохранено страниц: " + pageCount + (compressed ? " · сжато." : "."), Theme.OkGreen);
            try { Process.Start(outputPath); }
            catch { } // нет ассоциации PDF — файл всё равно сохранён
        }

        private void UpdateButtons()
        {
            bool one = !_busy && _grid.SelectedCount == 1;
            _compress.Enabled = !_busy;
            _btnAdd.Enabled = !_busy;
            _btnUp.Enabled = one;
            _btnDown.Enabled = one;
            _btnRemove.Enabled = !_busy && _grid.SelectedCount > 0;
            _btnSave.Enabled = !_busy && _order.Count > 0;
        }

    }
}
