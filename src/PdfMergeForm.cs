using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Инструмент «Объединение PDF»: список страниц выбранных документов,
    /// перестановка кнопками и перетаскиванием, удаление, сохранение в один PDF.
    /// Страницы копируются без переконвертации (PDFsharp).
    /// </summary>
    public class PdfMergeForm : Form
    {
        private const string Title = "Объединение PDF";

        private readonly PdfPageOrder _order = new PdfPageOrder();
        private ListView _list;
        private Button _btnAdd;
        private Button _btnUp;
        private Button _btnDown;
        private Button _btnRemove;
        private Button _btnSave;
        private Label _lblStatus;
        private ToolTip _tips;
        private bool _busy; // идёт сохранение (только UI-поток)

        public PdfMergeForm()
        {
            BuildUi();
            UpdateButtons();
        }

        private void BuildUi()
        {
            Text = Title;
            Font = new Font("Segoe UI", 9.75f);
            BackColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(720, 560);
            MinimumSize = new Size(620, 460);
            ShowInTaskbar = false;
            AllowDrop = true;
            DragEnter += OnFileDragEnter;
            DragDrop += OnFileDragDrop;
            _tips = new ToolTip();

            Ui.AccentBar(this, 0);
            Ui.Label(this, "Объединение PDF", 20, 16,
                new Font("Segoe UI", 14f, FontStyle.Bold), Color.FromArgb(40, 40, 40));
            Ui.Label(this, "Страницы копируются без изменений; порядок в итоговом файле — как в списке",
                22, 46, Font, Theme.TextMuted);

            int right = ClientSize.Width - 20;

            _list = new ListView();
            _list.SetBounds(20, 80, right - 20 - 150, ClientSize.Height - 80 - 88);
            _list.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _list.View = View.Details;
            _list.FullRowSelect = true;
            _list.HideSelection = false;
            _list.HeaderStyle = ColumnHeaderStyle.Nonclickable; // порядок задаёт пользователь, не сортировка
            _list.BorderStyle = BorderStyle.FixedSingle;
            _list.Columns.Add("№", 44);
            _list.Columns.Add("Файл", 340);
            _list.Columns.Add("Страница", 90);
            _list.SelectedIndexChanged += delegate { UpdateButtons(); };
            // Перетаскивание строк с системной меткой вставки.
            _list.AllowDrop = true;
            _list.ItemDrag += OnItemDrag;
            _list.DragOver += OnListDragOver;
            _list.DragDrop += OnListDragDrop;
            _list.DragLeave += delegate { _list.InsertionMark.Index = -1; };
            EnableDoubleBuffer(_list);
            Controls.Add(_list);

            int col = right - 130;
            _btnAdd = AddButton("Добавить PDF…", col, 80, 130, 32);
            _btnAdd.Click += OnAddClick;
            _tips.SetToolTip(_btnAdd, "Файлы также можно перетащить в окно");
            _btnUp = AddButton("▲ Вверх", col, 124, 130, 30);
            _btnUp.Click += delegate { MoveSelected(true); };
            _btnDown = AddButton("▼ Вниз", col, 160, 130, 30);
            _btnDown.Click += delegate { MoveSelected(false); };
            _btnRemove = AddButton("Удалить", col, 204, 130, 30);
            _btnRemove.Click += OnRemoveClick;

            var save = new RoundedButton(true);
            save.Text = "Сохранить PDF…";
            save.SetBounds(right - 190, ClientSize.Height - 58, 190, 38);
            save.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            save.Click += OnSaveClick;
            Controls.Add(save);
            _btnSave = save;
            AcceptButton = save;

            _lblStatus = Ui.Label(this, "Добавьте PDF-файлы — кнопкой или перетащив их в окно.",
                20, ClientSize.Height - 50, Font, Theme.TextMuted);
            _lblStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
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

        private static void EnableDoubleBuffer(ListView list)
        {
            var p = typeof(ListView).GetProperty("DoubleBuffered",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            if (p != null)
                p.SetValue(list, true, null);
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
            e.Effect = !_busy && ExtractPdfPaths(e).Length > 0
                ? DragDropEffects.Copy
                : DragDropEffects.None;
        }

        private void OnFileDragDrop(object sender, DragEventArgs e)
        {
            if (!_busy)
                AddFiles(ExtractPdfPaths(e));
        }

        private static string[] ExtractPdfPaths(DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return new string[0];
            var items = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (items == null)
                return new string[0];
            var pdfs = new System.Collections.Generic.List<string>();
            foreach (string item in items)
            {
                if (File.Exists(item) &&
                    string.Equals(Path.GetExtension(item), ".pdf", StringComparison.OrdinalIgnoreCase))
                    pdfs.Add(item);
            }
            return pdfs.ToArray();
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
                RefreshList();
                SetStatus("Страниц в списке: " + _order.Count + ".", Theme.TextMuted);
            }
            UpdateButtons();
        }

        // ---------- перестановка и удаление ----------

        private void MoveSelected(bool up)
        {
            if (_list.SelectedIndices.Count != 1)
                return;
            int index = _list.SelectedIndices[0];
            int moved = up ? _order.MoveUp(index) : _order.MoveDown(index);
            if (moved == index)
                return;
            RefreshList();
            _list.Items[moved].Selected = true;
            _list.Items[moved].EnsureVisible();
            _list.Focus();
        }

        private void OnRemoveClick(object sender, EventArgs e)
        {
            if (_list.SelectedIndices.Count == 0)
                return;
            var indices = new int[_list.SelectedIndices.Count];
            _list.SelectedIndices.CopyTo(indices, 0);
            _order.RemoveAt(indices);
            RefreshList();
            SetStatus("Страниц в списке: " + _order.Count + ".", Theme.TextMuted);
            UpdateButtons();
        }

        private void OnItemDrag(object sender, ItemDragEventArgs e)
        {
            if (!_busy)
                _list.DoDragDrop(e.Item, DragDropEffects.Move);
        }

        private void OnListDragOver(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                e.Effect = DragDropEffects.None;
                return;
            }
            e.Effect = DragDropEffects.Move;
            Point pt = _list.PointToClient(new Point(e.X, e.Y));
            int index = _list.InsertionMark.NearestIndex(pt);
            if (index >= 0)
            {
                Rectangle bounds = _list.GetItemRect(index);
                _list.InsertionMark.AppearsAfterItem = pt.Y > bounds.Top + bounds.Height / 2;
            }
            _list.InsertionMark.Index = index;
        }

        private void OnListDragDrop(object sender, DragEventArgs e)
        {
            int target = _list.InsertionMark.Index;
            bool after = _list.InsertionMark.AppearsAfterItem;
            _list.InsertionMark.Index = -1;
            var item = e.Data.GetData(typeof(ListViewItem)) as ListViewItem;
            if (item == null || target < 0)
                return;
            int from = item.Index;
            int to = after ? target + 1 : target;
            _order.Move(from, to);
            RefreshList();
            int landed = to > from ? to - 1 : to;
            if (landed >= 0 && landed < _list.Items.Count)
            {
                _list.Items[landed].Selected = true;
                _list.Items[landed].EnsureVisible();
            }
        }

        private void RefreshList()
        {
            _list.BeginUpdate();
            _list.Items.Clear();
            for (int i = 0; i < _order.Count; i++)
            {
                PdfPageRef page = _order[i];
                var item = new ListViewItem((i + 1).ToString());
                item.SubItems.Add(page.FileName);
                item.SubItems.Add("стр. " + (page.PageIndex + 1));
                item.ToolTipText = page.SourcePath;
                _list.Items.Add(item);
            }
            _list.EndUpdate();
        }

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
            _busy = true;
            UpdateButtons();
            SetStatus("Сохранение…", Theme.TextMuted);

            var thread = new Thread(delegate()
            {
                Exception error = null;
                try
                {
                    PdfMergeService.Merge(pages, outputPath);
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                try
                {
                    if (IsHandleCreated && !IsDisposed)
                        BeginInvoke((MethodInvoker)delegate { OnSaveFinished(error, outputPath, pages.Count); });
                }
                catch (InvalidOperationException) { }
            });
            thread.IsBackground = true;
            thread.Start();
        }

        private void OnSaveFinished(Exception error, string outputPath, int pageCount)
        {
            _busy = false;
            UpdateButtons();
            if (error != null)
            {
                SetStatus("PDF не сохранён.", Theme.ErrRed);
                Dialogs.Error(this, Title, "PDF не сохранён", error.Message);
                return;
            }
            SetStatus("✓ Сохранено страниц: " + pageCount + ".", Theme.OkGreen);
            try { Process.Start(outputPath); }
            catch { } // нет ассоциации PDF — файл всё равно сохранён
        }

        private void UpdateButtons()
        {
            bool hasSelection = _list.SelectedIndices.Count > 0;
            _btnAdd.Enabled = !_busy;
            _btnUp.Enabled = !_busy && _list.SelectedIndices.Count == 1;
            _btnDown.Enabled = !_busy && _list.SelectedIndices.Count == 1;
            _btnRemove.Enabled = !_busy && hasSelection;
            _btnSave.Enabled = !_busy && _order.Count > 0;
        }

        private void SetStatus(string text, Color color)
        {
            _lblStatus.Text = text;
            _lblStatus.ForeColor = color;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_busy)
            {
                SetStatus("Дождитесь завершения сохранения…", Theme.WarnOrange);
                e.Cancel = true;
                return;
            }
            base.OnFormClosing(e);
        }
    }
}
