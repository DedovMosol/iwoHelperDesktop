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
    /// Инструмент «PDF → Word»: извлекает текстовый слой цифрового PDF (сохранённого из
    /// Word, «Microsoft Print to PDF» и т.п.) в редактируемый .docx. Страницы источника
    /// показаны сеткой (<see cref="PdfPageGrid"/>). Отсканированные документы (без текстового
    /// слоя) в настоящее время недоступны — при попытке будет понятное сообщение (файл цел).
    /// Конвертация обёрнута try/catch (<see cref="OnConvertClick"/>): любая ошибка — диалог,
    /// не краш. На базе <see cref="PdfToolFormBase"/> (сетка/зум/статус/закрытие/освобождение — DRY).
    /// </summary>
    public class OcrForm : PdfToolFormBase
    {
        private const string Title = "PDF → Word";

        private readonly PdfPageOrder _order = new PdfPageOrder();
        private string _sourcePath;
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
            InitShell(Title, new Size(800, 620), new Size(700, 520), Theme.WordViolet);
            DragEnter += OnFileDragEnter;
            DragDrop += OnFileDragDrop;
            BuildHeaderWithHome(Title,
                "Извлеките текст цифрового PDF в редактируемый .docx. Порядок страниц можно менять перетаскиванием.",
                Theme.WordViolet, Theme.WordVioletDark, ShowHelp);

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;
            int panelW = 210;
            int gridBottom = ClientSize.Height - 112;

            _grid = new PdfPageGrid();
            _grid.AllowReorder = true; // перестановка страниц перетаскиванием
            _grid.SetBounds(20, m + 84, right - 20 - panelW, gridBottom - (m + 84));
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _grid.ReorderRequested += OnReorder;
            _grid.SelectionChanged += delegate { UpdateControls(); };
            Controls.Add(_grid);

            int px = right - panelW + 10;
            int pw = panelW - 10;
            _btnOpen = new RoundedButton(false);
            _btnOpen.Text = "Открыть PDF…";
            _btnOpen.SetBounds(px, m + 84, pw, 32);
            _btnOpen.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnOpen.Click += OnOpenClick;
            _tips.SetToolTip(_btnOpen, "Файл также можно перетащить в окно");
            Controls.Add(_btnOpen);

            _btnUp = AddPanelButton("◀ Раньше", px, m + 128, pw, "Переместить страницу раньше (Alt+←)");
            _btnUp.Click += delegate { MoveSelected(false); };
            _btnDown = AddPanelButton("Позже ▶", px, m + 164, pw, "Переместить страницу позже (Alt+→)");
            _btnDown.Click += delegate { MoveSelected(true); };
            _btnRemove = AddPanelButton("Удалить", px, m + 208, pw, "Убрать выбранные страницы из вывода (Delete)");
            _btnRemove.Click += OnRemoveClick;

            BuildBottomStrip(right, "Откройте цифровой PDF — кнопкой или перетащив его в окно.", false);

            // Действие — в правом нижнем углу (как «Сохранить PDF» в «Объединении»).
            _btnConvert = new RoundedButton(true);
            _btnConvert.Text = "Конвертировать в Word…";
            _btnConvert.SetBounds(right - 230, ClientSize.Height - 58, 230, 38);
            _btnConvert.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _btnConvert.Click += OnConvertClick;
            _tips.SetToolTip(_btnConvert, "Извлечь текст в редактируемый .docx");
            Controls.Add(_btnConvert);
            AcceptButton = _btnConvert;
        }

        private void ShowHelp()
        {
            Dialogs.Info(this, Title, "Как пользоваться",
                "1. Откройте PDF — кнопкой «Открыть PDF…» или перетащив его в окно.\n" +
                "2. При необходимости измените порядок страниц: перетащите миниатюру или выделите " +
                "её и нажмите «◀ Раньше»/«Позже ▶» (Alt+←/→). Лишние страницы уберите из вывода " +
                "кнопкой «Удалить» (Delete). В Word попадут страницы в показанном порядке.\n" +
                "3. Нажмите «Конвертировать в Word…» и укажите имя .docx.\n\n" +
                "Извлекается ТЕКСТОВЫЙ СЛОЙ цифровых PDF (например, сохранённых из Word, " +
                "«Microsoft Print to PDF», экспортированных из браузера). Текст переносится " +
                "абзацами в порядке чтения, с сохранением шрифта, размера, начертания, цвета, " +
                "выравнивания и красной строки.\n\n" +
                "Текущие ограничения перевода в Word:\n" +
                "• Отсканированные документы (страницы-изображения без текстового слоя) не " +
                "поддерживаются — появится сообщение, файл не пострадает.\n" +
                "• Если шрифт из PDF не установлен в системе, текст оформляется шрифтом " +
                "Times New Roman — начертание может немного отличаться от оригинала.\n" +
                "• Таблицы, врезки, несколько колонок и списки переносятся простыми абзацами " +
                "в одну колонку — их, возможно, придётся поправить вручную.\n" +
                "• Подчёркивание не переносится: в PDF это нарисованная линия, а не свойство " +
                "текста, поэтому его нельзя прочитать из текстового слоя.\n" +
                "• Если PDF сохранён с испорченной кодировкой текста (без корректного ToUnicode), " +
                "извлечённый текст будет нечитаемым — это дефект самого файла, а не конвертации; " +
                "проверить можно, скопировав текст в самом PDF (Ctrl+C).");
        }

        // ---------- открытие ----------

        private void OnOpenClick(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Документы PDF (*.pdf)|*.pdf";
                dialog.Title = "Выберите PDF";
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    LoadSource(dialog.FileName);
            }
        }

        private void OnFileDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = !_busy && PdfDrop.ExtractPaths(e).Length > 0 ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void OnFileDragDrop(object sender, DragEventArgs e)
        {
            if (_busy)
                return;
            string[] paths = PdfDrop.ExtractPaths(e);
            if (paths.Length > 0)
                LoadSource(paths[0]); // конвертация работает с одним документом
        }

        private void LoadSource(string path)
        {
            int pageCount;
            Cursor = Cursors.WaitCursor;
            try
            {
                pageCount = PdfMergeService.LoadPages(path).Count;
            }
            catch (MergeException ex)
            {
                Cursor = Cursors.Default;
                Dialogs.Error(this, Title, "Файл не открыт", ex.Message);
                return;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            _sourcePath = path;
            _order.Clear();
            _order.AddDocument(path, pageCount); // страницы в исходном порядке; пользователь переставит/удалит
            RefreshGrid();
            SetStatus("Открыт «" + Path.GetFileName(path) + "»: страниц " + pageCount + ".", Theme.TextMuted);
            UpdateControls();
        }

        // ---------- конвертация ----------

        private void OnConvertClick(object sender, EventArgs e)
        {
            if (_busy || _sourcePath == null || _order.Count == 0)
                return;
            string src = _sourcePath;
            string outPath;
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Документ Word (*.docx)|*.docx";
                dialog.FileName = Path.GetFileNameWithoutExtension(src) + ".docx";
                dialog.InitialDirectory = Path.GetDirectoryName(src);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                outPath = dialog.FileName;
            }

            // Порядок/подмножество страниц из сетки (индексы с нуля в текущем порядке).
            var order = new List<int>(_order.Count);
            foreach (PdfPageRef p in _order.ToList())
                order.Add(p.PageIndex);

            _busy = true;
            UpdateControls();
            SetStatus("Конвертация в Word…", Theme.TextMuted);
            BeginProgress();
            Action<int, int> onProgress = UiProgress();

            var thread = new Thread(delegate()
            {
                Exception error = null;
                ConvertResult result = null;
                try
                {
                    result = PdfToWordService.Convert(src, outPath, order, onProgress);
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
                SetStatus("Не выполнено.", Theme.ErrRed);
                Dialogs.Error(this, Title, "Конвертация не выполнена", error.Message);
                return;
            }
            UsageStats.RecordPdfToWord();
            SetStatus("✓ Готово: страниц " + result.Pages + " → Word (.docx).", Theme.OkGreen);
            try { Process.Start(outPath); }
            catch { } // нет ассоциации .docx — файл всё равно создан
        }

        private void UpdateControls()
        {
            bool one = !_busy && _grid.SelectedCount == 1;
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
            SetStatus("Страниц к переводу: " + _order.Count + ".", Theme.TextMuted);
            UpdateControls();
        }

        /// <summary>Горячие клавиши сетки страниц (та же раскладка, что в «Объединении»).</summary>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (_grid != null && _grid.ListFocused)
            {
                switch (PdfMergeForm.ClassifyPageKey(keyData))
                {
                    case PdfMergeForm.PageKeyAction.Remove: OnRemoveClick(this, EventArgs.Empty); return true;
                    case PdfMergeForm.PageKeyAction.MoveEarlier: MoveSelected(false); return true;
                    case PdfMergeForm.PageKeyAction.MoveLater: MoveSelected(true); return true;
                    case PdfMergeForm.PageKeyAction.SelectAll: _grid.SelectAll(); return true;
                    case PdfMergeForm.PageKeyAction.Swallow: return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
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
