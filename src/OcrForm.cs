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
        private const string Title = "PDF → Word";

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
                "Извлечение текста и таблиц из документов формата *.pdf с возможностью изменения порядка страниц.",
                Theme.WordViolet, Theme.WordVioletDark, ShowHelp);

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;
            int panelW = 210;
            int gridBottom = ClientSize.Height - 152;

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
            _btnOpen.Text = "Добавить PDF…";
            _btnOpen.SetBounds(px, m + 84, pw, 32);
            _btnOpen.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _btnOpen.Click += OnOpenClick;
            _tips.SetToolTip(_btnOpen, "Можно выбрать несколько файлов или перетащить их в окно");
            Controls.Add(_btnOpen);

            _btnUp = AddPanelButton("◀ Раньше", px, m + 128, pw, "Переместить страницу раньше (Alt+←)");
            _btnUp.Click += delegate { MoveSelected(false); };
            _btnDown = AddPanelButton("Позже ▶", px, m + 164, pw, "Переместить страницу позже (Alt+→)");
            _btnDown.Click += delegate { MoveSelected(true); };
            _btnRemove = AddPanelButton("Удалить", px, m + 208, pw, "Убрать выбранные страницы из вывода (Delete)");
            _btnRemove.Click += OnRemoveClick;

            BuildBottomStrip(right, "Добавьте цифровые PDF — кнопкой или перетащив их в окно.", 230, false);

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
                "1. Добавьте один или несколько PDF — кнопкой «Добавить PDF…» (можно выбрать сразу " +
                "несколько) или перетащив их в окно. Страницы всех файлов показываются одной сеткой.\n" +
                "2. При необходимости измените порядок страниц: перетащите миниатюру или выделите " +
                "её и нажмите «◀ Раньше»/«Позже ▶» (Alt+←/→). Лишние страницы уберите из вывода " +
                "кнопкой «Удалить» (Delete). В Word попадут страницы в показанном порядке.\n" +
                "3. Нажмите «Конвертировать в Word…» и укажите имя .docx — все выбранные страницы " +
                "соберутся в один документ.\n\n" +
                "Извлекается ТЕКСТОВЫЙ СЛОЙ цифровых PDF (например, сохранённых из Word, " +
                "«Microsoft Print to PDF», экспортированных из браузера). Переносятся: текст " +
                "абзацами в порядке чтения — с шрифтом, размером, начертанием, цветом, " +
                "подчёркиванием, выравниванием и красной строкой; таблицы с линиями (границами) " +
                "восстанавливаются ячейками, включая объединённые; книжная и альбомная " +
                "ориентация страниц сохраняется постранично; изображения и гиперссылки.\n\n" +
                "Текущие ограничения перевода в Word:\n" +
                "• Отсканированные документы (страницы-изображения без текстового слоя) не " +
                "поддерживаются — появится сообщение, файл не пострадает.\n" +
                "• Если шрифт из PDF не установлен в системе, текст оформляется шрифтом " +
                "Times New Roman — начертание может немного отличаться от оригинала.\n" +
                "• Таблицы БЕЗ линий (границ), врезки, несколько колонок и списки переносятся " +
                "простыми абзацами в одну колонку — их, возможно, придётся поправить вручную.\n" +
                "• Декоративные печати электронной подписи не воспроизводятся как графика; " +
                "их текст (сведения о сертификате) извлекается обычным текстом.\n" +
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
                dialog.Multiselect = true;
                dialog.Title = "Выберите PDF-файлы";
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

        /// <summary>Добавить PDF-файлы в конец списка страниц (порядок правится в сетке). Ошибка файла — диалог, остальные добавляются.</summary>
        private void AddFiles(string[] paths)
        {
            if (paths == null)
                return;
            int added = 0;
            Cursor = Cursors.WaitCursor;
            try
            {
                foreach (string path in paths)
                {
                    try
                    {
                        int pages = PdfMergeService.LoadPages(path).Count;
                        _order.AddDocument(path, pages); // страницы в исходном порядке; пользователь переставит/удалит
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
                SetStatus("Страниц к переводу: " + _order.Count + ".", Theme.TextMuted);
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
                dialog.Filter = "Документ Word (*.docx)|*.docx";
                dialog.FileName = DefaultOutputName(order);
                dialog.InitialDirectory = Path.GetDirectoryName(order[0].SourcePath);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                outPath = dialog.FileName;
            }

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
                : "Объединённый.docx";
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
