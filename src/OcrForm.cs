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

        private string _sourcePath;
        private int _pageCount;
        private Button _btnOpen;
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
                "Извлеките текст цифрового PDF (сохранённого из Word и т.п.) в редактируемый .docx.",
                Theme.WordViolet, Theme.WordVioletDark, ShowHelp);

            int m = HelpMenu.Height;
            int right = ClientSize.Width - 20;
            int panelW = 210;
            int gridBottom = ClientSize.Height - 112;

            _grid = new PdfPageGrid();
            _grid.AllowReorder = false; // только показ страниц источника
            _grid.SetBounds(20, m + 84, right - 20 - panelW, gridBottom - (m + 84));
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
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
                "2. Нажмите «Конвертировать в Word…» и укажите имя .docx.\n\n" +
                "Извлекается ТЕКСТОВЫЙ СЛОЙ цифровых PDF (например, сохранённых из Word, " +
                "«Microsoft Print to PDF», экспортированных из браузера). Текст переносится " +
                "абзацами в порядке чтения; сложную/многоколоночную вёрстку, возможно, " +
                "придётся поправить вручную.\n\n" +
                "Поддержка отсканированных документов (страниц-изображений без текстового слоя) " +
                "в настоящее время недоступна — при попытке появится понятное сообщение, " +
                "файл не пострадает.");
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
            Cursor = Cursors.WaitCursor;
            try
            {
                _pageCount = PdfMergeService.LoadPages(path).Count;
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
            var pages = new List<PdfPageRef>();
            for (int i = 0; i < _pageCount; i++)
                pages.Add(new PdfPageRef { SourcePath = path, PageIndex = i });
            _grid.SetPages(pages);
            SetStatus("Открыт «" + Path.GetFileName(path) + "»: страниц " + _pageCount + ".", Theme.TextMuted);
            UpdateControls();
        }

        // ---------- конвертация ----------

        private void OnConvertClick(object sender, EventArgs e)
        {
            if (_busy || _sourcePath == null)
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

            _busy = true;
            UpdateControls();
            SetStatus("Извлечение текста…", Theme.TextMuted);

            var thread = new Thread(delegate()
            {
                Exception error = null;
                ConvertResult result = null;
                try
                {
                    result = PdfToWordService.Convert(src, outPath);
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
            _btnOpen.Enabled = !_busy;
            _btnConvert.Enabled = !_busy && _sourcePath != null;
        }
    }
}
