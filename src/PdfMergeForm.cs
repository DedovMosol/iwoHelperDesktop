using System;
using System.Collections.Generic;
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
        private CheckBox _chkPadEven; // добить документы до чётного числа страниц (двусторонняя печать)

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
            var interleave = new ToolStripMenuItem(Loc.T("pdf.menu.interleave"), null,
                delegate { InterleavePages(); });
            BuildHeaderWithHome(Title,
                Loc.T("pdf.header.subtitle"),
                Theme.PdfRed, Theme.PdfRedDark, ShowHelp, interleave);

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

            _chkPadEven = new AccentCheckBox();
            _chkPadEven.Text = Loc.T("pdf.chk.padEven");
            // Под кнопками и по их ширине. Подпись длиннее колонки, поэтому переносится на
            // две строки (см. AccentCheckBox) — обрезать её ради одной строки нельзя.
            _chkPadEven.SetBounds(col, m + 248, 130, 40);
            _chkPadEven.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _chkPadEven.ForeColor = Theme.TextPrimary;
            _tips.SetToolTip(_chkPadEven, Loc.T("pdf.tip.padEven"));
            Controls.Add(_chkPadEven);

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
            PickAndAddFiles();
        }

        /// <summary>
        /// Чередование страниц добавленных документов — сборка пачек одностороннего сканера:
        /// лицевые стороны в одном файле, оборотные в другом и, как правило, задом наперёд.
        /// Берём по странице из каждой пачки по кругу.
        ///
        /// Вопрос про обратный порядок задаём только для ДВУХ пачек: это и есть случай
        /// двусторонней печати, а для трёх и более «которая из них перевёрнута» без отдельного
        /// диалога не выяснить — там просто чередуем по порядку. Перестановка идёт под
        /// снимком отмены, поэтому Ctrl+Z возвращает как было.
        /// </summary>
        private void InterleavePages()
        {
            if (Working || _order.Count == 0)
                return;
            List<PdfPageRef> pages = _order.ToList();
            List<PageRun> runs = PageInterleave.SplitIntoRuns(pages);
            if (!PageInterleave.CanInterleave(runs))
            {
                Dialogs.Info(this, Title, Loc.T("pdf.interleave.needTwo.title"), Loc.T("pdf.interleave.needTwo.body"));
                return;
            }
            if (runs.Count == 2 &&
                Dialogs.ConfirmWarning(this, Title, Loc.T("pdf.interleave.reverse.title"),
                    Loc.T("pdf.interleave.reverse.body")))
                runs[runs.Count - 1].Reverse = true;

            _order.Checkpoint(); // отмена одним Ctrl+Z
            _order.SetOrder(PageInterleave.Interleave(pages, runs, 1));
            RefreshGrid();
            RefreshRestingStatus();
            SyncControls();
            SetStatus(SuccessStatus(string.Format(Loc.T("pdf.interleave.done"), runs.Count)), Theme.OkGreen);
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
            bool padEven = _chkPadEven.Checked;
            BeginOperation(Loc.T("common.status.saving"), pages.Count, Loc.T("pdf.status.savingPage"));
            Action<int, int> onProgress = UiProgress();
            Func<bool> cancel = CancelToken();

            Ui.RunWorker(delegate()
            {
                Exception error = null;
                bool compressed = false;
                try
                {
                    PdfMergeService.Merge(pages, outputPath, onProgress, cancel, padEven);
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
                    // Ход сжатия неизмерим, поэтому на его время полоса становится бегущей, а
                    // статус называет фазу: иначе окно с полной полосой и без «Отмены»
                    // выглядело зависшим всё время работы Ghostscript (до пяти минут).
                    bool willCompress = level != CompressionLevel.None;
                    OnUi(delegate
                    {
                        StopOfferingCancel();
                        if (willCompress)
                            BeginIndeterminate(Loc.T("common.status.compressing"));
                    });
                    compressed = PdfCompression.Compress(outputPath, level);
                }
                bool didCompress = compressed;
                OnUi(delegate { OnSaveFinished(error, outputPath, pages.Count, didCompress, level); });
            });
        }

        private void OnSaveFinished(Exception error, string outputPath, int pageCount, bool compressed, CompressionLevel level)
        {
            if (!FinishOperation(error, Loc.T("pdf.status.saveFailed"), Loc.T("pdf.err.saveFailed")))
                return; // отмена или ошибка — статус и диалог уже показаны базой
            UsageStats.RecordPdfMerge();
            if (compressed)
                UsageStats.RecordPdfCompress();
            SetStatus(DoneStatus(pageCount, compressed, level), Theme.OkGreen);
            Ui.OpenPath(outputPath); // авто-открытие результата; молча, если нет ассоциации PDF
        }

        /// <summary>
        /// Строка успешного сохранения: число записанных страниц и, если сжатие сработало,
        /// разрешение, до которого уменьшены изображения. Отдельным методом, чтобы подстановка
        /// значений проверялась тестом, а не только глазами. Чистая — под тест.
        /// </summary>
        internal static string DoneStatus(int pageCount, bool compressed, CompressionLevel level)
        {
            return SuccessStatus(string.Format(Loc.T("pdf.status.pagesSaved"), pageCount),
                CompressedPart(compressed, level));
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
