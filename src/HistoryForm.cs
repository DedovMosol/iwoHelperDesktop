using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Окно «История операций»: что и когда сделано и где лежит результат.
    ///
    /// Отдельным окном, а не разделом настроек, по двум причинам. Первая — арифметика: окно
    /// настроек занимает 580 точек по высоте, а на минимальном поддерживаемом экране
    /// 1024×768 рабочая область с панелью задач около 728 — список туда не помещается.
    /// Вторая — обязанности: настройки настраивают, список показывает, и смешивать их в одном
    /// окне значит получить окно, которое делает два разных дела.
    ///
    /// Сами файлы здесь не открываются вслепую: путь мог устареть, файл — переехать или быть
    /// удалённым, и «Открыть» обязано это заметить, а не бросить исключение оболочки.
    /// </summary>
    public class HistoryForm : Form
    {
        private const int Pad = 20;
        private const int BtnH = 34;

        private readonly ListView _list;
        private readonly RoundedButton _openFile, _openFolder;

        public HistoryForm()
        {
            Ui.InitDialog(this, Loc.T("history.title"));
            ClientSize = new Size(720, 460);
            WindowChrome.Enable(this, Theme.HubBlue);

            Ui.AccentBar(this, 0, Theme.HubBlue);
            Ui.Label(this, Loc.T("history.title"), Pad, 18, Ui.Font(14f, FontStyle.Bold),
                Color.FromArgb(40, 40, 40));

            _list = new ListView();
            _list.View = View.Details;
            _list.FullRowSelect = true;
            _list.MultiSelect = false; // «Открыть» относится к ОДНОЙ строке
            _list.HideSelection = false;
            _list.GridLines = false;
            _list.SetBounds(Pad, 58, ClientSize.Width - 2 * Pad, ClientSize.Height - 58 - BtnH - 2 * Pad);
            _list.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _list.Columns.Add(Loc.T("history.col.when"), 130);
            _list.Columns.Add(Loc.T("history.col.what"), 160);
            // Путь — последней и самой широкой: он длиннее всего и его читают чаще прочего.
            _list.Columns.Add(Loc.T("history.col.path"), _list.Width - 130 - 160 - 28);
            _list.SelectedIndexChanged += delegate { SyncButtons(); };
            _list.DoubleClick += delegate { OpenSelected(false); };
            Controls.Add(_list);

            int y = ClientSize.Height - Pad - BtnH;
            _openFile = Button(Loc.T("excel.link.openFile"), Pad, y);
            _openFile.Click += delegate { OpenSelected(false); };

            _openFolder = Button(Loc.T("excel.link.openFolder"), _openFile.Right + 10, y);
            _openFolder.Click += delegate { OpenSelected(true); };

            var close = new RoundedButton(true);
            close.SetBounds(ClientSize.Width - Pad - 100, y, 100, BtnH);
            close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Controls.Add(close);
            close.Text = Loc.T("common.close");
            close.Click += delegate { Close(); };
            AcceptButton = close;
            CancelButton = close;

            LoadAndShow();
        }

        private RoundedButton Button(string text, int x, int y)
        {
            var b = new RoundedButton(false);
            b.SetBounds(x, y, 150, BtnH);
            b.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            Controls.Add(b); // шрифт наследуется от формы, до добавления мерить нечем
            b.Text = text;
            b.Width = Math.Max(150, TextRenderer.MeasureText(text, b.Font).Width + 28);
            return b;
        }

        /// <summary>Записи от новых к старым: свежее наверху — его и ищут.</summary>
        private void LoadAndShow()
        {
            OperationHistory.Data d = OperationHistory.Load();
            _list.BeginUpdate();
            _list.Items.Clear();
            for (int i = d.Entries.Count - 1; i >= 0; i--)
            {
                HistoryEntry e = d.Entries[i];
                // Время храним в UTC, показываем местное: файл переживает смену часового пояса.
                var item = new ListViewItem(e.WhenUtc.ToLocalTime().ToString("dd.MM.yyyy HH:mm"));
                item.SubItems.Add(Loc.T(e.Operation)); // ключ → подпись на ТЕКУЩЕМ языке
                item.SubItems.Add(e.Path);
                item.Tag = e.Path;
                _list.Items.Add(item);
            }
            _list.EndUpdate();
            SyncButtons();
        }

        private bool _checkingPath;

        private void SyncButtons()
        {
            bool has = !_checkingPath && _list.SelectedItems.Count == 1;
            _openFile.Enabled = has;
            _openFolder.Enabled = has;
        }

        /// <summary>
        /// Открыть выбранное. Существование проверяется ЗАРАНЕЕ: история хранит путь, а файл
        /// могли удалить или перенести, и молчаливый отказ оболочки выглядел бы как «кнопка
        /// не работает». Папку показываем с выделенным файлом, если он ещё на месте.
        /// </summary>
        private void OpenSelected(bool folder)
        {
            if (_checkingPath || _list.SelectedItems.Count != 1)
                return;
            string path = (string)_list.SelectedItems[0].Tag;
            _checkingPath = true;
            SyncButtons();
            var owner = new WeakReference(this);
            Ui.RunWorker(delegate
            {
                bool isFile = false, isFolder = false;
                try
                {
                    isFile = File.Exists(path);
                    if (!isFile)
                        isFolder = Directory.Exists(path);
                }
                catch { }
                HistoryForm form = owner.Target as HistoryForm;
                if (form == null)
                    return;
                Ui.OnUi(form, delegate
                {
                    form._checkingPath = false;
                    form.SyncButtons();
                    if (!isFile && !isFolder)
                    {
                        Dialogs.Error(form, Loc.T("history.title"),
                            Loc.T("history.err.gone.title"),
                            string.Format(Loc.T("history.err.gone.body"), path));
                        return;
                    }
                    Ui.OpenPathOrWarn(form, Loc.T("history.title"), path,
                        folder && isFile);
                });
            });
        }
    }
}
