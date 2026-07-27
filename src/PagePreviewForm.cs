using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Модальный предпросмотр одной страницы PDF в полный размер (двойной клик по плитке).
    /// Рендер — системным Windows.Data.Pdf в ФОНЕ (окно не подвисает), своей копией
    /// <see cref="PdfThumbnailRenderer"/> на своём потоке (тот не потокобезопасен), поэтому
    /// сетке он не мешает. Фон отдаёт страницу БЕЗ поворота, а поворот накладывает уже
    /// UI-поток — так рендер не зависит от того, крутил ли пользователь страницу, пока тот шёл.
    ///
    /// Правый клик открывает поворот вправо и влево (как в Acrobat, те же Ctrl+Shift+«+»/«−»).
    /// Поворот применяется НЕ здесь: форма зовёт переданный колбэк сетки, чтобы поворот прошёл
    /// штатным путём (чекпойнт для Ctrl+Z, чистка лишних плиток, перерисовка миниатюр), и лишь
    /// затем догоняет картинку под новое значение <see cref="PdfPageRef.Rotation"/>. Размер и
    /// положение окна запоминаются между запусками (<see cref="WindowPlacement"/>).
    ///
    /// Bitmap освобождается при закрытии, поздний результат после закрытия отбрасывается.
    /// Только UI-поток, кроме тела фонового рендера.
    /// </summary>
    internal sealed class PagePreviewForm : Form, IMessageFilter
    {
        private const int RenderWidthCap = 1600; // верхняя граница ширины рендера (память/скорость)

        private readonly PdfPageRef _page;      // тот же объект, что и в сетке — Rotation общий
        private readonly Action<int> _rotate;   // повернуть страницу в сетке на ±90, null — поворот запрещён
        private readonly Panel _viewport;       // область с прокруткой — вмещает увеличенную страницу
        private readonly PictureBox _picture;
        private readonly Label _loading;
        private Label _zoomLabel;               // текущий масштаб в процентах
        private ContextMenuStrip _menu; // не дочерний контрол, освобождаем сами
        private Bitmap _image;          // показанный рендер — освобождаем при закрытии
        private int _appliedRotation;   // поворот, УЖЕ впечённый в _image (в градусах по часовой)
        private Thread _worker;
        private volatile bool _closed;  // окно закрыто: поздний рендер не применяем

        // Масштаб. Пока пользователь не трогал лупу, страница подгоняется под окно и следует
        // за его размером; первое же ручное изменение эту привязку снимает — иначе окно
        // «отбирало» бы у пользователя выбранный им масштаб при каждом изменении размера.
        private double _scale = 1.0;
        private bool _fitToWindow = true;
        private Point _dragFrom;        // откуда начали тащить (экранные координаты)
        private Point _dragScroll;      // прокрутка на момент начала перетаскивания
        private bool _dragging;
        private bool _dragged;          // движение вышло за порог — это перетаскивание, а не клик

        private PagePreviewForm(PdfPageRef page, string caption, Size target, Action<int> rotate)
        {
            _page = page;
            _rotate = rotate;

            Text = caption;
            Icon = Ui.AppIcon();
            Font = Ui.Font(9.75f); // общий кэшированный шрифт (не освобождать)
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoScaleMode = AutoScaleMode.None; // размер задаём в пикселях под рабочую область
            ClientSize = target;
            MinimumSize = new Size(320, 240);
            BackColor = Color.FromArgb(37, 37, 38); // тёмная подложка — страница читается контрастно
            KeyPreview = true;

            // Область просмотра с прокруткой: при увеличении страница больше окна, и её
            // нужно и таскать рукой, и прокручивать колесом/полосами как обычно.
            _viewport = new Panel();
            _viewport.Dock = DockStyle.Fill;
            _viewport.AutoScroll = true;
            _viewport.BackColor = BackColor;
            Controls.Add(_viewport);

            _picture = new PictureBox();
            _picture.SizeMode = PictureBoxSizeMode.Zoom; // размер задаём сами, пропорции сохраняются
            _picture.BackColor = BackColor;
            // Именно MouseDown/Up с проверкой кнопки: событие Click у PictureBox приходит и по
            // ПРАВОЙ кнопке, поэтому на нём окно закрывалось бы прямо под открывающимся меню.
            _picture.MouseDown += OnSurfaceMouseDown;
            _picture.MouseMove += OnSurfaceMouseMove;
            _picture.MouseUp += OnSurfaceMouseUp;
            _viewport.Controls.Add(_picture);

            _loading = new Label();
            _loading.Dock = DockStyle.Fill;
            _loading.TextAlign = ContentAlignment.MiddleCenter;
            _loading.ForeColor = Color.White;
            _loading.BackColor = BackColor;
            _loading.Text = Loc.T("preview.loading");
            _loading.MouseUp += OnSurfaceMouseUp;
            Controls.Add(_loading);
            _loading.BringToFront();

            BuildZoomBar();

            if (_rotate != null)
                BuildContextMenu();
            WindowPlacement.Attach(this); // размер и положение окна между запусками
        }

        /// <summary>Показать предпросмотр страницы модально над owner. caption — подпись окна (номер
        /// страницы). rotate — повернуть страницу на ±90 средствами сетки, null запрещает поворот.</summary>
        public static void Show(IWin32Window owner, PdfPageRef page, string caption, Action<int> rotate)
        {
            if (page == null || string.IsNullOrEmpty(page.SourcePath))
                return;
            Rectangle wa = Screen.FromPoint(Control.MousePosition).WorkingArea;
            var target = new Size((int)(wa.Width * 0.82), (int)(wa.Height * 0.88));
            using (var f = new PagePreviewForm(page, caption, target, rotate))
                f.ShowDialog(owner);
        }

        private void BuildContextMenu()
        {
            _menu = new ContextMenuStrip();
            _menu.Items.Add(Loc.T("preview.menu.rotateRight"), null, delegate { Rotate(90); });
            _menu.Items.Add(Loc.T("preview.menu.rotateLeft"), null, delegate { Rotate(-90); });
            _menu.Items.Add(new ToolStripSeparator());
            // Печать той страницы, которую пользователь прямо сейчас рассматривает: смотреть
            // на лист и не иметь возможности его напечатать — самое очевидное, чего не хватало.
            _menu.Items.Add(Loc.T("preview.menu.print"), null, delegate { PrintThisPage(); });
            // Меню на форме: подсказка «Загрузка…» перекрывает картинку, а WM_CONTEXTMENU
            // с контрола без своего меню всплывает к родителю — правый клик работает всюду.
            ContextMenuStrip = _menu;
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            Application.AddMessageFilter(this); // снимается в OnFormClosed — иначе удержит окно
            int width = Math.Min(RenderWidthCap, Math.Max(ClientSize.Width, 400));
            _worker = new Thread(delegate() { RenderInBackground(width); });
            _worker.IsBackground = true;
            _worker.Name = "pdf-preview";
            _worker.Start();
        }

        private void RenderInBackground(int width)
        {
            PdfThumbnailRenderer renderer = null;
            Bitmap page = null;
            try
            {
                renderer = new PdfThumbnailRenderer();
                page = renderer.Render(_page.SourcePath, _page.PageIndex, width); // без поворота, его наложит UI
            }
            catch { page = null; } // предпросмотр не критичен — покажем сообщение
            finally { if (renderer != null) renderer.Dispose(); }
            Post(page);
        }

        private void Post(Bitmap page)
        {
            try
            {
                if (_closed || !IsHandleCreated || IsDisposed)
                {
                    if (page != null) page.Dispose();
                    return;
                }
                BeginInvoke((MethodInvoker)delegate { Apply(page); });
            }
            catch (InvalidOperationException)
            {
                if (page != null) page.Dispose();
            }
        }

        private void Apply(Bitmap page)
        {
            if (_closed)
            {
                if (page != null) page.Dispose();
                return;
            }
            if (page == null)
            {
                _loading.Text = Loc.T("preview.unavailable");
                return;
            }
            _image = page;
            _appliedRotation = 0; // фон отдаёт страницу как есть
            _loading.Visible = false;
            _picture.Image = _image;
            SyncRotation();  // пока шёл рендер, страницу могли повернуть — догоняем
            LayoutImage();   // первый показ — вписываем страницу в окно
        }

        /// <summary>
        /// Напечатать показанную страницу. Диалог принтера — на UI-потоке, печать — в фоне,
        /// как и везде: рендер под печать идёт при 200 dpi и мгновенным не бывает.
        /// Ошибку показываем прямо здесь: окно модальное, идти ей больше некуда.
        /// </summary>
        private void PrintThisPage()
        {
            System.Drawing.Printing.PrinterSettings settings;
            using (var dialog = new PrintDialog())
            {
                dialog.AllowSomePages = false;
                dialog.UseEXDialog = true;
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                settings = dialog.PrinterSettings;
            }
            string path = _page.SourcePath;
            var pages = new System.Collections.Generic.List<int> { _page.PageIndex };
            Ui.RunWorker(delegate()
            {
                Exception error = null;
                try { PdfPrintService.Print(path, pages, settings); }
                catch (Exception ex) { error = ex; }
                Exception shown = error;
                if (shown != null)
                    Ui.OnUi(this, delegate
                    {
                        Dialogs.Error(this, Text, Loc.T("preview.err.printFailed"), shown.Message);
                    });
            });
        }

        /// <summary>
        /// Повернуть страницу на delta градусов: сначала штатным путём сетки (чекпойнт отмены и
        /// перерисовка миниатюр), затем догнать показанную картинку. Если сетка поворот не
        /// приняла (например форма занята операцией), Rotation не изменится и картинка тоже.
        /// </summary>
        private void Rotate(int delta)
        {
            if (_rotate == null)
                return;
            _rotate(delta);
            SyncRotation();
        }

        /// <summary>
        /// Довернуть показанную картинку до текущего <see cref="PdfPageRef.Rotation"/>.
        /// Поворот на кратный 90° угол точен, поэтому крутим ТОТ ЖЕ bitmap на разницу, без
        /// копии и без повторного рендера. Ссылку из PictureBox снимаем на время поворота:
        /// RotateFlip меняет размеры прямо в объекте, который иначе рисуется на экране.
        /// </summary>
        private void SyncRotation()
        {
            if (_image == null)
                return;
            int desired = _page.Rotation;
            int delta = PdfPageRef.ComposeRotation(desired, -_appliedRotation);
            if (delta == 0)
                return;
            _picture.Image = null;
            _image.RotateFlip(PageRotation.FlipFor(delta));
            _picture.Image = _image;
            _appliedRotation = desired;
            LayoutImage(); // поворот меняет пропорции — пересчитываем размер и прокрутку
        }

        // ---------- масштаб и панорама ----------

        /// <summary>Полоса лупы внизу окна: «−», проценты, «+» и «по окну».</summary>
        private void BuildZoomBar()
        {
            var bar = new Panel();
            bar.Dock = DockStyle.Bottom;
            bar.Height = 34;
            bar.BackColor = Color.FromArgb(52, 52, 54);
            Controls.Add(bar);

            AddZoomButton(bar, "−", 8, Loc.T("preview.tip.zoomOut"), delegate { StepZoom(-1); });
            _zoomLabel = new Label();
            _zoomLabel.SetBounds(44, 8, 60, 20);
            _zoomLabel.TextAlign = ContentAlignment.MiddleCenter;
            _zoomLabel.ForeColor = Color.White;
            bar.Controls.Add(_zoomLabel);
            AddZoomButton(bar, "+", 108, Loc.T("preview.tip.zoomIn"), delegate { StepZoom(+1); });

            var print = new Button();
            print.Text = Loc.T("preview.print");
            print.SetBounds(264, 5, 110, 24);
            print.FlatStyle = FlatStyle.Flat;
            print.ForeColor = Color.White;
            print.FlatAppearance.BorderColor = Color.FromArgb(90, 90, 92);
            print.Click += delegate { PrintThisPage(); };
            bar.Controls.Add(print);

            var fit = new Button();
            fit.Text = Loc.T("preview.fit");
            fit.SetBounds(148, 5, 110, 24);
            fit.FlatStyle = FlatStyle.Flat;
            fit.ForeColor = Color.White;
            fit.FlatAppearance.BorderColor = Color.FromArgb(90, 90, 92);
            fit.Click += delegate { FitToWindow(); };
            bar.Controls.Add(fit);
        }

        private void AddZoomButton(Control parent, string text, int x, string tip, EventHandler onClick)
        {
            var b = new Button();
            b.Text = text;
            b.SetBounds(x, 5, 32, 24);
            b.FlatStyle = FlatStyle.Flat;
            b.ForeColor = Color.White;
            b.FlatAppearance.BorderColor = Color.FromArgb(90, 90, 92);
            b.Font = Ui.Font(11f, FontStyle.Bold);
            b.AccessibleName = tip;
            b.Click += onClick;
            parent.Controls.Add(b);
        }

        /// <summary>Шаг лупы кнопкой или клавишей — от центра области, там курсора нет.</summary>
        private void StepZoom(int direction)
        {
            var center = new Point(_viewport.ClientSize.Width / 2, _viewport.ClientSize.Height / 2);
            ApplyScale(PreviewZoom.Next(_scale, direction), center);
        }

        /// <summary>Вернуть подгонку по окну: страница снова видна целиком и следует за размером окна.</summary>
        private void FitToWindow()
        {
            _fitToWindow = true;
            LayoutImage();
        }

        /// <summary>
        /// Задать масштаб, оставив точку под anchor на месте. anchor — в координатах области
        /// просмотра. Ручной масштаб снимает привязку к размеру окна.
        /// </summary>
        private void ApplyScale(double scale, Point anchor)
        {
            if (_image == null)
                return;
            double old = _scale;
            _scale = scale;
            _fitToWindow = false;
            LayoutImage(old, anchor);
        }

        /// <summary>
        /// Пересчитать размер картинки и прокрутку под текущий масштаб. Когда oldScale задан,
        /// прокрутка подбирается так, чтобы точка под anchor не сдвинулась (Ctrl+колесо).
        /// </summary>
        private void LayoutImage(double oldScale = 0, Point anchor = default(Point))
        {
            if (_image == null)
                return;
            Size viewport = _viewport.ClientSize;
            if (_fitToWindow)
                _scale = PreviewZoom.Fit(_image.Size, viewport);

            Size scaled = PreviewZoom.Scaled(_image.Size, _scale);
            _picture.Size = scaled;
            // Меньше области — картинка стоит по центру, как в любом просмотрщике.
            _picture.Location = new Point(
                Math.Max(0, (viewport.Width - scaled.Width) / 2),
                Math.Max(0, (viewport.Height - scaled.Height) / 2));

            if (oldScale > 0)
            {
                // ВНИМАНИЕ: AutoScrollPosition ЧИТАЕТСЯ отрицательным, а ЗАДАЁТСЯ положительным —
                // это давняя особенность WinForms и частый источник «прокрутка прыгает не туда».
                Point current = new Point(-_viewport.AutoScrollPosition.X, -_viewport.AutoScrollPosition.Y);
                _viewport.AutoScrollPosition = new Point(
                    PreviewZoom.Anchor(current.X, anchor.X, oldScale, _scale),
                    PreviewZoom.Anchor(current.Y, anchor.Y, oldScale, _scale));
            }
            UpdateZoomUi();
        }

        private void UpdateZoomUi()
        {
            if (_zoomLabel != null)
                _zoomLabel.Text = PreviewZoom.Percent(_scale) + " %";
            _picture.Cursor = CanPan ? Cursors.Hand : Cursors.Default;
        }

        /// <summary>Есть ли что таскать: картинка не помещается в область целиком.</summary>
        private bool CanPan
        {
            get { return _image != null && !PreviewZoom.FitsEntirely(_picture.Size, _viewport.ClientSize); }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (_fitToWindow)
                LayoutImage(); // подгонка следует за окном, ручной масштаб — нет
            else
                UpdateZoomUi();
        }

        // Колесо мыши. Windows отправляет его контролу С ФОКУСОМ, а не тому, над которым
        // курсор, поэтому обычная подписка на событие вела бы себя по-разному в зависимости
        // от того, где сейчас фокус: над панелью с прокруткой колесо могло бы и прокручивать,
        // и не доходить вовсе. Фильтр сообщений снимает эту зависимость: пока окно открыто,
        // колесо над областью просмотра обрабатываем сами — с Ctrl это масштаб к точке под
        // курсором (как в браузерах), без Ctrl обычная вертикальная прокрутка.
        private const int WmMouseWheel = 0x020A;
        private const int WheelNotch = 120;   // стандартный «щелчок» колеса
        private const int ScrollPerNotch = 60; // на сколько пикселей прокручивать за щелчок

        bool IMessageFilter.PreFilterMessage(ref Message m)
        {
            if (m.Msg != WmMouseWheel || _closed || _image == null)
                return false;
            Point screen = Cursor.Position;
            if (!_viewport.RectangleToScreen(_viewport.ClientRectangle).Contains(screen))
                return false; // курсор вне области просмотра — пусть обрабатывают как обычно
            int delta = (short)((long)m.WParam >> 16);
            Point inViewport = _viewport.PointToClient(screen);
            if ((ModifierKeys & Keys.Control) != 0)
            {
                ApplyScale(PreviewZoom.Next(_scale, delta > 0 ? +1 : -1), inViewport);
                return true;
            }
            ScrollBy(0, -delta * ScrollPerNotch / WheelNotch);
            return true;
        }

        /// <summary>
        /// Сдвинуть прокрутку на dx/dy пикселей. ВНИМАНИЕ: AutoScrollPosition ЧИТАЕТСЯ
        /// отрицательным, а ЗАДАЁТСЯ положительным — давняя особенность WinForms.
        /// </summary>
        private void ScrollBy(int dx, int dy)
        {
            _viewport.AutoScrollPosition = new Point(
                -_viewport.AutoScrollPosition.X + dx,
                -_viewport.AutoScrollPosition.Y + dy);
        }

        private void OnSurfaceMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;
            _dragging = true;
            _dragged = false;
            _dragFrom = Control.MousePosition;
            _dragScroll = new Point(-_viewport.AutoScrollPosition.X, -_viewport.AutoScrollPosition.Y);
        }

        private void OnSurfaceMouseMove(object sender, MouseEventArgs e)
        {
            if (!_dragging)
                return;
            Point now = Control.MousePosition;
            if (!_dragged && !PreviewZoom.IsDrag(_dragFrom, now, SystemInformation.DragSize))
                return; // ещё дрожание руки, а не перетаскивание
            _dragged = true;
            if (!CanPan)
                return;
            // Тащим содержимое за курсором: ушли мышью вправо — содержимое поехало вправо,
            // значит смотрим левее, поэтому прокрутка уменьшается.
            _viewport.AutoScrollPosition = new Point(
                _dragScroll.X - (now.X - _dragFrom.X),
                _dragScroll.Y - (now.Y - _dragFrom.Y));
        }

        private void OnSurfaceMouseUp(object sender, MouseEventArgs e)
        {
            _dragging = false;
            _dragged = false;
            // Клик по странице окно НЕ закрывает. Пока просмотр был просто картинкой, это было
            // удобно, но теперь по странице возят лупой и таскают её рукой — закрываться от
            // касания рабочей области стало мешать. Закрывают Esc и крестик, как в любом
            // просмотрщике, и об этом написано прямо в окне.
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
                return;
            }
            if (PdfToolFormBase.IsResetZoomKey(e.KeyData)) // Ctrl+0 — натуральная величина, как в сетке
            {
                ApplyScale(1.0, new Point(_viewport.ClientSize.Width / 2, _viewport.ClientSize.Height / 2));
                e.Handled = true;
                return;
            }
            // Те же сочетания поворота, что и в сетке — разбор один на всё приложение.
            // Когда поворот запрещён, клавишу НЕ съедаем: пусть уходит дальше как обычно.
            if (_rotate != null)
            {
                switch (PdfToolFormBase.ClassifyPageKey(e.KeyData))
                {
                    case PdfToolFormBase.PageKeyAction.RotateRight:
                        Rotate(90);
                        e.Handled = true;
                        return;
                    case PdfToolFormBase.PageKeyAction.RotateLeft:
                        Rotate(-90);
                        e.Handled = true;
                        return;
                }
            }
            base.OnKeyDown(e);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _closed = true;
            Application.RemoveMessageFilter(this); // снимаем всегда: фильтр держал бы ссылку на окно
            base.OnFormClosed(e);
            _picture.Image = null; // снять ссылку до Dispose bitmap
            if (_image != null)
            {
                _image.Dispose();
                _image = null;
            }
        }

        protected override void Dispose(bool disposing)
        {
            // ContextMenuStrip назначен свойством, а не добавлен в Controls: сам он не освободится.
            // Освобождаем здесь, а не в Closed — там меню ещё может обрабатывать свой клик.
            // Сначала отвязываем от формы, чтобы база не трогала уже освобождённый объект.
            if (disposing && _menu != null)
            {
                ContextMenuStrip = null;
                _menu.Dispose();
                _menu = null;
            }
            base.Dispose(disposing);
        }
    }
}
