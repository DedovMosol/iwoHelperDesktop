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
    internal sealed class PagePreviewForm : Form
    {
        private const int RenderWidthCap = 1600; // верхняя граница ширины рендера (память/скорость)

        private readonly PdfPageRef _page;      // тот же объект, что и в сетке — Rotation общий
        private readonly Action<int> _rotate;   // повернуть страницу в сетке на ±90, null — поворот запрещён
        private readonly PictureBox _picture;
        private readonly Label _loading;
        private ContextMenuStrip _menu; // не дочерний контрол, освобождаем сами
        private Bitmap _image;          // показанный рендер — освобождаем при закрытии
        private int _appliedRotation;   // поворот, УЖЕ впечённый в _image (в градусах по часовой)
        private Thread _worker;
        private volatile bool _closed;  // окно закрыто: поздний рендер не применяем

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

            _picture = new PictureBox();
            _picture.Dock = DockStyle.Fill;
            _picture.SizeMode = PictureBoxSizeMode.Zoom; // вписать с сохранением пропорций (letterbox)
            _picture.BackColor = BackColor;
            // Именно MouseClick с проверкой кнопки: событие Click у PictureBox приходит и по
            // ПРАВОЙ кнопке, поэтому на нём окно закрывалось бы прямо под открывающимся меню.
            _picture.MouseClick += OnSurfaceMouseClick;
            Controls.Add(_picture);

            _loading = new Label();
            _loading.Dock = DockStyle.Fill;
            _loading.TextAlign = ContentAlignment.MiddleCenter;
            _loading.ForeColor = Color.White;
            _loading.BackColor = BackColor;
            _loading.Text = Loc.T("preview.loading");
            _loading.MouseClick += OnSurfaceMouseClick;
            Controls.Add(_loading);
            _loading.BringToFront();

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
            // Меню на форме: подсказка «Загрузка…» перекрывает картинку, а WM_CONTEXTMENU
            // с контрола без своего меню всплывает к родителю — правый клик работает всюду.
            ContextMenuStrip = _menu;
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
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
            SyncRotation(); // пока шёл рендер, страницу могли повернуть — догоняем
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
        }

        /// <summary>
        /// Закрывает ли клик окно предпросмотра. Окно закрывает только ЛЕВАЯ кнопка: правая
        /// принадлежит контекстному меню, и раньше она закрывала окно прямо под открывающимся
        /// меню, потому что событие Click у PictureBox приходит по любой кнопке. Чистая — под тест.
        /// </summary>
        internal static bool ClosesOnClick(MouseButtons button)
        {
            return button == MouseButtons.Left;
        }

        private void OnSurfaceMouseClick(object sender, MouseEventArgs e)
        {
            if (ClosesOnClick(e.Button))
                Close();
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
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
