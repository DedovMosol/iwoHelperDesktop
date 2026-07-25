using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Модальный предпросмотр одной страницы PDF в полный размер (двойной клик по плитке).
    /// Рендер — системным Windows.Data.Pdf в ФОНЕ (окно не подвисает), с учётом поворота
    /// страницы; закрытие по Esc/клику/кнопке. Своя копия <see cref="PdfThumbnailRenderer"/>
    /// на своём фоновом потоке (тот не потокобезопасен) — не мешает сетке. Bitmap
    /// освобождается при закрытии; поздний результат после закрытия отбрасывается.
    /// </summary>
    internal sealed class PagePreviewForm : Form
    {
        private const int RenderWidthCap = 1600; // верхняя граница ширины рендера (память/скорость)

        private readonly string _path;
        private readonly int _pageIndex;
        private readonly int _rotation;
        private readonly PictureBox _picture;
        private readonly Label _loading;
        private Bitmap _image;         // показанный рендер — освобождаем при закрытии
        private Thread _worker;
        private volatile bool _closed; // окно закрыто: поздний рендер не применяем

        private PagePreviewForm(PdfPageRef page, string caption, Size target)
        {
            _path = page.SourcePath;
            _pageIndex = page.PageIndex;
            _rotation = page.Rotation;

            Text = caption;
            Icon = Ui.AppIcon();
            Font = new Font("Segoe UI", 9.75f);
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
            _picture.Click += delegate { Close(); }; // клик по превью закрывает
            Controls.Add(_picture);

            _loading = new Label();
            _loading.Dock = DockStyle.Fill;
            _loading.TextAlign = ContentAlignment.MiddleCenter;
            _loading.ForeColor = Color.White;
            _loading.BackColor = BackColor;
            _loading.Text = Loc.T("preview.loading");
            Controls.Add(_loading);
            _loading.BringToFront();
        }

        /// <summary>Показать предпросмотр страницы модально над owner. caption — подпись окна (номер страницы).</summary>
        public static void Show(IWin32Window owner, PdfPageRef page, string caption)
        {
            if (page == null || string.IsNullOrEmpty(page.SourcePath))
                return;
            Rectangle wa = Screen.FromPoint(Control.MousePosition).WorkingArea;
            var target = new Size((int)(wa.Width * 0.82), (int)(wa.Height * 0.88));
            using (var f = new PagePreviewForm(page, caption, target))
                f.ShowDialog(owner);
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
                page = renderer.Render(_path, _pageIndex, width);
                if (page != null && _rotation != 0)
                {
                    // RotateFlip мутирует — поворачиваем этот же (наш) bitmap, чужого кэша тут нет.
                    page.RotateFlip(PageRotation.FlipFor(_rotation));
                }
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
            _loading.Visible = false;
            _picture.Image = _image;
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
                return;
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
    }
}
