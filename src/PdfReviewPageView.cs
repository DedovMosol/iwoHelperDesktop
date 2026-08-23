using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Одна read-only страница для бокового просмотра Review. Рендерит в фоне,
    /// масштабирует существующий bitmap без повторного чтения PDF, отбрасывает
    /// поздний результат по generation и поверх растра рисует подсветку изменений
    /// (рамки слов из ворд-диффа). Не владеет общим grid renderer.
    /// </summary>
    internal sealed class PdfReviewPageView : UserControl
    {
        private readonly Panel _viewport;
        private readonly PictureBox _picture;
        private readonly Label _status;
        private readonly RoundedButton _minus, _plus, _fit;
        private Bitmap _bitmap;
        private PdfPageRef _page;
        private double _scale = 1.0;
        private int _generation;
        private readonly object _renderGate = new object();
        private RenderRequest _pending;
        private bool _renderWorker;

        private sealed class RenderRequest
        {
            public int Generation;
            public PdfPageRef Page;
            public string Caption;
            public PdfReviewHighlight Highlight;
        }

        public PdfReviewPageView()
        {
            BackColor = Theme.DarkBarFill;
            _minus = Button("−", 8);
            _plus = Button("+", 46);
            _fit = Button(Loc.T("preview.fit"), 84, 100);
            _minus.Click += delegate { Step(-1); };
            _plus.Click += delegate { Step(1); };
            _fit.Click += delegate { Fit(); };

            _status = new Label();
            _status.ForeColor = Color.White;
            _status.BackColor = Color.Transparent;
            _status.AutoEllipsis = true;
            _status.TextAlign = ContentAlignment.MiddleLeft;
            _status.SetBounds(194, 8, 200, 28);
            _status.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(_status);

            _viewport = new Panel();
            _viewport.AutoScroll = true;
            _viewport.BackColor = Color.FromArgb(55, 55, 58);
            _viewport.SetBounds(0, 44, Width, Height - 44);
            _viewport.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _viewport.Resize += delegate { PlacePicture(); };
            Controls.Add(_viewport);

            _picture = new PictureBox();
            _picture.BackColor = Color.White;
            _picture.SizeMode = PictureBoxSizeMode.StretchImage;
            _picture.DoubleClick += delegate
            {
                if (_page != null)
                    PagePreviewForm.Show(FindForm(), _page,
                        string.Format(Loc.T("preview.title"), _page.PageIndex + 1), null);
            };
            _viewport.Controls.Add(_picture);
            UpdateButtons();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (_status == null || _viewport == null)
                return; // OnResize в середине конструктора (контролы ещё не созданы)
            // На узкой панели (разделение сплиттера) строка статуса может не поместиться
            // после кнопок — прижимаем её к правому краю и не даём вылезти за ширину.
            int statusX = Math.Min(194, Math.Max(0, Width - 10));
            _status.SetBounds(statusX, 8, Math.Max(1, Width - statusX - 8), 28);
            _viewport.SetBounds(0, 44, Width, Math.Max(1, Height - 44));
            PlacePicture();
        }

        /// <summary>Показать страницу <paramref name="page"/>; null — «страницы нет».</summary>
        public void ShowPage(PdfPageRef page, string caption, PdfReviewHighlight highlight)
        {
            int generation = ++_generation;
            DisposeBitmap();
            _picture.Visible = false;
            _status.Text = caption ?? "";
            if (page == null)
            {
                lock (_renderGate) _pending = null;
                _status.Text = Loc.T("review.source.missing");
                UpdateButtons();
                return;
            }

            bool start = false;
            lock (_renderGate)
            {
                // Только последний запрос заслуживает рендера. Один worker на pane не даёт
                // быстрому листанию плодить десятки WinRT-документов и растров параллельно.
                _pending = new RenderRequest
                {
                    Generation = generation,
                    Page = page.Clone(),
                    Caption = caption,
                    Highlight = highlight
                };
                if (!_renderWorker)
                {
                    _renderWorker = true;
                    start = true;
                }
            }
            _status.Text = Loc.T("preview.loading");
            if (start)
                Ui.RunWorker(RenderLoop);
        }

        private void RenderLoop()
        {
            while (true)
            {
                RenderRequest request;
                lock (_renderGate)
                {
                    request = _pending;
                    _pending = null;
                    if (request == null)
                    {
                        _renderWorker = false;
                        return;
                    }
                }

                Bitmap rendered = null;
                try
                {
                    using (var renderer = new PdfThumbnailRenderer())
                        rendered = renderer.Render(request.Page.SourcePath,
                            request.Page.PageIndex, 1200, 20000); // ≤24 млн пикселей
                    if (rendered != null && request.Page.Rotation != 0)
                        rendered.RotateFlip(PageRotation.FlipFor(request.Page.Rotation));
                    // Подсветка рисуется СРАЗУ на копии растра в воркере: переключение пар
                    // и зум тогда ничего не перерисовывают (картинка уже готова).
                    if (rendered != null)
                        rendered = DrawHighlight(rendered, request.Highlight);
                }
                catch
                {
                    // Сбой рендера одной страницы не должен ронять воркер: «нет картинки»
                    // покажет сам ApplyRendered (null вместо битмапа), а следующая пара
                    // всё равно перерисует панель.
                }
                Bitmap ready = rendered;
                if (!Ui.OnUi(this, delegate
                {
                    ApplyRendered(request.Generation, request.Page, ready, request.Caption);
                }))
                    if (ready != null) ready.Dispose();
            }
        }

        /// <summary>
        /// Возвращает растр с нарисованной подсветкой (свой — входной освобождается).
        /// Пустая подсветка возвращает входной растр без копирования.
        /// </summary>
        private static Bitmap DrawHighlight(Bitmap source, PdfReviewHighlight highlight)
        {
            if (highlight == null || highlight.Boxes.Count == 0 ||
                highlight.ViewWidthPt <= 0 || highlight.ViewHeightPt <= 0)
                return source;
            var copy = new Bitmap(source);
            source.Dispose();
            using (Graphics g = Graphics.FromImage(copy))
            using (var fill = new SolidBrush(Color.FromArgb(64, highlight.Color)))
            using (var border = new Pen(Color.FromArgb(160, highlight.Color), 1f))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                foreach (PdfReviewBox box in highlight.Boxes)
                {
                    RectangleF rect = PdfReviewGeometry.ToPixelRect(box,
                        highlight.ViewWidthPt, highlight.ViewHeightPt, copy.Width, copy.Height);
                    if (rect.Width < 1 || rect.Height < 1)
                        continue; // рамка меньше пикселя: не рисуем, чтобы не засорять растр
                    g.FillRectangle(fill, rect);
                    g.DrawRectangle(border, rect.X, rect.Y, rect.Width, rect.Height);
                }
            }
            return copy;
        }

        private void ApplyRendered(int generation, PdfPageRef request, Bitmap rendered, string caption)
        {
            if (generation != _generation || IsDisposed)
            {
                if (rendered != null) rendered.Dispose();
                return;
            }
            DisposeBitmap();
            _bitmap = rendered;
            _page = request;
            if (_bitmap == null)
            {
                _status.Text = Loc.T("preview.unavailable");
                _picture.Visible = false;
            }
            else
            {
                _status.Text = caption ?? "";
                _picture.Image = _bitmap;
                _picture.Visible = true;
                Fit();
            }
            UpdateButtons();
        }

        private RoundedButton Button(string text, int x, int width = 32)
        {
            var button = new RoundedButton(ButtonLook.OnDark);
            button.Text = text;
            button.SetBounds(x, 8, width, 28);
            button.AccessibleName = text;
            Controls.Add(button);
            return button;
        }

        private void Step(int direction)
        {
            if (_bitmap == null) return;
            _scale = PreviewZoom.Next(_scale, direction);
            PlacePicture();
        }

        private void Fit()
        {
            if (_bitmap == null) return;
            _scale = PreviewZoom.Fit(_bitmap.Size,
                new Size(Math.Max(1, _viewport.ClientSize.Width - 24),
                         Math.Max(1, _viewport.ClientSize.Height - 24)));
            PlacePicture();
        }

        private void PlacePicture()
        {
            if (_bitmap == null || !_picture.Visible) return;
            Size size = PreviewZoom.Scaled(_bitmap.Size, _scale);
            _picture.Size = size;
            _picture.Location = PreviewZoom.Centered(size, _viewport.ClientSize,
                _viewport.AutoScrollPosition);
            _status.Text = (_page == null ? "" : (_page.PageIndex + 1).ToString()) +
                " · " + PreviewZoom.Percent(_scale) + " %";
            UpdateButtons();
        }

        private void UpdateButtons()
        {
            bool ready = _bitmap != null;
            _minus.Enabled = ready && _scale > PreviewZoom.Min;
            _plus.Enabled = ready && _scale < PreviewZoom.Max;
            _fit.Enabled = ready;
        }

        private void DisposeBitmap()
        {
            _picture.Image = null;
            if (_bitmap != null) _bitmap.Dispose();
            _bitmap = null;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _generation++;
                lock (_renderGate) _pending = null;
                DisposeBitmap();
            }
            base.Dispose(disposing);
        }
    }
}
