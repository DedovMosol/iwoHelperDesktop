using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Единый redline: поздняя/правая страница — единственный растр. Добавления остаются
    /// зелёными; удаление отмечается красной кареткой в точке ПРАВОЙ страницы, где стоял
    /// удалённый текст (край ближайшего контрагента или вставки замены), и рейкой «−» у
    /// левого поля — той же грамматикой, что левая панель «Бок о бок». Точки нет — остаётся
    /// рейка у верхнего поля: правку не прячем. Рамки удалённых слов в координатах левой
    /// страницы на правый растр не переносятся: после перевёрстки они ложились на
    /// неизменённый текст. Точный удалённый текст остаётся в режиме «Бок о бок».
    /// </summary>
    internal static class PdfReviewUnifiedRenderer
    {
        // Ширина каретки — доля высоты слова-соседа (не меньше минимума в пунктах), чтобы
        // полоса изменений получила строку даже на мелком растре.
        private const double CaretWidthShare = 0.08;
        private const double CaretMinWidthPt = 1.5;
        // Рейка без точки: строка высотой RailLineHeightPt под верхним полем страницы.
        private const double RailTopMarginPt = 36;
        private const double RailLineHeightPt = 12;

        /// <summary>
        /// Каретки удалений пары в координатах правой страницы. shownDeletions — удалённые
        /// слова, которые показала бы левая панель (BuildHighlight уже отфильтровал, например,
        /// стабильный номер страницы). Якорь ищется одним проходом по Operations среди слов
        /// этой правой страницы: правый край предыдущего контрагента (любого вида
        /// сопоставления — ближайший в порядке чтения), иначе левый край первой вставки той же
        /// правки (замена: вставленный текст стоит там, где был удалённый), иначе левый край
        /// следующего контрагента. Якоря нет — одна рейка «−» у верхнего поля без каретки.
        /// null — на левой странице пары нет показываемых удалений.
        /// </summary>
        internal static PdfReviewHighlight DeletionAnchors(PdfReviewResult result,
            PdfReviewPagePair pair, ICollection<PdfReviewWord> shownDeletions)
        {
            if (result == null || pair == null || pair.LeftPageIndex < 0 ||
                pair.RightPageIndex < 0 || shownDeletions == null || shownDeletions.Count == 0)
                return null;
            PdfReviewPage rightPage = PdfReviewDiff.PageAt(result.Right, pair.RightPageIndex);
            if (rightPage == null)
                return null;
            var highlight = new PdfReviewHighlight
            {
                Color = Theme.ReviewDeleteFill,
                EdgeColor = Theme.ReviewDeleteMarker,
                Style = PdfReviewHighlightStyle.Removed,
                ChangeBarSide = PdfReviewChangeBarSide.Left,
                ViewWidthPt = rightPage.ViewWidthPt,
                ViewHeightPt = rightPage.ViewHeightPt
            };
            int page = pair.RightPageIndex;
            var anchors = new HashSet<PdfReviewBox>();
            PdfReviewWord previous = null;
            PdfReviewWord replacement = null; // первая вставка текущей правки на этой странице
            bool waiting = false, unanchored = false;
            foreach (PdfReviewWordOp op in result.Operations)
            {
                if (op == null)
                    continue;
                if (op.Kind == PdfReviewDiffKind.Equal)
                {
                    if (op.RightWords.Count == 0)
                        continue; // пустой sentinel не разделяет правки
                    if (waiting && !AddCaret(highlight, anchors,
                            replacement ?? OnPage(op.RightWords[0], page), false))
                        unanchored = true;
                    waiting = false;
                    replacement = null;
                    previous = op.RightWords[op.RightWords.Count - 1];
                }
                else if (op.Kind == PdfReviewDiffKind.Insert)
                {
                    if (replacement == null)
                        replacement = FirstOnPage(op.RightWords, page);
                }
                else if (op.Kind == PdfReviewDiffKind.Delete && Shows(op.LeftWords, shownDeletions))
                {
                    if (!AddCaret(highlight, anchors, OnPage(previous, page), true))
                        waiting = true;
                }
            }
            if (waiting && !AddCaret(highlight, anchors, replacement, false))
                unanchored = true;
            if (unanchored)
                AddPageRail(highlight);
            return highlight;
        }

        private static bool Shows(IList<PdfReviewWord> words, ICollection<PdfReviewWord> shown)
        {
            foreach (PdfReviewWord word in words)
                if (word != null && shown.Contains(word))
                    return true;
            return false;
        }

        private static PdfReviewWord OnPage(PdfReviewWord word, int pageIndex)
        {
            return word != null && word.PageIndex == pageIndex ? word : null;
        }

        private static PdfReviewWord FirstOnPage(IList<PdfReviewWord> words, int pageIndex)
        {
            foreach (PdfReviewWord word in words)
                if (OnPage(word, pageIndex) != null)
                    return word;
            return null;
        }

        /// <summary>
        /// Узкая рамка, начинающаяся в точке каретки: Left — сама точка. Несколько удалений
        /// у одной точки дают одну каретку. false — якоря нет или у него нет рамки.
        /// </summary>
        private static bool AddCaret(PdfReviewHighlight highlight, HashSet<PdfReviewBox> anchors,
            PdfReviewWord neighbour, bool afterNeighbour)
        {
            if (neighbour == null)
                return false;
            double height = neighbour.Box.Top - neighbour.Box.Bottom;
            double x = afterNeighbour ? neighbour.Box.Right : neighbour.Box.Left;
            if (!(height > 0) || double.IsNaN(x) || double.IsInfinity(x))
                return false;
            double width = Math.Max(CaretMinWidthPt, height * CaretWidthShare);
            var caret = new PdfReviewBox
            {
                Left = x,
                Right = x + width,
                Bottom = neighbour.Box.Bottom,
                Top = neighbour.Box.Top
            };
            if (anchors.Add(caret))
                highlight.Boxes.Add(caret);
            return true;
        }

        /// <summary>
        /// Удаление, которому на этой странице не нашлось места: строка у верхнего поля,
        /// целиком левее страницы. Каретки за краем растра DrawDeletedMarkers не рисует, а
        /// рейка «−» её строки остаётся — пометка «здесь удаляли, смотрите «Бок о бок»».
        /// </summary>
        private static void AddPageRail(PdfReviewHighlight highlight)
        {
            double top = highlight.ViewHeightPt -
                Math.Min(RailTopMarginPt, highlight.ViewHeightPt / 4);
            highlight.Boxes.Add(new PdfReviewBox
            {
                Left = -RailLineHeightPt,
                Right = 0,
                Top = top,
                Bottom = Math.Max(0, top - RailLineHeightPt)
            });
        }

        /// <summary>
        /// Пробельные удаления левой страницы в координатах правой. Каждое изменение пробелов
        /// привязано к Exact-соседям обеих сторон (change.Right), поэтому маркер ставится на
        /// правой границе тех же слов; текст, описание и стиль берутся у authoritative маркера.
        /// Маркер сопоставляется с изменением по рамке левой границы в порядке публикации.
        /// </summary>
        internal static List<PdfReviewWhitespaceMarker> DeletedWhitespaceOnLaterPage(
            PdfReviewResult result, PdfReviewPagePair pair)
        {
            var moved = new List<PdfReviewWhitespaceMarker>();
            List<PdfReviewWhitespaceMarker> markers;
            if (result == null || pair == null || pair.LeftPageIndex < 0 ||
                pair.RightPageIndex < 0 ||
                !result.DeletedWhitespaceByPage.TryGetValue(pair.LeftPageIndex, out markers) ||
                markers == null)
                return moved;
            var used = new bool[markers.Count];
            foreach (PdfReviewWhitespaceChange change in result.WhitespaceChanges)
            {
                if (change == null || change.Left == null || change.Right == null ||
                    change.DeletedAtoms.Count == 0 || change.Left.PageIndex != pair.LeftPageIndex)
                    continue;
                int index = FindMarker(markers, used, change.Left.MarkerBox);
                if (index < 0)
                    continue;
                used[index] = true;
                if (change.Right.PageIndex != pair.RightPageIndex)
                    continue;
                PdfReviewWhitespaceMarker source = markers[index];
                moved.Add(new PdfReviewWhitespaceMarker
                {
                    Box = change.Right.MarkerBox,
                    Text = source.Text,
                    AccessibleDescription = source.AccessibleDescription,
                    Style = source.Style
                });
            }
            return moved;
        }

        private static int FindMarker(IList<PdfReviewWhitespaceMarker> markers, bool[] used,
            PdfReviewBox box)
        {
            for (int i = 0; i < markers.Count; i++)
            {
                PdfReviewWhitespaceMarker marker = markers[i];
                if (!used[i] && marker != null && marker.Box.Left == box.Left &&
                    marker.Box.Bottom == box.Bottom && marker.Box.Right == box.Right &&
                    marker.Box.Top == box.Top)
                    return i;
            }
            return -1;
        }

        /// <summary>
        /// Рисует каретки удалений (левый край каждой рамки внутри растра) и рейки «−»
        /// общей раскладкой полос PdfReviewPageView. Пробельные маркеры здесь не рисуются:
        /// они уже в базовой подсветке страницы, второй проход дублировал бы их.
        /// </summary>
        internal static void DrawDeletedMarkers(Bitmap target, PdfReviewHighlight deleted)
        {
            if (target == null || deleted == null || deleted.ViewWidthPt <= 0 ||
                deleted.ViewHeightPt <= 0 || deleted.Boxes.Count == 0)
                return;
            Color color, edge;
            PdfReviewPageView.ResolveHighlightColors(deleted,
                SystemInformation.HighContrast, out color, out edge);
            var carets = new List<RectangleF>();
            foreach (PdfReviewBox box in deleted.Boxes)
            {
                RectangleF rect = PdfReviewGeometry.ToPixelRect(box, deleted.ViewWidthPt,
                    deleted.ViewHeightPt, target.Width, target.Height);
                if (rect.Height >= 1 && rect.Left >= 0 && rect.Left <= target.Width)
                    carets.Add(rect);
            }
            // Раскладка — до Graphics: поиск чистого поля читает пиксели растра.
            List<PdfReviewChangeBar> bars = PdfReviewPageView.LayoutChangeBars(deleted,
                target.Width, target.Height, target);
            if (carets.Count == 0 && bars.Count == 0)
                return;
            using (Graphics graphics = Graphics.FromImage(target))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                // Каретка — той же пары цветов, что рейка: тёмный край и яркая середина.
                using (var outline = new Pen(edge, 4f))
                using (var primary = new Pen(color, 2f))
                    foreach (RectangleF rect in carets)
                    {
                        float x = rect.Left;
                        float top = Math.Max(0f, rect.Top - 2f);
                        float bottom = Math.Min(target.Height - 1f, rect.Bottom + 2f);
                        float serif = Math.Max(3f, rect.Height * 0.2f);
                        foreach (Pen pen in new[] { outline, primary })
                        {
                            graphics.DrawLine(pen, x, top, x, bottom);
                            graphics.DrawLine(pen, x - serif, top, x + serif, top);
                            graphics.DrawLine(pen, x - serif, bottom, x + serif, bottom);
                        }
                    }
                PdfReviewPageView.DrawChangeBars(graphics, target, bars, deleted, edge, color);
            }
        }
    }
}
