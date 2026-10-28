using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    /// <summary>Аудит F2: окно и сервис сравнения PDF (Review).</summary>
    internal static partial class UnitTests
    {
        private const int AuditWheelMessage = 0x020A;

        /// <summary>WM_MOUSEWHEEL как его шлёт Windows: delta и MK_CONTROL в wParam, экранная точка в lParam.</summary>
        private static Message AuditWheel(IntPtr hwnd, Point screen, int delta, bool control)
        {
            int wParam = unchecked((int)(((uint)(delta & 0xFFFF) << 16) | (control ? 0x0008u : 0u)));
            int lParam = unchecked((int)(((uint)(screen.Y & 0xFFFF) << 16) | (uint)(screen.X & 0xFFFF)));
            return Message.Create(hwnd, AuditWheelMessage, new IntPtr(wParam), new IntPtr(lParam));
        }

        /// <summary>Готовый к показу результат на реальных PDF: три одинаковые пары страниц.</summary>
        private static PdfReviewResult AuditLiveResult(TempDir dir)
        {
            string leftPath = dir.File("left.pdf");
            string rightPath = dir.File("right.pdf");
            WriteTextPdf(leftPath, new[] { "Left zero", "Left one", "Left two" });
            WriteTextPdf(rightPath, new[] { "Right zero", "Right one", "Right two" });
            var result = new PdfReviewResult
            {
                Left = ReviewDocument(leftPath, "left zero", "left one", "left two"),
                Right = ReviewDocument(rightPath, "right zero", "right one", "right two")
            };
            for (int i = 0; i < 3; i++)
                result.Pairs.Add(new PdfReviewPagePair { LeftPageIndex = i, RightPageIndex = i });
            PdfReviewDiff.Project(result);
            return result;
        }

        private static bool AuditViewIdle(PdfReviewPageView view)
        {
            return view.ViewState == PdfReviewPageViewState.Ready &&
                !view.RenderWorkerActive;
        }

        [Test("Аудит F2: колесо над чужим окном не листает и не масштабирует Review")]
        private static void AuditReviewWheelIgnoresForeignWindows()
        {
            using (var dir = new TempDir("iwo_f2_wheel_"))
            {
                PdfReviewResult result = AuditLiveResult(dir);
                InIsolatedSettings("iwo_f2_wheel_settings_", delegate
                {
                    using (var form = new PdfReviewForm(null))
                    using (var other = new Form())
                    {
                        form.Show();
                        Application.DoEvents();
                        InvokeReviewApplyResult(form, result);
                        PdfReviewPageView unified = form.UnifiedSourceView;
                        AssertTrue(WaitFor(delegate { return AuditViewIdle(unified); }),
                            "единый canvas отрендерил первую строку");
                        Panel viewport = unified.ViewportPanel;
                        ListBox rows = form.PairsList;
                        Point point = ReviewScreenCenter(viewport);
                        var filter = (IMessageFilter)form;

                        double fit = unified.ZoomScale;
                        Message own = AuditWheel(viewport.Handle, point, 120, true);
                        AssertTrue(filter.PreFilterMessage(ref own),
                            "Ctrl+колесо своей панели (точка и Ctrl из сообщения) обработано");
                        AssertTrue(unified.ZoomScale > fit,
                            "масштаб изменён по координатам lParam: " + fit + " → " + unified.ZoomScale);

                        other.StartPosition = FormStartPosition.Manual;
                        other.ShowInTaskbar = false;
                        other.Bounds = form.Bounds;
                        var otherButton = new Button { Dock = DockStyle.Fill };
                        other.Controls.Add(otherButton);
                        other.Show();
                        Application.DoEvents();
                        double scale = unified.ZoomScale;
                        int row = rows.SelectedIndex;
                        Message foreign = AuditWheel(other.Handle, point, 120, true);
                        AssertTrue(!filter.PreFilterMessage(ref foreign),
                            "колесо другого окна над той же точкой экрана не поглощается Review");
                        Message foreignChild = AuditWheel(otherButton.Handle, point, -120, false);
                        AssertTrue(!filter.PreFilterMessage(ref foreignChild),
                            "колесо дочернего контрола другого окна не поглощается Review");
                        AssertEq(scale, unified.ZoomScale, "масштаб Review не изменился");
                        AssertEq(row, rows.SelectedIndex, "строка сопоставления не переключилась");
                        other.Hide();

                        PdfReviewViewContent content = PdfReviewForm.BuildUnifiedContent(result,
                            result.Pairs[0], 1L);
                        using (var full = new PdfReviewFullScreenForm(content))
                        {
                            full.WindowState = FormWindowState.Normal;
                            full.StartPosition = FormStartPosition.Manual;
                            full.Bounds = form.Bounds;
                            full.Show(form);
                            AssertTrue(WaitFor(delegate { return AuditViewIdle(full.View); }),
                                "полноэкранный canvas отрендерил страницу");
                            Panel fullViewport = full.View.ViewportPanel;
                            Point fullPoint = ReviewScreenCenter(fullViewport);
                            double fullFit = full.View.ZoomScale;
                            scale = unified.ZoomScale;
                            row = rows.SelectedIndex;
                            Message fullWheel = AuditWheel(fullViewport.Handle, fullPoint, 120, true);
                            AssertTrue(!filter.PreFilterMessage(ref fullWheel),
                                "Review не перехватывает колесо полноэкранного окна");
                            AssertTrue(((IMessageFilter)full).PreFilterMessage(ref fullWheel),
                                "полноэкранное окно обрабатывает своё колесо само");
                            AssertTrue(full.View.ZoomScale > fullFit,
                                "масштабировался полноэкранный canvas");
                            AssertEq(scale, unified.ZoomScale, "скрытая панель Review не тронута");
                            AssertEq(row, rows.SelectedIndex, "строка Review не переключилась");
                            full.Close();
                        }
                        form.Close();
                    }
                });
            }
        }

        [Test("Аудит F2: Ctrl+колесо на пределе масштаба поглощается и не прокручивает")]
        private static void AuditReviewCtrlWheelAtZoomLimitIsSwallowed()
        {
            using (var dir = new TempDir("iwo_f2_zoomlimit_"))
            {
                PdfReviewResult result = AuditLiveResult(dir);
                InIsolatedSettings("iwo_f2_zoomlimit_settings_", delegate
                {
                    using (var form = new PdfReviewForm(null))
                    {
                        form.Show();
                        Application.DoEvents();
                        InvokeReviewApplyResult(form, result);
                        PdfReviewPageView unified = form.UnifiedSourceView;
                        AssertTrue(WaitFor(delegate { return AuditViewIdle(unified); }),
                            "страница готова");
                        Point point = ReviewScreenCenter(unified.ViewportPanel);
                        for (int i = 0; i < 64 && unified.ZoomScale < PreviewZoom.Max - 0.000001; i++)
                            form.RouteWheel(point, 120, true);
                        AssertEq(PreviewZoom.Max, unified.ZoomScale, "дошли до максимума");
                        Point offset = unified.ScrollOffset;
                        AssertTrue(form.RouteWheel(point, 120, true),
                            "на 400% Ctrl+колесо поглощено, а не отдано AutoScroll");
                        AssertEq(offset, unified.ScrollOffset, "страница не прокрутилась");
                        AssertEq(PreviewZoom.Max, unified.ZoomScale, "масштаб остался предельным");

                        for (int i = 0; i < 64 && unified.ZoomScale > PreviewZoom.Min + 0.000001; i++)
                            form.RouteWheel(point, -120, true);
                        AssertEq(PreviewZoom.Min, unified.ZoomScale, "дошли до минимума");
                        AssertTrue(form.RouteWheel(point, -120, true),
                            "на 25% Ctrl+колесо тоже поглощено");
                        AssertEq(0, form.PairsList.SelectedIndex,
                            "Ctrl+колесо на пределе не листает строки");
                        form.Close();
                    }
                });
            }
        }

        private static PdfReviewWord AuditBoxWord(string text, int pageIndex, double left,
            double bottom, double right, double top)
        {
            return new PdfReviewWord
            {
                Text = text,
                Key = text,
                PageIndex = pageIndex,
                Box = new PdfReviewBox { Left = left, Bottom = bottom, Right = right, Top = top }
            };
        }

        private static PdfReviewPage AuditBoxPage(int index, params PdfReviewWord[] words)
        {
            var page = new PdfReviewPage
            {
                PageIndex = index,
                WidthPt = 600,
                HeightPt = 800,
                ViewWidthPt = 600,
                ViewHeightPt = 800,
                Text = "",
                NormalizedText = ""
            };
            page.Words.AddRange(words);
            return page;
        }

        /// <summary>Пара страниц, у которой левый номер «1» числится удалённым словом.</summary>
        private static PdfReviewResult AuditPageNumberResult(PdfReviewWord leftNumber,
            params PdfReviewWord[] rightWords)
        {
            var left = new PdfReviewDocument { Path = "left.pdf" };
            var right = new PdfReviewDocument { Path = "right.pdf" };
            left.Pages.Add(AuditBoxPage(0, leftNumber));
            right.Pages.Add(AuditBoxPage(0, rightWords));
            var result = new PdfReviewResult { Left = left, Right = right };
            result.Pairs.Add(new PdfReviewPagePair { LeftPageIndex = 0, RightPageIndex = 0 });
            result.DeletedWordsByPage[0] = new List<PdfReviewWord> { leftNumber };
            return result;
        }

        [Test("Аудит F2: стабильный номер страницы гасится только в нижнем колонтитуле")]
        private static void AuditReviewFooterPageNumberUsesBottomBand()
        {
            // Ось Y рамок — вверх: нижний колонтитул имеет малые Bottom/Top.
            PdfReviewWord footer = AuditBoxWord("1", 0, 295, 20, 305, 32);
            PdfReviewResult result = AuditPageNumberResult(footer,
                AuditBoxWord("1", 0, 295, 20, 305, 32));
            PdfReviewHighlight highlight = PdfReviewForm.BuildHighlight(result,
                result.Pairs[0], true);
            AssertEq(0, highlight.Boxes.Count,
                "номер внизу на том же месте обеих страниц — стабильный колонтитул");

            PdfReviewWord header = AuditBoxWord("1", 0, 295, 770, 305, 782);
            result = AuditPageNumberResult(header, AuditBoxWord("1", 0, 295, 770, 305, 782));
            highlight = PdfReviewForm.BuildHighlight(result, result.Pairs[0], true);
            AssertEq(1, highlight.Boxes.Count,
                "число в верхней части страницы — не нижний колонтитул и остаётся правкой");

            result = AuditPageNumberResult(AuditBoxWord("1", 0, 295, 20, 305, 32),
                AuditBoxWord("1", 0, 40, 20, 50, 32),
                AuditBoxWord("1", 0, 295, 20, 305, 32));
            highlight = PdfReviewForm.BuildHighlight(result, result.Pairs[0], true);
            AssertEq(0, highlight.Boxes.Count,
                "проверяются все кандидаты: совпал по месту второй номер колонтитула");

            result = AuditPageNumberResult(AuditBoxWord("1", 0, 295, 20, 305, 32),
                AuditBoxWord("1", 0, 40, 20, 50, 32));
            highlight = PdfReviewForm.BuildHighlight(result, result.Pairs[0], true);
            AssertEq(1, highlight.Boxes.Count,
                "номер в другом месте колонтитула не считается стабильным");
        }

        private static string AuditPairSignature(IList<PdfReviewPagePair> pairs)
        {
            var parts = new List<string>();
            foreach (PdfReviewPagePair pair in pairs)
                parts.Add(pair.LeftPageIndex + ":" + pair.RightPageIndex + ":" + pair.Status);
            return string.Join("|", parts.ToArray());
        }

        [Test("Аудит F2: ручная пара проецируется на копию и выделяет созданную строку")]
        private static void AuditReviewManualPairIsAtomicAndSelectsCreatedRow()
        {
            PdfReviewDocument left = ReviewDocument("left.pdf", "alpha one", "beta two", "gamma three");
            PdfReviewDocument right = ReviewDocument("right.pdf", "alpha one", "beta two", "gamma three");
            PdfReviewResult original = PdfReviewDiff.Compare(left, right, PdfReviewLimits.Default());
            AssertEq("0:0:Unchanged|1:1:Unchanged|2:2:Unchanged", AuditPairSignature(original.Pairs),
                "исходные строки один к одному");
            var originalPairs = new List<PdfReviewPagePair>(original.Pairs);

            PdfReviewResult paired = PdfReviewForm.ProjectManualPair(original, 2, 1);
            AssertEq("0:0:Unchanged|1:1:Unchanged|2:2:Unchanged", AuditPairSignature(original.Pairs),
                "показанный снимок не меняется до подмены");
            for (int i = 0; i < originalPairs.Count; i++)
                AssertTrue(ReferenceEquals(originalPairs[i], original.Pairs[i]),
                    "строки исходного результата остались теми же объектами: " + i);
            AssertEq(1, PdfReviewForm.FindPairRow(paired.Pairs, 2, 1),
                "созданная пара (3 ↔ 2) стоит на месте первой затронутой строки");
            foreach (PdfReviewPagePair pair in paired.Pairs)
                AssertTrue(!originalPairs.Contains(pair),
                    "копия не делит объекты строк со старым снимком");
            AssertTrue(ReferenceEquals(original.Operations, paired.Operations),
                "семантика не пересчитывается: тот же поток операций");
            AssertThrowsExactly<ArgumentOutOfRangeException>("несуществующая страница отвергнута",
                delegate { PdfReviewForm.ProjectManualPair(original, 9, 0); });
            AssertEq("0:0:Unchanged|1:1:Unchanged|2:2:Unchanged", AuditPairSignature(original.Pairs),
                "отказ не оставляет полуизменённый снимок");

            using (var dir = new TempDir("iwo_f2_manual_"))
            {
                PdfReviewResult live = AuditLiveResult(dir);
                InIsolatedSettings("iwo_f2_manual_settings_", delegate
                {
                    using (var form = new PdfReviewForm(null))
                    {
                        form.Show();
                        Application.DoEvents();
                        InvokeReviewApplyResult(form, live);
                        form.ApplyManualPair(2, 1);
                        ListBox rows = form.PairsList;
                        PdfReviewResult shown = form.ReviewResult;
                        AssertTrue(!ReferenceEquals(live, shown), "окно подменило снимок целиком");
                        AssertEq(PdfReviewForm.FindPairRow(shown.Pairs, 2, 1), rows.SelectedIndex,
                            "выделена созданная строка, а не первая");
                        AssertEq(1, rows.SelectedIndex, "созданная строка — вторая");
                        PdfReviewPageView unified = form.UnifiedSourceView;
                        AssertTrue(WaitFor(delegate { return AuditViewIdle(unified); }),
                            "созданная строка отрисована");
                        form.Close();
                    }
                });
            }
        }

        [Test("Аудит F2: исчерпанный бюджет сравнения показан предупреждением в окне")]
        private static void AuditReviewWorkExhaustedWarning()
        {
            var result = new PdfReviewResult
            {
                Left = ReviewDocument("left.pdf", "a"),
                Right = ReviewDocument("right.pdf", "b"),
                Stats = new PdfReviewStats { ChangedPages = 1 }
            };
            string stats = PdfReviewForm.SummaryText(result);
            AssertTrue(stats.Length > 0 && !stats.Contains(Loc.T("review.warn.workExhausted")),
                "полное сравнение показывает только статистику");
            result.WorkExhausted = true;
            string warned = PdfReviewForm.SummaryText(result);
            AssertTrue(warned.StartsWith(Loc.T("review.warn.workExhausted"), StringComparison.Ordinal),
                "предупреждение стоит первым: " + warned);
            AssertTrue(warned.EndsWith(stats, StringComparison.Ordinal), "статистика сохранена");

            Lang before = Loc.Current;
            try
            {
                Loc.Init(Lang.Ru);
                string ru = Loc.T("review.warn.workExhausted");
                Loc.Init(Lang.En);
                string en = Loc.T("review.warn.workExhausted");
                AssertTrue(ru.Contains("грубо") && en.Contains("coarsely") && ru != en,
                    "ключ переведён на оба языка: «" + ru + "» / «" + en + "»");
            }
            finally { Loc.Init(before); }

            using (var dir = new TempDir("iwo_f2_exhausted_"))
            {
                PdfReviewResult live = AuditLiveResult(dir);
                live.WorkExhausted = true;
                InIsolatedSettings("iwo_f2_exhausted_settings_", delegate
                {
                    using (var form = new PdfReviewForm(null))
                    {
                        form.Show();
                        Application.DoEvents();
                        InvokeReviewApplyResult(form, live);
                        Label summary = form.SummaryLabel;
                        AssertTrue(summary.Text.StartsWith(Loc.T("review.warn.workExhausted"),
                            StringComparison.Ordinal), "окно показывает предупреждение: " + summary.Text);
                        AssertEq(Theme.WarnOrange.ToArgb(), summary.ForeColor.ToArgb(),
                            "предупреждение выделено цветом предупреждения");
                        AssertEq(summary.Text, summary.AccessibleDescription,
                            "предупреждение доступно экранному диктору");
                        AssertTrue(WaitFor(delegate
                        {
                            return AuditViewIdle(form.UnifiedSourceView);
                        }), "страница отрисована до закрытия");
                        form.Close();
                    }
                });
            }
        }

        [Test("Аудит F2: итог и сообщение панели — живые области, ошибка перехода не стирает предупреждение")]
        private static void AuditReviewSummaryIsLiveAndKeepsWarning()
        {
            string error = Loc.T("review.page.err.number");
            var complete = new PdfReviewResult();
            AssertEq(error, PdfReviewForm.NavigationErrorText(error, complete),
                "без грубых мест ошибка перехода показывается как есть");
            AssertEq(error, PdfReviewForm.NavigationErrorText(error, null), "результата нет");
            using (var dir = new TempDir("iwo_f2_live_"))
            {
                PdfReviewResult live = AuditLiveResult(dir);
                live.WorkExhausted = true;
                InIsolatedSettings("iwo_f2_live_settings_", delegate
                {
                    using (var form = new PdfReviewForm(null))
                    {
                        form.Show();
                        Application.DoEvents();
                        InvokeReviewApplyResult(form, live);
                        Label summary = form.SummaryLabel;
                        AssertEq(System.Windows.Forms.Automation.AutomationLiveSetting.Polite,
                            summary.LiveSetting, "итог сравнения объявляется диктором при смене");
                        // Не живая область бросила бы InvalidOperationException.
                        summary.AccessibilityObject.RaiseLiveRegionChanged();
                        PdfReviewPageView unified = form.UnifiedSourceView;
                        AssertEq(System.Windows.Forms.Automation.AutomationLiveSetting.Polite,
                            unified.MessageLabel.LiveSetting,
                            "«недоступно» и «слишком большой» панели тоже объявляются");

                        TextBox input = form.LeftPageBox;
                        string warning = Loc.T("review.warn.workExhausted");
                        input.Text = "abc";
                        AssertTrue(!form.NavigateToPhysicalPage(true), "нечисловой номер отвергнут");
                        AssertTrue(summary.Text.StartsWith(error, StringComparison.Ordinal) &&
                                   summary.Text.Contains(warning),
                            "ошибка первой, предупреждение о грубом сравнении осталось: " + summary.Text);
                        input.Text = "99";
                        AssertTrue(!form.NavigateToPhysicalPage(true), "номер вне диапазона отвергнут");
                        AssertTrue(summary.Text.Contains(warning),
                            "и после ошибки диапазона предупреждение видно: " + summary.Text);
                        input.Text = "2";
                        AssertTrue(form.NavigateToPhysicalPage(true), "верный номер принят");
                        AssertTrue(summary.Text.StartsWith(warning, StringComparison.Ordinal),
                            "удачный переход возвращает итог с предупреждением первым: " + summary.Text);
                        AssertTrue(WaitFor(delegate { return AuditViewIdle(unified); }),
                            "страница отрисована до закрытия");
                        form.Close();
                    }
                });
            }
        }

        private static PdfReviewPage AuditTextPage(int index, params PdfReviewWord[] words)
        {
            PdfReviewPage page = AuditBoxPage(index, words);
            var text = new List<string>();
            foreach (PdfReviewWord word in words)
                text.Add(word.Text);
            page.Text = string.Join(" ", text.ToArray());
            page.NormalizedText = PdfReviewDiff.Normalize(page.Text);
            page.Fingerprint = PdfReviewDiff.Fingerprint(page.NormalizedText, page.WidthPt,
                page.HeightPt);
            return page;
        }

        /// <summary>
        /// Перевёрстка: удалённое слово «removed» стояло слева там, где справа теперь стоит
        /// неизменённое «omega». Страница 600×800 pt рендерится 1:1 в пиксели.
        /// </summary>
        private static PdfReviewResult AuditReflowedDeletion()
        {
            var left = new PdfReviewDocument { Path = "left.pdf" };
            var right = new PdfReviewDocument { Path = "right.pdf" };
            left.Pages.Add(AuditTextPage(0,
                AuditBoxWord("alpha", 0, 100, 700, 150, 712),
                AuditBoxWord("removed", 0, 200, 700, 260, 712),
                AuditBoxWord("omega", 0, 300, 700, 350, 712)));
            right.Pages.Add(AuditTextPage(0,
                AuditBoxWord("alpha", 0, 100, 700, 150, 712),
                AuditBoxWord("omega", 0, 200, 700, 260, 712)));
            return PdfReviewDiff.Compare(left, right, PdfReviewLimits.Default());
        }

        private static Bitmap AuditUnifiedComposite(PdfReviewViewContent content)
        {
            var page = new Bitmap(600, 800);
            using (Graphics graphics = Graphics.FromImage(page))
                graphics.Clear(Color.White);
            // Тот же порядок, что в рендер-воркере: базовая подсветка, затем слой удалений.
            Bitmap composed = PdfReviewPageView.DrawHighlight(page, content.BaseHighlight, false);
            if (content.IsComposite)
                PdfReviewUnifiedRenderer.DrawDeletedMarkers(composed, content.OverlayHighlight);
            return composed;
        }

        private static bool AuditRedPixel(Bitmap bitmap, int left, int right, int top, int bottom)
        {
            for (int y = Math.Max(0, top); y <= Math.Min(bitmap.Height - 1, bottom); y++)
                for (int x = Math.Max(0, left); x <= Math.Min(bitmap.Width - 1, right); x++)
                {
                    Color pixel = bitmap.GetPixel(x, y);
                    if (pixel.R > 150 && pixel.G < 120 && pixel.B < 120)
                        return true;
                }
            return false;
        }

        [Test("Аудит F2: единый режим не рисует рамки левой страницы на правом растре")]
        private static void AuditReviewUnifiedKeepsLeftGeometryOffRightRaster()
        {
            PdfReviewResult result = AuditReflowedDeletion();
            PdfReviewViewContent content = PdfReviewForm.BuildUnifiedContent(result,
                result.Pairs[0], 1L);
            AssertTrue(content != null && content.IsComposite,
                "удаление на странице даёт слой пометок удаления");
            using (Bitmap composed = AuditUnifiedComposite(content))
            {
                // «omega» справа (x 200..260, y 88..100 px) не изменилось: на нём не должно
                // быть ни рамки, ни перечёркивания удалённого «removed» из левой страницы.
                AssertTrue(!AuditRedPixel(composed, 196, 264, 84, 104),
                    "неизменённое слово правой страницы осталось без красной рамки");
                AssertTrue(AuditRedPixel(composed, 148, 153, 86, 102),
                    "каретка удаления стоит на правом краю Exact-соседа «alpha» (x=150)");
                AssertTrue(AuditRedPixel(composed, 0, 30, 82, 106),
                    "у левого поля той же строки стоит рейка «−»");
            }
        }

        /// <summary>
        /// Три страницы: первая и последняя совпадают, средняя переписана целиком
        /// (rightMiddle — её слова справа, пусто — пустая страница). Пара 2 ↔ 2 — ручная,
        /// если выравнивание её не составило само.
        /// </summary>
        private static PdfReviewResult AuditRewrittenMiddlePage(out PdfReviewPagePair pair,
            params PdfReviewWord[] rightMiddle)
        {
            var left = new PdfReviewDocument { Path = "left.pdf" };
            var right = new PdfReviewDocument { Path = "right.pdf" };
            foreach (PdfReviewDocument document in new[] { left, right })
                document.Pages.Add(AuditTextPage(0,
                    AuditBoxWord("alpha", 0, 100, 700, 150, 712),
                    AuditBoxWord("one", 0, 160, 700, 190, 712)));
            left.Pages.Add(AuditTextPage(1,
                AuditBoxWord("removed", 1, 100, 700, 160, 712),
                AuditBoxWord("text", 1, 170, 700, 200, 712),
                AuditBoxWord("here", 1, 210, 700, 240, 712)));
            right.Pages.Add(AuditTextPage(1, rightMiddle));
            foreach (PdfReviewDocument document in new[] { left, right })
                document.Pages.Add(AuditTextPage(2,
                    AuditBoxWord("gamma", 2, 100, 700, 150, 712),
                    AuditBoxWord("three", 2, 160, 700, 190, 712)));
            PdfReviewResult result = PdfReviewDiff.Compare(left, right, PdfReviewLimits.Default());
            if (PdfReviewForm.FindPairRow(result.Pairs, 1, 1) < 0)
                result = PdfReviewForm.ProjectManualPair(result, 1, 1);
            pair = result.Pairs[PdfReviewForm.FindPairRow(result.Pairs, 1, 1)];
            return result;
        }

        [Test("Аудит F2: единый режим не прячет удаление без контрагента на правой странице пары")]
        private static void AuditReviewUnifiedDeletionWithoutCounterpartIsVisible()
        {
            PdfReviewPagePair pair;
            PdfReviewResult result = AuditRewrittenMiddlePage(out pair,
                AuditBoxWord("inserted", 1, 300, 700, 360, 712),
                AuditBoxWord("words", 1, 370, 700, 400, 712),
                AuditBoxWord("only", 1, 410, 700, 440, 712));
            AssertEq(3, PdfReviewForm.BuildHighlight(result, pair, true).Boxes.Count,
                "«Бок о бок» показывает все три удалённых слова");
            PdfReviewViewContent content = PdfReviewForm.BuildUnifiedContent(result, pair, 1L);
            AssertTrue(content != null && content.IsComposite, "у пары есть слой удалений");
            using (Bitmap composed = AuditUnifiedComposite(content))
            {
                // «inserted» справа: x 300..360, y 88..100 px. Удалённый текст заменён им —
                // каретка встаёт у его левого края, а не исчезает.
                AssertTrue(AuditRedPixel(composed, 297, 303, 86, 102),
                    "каретка удаления стоит у первого вставленного слова замены");
                AssertTrue(AuditRedPixel(composed, 0, 30, 82, 106),
                    "у левого поля той же строки — рейка «−»");
            }

            result = AuditRewrittenMiddlePage(out pair);
            content = PdfReviewForm.BuildUnifiedContent(result, pair, 2L);
            AssertTrue(content != null && content.IsComposite,
                "пустая правая страница: удаление всё равно попадает на холст");
            using (Bitmap composed = AuditUnifiedComposite(content))
            {
                AssertTrue(AuditRedPixel(composed, 0, 30, 0, 120),
                    "места нет — рейка «−» у верхнего поля страницы");
                AssertTrue(!AuditRedPixel(composed, 40, 599, 0, 799),
                    "без якоря каретка не рисуется в случайной точке страницы");
            }
        }

        [Test("Аудит F2: каретка удаления встаёт у ближайшего контрагента, даже не Exact")]
        private static void AuditReviewUnifiedCaretUsesNearestCounterpart()
        {
            // Строка 1: «far» (Exact). Строка 2: «near» — склейка (SplitJoin) из «ne ar»,
            // сразу за ней удалено «gone». Каретка должна встать за «near» (x=160, строка 2),
            // а не за далёким Exact-словом другой строки.
            PdfReviewWord far = AuditBoxWord("far", 0, 100, 740, 150, 752);
            PdfReviewWord near = AuditBoxWord("near", 0, 100, 700, 160, 712);
            PdfReviewWord gone = AuditBoxWord("gone", 0, 170, 700, 200, 712);
            var result = new PdfReviewResult
            {
                Left = new PdfReviewDocument { Path = "left.pdf" },
                Right = new PdfReviewDocument { Path = "right.pdf" }
            };
            result.Left.Pages.Add(AuditBoxPage(0, far, gone));
            result.Right.Pages.Add(AuditBoxPage(0, far, near));
            var exact = new PdfReviewWordOp
                { Kind = PdfReviewDiffKind.Equal, MatchKind = PdfReviewMatchKind.Exact };
            exact.LeftWords.Add(far);
            exact.RightWords.Add(far);
            exact.Matches.Add(new PdfReviewWordMatch
                { Left = far, Right = far, Kind = PdfReviewMatchKind.Exact });
            var joined = new PdfReviewWordOp
                { Kind = PdfReviewDiffKind.Equal, MatchKind = PdfReviewMatchKind.SplitJoin };
            joined.LeftWords.Add(AuditBoxWord("ne", 0, 100, 700, 125, 712));
            joined.LeftWords.Add(AuditBoxWord("ar", 0, 130, 700, 160, 712));
            joined.RightWords.Add(near);
            var deletion = new PdfReviewWordOp { Kind = PdfReviewDiffKind.Delete };
            deletion.LeftWords.Add(gone);
            result.Operations.Add(exact);
            result.Operations.Add(joined);
            result.Operations.Add(deletion);
            var pair = new PdfReviewPagePair { LeftPageIndex = 0, RightPageIndex = 0 };

            PdfReviewHighlight anchors = PdfReviewUnifiedRenderer.DeletionAnchors(result, pair,
                new HashSet<PdfReviewWord> { gone });
            AssertEq(1, anchors.Boxes.Count, "одно удаление — одна каретка");
            AssertTrue(anchors.Boxes[0].Left == near.Box.Right &&
                       anchors.Boxes[0].Bottom == near.Box.Bottom,
                "каретка за склеенным «near» на той же строке: Left=" + anchors.Boxes[0].Left +
                " Bottom=" + anchors.Boxes[0].Bottom);
        }

        [Test("Аудит F2: пробельное удаление в едином режиме — один маркер в координатах правой страницы")]
        private static void AuditReviewUnifiedWhitespaceDeletionDrawnOnce()
        {
            PdfReviewResult result = AuditReflowedDeletion();
            var leftBoundary = new PdfReviewBox { Left = 150, Bottom = 700, Right = 200, Top = 712 };
            var rightBoundary = new PdfReviewBox { Left = 150, Bottom = 650, Right = 200, Top = 662 };
            var marker = new PdfReviewWhitespaceMarker
            {
                Box = leftBoundary,
                Text = "− ␠",
                AccessibleDescription = "removed space",
                Style = PdfReviewHighlightStyle.Removed
            };
            var change = new PdfReviewWhitespaceChange
            {
                Left = new PdfReviewWhitespaceEvidence { PageIndex = 0, MarkerBox = leftBoundary },
                Right = new PdfReviewWhitespaceEvidence { PageIndex = 0, MarkerBox = rightBoundary }
            };
            change.DeletedAtoms.Add(new PdfReviewWhitespaceAtom
                { Kind = PdfReviewWhitespaceAtomKind.Space, RawText = " " });
            result.ReplaceWhitespaceChanges(new List<PdfReviewWhitespaceChange> { change });
            result.DeletedWhitespaceByPage[0] = new List<PdfReviewWhitespaceMarker> { marker };

            PdfReviewViewContent content = PdfReviewForm.BuildUnifiedContent(result,
                result.Pairs[0], 2L);
            int removedMarkers = 0;
            PdfReviewWhitespaceMarker shown = null;
            foreach (PdfReviewWhitespaceMarker item in content.BaseHighlight.WhitespaceMarkers)
                if (item.Style == PdfReviewHighlightStyle.Removed)
                {
                    removedMarkers++;
                    shown = item;
                }
            AssertEq(1, removedMarkers, "удаление пробела попадает на холст ровно один раз");
            AssertTrue(content.OverlayHighlight == null ||
                       content.OverlayHighlight.WhitespaceMarkers.Count == 0,
                "слой удалений слов не рисует пробельные маркеры второй раз");
            AssertTrue(shown.Box.Bottom == rightBoundary.Bottom && shown.Box.Top == rightBoundary.Top,
                "маркер стоит на правой границе тех же слов, а не на левой");
            AssertEq(marker.Text, shown.Text, "текст маркера — authoritative");
        }

        /// <summary>
        /// Строки растра в полосе x ∈ [left, right] с пикселями цвета маркера. Сравнение — с
        /// самим цветом, а не «краснее/зеленее»: цветные края сглаживания текста иначе
        /// засчитывались бы обоим маркерам.
        /// </summary>
        private static List<int> AuditMarkerRows(Bitmap bitmap, int left, int right, Color ink)
        {
            var rows = new List<int>();
            for (int y = 0; y < bitmap.Height; y++)
                for (int x = Math.Max(0, left); x <= Math.Min(bitmap.Width - 1, right); x++)
                {
                    Color pixel = bitmap.GetPixel(x, y);
                    if (Math.Abs(pixel.R - ink.R) + Math.Abs(pixel.G - ink.G) +
                        Math.Abs(pixel.B - ink.B) < 60)
                    {
                        rows.Add(y);
                        break;
                    }
                }
            return rows;
        }

        [Test("Аудит F2: замена пробела в едином режиме — «−» и «+» не накладываются")]
        private static void AuditReviewUnifiedWhitespaceReplacementMarkersApart()
        {
            PdfReviewResult result = ReviewWhitespaceComparison("\t", " ");
            PdfReviewViewContent content = PdfReviewForm.BuildUnifiedContent(result,
                result.Pairs[0], 1L);
            PdfReviewWhitespaceMarker added = null, removed = null;
            foreach (PdfReviewWhitespaceMarker marker in content.BaseHighlight.WhitespaceMarkers)
                if (marker.Style == PdfReviewHighlightStyle.Added) added = marker;
                else removed = marker;
            AssertTrue(added != null && removed != null, "на холсте оба маркера замены");
            AssertTrue(added.Box.Left == removed.Box.Left && added.Box.Top == removed.Box.Top,
                "оба привязаны к одной границе слов правой страницы");

            // Страница 200×200 pt в масштабе 1: маркеры (≈16 px) выше рамки границы (10 px).
            var page = new Bitmap(200, 200);
            using (Graphics graphics = Graphics.FromImage(page))
                graphics.Clear(Color.White);
            using (Bitmap drawn = PdfReviewPageView.DrawHighlight(page, content.BaseHighlight, false))
            {
                int left = (int)removed.Box.Left - 30, right = (int)removed.Box.Right + 30;
                List<int> red = AuditMarkerRows(drawn, left, right, Theme.ReviewDeleteMarker);
                List<int> teal = AuditMarkerRows(drawn, left, right, Theme.ReviewInsertMarker);
                AssertTrue(red.Count > 0 && teal.Count > 0, "оба маркера нарисованы");
                bool apart = red[red.Count - 1] < teal[0] || teal[teal.Count - 1] < red[0];
                AssertTrue(apart, "«−» (строки " + red[0] + "–" + red[red.Count - 1] +
                    ") и «+» (строки " + teal[0] + "–" + teal[teal.Count - 1] +
                    ") разнесены по вертикали, а не наложены");
            }

            // У верхнего края места над маркером нет — второй уходит под первый.
            var top = new PdfReviewHighlight
            {
                ViewWidthPt = 200,
                ViewHeightPt = 200,
                Style = PdfReviewHighlightStyle.Added,
                ChangeBarSide = PdfReviewChangeBarSide.None
            };
            var edge = new PdfReviewBox { Left = 39.5, Bottom = 190, Right = 40.5, Top = 200 };
            top.WhitespaceMarkers.Add(new PdfReviewWhitespaceMarker
                { Box = edge, Text = added.Text, Style = PdfReviewHighlightStyle.Added });
            top.WhitespaceMarkers.Add(new PdfReviewWhitespaceMarker
                { Box = edge, Text = removed.Text, Style = PdfReviewHighlightStyle.Removed });
            var blank = new Bitmap(200, 200);
            using (Graphics graphics = Graphics.FromImage(blank))
                graphics.Clear(Color.White);
            using (Bitmap drawn = PdfReviewPageView.DrawHighlight(blank, top, false))
            {
                List<int> red = AuditMarkerRows(drawn, 0, 100, Theme.ReviewDeleteMarker);
                List<int> teal = AuditMarkerRows(drawn, 0, 100, Theme.ReviewInsertMarker);
                AssertTrue(red.Count > 0 && teal.Count > 0 && teal[teal.Count - 1] < red[0],
                    "у верхнего края «−» встал под «+», оба видны: красные " + red.Count +
                    " строк, бирюзовые " + teal.Count);
            }
        }

        [Test("Аудит F2: сбой композиции подсветки не публикует страницу без пометок")]
        private static void AuditReviewCompositionFailureIsUnavailable()
        {
            var highlight = new PdfReviewHighlight
            {
                ViewWidthPt = 20,
                ViewHeightPt = 12,
                Color = Theme.ReviewInsertFill,
                EdgeColor = Theme.ReviewInsertMarker,
                Style = PdfReviewHighlightStyle.Added,
                ChangeBarSide = PdfReviewChangeBarSide.Right
            };
            highlight.Boxes.Add(new PdfReviewBox { Left = 2, Bottom = 4, Right = 8, Top = 10 });
            using (var source = new Bitmap(20, 12))
            {
                // Заблокированные биты: размеры известны, а рисование источника GDI+ отвергает
                // («Bitmap region is already locked») — детерминированный сбой композиции.
                System.Drawing.Imaging.BitmapData locked = source.LockBits(
                    new Rectangle(0, 0, source.Width, source.Height),
                    System.Drawing.Imaging.ImageLockMode.ReadWrite,
                    System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                Bitmap result;
                List<PdfReviewChangeBar> bars;
                try
                {
                    result = PdfReviewPageView.DrawHighlight(source, highlight, false, out bars);
                }
                finally { source.UnlockBits(locked); }
                AssertTrue(result == null,
                    "сбой GDI+ возвращает null («недоступно»), а не исходную страницу без подсветки");
                AssertEq(0, bars.Count, "полосы несостоявшейся композиции не публикуются");
                source.GetPixel(0, 0); // источник не освобождён: им по-прежнему владеет вызывающий
            }
        }

        /// <summary>Две колонки: «L1a» (слева, строка 1), «L2» (слева, строка 2), «R1» (справа, строка 1).</summary>
        private static PdfReviewPage AuditTwoColumnPage()
        {
            PdfReviewPage page = ReviewTrustedPage(0, new[] { "L1a", "L2", "R1" },
                new[] { " ", " " });
            page.Words[0].Box = new PdfReviewBox { Left = 5, Bottom = 70, Right = 20, Top = 80 };
            page.Words[1].Box = new PdfReviewBox { Left = 5, Bottom = 50, Right = 20, Top = 60 };
            page.Words[2].Box = new PdfReviewBox { Left = 60, Bottom = 70, Right = 80, Top = 80 };
            return page;
        }

        private static PdfReviewHighlight AuditRailHighlight(PdfReviewPage page, params int[] changed)
        {
            var highlight = new PdfReviewHighlight
            {
                ViewWidthPt = page.ViewWidthPt,
                ViewHeightPt = page.ViewHeightPt,
                Color = Theme.ReviewDeleteFill,
                EdgeColor = Theme.ReviewDeleteMarker,
                ChangeBarSide = PdfReviewChangeBarSide.Left,
                Style = PdfReviewHighlightStyle.Removed
            };
            foreach (int index in changed)
            {
                highlight.Words.Add(page.Words[index]);
                highlight.Boxes.Add(page.Words[index].Box);
            }
            return highlight;
        }

        [Test("Аудит F2: клик по полосе двух колонок выбирает только изменённые слова строки")]
        private static void AuditReviewChangeBarSelectsOnlyLineWords()
        {
            RunSta(delegate
            {
                using (var host = new Form())
                using (var image = new Bitmap(200, 100))
                {
                    using (Graphics graphics = Graphics.FromImage(image))
                        graphics.Clear(Color.White);
                    host.ClientSize = new Size(220, 120);
                    var surface = new PdfReviewPageSurface
                    {
                        Location = new Point(10, 10),
                        Size = new Size(200, 100),
                        Image = image,
                        SizeMode = PictureBoxSizeMode.Zoom
                    };
                    host.Controls.Add(surface);
                    PdfReviewPage page = AuditTwoColumnPage();
                    surface.AttachPage(page, AuditRailHighlight(page, 0, 2));
                    host.Show();
                    Application.DoEvents();

                    InvokeReviewSurfaceMouse(surface, "OnMouseDown",
                        new MouseEventArgs(MouseButtons.Left, 1, 2, 25, 0));
                    AssertEq(2, surface.SelectedWordCount,
                        "выбраны «L1a» и «R1», а не диапазон чтения через «L2» другой строки");
                    PdfReviewCopyText copy = surface.SelectionModel.BuildCopyText();
                    AssertEq("L1a R1", copy.Text, "копируются только слова полосы");
                    AssertTrue(copy.UsedFallbackSeparator,
                        "между колонками нет доказанного разделителя — честный fallback-пробел");
                    AssertTrue(!surface.SelectionModel.IsSelected(1),
                        "слово другой строки не выделено");
                    host.Close();
                }
            });
        }

        [Test("Аудит F2: попадание по полосе — там, где она нарисована, и с учётом масштаба")]
        private static void AuditReviewChangeBarHitFollowsDrawnRail()
        {
            RunSta(delegate
            {
                PdfReviewPage page = ReviewTrustedPage(0, new[] { "one", "two", "three" },
                    new[] { " ", " " });
                // Слова правее обычного: поле у края занято тёмной заливкой страницы.
                for (int i = 0; i < page.Words.Count; i++)
                    page.Words[i].Box = new PdfReviewBox
                        { Left = 30 + i * 12, Bottom = 70, Right = 38 + i * 12, Top = 80 };
                PdfReviewHighlight highlight = AuditRailHighlight(page, 0, 2);
                var source = new Bitmap(200, 100);
                using (Graphics graphics = Graphics.FromImage(source))
                {
                    graphics.Clear(Color.White);
                    using (var ink = new SolidBrush(Color.FromArgb(40, 40, 40)))
                        graphics.FillRectangle(ink, 0, 0, 11, 100);
                }
                List<PdfReviewChangeBar> bars;
                Bitmap drawn = PdfReviewPageView.DrawHighlight(source, highlight, false, out bars);
                using (var host = new Form())
                using (drawn)
                {
                    AssertEq(1, bars.Count, "одна строка — одна полоса");
                    float railX = bars[0].X;
                    AssertTrue(railX >= 14, "рейка обошла тёмное поле: X=" + railX);
                    host.ClientSize = new Size(440, 240);
                    var surface = new PdfReviewPageSurface
                    {
                        Location = new Point(10, 10),
                        Size = new Size(400, 200),
                        Image = drawn,
                        SizeMode = PictureBoxSizeMode.Zoom
                    };
                    host.Controls.Add(surface);
                    surface.AttachPage(page, highlight, bars);
                    host.Show();
                    Application.DoEvents();

                    // Масштаб 2: рейка растра X → X·2 на поверхности.
                    var onRail = new Point((int)Math.Round(railX * 2), 50);
                    AssertTrue(surface.HitTestChangeBar(onRail) >= 0,
                        "клик по нарисованной рейке попадает в полосу");
                    AssertEq(-1, surface.HitTestWord(onRail, false),
                        "клик по рейке — не выбор слова");
                    var oldZone = new Point(4, 50);
                    AssertEq(-1, surface.HitTestChangeBar(oldZone),
                        "прежняя фиксированная зона 24 px у края, где рейки нет, не срабатывает");
                    int wordX = (int)Math.Round((page.Words[1].Box.Left + 2) * 4);
                    AssertEq(1, surface.HitTestWord(new Point(wordX, 50), false),
                        "реальный hit-test поверхности находит слово под точкой");
                    AssertEq(-1, surface.HitTestChangeBar(new Point(wordX, 50)),
                        "точка над словом не считается полосой");

                    InvokeReviewSurfaceMouse(surface, "OnMouseDown",
                        new MouseEventArgs(MouseButtons.Left, 1, onRail.X, onRail.Y, 0));
                    AssertEq(3, surface.SelectedWordCount,
                        "полоса выбирает изменённые слова строки и неизменённое между ними");
                    InvokeReviewSurfaceMouse(surface, "OnMouseDown",
                        new MouseEventArgs(MouseButtons.Left, 1, oldZone.X, oldZone.Y, 0));
                    AssertTrue(!surface.HasSelection, "клик мимо рейки и слов снимает выделение");
                    host.Close();
                }
            });
        }

        [Test("Аудит F2: поиск чистого поля под рейку читает растр одним LockBits")]
        private static void AuditReviewRailClearanceIsFast()
        {
            const int width = 2400, height = 3400, lines = 76, step = 44, lineHeight = 30;
            var highlight = new PdfReviewHighlight
            {
                ViewWidthPt = width,
                ViewHeightPt = height,
                Color = Theme.ReviewDeleteFill,
                EdgeColor = Theme.ReviewDeleteMarker,
                ChangeBarSide = PdfReviewChangeBarSide.Left,
                Style = PdfReviewHighlightStyle.Removed
            };
            var source = new Bitmap(width, height);
            using (Graphics graphics = Graphics.FromImage(source))
            {
                graphics.Clear(Color.White);
                using (var ink = new SolidBrush(Color.Black))
                    for (int i = 0; i < lines; i++)
                    {
                        int top = 20 + i * step;
                        highlight.Boxes.Add(new PdfReviewBox
                        {
                            Left = 400, Right = 460,
                            Top = height - top, Bottom = height - top - lineHeight
                        });
                        // Тонкая линия у нижнего края полосы рейки, ниже самой строки: каждый
                        // кандидат X приходится просматривать почти целиком — худший случай.
                        graphics.FillRectangle(ink, 0, top + lineHeight + 4, 120, 1);
                    }
            }
            using (source)
            {
                // Детерминированные меры вместо секундомера (TS-5). Прежний GetPixel-обход —
                // ~0,9 млн вызовов GDI+ на раскладку; одно чтение LockBits — полоса у края листа.
                // Выделения: полоса одна на раскладку (замер 28.10.2026: 1,21 МБ на x64 и x86,
                // 88 px × 3400 строк × 4 байта); потолок — две полосы. Копия всего листа (~32 МБ)
                // или полоса на каждую рейку выходят за него.
                List<PdfReviewChangeBar> bars = null;
                long allocated = AllocatedBytes(delegate
                {
                    bars = PdfReviewPageView.LayoutChangeBars(highlight, width, height, source);
                });
                AssertEq(lines, bars.Count, "каждая строка получила полосу");
                foreach (PdfReviewChangeBar bar in bars)
                    AssertEq(12f, bar.X,
                        "все кандидаты заняты чернилами — рейка на предпочтительном отступе");
                AssertTrue(allocated < 2500000, "раскладка выделила " + allocated +
                    " байт: растр читается не одной полосой у края");
                // Попиксельное чтение следа в памяти не оставляет — его запрещает правило по коду.
                foreach (string name in new[] { "GetPixel", "SetPixel" })
                {
                    List<string> callers = CallersOf(typeof(PdfReviewPageView), typeof(Bitmap).GetMethod(name));
                    AssertEq(0, callers.Count, "PdfReviewPageView читает пиксели только через LockBits, а " +
                        name + " зовут: " + string.Join(", ", callers.ToArray()));
                }

                var clear = new Bitmap(width, height);
                using (Graphics graphics = Graphics.FromImage(clear))
                {
                    graphics.Clear(Color.White);
                    using (var ink = new SolidBrush(Color.Black))
                        graphics.FillRectangle(ink, 0, 0, 30, height);
                }
                using (clear)
                    AssertEq(33f, PdfReviewPageView.LayoutChangeBars(highlight, width, height,
                        clear)[0].X, "рейка встаёт в первую чистую полосу за тёмным полем");
            }
        }

        /// <summary>Отдельный вид в окне-хосте (как панель Review) с реальным PDF.</summary>
        private static void AuditWithView(Action<PdfReviewPageView> body)
        {
            RunSta(delegate
            {
                using (var host = new Form())
                {
                    host.ClientSize = new Size(640, 480);
                    var view = new PdfReviewPageView { Dock = DockStyle.Fill };
                    host.Controls.Add(view);
                    host.Show();
                    Application.DoEvents();
                    body(view);
                    host.Close();
                }
            });
        }

        private static bool AuditViewSettled(PdfReviewPageView view, PdfReviewPageViewState state)
        {
            return view.ViewState == state && !view.RenderWorkerActive;
        }

        [Test("Аудит F2: панели окна делят один рендерер, документ не перечитывается на каждый показ")]
        private static void AuditReviewPanelsShareOneRenderer()
        {
            using (var dir = new TempDir("iwo_f2_session_"))
            {
                PdfReviewResult result = AuditLiveResult(dir);
                InIsolatedSettings("iwo_f2_session_settings_", delegate
                {
                    using (var form = new PdfReviewForm(null))
                    {
                        form.Show();
                        form.SideModeButton.PerformClick();
                        Application.DoEvents();
                        PdfReviewPageView left = form.LeftSourceView;
                        PdfReviewPageView right = form.RightSourceView;
                        PdfReviewPageView unified = form.UnifiedSourceView;
                        PdfReviewRenderSession session = form.RenderSession;
                        AssertTrue(ReferenceEquals(session, left.RenderSession) &&
                                   ReferenceEquals(session, right.RenderSession) &&
                                   ReferenceEquals(session, unified.RenderSession),
                            "все панели окна рендерят через один рендерер");
                        InvokeReviewApplyResult(form, result);
                        ListBox rows = form.PairsList;
                        for (int row = 0; row < 3; row++)
                        {
                            rows.SelectedIndex = row;
                            AssertTrue(WaitFor(delegate
                            {
                                return AuditViewIdle(left) && AuditViewIdle(right);
                            }), "строка " + row + " отрисована с обеих сторон");
                        }
                        AssertEq(1, session.RendererCreations,
                            "шесть показов двух документов — один рендерер (документы не перечитываются)");
                        form.Close();
                    }
                });
            }
        }

        [Test("Аудит F2: рендерер освобождается после простоя и при закрытии")]
        private static void AuditReviewRendererIdleRelease()
        {
            using (var dir = new TempDir("iwo_f2_idle_"))
            {
                string pdf = dir.File("doc.pdf");
                WriteTextPdf(pdf, new[] { "page one", "page two" });
                AuditWithView(delegate(PdfReviewPageView view)
                {
                    // Повторное использование — на сессии с простоем дольше теста: между
                    // показами идут UI-очередь и старт воркера, и короткий простой на
                    // загруженной машине успел бы освободить рендерер посередине.
                    var session = new PdfReviewRenderSession(TimeSpan.FromMinutes(10));
                    try
                    {
                        view.RenderSession = session;
                        view.ShowPage(new PdfPageRef { SourcePath = pdf, PageIndex = 0 }, "один", null);
                        AssertTrue(WaitFor(delegate
                        {
                            return AuditViewSettled(view, PdfReviewPageViewState.Ready);
                        }), "страница отрисована");
                        view.ShowPage(new PdfPageRef { SourcePath = pdf, PageIndex = 1 }, "два", null);
                        AssertTrue(WaitFor(delegate
                        {
                            return AuditViewSettled(view, PdfReviewPageViewState.Ready);
                        }), "вторая страница отрисована");
                        AssertEq(1, session.RendererCreations, "второй показ — тот же рендерер");
                        session.Release();
                        AssertTrue(!session.HasRenderer, "освобождение отпускает рендерер и документы");
                        view.ShowPage(new PdfPageRef { SourcePath = pdf, PageIndex = 0 }, "снова", null);
                        AssertTrue(WaitFor(delegate
                        {
                            return AuditViewSettled(view, PdfReviewPageViewState.Ready);
                        }), "после освобождения страница снова рендерится");
                        AssertEq(2, session.RendererCreations, "после освобождения создан новый рендерер");
                    }
                    finally
                    {
                        session.Dispose();
                        session.Dispose();
                    }
                    AssertTrue(!session.HasRenderer, "закрытая сессия не держит рендерер");
                    AssertTrue(session.Use(pdf, delegate(PdfThumbnailRenderer renderer)
                    {
                        return new object();
                    }) == null, "закрытая сессия не создаёт рендерер заново");

                    // Простой — отдельной сессией: сколько бы ни шёл показ, после него
                    // рендерер освобождается сам, без явного Release.
                    var idle = new PdfReviewRenderSession(TimeSpan.FromMilliseconds(50));
                    try
                    {
                        view.RenderSession = idle;
                        view.ShowPage(new PdfPageRef { SourcePath = pdf, PageIndex = 1 }, "простой", null);
                        AssertTrue(WaitFor(delegate
                        {
                            return AuditViewSettled(view, PdfReviewPageViewState.Ready);
                        }), "страница отрисована через вторую сессию");
                        AssertTrue(WaitFor(delegate { return !idle.HasRenderer; }),
                            "после простоя рендерер и документы освобождены");
                    }
                    finally { idle.Dispose(); }
                });
            }
        }

        [Test("Аудит F2: «недоступно» не считается показом и повторяется после освобождения памяти")]
        private static void AuditReviewUnavailableRetries()
        {
            using (var dir = new TempDir("iwo_f2_retry_"))
            {
                string pdf = dir.File("doc.pdf");
                WriteTextPdf(pdf, new[] { "page one" });
                AuditWithView(delegate(PdfReviewPageView view)
                {
                    var missing = new PdfPageRef { SourcePath = dir.File("missing.pdf"), PageIndex = 0 };
                    view.ShowPage(missing, null, 7L, "нет файла", null);
                    AssertTrue(WaitFor(delegate
                    {
                        return AuditViewSettled(view, PdfReviewPageViewState.Unavailable);
                    }), "несуществующий файл — «недоступно»");
                    AssertTrue(!view.IsShowing(missing, 7L),
                        "недоступная страница не считается показанной: повторный клик попробует снова");

                    var page = new PdfPageRef { SourcePath = pdf, PageIndex = 0 };
                    PdfMemoryLease hog;
                    AssertTrue(PdfMemoryBudget.TryAcquire(
                        PdfMemoryBudget.LimitBytes - PdfMemoryBudget.Used - 1024, out hog),
                        "тест занял почти весь бюджет");
                    try
                    {
                        view.ShowPage(page, null, 8L, "без памяти", null);
                        AssertTrue(WaitFor(delegate
                        {
                            return AuditViewSettled(view, PdfReviewPageViewState.Unavailable);
                        }), "без бюджета страница недоступна");
                    }
                    finally { hog.Dispose(); }
                    AssertTrue(WaitFor(delegate
                    {
                        return AuditViewSettled(view, PdfReviewPageViewState.Ready);
                    }), "после освобождения памяти страница показана без клика пользователя");
                    AssertTrue(view.IsShowing(page, 8L), "показ той же публикации");
                });
            }
        }

        [Test("Аудит F2: файл больше бюджета показа — «слишком большой», а не «недоступно»")]
        private static void AuditReviewTooLargeToShow()
        {
            long x86 = PdfReviewRenderSession.MaxShowableFileBytes(128L << 20, 1200);
            long x86Full = PdfReviewRenderSession.MaxShowableFileBytes(128L << 20, 1600);
            long x64Full = PdfReviewRenderSession.MaxShowableFileBytes(768L << 20, 2400);
            AssertTrue(x86 > (40L << 20) && x86 < (48L << 20),
                "на x86 панель показывает файлы до ~43 МБ: " + x86);
            AssertTrue(x86Full > (24L << 20) && x86Full < x86,
                "полный экран x86 (1600 px) шире панели — его предел ниже: " + x86Full);
            AssertTrue(x64Full > PdfReviewLimits.Default().MaxFileBytes,
                "на x64 любой допустимый для сравнения файл показывается и на полном экране: " +
                x64Full);
            // Предел согласован с бюджетом: один документ, растр показа и страница соседней
            // панели помещаются, а файл на байт больше — уже нет.
            long own = RasterBudget.BitmapWorkingSetBytes(1600, RasterBudget.ExpectedHeight(1600,
                297.0 / 210.0), 4);
            long neighbour = PdfMemoryBudget.EstimateBitmapBytes(1200,
                RasterBudget.ExpectedHeight(1200, 297.0 / 210.0));
            AssertTrue(2 * x86Full + own + neighbour <= (128L << 20) &&
                       2 * (x86Full + 1) + own + neighbour > (128L << 20),
                "предел полного экрана x86 — ровно то, что помещается в бюджет");
            AssertEq(0L, PdfReviewRenderSession.MaxShowableFileBytes(1L << 20, 1200),
                "бюджет меньше одного растра — показ невозможен вообще");
            AssertEq(PdfReviewRenderSession.MaxShowableFileBytes(128L << 20, 1),
                PdfReviewRenderSession.MaxShowableFileBytes(128L << 20, -5),
                "неположительная ширина не ломает расчёт");
            AssertEq(0L, PdfReviewRenderSession.MaxShowableFileBytes(128L << 20, int.MaxValue),
                "запредельная ширина не переполняет сумму растров и не выдаёт «всё покажем»");

            Lang before = Loc.Current;
            try
            {
                Loc.Init(Lang.En);
                string en = Loc.T("review.source.tooLarge");
                Loc.Init(Lang.Ru);
                string ru = Loc.T("review.source.tooLarge");
                AssertTrue(en.Contains("too large") && ru.Contains("слишком большой"),
                    "сообщение переведено: «" + ru + "» / «" + en + "»");
            }
            finally { Loc.Init(before); }

            using (var dir = new TempDir("iwo_f2_toolarge_"))
            {
                string pdf = dir.File("doc.pdf");
                WriteTextPdf(pdf, new[] { "page one" });
                AuditWithView(delegate(PdfReviewPageView view)
                {
                    view.RenderSession.ShowBudgetBytes = 1L << 20;
                    var page = new PdfPageRef { SourcePath = pdf, PageIndex = 0 };
                    view.ShowPage(page, null, 9L, "большой", null);
                    AssertTrue(WaitFor(delegate
                    {
                        return AuditViewSettled(view, PdfReviewPageViewState.TooLargeToShow);
                    }), "файл больше предела показа — отдельное состояние");
                    AssertEq(Loc.T("review.source.tooLarge"),
                        view.MessageLabel.Text,
                        "пользователь видит «слишком большой», а не «недоступно»");
                    AssertTrue(view.IsShowing(page, 9L),
                        "детерминированный отказ не перерендеривается на каждый клик");
                    AssertTrue(!view.RenderSession.HasRenderer,
                        "ради заведомо неподъёмного файла рендерер даже не создавался");
                });
            }
        }

        /// <summary>Одностраничный PDF с балластом в словаре Info: размер файла задаёт резерв документа.</summary>
        private static void AuditPaddedPdf(string path, string text, int paddingChars)
        {
            using (var doc = new PdfSharp.Pdf.PdfDocument())
            {
                PdfSharp.Pdf.PdfPage page = doc.AddPage();
                using (PdfSharp.Drawing.XGraphics graphics = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
                    graphics.DrawString(text, new PdfSharp.Drawing.XFont("Times New Roman", 14),
                        PdfSharp.Drawing.XBrushes.Black, new PdfSharp.Drawing.XPoint(50, 100));
                doc.Info.Keywords = new string('x', paddingChars);
                doc.Save(path);
            }
        }

        [Test("Аудит F2: вторая панель получает бюджет растра, не дожидаясь простоя рендерера")]
        private static void AuditReviewSecondPanelDoesNotWaitForIdle()
        {
            using (var dir = new TempDir("iwo_f2_neighbour_"))
            {
                string leftPdf = dir.File("left.pdf");
                string rightPdf = dir.File("right.pdf");
                AuditPaddedPdf(leftPdf, "Left page", 1 << 20);
                AuditPaddedPdf(rightPdf, "Right page", 1 << 20);
                RunSta(delegate
                {
                    // Простой дольше теста: показ правой панели не может прийти от таймера.
                    var session = new PdfReviewRenderSession(TimeSpan.FromMinutes(10));
                    PdfMemoryLease hog = null;
                    var host = new Form();
                    try
                    {
                        host.ClientSize = new Size(660, 480);
                        var left = new PdfReviewPageView
                            { Bounds = new Rectangle(0, 0, 320, 480), RenderSession = session };
                        var right = new PdfReviewPageView
                            { Bounds = new Rectangle(330, 0, 320, 480), RenderSession = session };
                        host.Controls.Add(left);
                        host.Controls.Add(right);
                        host.Show();
                        Application.DoEvents();
                        left.ShowPage(new PdfPageRef { SourcePath = leftPdf, PageIndex = 0 }, "L", null);
                        AssertTrue(WaitFor(delegate
                        {
                            return AuditViewSettled(left, PdfReviewPageViewState.Ready);
                        }), "левая панель показана");

                        // Панели одинаковы, страницы A4: растр правой той же ширины и высоты.
                        Bitmap shown = left.ShownBitmap;
                        long working = RasterBudget.BitmapWorkingSetBytes(shown.Width, shown.Height, 3);
                        long leftDocument = PdfMemoryBudget.EstimateDocumentBytes(
                            new FileInfo(leftPdf).Length);
                        long rightDocument = PdfMemoryBudget.EstimateDocumentBytes(
                            new FileInfo(rightPdf).Length);
                        // Свободно меньше, чем нужно правой рядом с документом левой, но хватает,
                        // если его отпустить: запас в обе стороны — половина резерва документа.
                        long free = rightDocument + working - leftDocument / 2;
                        AssertTrue(PdfMemoryBudget.TryAcquire(
                            PdfMemoryBudget.LimitBytes - PdfMemoryBudget.Used - free, out hog),
                            "тест занял бюджет: свободно " + free + " байт");

                        right.ShowPage(new PdfPageRef { SourcePath = rightPdf, PageIndex = 0 }, "R", null);
                        AssertTrue(WaitFor(delegate
                        {
                            return AuditViewSettled(right, PdfReviewPageViewState.Ready) ||
                                   AuditViewSettled(right, PdfReviewPageViewState.Unavailable);
                        }), "правая панель завершила рендер");
                        AssertEq(PdfReviewPageViewState.Ready, right.ViewState,
                            "правая панель показана сразу: рендерер отпустил документ левой, а не ждал простоя");
                        AssertEq(PdfReviewPageViewState.Ready, left.ViewState,
                            "показанная левая страница осталась на месте");
                    }
                    finally
                    {
                        host.Close();
                        host.Dispose();
                        session.Dispose();
                        if (hog != null) hog.Dispose();
                    }
                });
            }
        }

        [Test("Аудит F2: сессия повторяет неудавшийся показ один раз и только ради чужих документов")]
        private static void AuditReviewSessionRetriesOnceForOtherDocuments()
        {
            var session = new PdfReviewRenderSession(TimeSpan.FromMinutes(10));
            try
            {
                int calls = 0;
                Func<PdfThumbnailRenderer, object> fails = delegate(PdfThumbnailRenderer renderer)
                {
                    calls++;
                    return null;
                };
                AssertTrue(session.Use("a.pdf", fails) == null, "неудача отдаётся вызывающему");
                AssertEq(1, calls, "у свежего рендерера чужих документов нет — без повтора");
                calls = 0;
                session.Use("A.PDF", fails);
                AssertEq(1, calls, "тот же документ (регистр пути не важен) — без повтора");

                calls = 0;
                int creations = session.RendererCreations;
                AssertTrue(session.Use("b.pdf", fails) == null, "повтор тоже не удался — null");
                AssertEq(2, calls, "рендерер держал «a.pdf» — ровно один повтор");
                AssertEq(creations + 1, session.RendererCreations, "повтор идёт на свежем рендерере");
                calls = 0;
                session.Use("b.pdf", fails);
                AssertEq(1, calls, "после перезапуска рендерер держит только «b.pdf»");

                calls = 0;
                object done = session.Use("c.pdf", delegate(PdfThumbnailRenderer renderer)
                {
                    calls++;
                    return new object();
                });
                AssertTrue(done != null && calls == 1, "удачная работа не повторяется: " + calls);

                calls = 0;
                session.Use<object>("d.pdf", delegate(PdfThumbnailRenderer renderer)
                {
                    calls++;
                    session.Dispose();
                    return null;
                });
                AssertEq(1, calls, "закрытие сессии посреди работы отменяет повтор");
                AssertTrue(!session.HasRenderer, "выход из работы отпустил рендерер закрытой сессии");
            }
            finally { session.Dispose(); }
        }

        private static PdfReviewDocument AuditSizedDocument(string path, int words, int characters)
        {
            return new PdfReviewDocument { Path = path, WordCount = words, CharacterCount = characters };
        }

        [Test("Аудит F2: кэш документов ограничен оценкой байт, на x86 — двумя документами")]
        private static void AuditReviewDocumentCacheIsByteBounded()
        {
            // По полям модели на x64 слово со свидетельством пробела занимает ~384 Б.
            long perWord = PdfReviewDocumentCache.BytesPerWord;
            AssertTrue(perWord >= 384 && perWord <= 480,
                "оценка слова не ниже расчёта по полям модели: " + perWord);
            PdfReviewDocument limit = AuditSizedDocument("big.pdf", 300000, 2000000);
            long estimate = PdfReviewDocumentCache.EstimateBytes(limit);
            AssertTrue(estimate > (110L << 20) && estimate < (150L << 20),
                "документ у предела 2 млн знаков оценивается в ~120 МБ: " + estimate);
            AssertEq(IntPtr.Size == 8 ? 4 : 2, PdfReviewService.CacheCapacity,
                "на x86 кэш держит не больше пары документов");
            AssertTrue(PdfReviewService.CacheMaxBytes < (IntPtr.Size == 8 ? 1L << 30 : 256L << 20),
                "суммарный потолок кэша соответствует разрядности: " + PdfReviewService.CacheMaxBytes);

            long unit = PdfReviewDocumentCache.EstimateBytes(AuditSizedDocument("x", 1000, 0));
            var cache = new PdfReviewDocumentCache(4, unit * 2 + unit / 2);
            cache.Add("a", AuditSizedDocument("a", 1000, 0));
            cache.Add("b", AuditSizedDocument("b", 1000, 0));
            PdfReviewDocument hit;
            AssertTrue(cache.TryGet("a", out hit), "a свежий после обращения");
            cache.Add("c", AuditSizedDocument("c", 1000, 0));
            AssertEq(2, cache.Count, "третий документ вытеснил по байтам, хотя место по числу есть");
            AssertTrue(!cache.TryGet("b", out hit), "вытеснен наименее недавно использованный");
            AssertTrue(cache.TryGet("a", out hit) && cache.TryGet("c", out hit),
                "свежие документы остались");
            AssertEq(unit * 2, cache.Bytes, "учёт байт совпадает с содержимым");

            cache.Add("huge", AuditSizedDocument("huge", 3000, 0));
            AssertTrue(!cache.TryGet("huge", out hit) && cache.Count == 2,
                "документ больше всего потолка не кэшируется и ничего не вытесняет");
            cache.Add("a", AuditSizedDocument("a2", 1000, 0));
            AssertEq(unit * 2, cache.Bytes, "повторный ключ не учитывается дважды");

            var small = new PdfReviewDocumentCache(2, long.MaxValue);
            small.Add("1", AuditSizedDocument("1", 1, 1));
            small.Add("2", AuditSizedDocument("2", 1, 1));
            small.Add("3", AuditSizedDocument("3", 1, 1));
            AssertEq(2, small.Count, "потолок по числу сохраняется");
            AssertEq(PdfReviewDocumentCache.EstimateBytes(AuditSizedDocument("x", 1, 1)) * 2,
                small.Bytes, "вытеснение по числу тоже вычитает байты");
        }

        /// <summary>count пар страниц, на каждой одна замена слова — отдельная группа уточнения.</summary>
        private static PdfReviewResult AuditManyChangedPairs(int count)
        {
            var leftPages = new string[count];
            var rightPages = new string[count];
            for (int i = 0; i < count; i++)
            {
                leftPages[i] = "common page" + i + " old" + i;
                rightPages[i] = "common page" + i + " new" + i;
            }
            return PdfReviewDiff.Compare(ReviewDocument("left.pdf", leftPages),
                ReviewDocument("right.pdf", rightPages), PdfReviewLimits.Default());
        }

        [Test("Аудит F2: растровое уточнение ограничено числом пар и сообщает прогресс")]
        private static void AuditReviewVisualDiffPairCap()
        {
            int pairs = PdfReviewVisualDiff.MaximumRenderedPairs + 8;
            PdfReviewResult result = AuditManyChangedPairs(pairs);
            AssertTrue(!result.WorkExhausted, "текстовое сравнение уложилось в бюджет");
            int renders = 0;
            var reports = new List<int>();
            int reportedTotal = 0;
            PdfReviewVisualDiff.Refine(result, PdfReviewLimits.Default(), null,
                delegate(string path, int pageIndex, int width, int maxHeight)
                {
                    renders++;
                    var bitmap = new Bitmap(width, Math.Max(1, Math.Min(maxHeight, width * 3 / 2)));
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                        graphics.Clear(Color.White);
                    return bitmap;
                },
                delegate(int done, int total)
                {
                    reports.Add(done);
                    reportedTotal = total;
                });
            AssertEq(PdfReviewVisualDiff.MaximumRenderedPairs * 2, renders,
                "растеризовано не больше " + PdfReviewVisualDiff.MaximumRenderedPairs + " пар страниц");
            AssertTrue(result.WorkExhausted,
                "непроверенные пары помечают результат как сравнённый грубо");
            AssertEq(pairs, reportedTotal, "прогресс считает группы уточнения");
            for (int i = 1; i < reports.Count; i++)
                AssertTrue(reports[i] >= reports[i - 1], "прогресс не убывает");
            AssertTrue(reports.Count > 2 && reports[reports.Count - 1] == pairs,
                "прогресс дошёл до конца стадии по группам");

            PdfReviewResult few = AuditManyChangedPairs(3);
            renders = 0;
            PdfReviewVisualDiff.Refine(few, PdfReviewLimits.Default(), null,
                delegate(string path, int pageIndex, int width, int maxHeight)
                {
                    renders++;
                    return new Bitmap(width, Math.Max(1, Math.Min(maxHeight, width)));
                });
            AssertEq(6, renders, "до потолка рендерится каждая пара");
            AssertTrue(!few.WorkExhausted, "без потолка грубых мест нет");
        }

        [Test("Аудит F2: пара без растра (нехватка бюджета) помечает сравнение как грубое")]
        private static void AuditReviewVisualDiffMissingRasterIsExhausted()
        {
            PdfReviewResult result = AuditManyChangedPairs(3);
            AssertTrue(!result.WorkExhausted, "текстовое сравнение уложилось в бюджет");
            int renders = 0;
            PdfReviewVisualDiff.Refine(result, PdfReviewLimits.Default(), null,
                delegate(string path, int pageIndex, int width, int maxHeight)
                {
                    renders++;
                    return null;
                });
            AssertTrue(renders > 0, "уточнение пыталось растрировать пары");
            AssertTrue(result.WorkExhausted,
                "ни одна пара не растрирована — окно обязано предупредить: рендеров " + renders);

            PdfReviewResult oneSide = AuditManyChangedPairs(3);
            PdfReviewVisualDiff.Refine(oneSide, PdfReviewLimits.Default(), null,
                delegate(string path, int pageIndex, int width, int maxHeight)
                {
                    if (path == "right.pdf")
                        return null;
                    var bitmap = new Bitmap(width, Math.Max(1, Math.Min(maxHeight, width)));
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                        graphics.Clear(Color.White);
                    return bitmap;
                });
            AssertTrue(oneSide.WorkExhausted, "не хватило растра одной стороны — тоже грубо");
        }

        [Test("Аудит F2: шкала прогресса сравнения не стоит на 100% до конца")]
        private static void AuditReviewCompareProgressScale()
        {
            AssertEq(0, PdfReviewService.ProgressPoint(0, 10, 0, 400), "начало стадии");
            AssertEq(200, PdfReviewService.ProgressPoint(5, 10, 0, 400), "середина стадии");
            AssertEq(400, PdfReviewService.ProgressPoint(10, 10, 0, 400), "конец стадии");
            AssertEq(850, PdfReviewService.ProgressPoint(3, 0, 850, 990), "пустая стадия — её начало");
            AssertEq(990, PdfReviewService.ProgressPoint(int.MaxValue, 7, 850, 990),
                "переполнение done зажато в конец стадии");

            using (var dir = new TempDir("iwo_f2_progress_"))
            {
                string left = dir.File("left.pdf");
                string right = dir.File("right.pdf");
                WriteTextPdf(left, new[] { "Alpha one", "Beta two", "Gamma three" });
                WriteTextPdf(right, new[] { "Alpha one", "Beta changed", "Gamma three" });
                var percents = new List<int>();
                PdfReviewService.Compare(left, right, delegate(int done, int total)
                {
                    percents.Add(PdfToolFormBase.ProgressPercent(done, total));
                });
                AssertTrue(percents.Count > 0, "прогресс сообщается");
                int full = 0;
                bool comparing = false;
                for (int i = 0; i < percents.Count; i++)
                {
                    if (percents[i] == 100) full++;
                    if (percents[i] > 80 && percents[i] < 100) comparing = true;
                    if (i > 0)
                        AssertTrue(percents[i] >= percents[i - 1],
                            "прогресс не убывает: " + string.Join(",", ToStrings(percents)));
                }
                AssertEq(1, full, "100% — один раз, когда результат готов: " +
                    string.Join(",", ToStrings(percents)));
                AssertEq(100, percents[percents.Count - 1], "последний отчёт — 100%");
                AssertTrue(comparing, "сравнение и уточнение видны на шкале между 80% и 100%: " +
                    string.Join(",", ToStrings(percents)));
            }
        }

        private static string[] ToStrings(List<int> values)
        {
            var result = new string[values.Count];
            for (int i = 0; i < values.Count; i++)
                result[i] = values[i].ToString(System.Globalization.CultureInfo.InvariantCulture);
            return result;
        }

        [Test("Аудит F2: загрузка без лишних открытий сохраняет причины отказа")]
        private static void AuditReviewLoadSingleProbeKeepsReasons()
        {
            using (var dir = new TempDir("iwo_f2_load_"))
            {
                string text = dir.File("text.pdf");
                WriteTextPdf(text, new[] { "Visible text one", "Visible text two" });
                PdfReviewDocument loaded = PdfReviewService.Load(text, null, null,
                    PdfReviewLimits.Default());
                AssertEq(2, loaded.Pages.Count, "born-digital PDF загружен целиком");

                string blank = dir.File("blank.pdf");
                using (var doc = new PdfSharp.Pdf.PdfDocument())
                {
                    doc.AddPage();
                    doc.AddPage();
                    doc.Save(blank);
                }
                PdfReviewException noText = AssertThrowsExactly<PdfReviewException>(
                    "страницы без букв", delegate
                    {
                        PdfReviewService.Load(blank, null, null, PdfReviewLimits.Default());
                    });
                AssertEq(PdfReviewFailure.NoText, noText.Reason,
                    "документ без текстового слоя по-прежнему отвергнут как скан");

                var fewPages = PdfReviewLimits.Default();
                fewPages.MaxPages = 1;
                string other = dir.File("other.pdf");
                WriteTextPdf(other, new[] { "Another one", "Another two" });
                PdfReviewException tooMany = AssertThrowsExactly<PdfReviewException>(
                    "слишком много страниц", delegate
                    {
                        PdfReviewService.Load(other, null, null, fewPages);
                    });
                AssertEq(PdfReviewFailure.TooLarge, tooMany.Reason,
                    "предел страниц проверяется до поиска текста");

                AssertThrowsExactly<OperationCanceledException>("отмена во время поиска текста",
                    delegate
                    {
                        PdfReviewService.Load(dir.File("blank.pdf"), null,
                            delegate { return true; }, PdfReviewLimits.Default());
                    });
            }
        }

        [Test("Аудит F2: путь проверяется без обращения к сети на UI-потоке")]
        private static void AuditReviewSourceChecksAvoidUiBlocking()
        {
            // Сервера нет: если бы Resolve проверял существование файла, ответ был бы «нет файла».
            // «Принят» доказывает, что обращения не было, без порога по секундомеру (TS-5).
            const string unc = @"\\iwo-f2-unreachable.invalid\share\source.pdf";
            PdfReviewSourceResult resolved = PdfReviewInput.Resolve(unc);
            AssertTrue(resolved.IsValid && resolved.Path == unc,
                "сетевой путь синтаксически принят, существование проверит фоновая проба");
            AssertTrue(!PdfReviewInput.IsLocalDiskPath(unc) &&
                       !PdfReviewInput.IsLocalDiskPath(@"\\?\C:\x.pdf"),
                "UNC и \\\\?\\-пути не локальные");
            AssertTrue(PdfReviewInput.IsLocalDiskPath(Path.GetFullPath(Path.GetTempPath())),
                "временная папка — локальный диск");

            using (var dir = new TempDir("iwo_f2_probe_"))
            {
                AssertEq(PdfReviewSourceError.Missing, PdfReviewInput.Probe(
                    new PdfReviewSourceResult { Path = dir.File("gone.pdf") }).Error,
                    "фоновая проба сообщает «не найден», а не «не читается»");
            }

            var checkedPaths = new List<string>();
            string[] picked = PdfDrop.Filter(new[] { "notes.txt", @"\\server\share", "doc.PDF", null },
                PdfDrop.PdfOnly, delegate(string path)
                {
                    checkedPaths.Add(path);
                    return true;
                });
            AssertEq(1, picked.Length, "принят только PDF");
            AssertEq(1, checkedPaths.Count, "диск проверяется только для подходящего расширения");
            AssertEq("doc.PDF", checkedPaths[0], "существование проверено у PDF");
        }

        /// <summary>IDataObject, считающий чтения списка файлов перетаскивания.</summary>
        private sealed class AuditCountingData : IDataObject
        {
            private readonly string[] _files;
            public int Reads;

            public AuditCountingData(params string[] files) { _files = files; }

            public object GetData(string format, bool autoConvert) { return GetData(format); }
            public object GetData(string format)
            {
                if (format != DataFormats.FileDrop)
                    return null;
                Reads++;
                return _files;
            }
            public object GetData(Type format) { return null; }
            public bool GetDataPresent(string format, bool autoConvert) { return GetDataPresent(format); }
            public bool GetDataPresent(string format) { return format == DataFormats.FileDrop; }
            public bool GetDataPresent(Type format) { return false; }
            public string[] GetFormats(bool autoConvert) { return GetFormats(); }
            public string[] GetFormats() { return new[] { DataFormats.FileDrop }; }
            public void SetData(string format, bool autoConvert, object data) { }
            public void SetData(string format, object data) { }
            public void SetData(Type format, object data) { }
            public void SetData(object data) { }
        }

        [Test("Аудит F2: DragOver не перечитывает и не проверяет пути на каждое движение")]
        private static void AuditReviewDragOverUsesCachedPaths()
        {
            using (var dir = new TempDir("iwo_f2_drag_"))
            {
                string pdf = dir.File("drop.pdf");
                WriteTextPdf(pdf, new[] { "Dropped" });
                InIsolatedSettings("iwo_f2_drag_settings_", delegate
                {
                    using (var form = new PdfReviewForm(null))
                    {
                        form.Show();
                        Application.DoEvents();
                        var data = new AuditCountingData(pdf);
                        System.Reflection.MethodInfo enter = typeof(PdfReviewForm).GetMethod(
                            "OnReviewDragEnter", System.Reflection.BindingFlags.Instance |
                            System.Reflection.BindingFlags.NonPublic);
                        System.Reflection.MethodInfo over = typeof(PdfReviewForm).GetMethod(
                            "OnReviewDragOver", System.Reflection.BindingFlags.Instance |
                            System.Reflection.BindingFlags.NonPublic);
                        Point point = form.PointToScreen(new Point(30, 30));
                        var args = new DragEventArgs(data, 0, point.X, point.Y,
                            DragDropEffects.Copy, DragDropEffects.None);
                        enter.Invoke(form, new object[] { form, args });
                        AssertEq(DragDropEffects.Copy, args.Effect, "PDF принят на входе");
                        for (int i = 0; i < 5; i++)
                        {
                            var move = new DragEventArgs(data, 0, point.X + i, point.Y,
                                DragDropEffects.Copy, DragDropEffects.None);
                            over.Invoke(form, new object[] { form, move });
                            AssertEq(DragDropEffects.Copy, move.Effect, "движение сохраняет эффект");
                        }
                        AssertEq(1, data.Reads,
                            "список файлов прочитан один раз на входе, а не на каждом DragOver");
                        form.SetDropCue(null);
                        form.Close();
                    }
                });
            }
        }

        [Test("Аудит F2: перетаскивание, начатое во время работы, принимается, когда работа закончилась")]
        private static void AuditReviewDragDuringWorkIsReevaluated()
        {
            using (var dir = new TempDir("iwo_f2_dragwork_"))
            {
                string pdf = dir.File("drop.pdf");
                WriteTextPdf(pdf, new[] { "Dropped" });
                InIsolatedSettings("iwo_f2_dragwork_settings_", delegate
                {
                    using (var form = new PdfReviewForm(null))
                    {
                        form.Show();
                        Application.DoEvents();
                        var data = new AuditCountingData(pdf);
                        System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.Instance |
                            System.Reflection.BindingFlags.NonPublic;
                        System.Reflection.MethodInfo enter = typeof(PdfReviewForm).GetMethod(
                            "OnReviewDragEnter", flags);
                        System.Reflection.MethodInfo over = typeof(PdfReviewForm).GetMethod(
                            "OnReviewDragOver", flags);
                        System.Reflection.FieldInfo busy = typeof(PdfToolFormBase).GetField("_busy",
                            flags);
                        Point point = form.PointToScreen(new Point(30, 30));
                        try
                        {
                            busy.SetValue(form, true);
                            var args = new DragEventArgs(data, 0, point.X, point.Y,
                                DragDropEffects.Copy, DragDropEffects.None);
                            enter.Invoke(form, new object[] { form, args });
                            AssertEq(DragDropEffects.None, args.Effect, "во время работы drop запрещён");
                            AssertEq(0, data.Reads, "во время работы список файлов даже не читается");
                        }
                        finally { busy.SetValue(form, false); }

                        var move = new DragEventArgs(data, 0, point.X + 1, point.Y,
                            DragDropEffects.Copy, DragDropEffects.None);
                        over.Invoke(form, new object[] { form, move });
                        AssertEq(DragDropEffects.Copy, move.Effect,
                            "работа закончилась — то же перетаскивание принимает PDF без повторного входа");
                        over.Invoke(form, new object[] { form, move });
                        AssertEq(1, data.Reads, "после работы список прочитан один раз и закэширован");
                        form.SetDropCue(null);
                        form.Close();
                    }
                });
            }
        }
    }
}
