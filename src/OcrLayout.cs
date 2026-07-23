using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelMerger
{
    /// <summary>Слово с рамкой (координаты PDF, ось Y направлена вверх) и форматом.</summary>
    internal class PdfWord
    {
        public string Text;
        public double Left;
        public double Right;
        public double Bottom;
        public double Top;
        public double FontSizePt; // кегль (pt); 0 — неизвестно
        public bool Bold;
        public bool Italic;
        public int ColorArgb;     // 0xRRGGBB цвет текста; 0 — чёрный
        public string FontName;   // семейство шрифта; null — по умолчанию
        public bool Super;        // надстрочный (мельче и приподнят над базовой линией)
        public bool Sub;          // подстрочный (мельче и опущен)
        public bool Underline;    // под словом проходит горизонтальная линовка (подчёркивание)
        public string Uri;        // гиперссылка, если слово внутри её рамки; иначе null

        public double MidY { get { return (Top + Bottom) / 2; } }
        public double Height { get { return Top - Bottom; } }
    }

    /// <summary>Выравнивание абзаца в выводе.</summary>
    public enum OcrAlignment { Left, Justify, Center }

    /// <summary>Ран — отрезок абзаца с единым форматом (кегль, полужирный, курсив, цвет).</summary>
    public class OcrRun
    {
        public string Text;
        public double FontSizePt; // 0 — неизвестно, писать кеглем по умолчанию
        public bool Bold;
        public bool Italic;
        public int ColorArgb;     // 0xRRGGBB; 0 — чёрный
        public string FontName;   // семейство шрифта; null — по умолчанию
        public bool Super;        // надстрочный
        public bool Sub;          // подстрочный
        public bool Underline;    // подчёркнутый
        public string Uri;        // гиперссылка рана; null — обычный текст
    }

    /// <summary>Абзац: раны с форматом + выравнивание. Text — склейка ранов (единый источник).</summary>
    public class OcrParagraph
    {
        public List<OcrRun> Runs = new List<OcrRun>();
        public OcrAlignment Alignment = OcrAlignment.Justify;
        public double TopPt;  // верх абзаца (Y первой строки, ось вверх) — для порядка с изображениями
        public double LeftPt; // левый край абзаца — вторичный порядок (левее — раньше в одной строке-полосе)

        // Список: вид маркера в начале абзаца и номер (для нумерованного). None — обычный абзац.
        // ListContentStart — индекс в Text, с которого идёт содержимое (маркер снимается при записи —
        // нативный список Word рисует свой маркер). Заполняется в AnalyzeLines через ListMarker.
        public ListKind ListKind = ListKind.None;
        public int ListNumber;       // номер нумерованного пункта; 0 для маркированного/обычного
        public int ListContentStart; // сколько ведущих символов Text занимает маркер (снять при записи)

        public string Text
        {
            get
            {
                var sb = new StringBuilder();
                for (int i = 0; i < Runs.Count; i++)
                    sb.Append(Runs[i].Text);
                return sb.ToString();
            }
        }
    }

    /// <summary>
    /// Порядок чтения born-digital PDF: слова с рамками → строки → абзацы
    /// (сверху вниз, слева направо). Рассчитан на одноколоночную вёрстку (типичный
    /// экспорт из Word); многоколоночную сюда не тащим — это отдельная задача
    /// (PdfPig DLA). Чистая логика без типов PdfPig — покрыта юнит-тестами.
    /// </summary>
    internal static class OcrLayout
    {
        // Слово в той же строке, если рамки перекрываются по вертикали не меньше чем на эту
        // долю меньшей высоты. Через перекрытие рамок (а не расстояние центров) — чтобы тонкая
        // пунктуация (тире «—», дефис) с крошечной высотой и смещённым центром не отрывалась
        // в отдельную строку.
        private const double SameLineFactor = 0.5;
        // Новый абзац — если вертикальный зазор между строками заметно больше обычного.
        private const double ParagraphGapFactor = 1.6;
        // Отступ первой строки (красная строка) — сигнал нового абзаца: строка начинается
        // правее общего левого поля больше чем на IndentFactor кегля (или IndentWidthFraction ширины).
        private const double IndentFactor = 0.75;
        private const double IndentWidthFraction = 0.03;
        // Выключка по ширине (justified): если столько строк достаёт до правого поля — считаем
        // абзац законченным после «короткой» строки (не дотянувшей до поля на ShortLineFraction ширины).
        private const double JustifiedShare = 0.6;
        private const double ShortLineFraction = 0.15;
        // Красную строку воссоздаём в выводе, только если ею оформлено не меньше этой доли абзацев
        // (иначе документ без отступов — не портим его). Отступ крупнее MaxIndentFraction ширины —
        // это центрирование (номер страницы и т.п.), в расчёт красной строки не берём.
        private const double IndentedShare = 0.6;
        private const double MaxIndentFraction = 0.25;
        // Центрированная строка: левый и правый зазоры до полей заметны (> CenterMinGapFraction
        // ширины) и почти равны (разница <= CenterBalanceFraction ширины) — так узнаём номер
        // страницы/заголовок по центру и не выключаем его по ширине с красной строкой.
        private const double CenterMinGapFraction = 0.12;
        private const double CenterBalanceFraction = 0.08;
        // Надстрочный/подстрочный: слово мельче доминирующего кегля строки и смещено по базовой
        // линии вверх (надстрочный) или вниз (подстрочный) заметно относительно кегля.
        private const double ScriptSizeFactor = 0.85;
        private const double ScriptRiseFactor = 0.1;
        // Между двумя соседними словами строки ставим пробел: PdfPig уже сгруппировал буквы в
        // слова, поэтому раздельные слова — это граница слова. Склеиваем без пробела ТОЛЬКО
        // почти соприкасающиеся токены (зазор < SpaceGapFactor кегля) — так PdfPig изредка дробит
        // один токен на куски. Порог выбран заметно НИЖЕ минимального настоящего межсловного
        // зазора: в узких шрифтах (напр. Calibri Light) пробел ≈ 0.18 кегля, поэтому прежние 0.2
        // роняли настоящие пробелы и слепляли слова («СЛОВОСЛОВО»). 0.08 надёжно ниже.
        private const double SpaceGapFactor = 0.08;
        // Разделение узкой левой колонки-сайдбара от тела. Внутристрочный зазор больше
        // ColumnGapFactor×em — граница колонок; сайдбар-сегмент начинается левее поля тела не
        // меньше чем на SidebarBodyGapFactor×em; активируется, только если так делятся ≥ MinSplitLines
        // строк (иначе одноколоночный текст — не трогаем) и обе части существенны.
        private const double ColumnGapFactor = 3.0;
        private const double SidebarBodyGapFactor = 2.0;
        private const int MinSplitLines = 2;

        /// <summary>Итог разбора страницы: абзацы и измеренный отступ первой строки (0 — без красной строки).</summary>
        internal sealed class OcrPageLayout
        {
            public List<OcrParagraph> Paragraphs = new List<OcrParagraph>();
            public double FirstLineIndentPt; // pt (PDF = pt), 0 если документ без отступов
        }

        /// <summary>Слова (в любом порядке) → абзацы в порядке чтения (только текст). Чистая — под тест.</summary>
        public static List<string> ToParagraphs(IList<PdfWord> words)
        {
            List<OcrParagraph> paras = Analyze(words).Paragraphs;
            var texts = new List<string>(paras.Count);
            foreach (OcrParagraph p in paras)
                texts.Add(p.Text);
            return texts;
        }

        /// <summary>
        /// Слова (в любом порядке) → абзацы в порядке чтения + отступ красной строки. Чистая — под тест.
        /// Граница абзаца определяется тремя независимыми сигналами (любой срабатывает):
        ///  • вертикальный зазор заметно больше обычного (абзацы разделены пустым местом);
        ///  • красная строка — текущая строка начата правее общего левого поля (отступ);
        ///  • в justified-тексте предыдущая строка «короткая» (не дотянулась до правого поля —
        ///    значит была последней в абзаце). Это ключ для Word-экспортов, где абзацы
        ///    разделены НЕ зазором, а отступом первой строки при равном межстрочном интервале.
        /// FirstLineIndentPt — медиана отступов первых строк, если красной строкой оформлено
        /// большинство абзацев; иначе 0 (документ без отступов не трогаем).
        /// </summary>
        public static OcrPageLayout Analyze(IList<PdfWord> words)
        {
            List<Line> lines = ToLines(words);
            // Левая узкая колонка-сайдбар (метки/даты резюме и т.п.) сбивает геометрию тела и
            // «влезает» в его строки. Если она распознана — тело и сайдбар разбираются ОТДЕЛЬНО
            // (у каждого своя геометрия), а абзацы сводятся по вертикали в OrderedBlocks (по TopPt).
            ColumnSplit split = TrySplitSidebar(lines);
            if (split != null)
            {
                OcrPageLayout body = AnalyzeLines(ToLines(split.Body));
                OcrPageLayout side = AnalyzeLines(ToLines(split.Sidebar));
                var merged = new OcrPageLayout { FirstLineIndentPt = body.FirstLineIndentPt };
                merged.Paragraphs.AddRange(body.Paragraphs);
                merged.Paragraphs.AddRange(side.Paragraphs); // порядок неважен — OrderedBlocks сортирует по TopPt
                return merged;
            }
            return AnalyzeLines(lines);
        }

        /// <summary>Разбор уже сгруппированных строк в абзацы (ядро Analyze без разделения колонок).</summary>
        private static OcrPageLayout AnalyzeLines(List<Line> lines)
        {
            var result = new OcrPageLayout();
            if (lines.Count == 0)
                return result;
            MarkScripts(lines); // пометить надстрочные/подстрочные слова

            // Геометрия страницы: левое/правое поле, типичный кегль, признак выключки.
            double bodyLeft = double.MaxValue, bodyRight = double.MinValue;
            var heights = new List<double>(lines.Count);
            foreach (Line ln in lines)
            {
                if (ln.Left < bodyLeft) bodyLeft = ln.Left;
                if (ln.Right > bodyRight) bodyRight = ln.Right;
                heights.Add(ln.Height);
            }
            double em = Median(heights);
            if (em <= 0) em = 1;
            double width = bodyRight - bodyLeft;
            if (width <= 0) width = 1;

            int reaching = 0; // строк, достающих до правого поля (полные строки justified-текста)
            foreach (Line ln in lines)
                if (ln.Right >= bodyRight - 0.5 * em) reaching++;
            bool justified = reaching >= (int)Math.Ceiling(lines.Count * JustifiedShare);

            double gapThreshold = ParagraphThreshold(lines);
            double indentTol = Math.Max(IndentFactor * em, IndentWidthFraction * width);
            double shortTol = ShortLineFraction * width;

            // Разбить на группы строк-абзацев. Дополнительный сигнал — строка начинается с
            // маркера списка («1.», «•»): каждый пункт становится своим абзацем даже в плотном
            // одностроковом списке (равный интервал, без отступа), где иначе строки слиплись бы
            // в один абзац и распознался бы лишь первый маркер.
            var groups = new List<List<Line>>();
            var current = new List<Line>();
            for (int i = 0; i < lines.Count; i++)
            {
                if (i > 0)
                {
                    bool gapBreak = lines[i - 1].MidY - lines[i].MidY > gapThreshold;
                    bool indentBreak = lines[i].Left - bodyLeft > indentTol;
                    bool shortBreak = justified && lines[i - 1].Right < bodyRight - shortTol;
                    bool listBreak = StartsWithListMarker(lines[i]);
                    if (gapBreak || indentBreak || shortBreak || listBreak)
                    {
                        groups.Add(current);
                        current = new List<Line>();
                    }
                }
                current.Add(lines[i]);
            }
            groups.Add(current);

            // Абзацы: раны с форматом + выравнивание; сбор отступов первых строк (исключая центрирование).
            double maxIndent = MaxIndentFraction * width;
            var indents = new List<double>();
            foreach (List<Line> g in groups)
            {
                var para = new OcrParagraph
                {
                    Runs = BuildRuns(g),
                    Alignment = DetectAlignment(g, bodyLeft, bodyRight, width, em),
                    TopPt = g[0].Top,
                    LeftPt = g[0].Left
                };
                if (para.Alignment != OcrAlignment.Center)
                {
                    ListMarker.Result m = ListMarker.Detect(para.Text);
                    para.ListKind = m.Kind;
                    para.ListNumber = m.Number;
                    para.ListContentStart = m.ContentStart;
                }
                result.Paragraphs.Add(para);
                double ind = g[0].Left - bodyLeft;
                if (ind > indentTol && ind <= maxIndent)
                    indents.Add(ind);
            }
            if (groups.Count >= 2 && indents.Count >= (int)Math.Ceiling(groups.Count * IndentedShare))
                result.FirstLineIndentPt = Median(indents);

            return result;
        }

        /// <summary>Слова левого сайдбара и тела после разделения колонок.</summary>
        private sealed class ColumnSplit
        {
            public List<PdfWord> Sidebar;
            public List<PdfWord> Body;
        }

        /// <summary>
        /// Распознать узкую левую колонку-сайдбар и вернуть слова тела/сайдбара, иначе null.
        /// Признак: в строках есть большие внутренние зазоры (сайдбар-сегмент слева + тело справа);
        /// поле тела — минимальный левый край сегментов ПОСЛЕ зазора; сайдбар — сегменты, начатые
        /// заметно левее поля тела. Активируется только при ≥ MinSplitLines делящихся строках и
        /// существенности обеих частей — иначе (одноколоночный текст) вернём null. Чистая — под тест.
        /// </summary>
        private static ColumnSplit TrySplitSidebar(List<Line> lines)
        {
            if (lines.Count < 4)
                return null;
            var heights = new List<double>(lines.Count);
            foreach (Line ln in lines) heights.Add(ln.Height);
            double em = Median(heights);
            if (em <= 0)
                return null;
            double gapTol = ColumnGapFactor * em;

            // Сегменты каждой строки (разрыв по зазору > gapTol) и левые края «правых» сегментов.
            var segments = new List<Segment>();     // все сегменты всех строк
            var afterGapLefts = new List<double>();  // левые края сегментов, перед которыми был зазор
            int splitLines = 0;
            foreach (Line ln in lines)
            {
                List<Segment> segs = SplitLineByGaps(ln, gapTol);
                if (segs.Count > 1)
                {
                    splitLines++;
                    for (int i = 1; i < segs.Count; i++)
                        afterGapLefts.Add(segs[i].Left);
                }
                segments.AddRange(segs);
            }
            if (splitLines < MinSplitLines || afterGapLefts.Count == 0)
                return null;

            double bodyLeft = Min(afterGapLefts);
            double threshold = bodyLeft - SidebarBodyGapFactor * em;

            var sidebar = new List<PdfWord>();
            var body = new List<PdfWord>();
            foreach (Segment s in segments)
                (s.Left < threshold ? sidebar : body).AddRange(s.Words);

            // Обе части должны быть существенны, иначе это не сайдбар-раскладка.
            if (sidebar.Count < 3 || body.Count < 10)
                return null;
            return new ColumnSplit { Sidebar = sidebar, Body = body };
        }

        /// <summary>Отрезок строки без больших внутренних зазоров: левый край и его слова.</summary>
        private sealed class Segment
        {
            public double Left;
            public readonly List<PdfWord> Words = new List<PdfWord>();
        }

        /// <summary>Разбить строку на сегменты там, где зазор между соседними словами больше gapTol.</summary>
        private static List<Segment> SplitLineByGaps(Line line, double gapTol)
        {
            var result = new List<Segment>();
            var cur = new Segment { Left = line.Words[0].Left };
            cur.Words.Add(line.Words[0]);
            for (int i = 1; i < line.Words.Count; i++)
            {
                double gap = line.Words[i].Left - line.Words[i - 1].Right;
                if (gap > gapTol)
                {
                    result.Add(cur);
                    cur = new Segment { Left = line.Words[i].Left };
                }
                cur.Words.Add(line.Words[i]);
            }
            result.Add(cur);
            return result;
        }

        private static double Min(List<double> values)
        {
            double m = double.MaxValue;
            for (int i = 0; i < values.Count; i++) if (values[i] < m) m = values[i];
            return m;
        }

        /// <summary>Медиана (нижняя при чётном числе). Чистая.</summary>
        private static double Median(List<double> values)
        {
            if (values.Count == 0)
                return 0;
            var copy = new List<double>(values);
            copy.Sort();
            return copy[(copy.Count - 1) / 2];
        }

        /// <summary>
        /// Абзац → раны: слова в порядке чтения склеиваются (пробел между словами; перенос по
        /// строкам склеивает, дефис-перенос снимается), а при смене формата (кегль/жирный/курсив/
        /// цвет) начинается новый ран — формат сохраняется пословно, а не на весь абзац.
        /// </summary>
        private static List<OcrRun> BuildRuns(List<Line> group)
        {
            // 1) Плоский список слов в порядке чтения: текст (с возможным снятием дефиса-переноса),
            //    слово-источник формата и признак пробела перед словом.
            var texts = new List<string>();
            var fmts = new List<PdfWord>();
            var spaceBefore = new List<bool>();
            for (int li = 0; li < group.Count; li++)
            {
                List<PdfWord> ws = group[li].Words;
                for (int wi = 0; wi < ws.Count; wi++)
                {
                    string text = ws[wi].Text ?? "";
                    bool space;
                    if (texts.Count == 0)
                        space = false;
                    else if (wi == 0) // первое слово новой строки — перенос
                    {
                        int last = texts.Count - 1;
                        if (EndsWithHyphenAfterLetter(texts[last]))
                        {
                            texts[last] = texts[last].Substring(0, texts[last].Length - 1); // снять дефис
                            space = false; // склеить перенос без пробела
                        }
                        else space = true;
                    }
                    else space = HasSpaceBetween(ws[wi - 1], ws[wi]); // по зазору, а не всегда пробел
                    texts.Add(text);
                    fmts.Add(ws[wi]);
                    spaceBefore.Add(space);
                }
            }

            // 2) Склейка в раны: смежные слова одного формата — в один ран; межсловный пробел
            //    относим к предыдущему рану (у пробела нет глифа — формат не важен).
            var runs = new List<OcrRun>();
            PdfWord runFmt = null;
            for (int i = 0; i < texts.Count; i++)
            {
                if (runs.Count > 0 && runFmt != null && SameFormat(runFmt, fmts[i]))
                {
                    runs[runs.Count - 1].Text += (spaceBefore[i] ? " " : "") + texts[i];
                }
                else
                {
                    if (runs.Count > 0 && spaceBefore[i])
                        runs[runs.Count - 1].Text += " ";
                    runs.Add(new OcrRun
                    {
                        Text = texts[i],
                        FontSizePt = fmts[i].FontSizePt,
                        Bold = fmts[i].Bold,
                        Italic = fmts[i].Italic,
                        ColorArgb = fmts[i].ColorArgb,
                        FontName = fmts[i].FontName,
                        Super = fmts[i].Super,
                        Sub = fmts[i].Sub,
                        Underline = fmts[i].Underline,
                        Uri = fmts[i].Uri
                    });
                    runFmt = fmts[i];
                }
            }

            // 3) Обрезать края абзаца; выкинуть опустевшие раны.
            if (runs.Count > 0)
            {
                runs[0].Text = runs[0].Text.TrimStart();
                runs[runs.Count - 1].Text = runs[runs.Count - 1].Text.TrimEnd();
            }
            runs.RemoveAll(delegate(OcrRun r) { return r.Text.Length == 0; });
            return runs;
        }

        /// <summary>
        /// Есть ли пробел между соседними словами строки. По умолчанию — да (PdfPig отдал их
        /// как разные слова, значит это граница слова). Нет — только если токены почти
        /// соприкасаются (зазор &lt; SpaceGapFactor кегля): такой мизерный зазор означает, что
        /// PdfPig раздробил один токен на куски, их склеиваем без пробела. refSize — кегль
        /// текущего слова (иначе высота рамки), чтобы порог был относительным.
        /// </summary>
        private static bool HasSpaceBetween(PdfWord prev, PdfWord cur)
        {
            double refSize = cur.FontSizePt > 0 ? cur.FontSizePt : cur.Height;
            if (refSize <= 0) refSize = prev.Height;
            if (refSize <= 0) refSize = 1;
            return cur.Left - prev.Right >= SpaceGapFactor * refSize;
        }

        private static bool SameFormat(PdfWord a, PdfWord b)
        {
            return a.FontSizePt == b.FontSizePt && a.Bold == b.Bold
                && a.Italic == b.Italic && a.ColorArgb == b.ColorArgb
                && a.FontName == b.FontName && a.Super == b.Super && a.Sub == b.Sub
                && a.Underline == b.Underline && a.Uri == b.Uri;
        }

        /// <summary>
        /// Пометить надстрочные/подстрочные: слово мельче доминирующего кегля строки и его
        /// базовая линия заметно выше (над-) или ниже (под-) доминирующей. Медианы по строке
        /// устойчивы к самим скриптам. Меняет только флаги слов — чистая по смыслу.
        /// </summary>
        private static void MarkScripts(List<Line> lines)
        {
            foreach (Line ln in lines)
            {
                if (ln.Words.Count < 2)
                    continue;
                var heights = new List<double>(ln.Words.Count);
                var bottoms = new List<double>(ln.Words.Count);
                foreach (PdfWord w in ln.Words) { heights.Add(w.Height); bottoms.Add(w.Bottom); }
                double domH = Median(heights);
                double domBottom = Median(bottoms);
                if (domH <= 0)
                    continue;
                foreach (PdfWord w in ln.Words)
                {
                    if (w.Height >= ScriptSizeFactor * domH)
                        continue; // не мельче — не скрипт
                    double rise = w.Bottom - domBottom;
                    if (rise > ScriptRiseFactor * domH) w.Super = true;
                    else if (rise < -ScriptRiseFactor * domH) w.Sub = true;
                }
            }
        }

        /// <summary>
        /// Выравнивание абзаца из геометрии:
        ///  • Center — КАЖДАЯ строка с заметными и почти равными полями слева/справа
        ///    (номер страницы, заголовок по центру);
        ///  • Justify — многострочный абзац, где НЕпоследние строки достают до правого поля
        ///    (последняя строка justified-абзаца всегда рваная, её не учитываем);
        ///  • иначе Left (рваный справа / одиночная строка — визуально как по левому краю).
        /// </summary>
        private static OcrAlignment DetectAlignment(List<Line> group, double bodyLeft, double bodyRight, double width, double em)
        {
            double minGap = CenterMinGapFraction * width;
            double balance = CenterBalanceFraction * width;
            bool allCentered = true;
            foreach (Line ln in group)
            {
                double lg = ln.Left - bodyLeft, rg = bodyRight - ln.Right;
                if (!(lg > minGap && rg > minGap && Math.Abs(lg - rg) <= balance)) { allCentered = false; break; }
            }
            if (allCentered)
                return OcrAlignment.Center;

            if (group.Count >= 2)
            {
                double reachTol = 0.5 * em;
                int nonLast = group.Count - 1, reach = 0;
                for (int i = 0; i < nonLast; i++)
                    if (group[i].Right >= bodyRight - reachTol) reach++;
                if (reach >= (int)Math.Ceiling(nonLast * JustifiedShare))
                    return OcrAlignment.Justify;
            }
            return OcrAlignment.Left;
        }

        private sealed class Line
        {
            public readonly List<PdfWord> Words = new List<PdfWord>();
            public double MidY; // центр строки — по слову-затравке (самому верхнему)

            // Геометрия строки. Валидна после сортировки слов слева направо в ToLines:
            // Words[0] — самое левое; строка всегда непуста.
            public double Left { get { return Words[0].Left; } }

            public double Right
            {
                get
                {
                    double r = double.MinValue;
                    for (int i = 0; i < Words.Count; i++)
                        if (Words[i].Right > r) r = Words[i].Right;
                    return r;
                }
            }

            public double Height // кегль строки (высота самого крупного слова)
            {
                get
                {
                    double h = 0;
                    for (int i = 0; i < Words.Count; i++)
                        if (Words[i].Height > h) h = Words[i].Height;
                    return h;
                }
            }

            public double Top // верх строки (максимум по словам)
            {
                get
                {
                    double t = double.MinValue;
                    for (int i = 0; i < Words.Count; i++)
                        if (Words[i].Top > t) t = Words[i].Top;
                    return t;
                }
            }

            public double Bottom // низ строки (минимум по словам)
            {
                get
                {
                    double b = double.MaxValue;
                    for (int i = 0; i < Words.Count; i++)
                        if (Words[i].Bottom < b) b = Words[i].Bottom;
                    return b;
                }
            }
        }

        /// <summary>
        /// Принадлежит ли слово строке: рамки перекрываются по вертикали не меньше чем на
        /// половину меньшей высоты. Устойчиво к тонкой пунктуации (её маленькая рамка целиком
        /// лежит внутри рамки текста, поэтому перекрытие большое относительно её высоты).
        /// </summary>
        private static bool SameLine(Line line, PdfWord w)
        {
            double overlap = Math.Min(line.Top, w.Top) - Math.Max(line.Bottom, w.Bottom);
            double minH = Math.Min(line.Height, w.Height);
            if (minH <= 0) minH = 1;
            return overlap >= SameLineFactor * minH;
        }

        private static List<Line> ToLines(IList<PdfWord> words)
        {
            var result = new List<Line>();
            if (words == null || words.Count == 0)
                return result;

            var sorted = new List<PdfWord>(words);
            // Сверху вниз (MidY убывает), затем слева направо, затем по тексту (детерминизм).
            sorted.Sort(delegate(PdfWord a, PdfWord b)
            {
                int c = b.MidY.CompareTo(a.MidY);
                if (c != 0) return c;
                c = a.Left.CompareTo(b.Left);
                if (c != 0) return c;
                return string.CompareOrdinal(a.Text ?? "", b.Text ?? "");
            });

            Line line = null;
            foreach (PdfWord w in sorted)
            {
                if (line != null && SameLine(line, w))
                {
                    line.Words.Add(w);
                }
                else
                {
                    line = new Line { MidY = w.MidY };
                    line.Words.Add(w);
                    result.Add(line);
                }
            }

            // Слова внутри строки — строго слева направо.
            foreach (Line ln in result)
                ln.Words.Sort(delegate(PdfWord a, PdfWord b)
                {
                    int c = a.Left.CompareTo(b.Left);
                    return c != 0 ? c : string.CompareOrdinal(a.Text ?? "", b.Text ?? "");
                });
            return result;
        }

        /// <summary>
        /// Порог зазора для разрыва абзаца: типичный межстрочный зазор × коэффициент.
        /// Типичный — нижняя медиана зазоров (обычные строки плотнее абзацных разрывов).
        /// </summary>
        private static double ParagraphThreshold(List<Line> lines)
        {
            if (lines.Count < 2)
                return double.MaxValue; // одна строка — один абзац
            var gaps = new List<double>(lines.Count - 1);
            for (int i = 1; i < lines.Count; i++)
                gaps.Add(lines[i - 1].MidY - lines[i].MidY);
            gaps.Sort();
            double typical = gaps[(gaps.Count - 1) / 2];
            return typical * ParagraphGapFactor;
        }

        /// <summary>Текст кончается дефисом-переносом (дефис сразу после буквы).</summary>
        private static bool EndsWithHyphenAfterLetter(string s)
        {
            return s.Length >= 2 && s[s.Length - 1] == '-' && char.IsLetter(s[s.Length - 2]);
        }

        /// <summary>Начинается ли строка с маркера списка («1.», «•»). По первым словам (маркеру
        /// хватает начала строки), с одним пробелом между словами — маркеру этого достаточно.</summary>
        private static bool StartsWithListMarker(Line line)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < line.Words.Count; i++)
            {
                if (i > 0) sb.Append(' ');
                sb.Append(line.Words[i].Text ?? "");
                if (sb.Length > 8) break; // маркер распознаётся по самому началу — большего не нужно
            }
            return ListMarker.Detect(sb.ToString()).Kind != ListKind.None;
        }
    }
}
