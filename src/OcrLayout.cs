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
        public double TopPt;    // верх абзаца (Y первой строки, ось вверх) — для порядка с изображениями
        public double LeftPt;   // левый край абзаца — вторичный порядок (левее — раньше в одной строке-полосе)
        public double RightPt;  // правый край рамки абзаца — для XY-порядка блоков страницы (колонки)
        public double BottomPt; // нижний край рамки абзаца

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
    /// Порядок чтения born-digital PDF: слова с рамками → блоки страницы (XY-разрез:
    /// этажи и колонки, <see cref="XyCut"/>) → строки → абзацы (сверху вниз, слева направо;
    /// левая колонка целиком раньше правой). Одноколоночная страница — один блок, разбор
    /// эквивалентен прежнему построчному. Чистая логика без типов PdfPig — покрыта юнит-тестами.
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
        // Центрированная строка: отступает от ОБОИХ полей (больше CenterInsetEmFactor кегля И
        // CenterMinInsetFraction ширины) и почти симметрична (|левый−правый| <= CenterBalanceFraction
        // ширины). Так узнаём и узкую (номер страницы), и ШИРОКУЮ (строка титула) центрированную
        // строку: у широкой зазоры малы в долях ширины, но симметричны, поэтому порог задан ещё и в
        // долях кегля. Отличие от красной строки — та упирается в правое поле (правый зазор ≈ 0); от
        // рваной левой — та стоит у левого поля (левый зазор ≈ 0).
        private const double CenterInsetEmFactor = 0.5;
        private const double CenterMinInsetFraction = 0.02;
        private const double CenterBalanceFraction = 0.08;
        // Прогон по общей оси центра: соседние строки соосны, если их midX расходится не больше чем на
        // эту долю кегля. Красная строка/короткий хвост justified-абзаца смещают midX сильнее — прогон
        // на них обрывается, поэтому тело не «прилипает» к центрированному титулу.
        private const double MidAxisTolEmFactor = 0.4;
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
        // Разрез страницы на блоки (XyCut). Просвет колонок — в долях кегля (у шапки письма
        // канал ≈29–42pt при кегле ~11, у justified-текста межсловные зазоры < 0.4 кегля —
        // запас многократный). Колонка обязана быть высокой (не «подпись … дата» одной строки)
        // и непустой. Нижний порог пустой полосы этажа — от вырожденных метрик.
        private const double ColumnGapEmFactor = 2.5;
        private const double ColumnMinExtentEmFactor = 2.2;
        private const int ColumnMinWords = 2;
        private const double MinCutGapPt = 2.0;
        // Порог пустой полосы этажа — от ТИПИЧНОГО просвета между рамками соседних строк
        // (не от шага базовых линий: при разнородных кеглях шапки и тела связь «шаг − высота»
        // рассыпается, и порог задирается выше зазоров между зонами шапки — у письма зоны
        // 18.7–26.4pt при межстрочных 12.8pt). Плюс нижняя планка в долях кегля: этаж,
        // нарезанный ПО СТРОКАМ плотного текста, схлопнул бы зону колонок (в этаже из одной
        // строки вертикальному просвету не хватает высоты MinColumnExtent).
        private const double FloorGapFactor = 1.35;
        private const double FloorMinEmFactor = 1.2;

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
            return Analyze(words, true);
        }

        /// <summary>
        /// Разбор с управлением колонками: содержимое ЯЧЕЙКИ таблицы разбирается БЕЗ поиска
        /// колонок (splitColumns=false) — у ячейки многоколоночной вёрстки не бывает, а «метка
        /// слева … число справа» через широкий зазор — одна строка, которую вертикальный разрез
        /// растащил бы на столбики («Итого:»-метки отдельно от сумм). Этажи (горизонтальные
        /// разрезы) остаются — они совпадают с разрывами абзацев и безвредны.
        /// </summary>
        internal static OcrPageLayout Analyze(IList<PdfWord> words, bool splitColumns)
        {
            var result = new OcrPageLayout();
            if (words == null || words.Count == 0)
                return result;

            // Метрики страницы для порогов разреза — по строкам ВСЕЙ страницы. В многоколоночной
            // вёрстке строки соседних колонок перемежаются и типичный шаг занижается — это лишь
            // добавляет горизонтальных разрезов (безвредно: порядок этажей совпадает с «сверху
            // вниз»), а колонки ищутся внутри каждого этажа по своим просветам.
            List<Line> pageLines = ToLines(words);
            var pageHeights = new List<double>(pageLines.Count);
            foreach (Line ln in pageLines)
                pageHeights.Add(ln.Height);
            double em = MathUtil.Median(pageHeights);
            if (em <= 0) em = 1;
            // Порог этажа — по типичному просвету между рамками строк: этаж появляется на
            // полосах заметно шире межстрочных (границы зон и абзацев). Разрез, совпадающий с
            // разрывом абзаца, безвреден — те же абзацы получаются поблочно (покрыто тестами).
            double hGapMin = pageLines.Count < 2
                ? double.MaxValue
                : Math.Max(MinCutGapPt,
                    Math.Max(FloorGapFactor * TypicalLineGap(pageLines), FloorMinEmFactor * em));

            var boxes = new CutBox[words.Count];
            for (int i = 0; i < words.Count; i++)
            {
                PdfWord w = words[i];
                boxes[i] = new CutBox { Left = w.Left, Right = w.Right, Bottom = w.Bottom, Top = w.Top, Tag = i };
            }
            // Запрет колонок = «бесконечный» просвет: этажи режутся как обычно, колонки — никогда.
            double vGapMin = splitColumns ? ColumnGapEmFactor * em : double.MaxValue;
            List<CutLeaf> leaves = XyCut.Order(boxes, hGapMin, vGapMin,
                ColumnMinExtentEmFactor * em, ColumnMinWords);

            // Красная строка решается по ВСЕЙ странице (группы всех блоков разом): у отдельного
            // этажа/колонки абзацев мало, порознь порог IndentedShare не набирался бы.
            var stats = new IndentStats();
            foreach (CutLeaf leaf in leaves)
            {
                var regionWords = new List<PdfWord>(leaf.Tags.Count);
                foreach (int tag in leaf.Tags)
                    regionWords.Add(words[tag]);
                AnalyzeRegion(ToLines(regionWords), leaf.ColumnLeft, leaf.ColumnRight,
                    result.Paragraphs, stats);
            }
            result.FirstLineIndentPt = stats.DocumentIndent();
            return result;
        }

        /// <summary>
        /// Сбор отступов первых строк по странице для решения о красной строке. Основной
        /// признак — доля отступов среди JUSTIFIED-групп: красная строка — атрибут выключенного
        /// по ширине тела, а шапки/адресаты (короткие Left-строки) её долю иначе размывают —
        /// двухколоночное письмо теряло отступ тела из-за полутора десятков строк бланка.
        /// Фолбэк — прежняя доля по всем группам: рваный справа (ragged) документ с отступами
        /// не имеет justified-групп, но отступ у него настоящий.
        /// </summary>
        private sealed class IndentStats
        {
            public int Groups;            // все группы страницы
            public int JustifiedGroups;   // из них — с выключкой по ширине
            public readonly List<double> Indents = new List<double>();          // отступы некентрированных групп
            public readonly List<double> JustifiedIndents = new List<double>(); // из них — у justified-групп

            /// <summary>Отступ красной строки страницы; 0 — не воспроизводим (см. описание класса).</summary>
            public double DocumentIndent()
            {
                if (JustifiedGroups >= 2 && JustifiedIndents.Count >= (int)Math.Ceiling(JustifiedGroups * IndentedShare))
                    return MathUtil.Median(JustifiedIndents);
                if (Groups >= 2 && Indents.Count >= (int)Math.Ceiling(Groups * IndentedShare))
                    return MathUtil.Median(Indents);
                return 0;
            }
        }

        /// <summary>
        /// Разбор строк одного блока (этаж/колонка XY-разреза) в абзацы (порядок чтения внутри
        /// блока). colLeft/colRight — рамка КОЛОНКИ блока: центрирование, красная строка и
        /// выключка считаются от полей колонки, а не от рамки самого блока — узкий блок (одна
        /// центрированная строка титула, отрезанная этажом) иначе терял бы центрирование.
        /// Абзацы добавляются в paragraphs; отступы первых строк и число групп копятся в stats
        /// (красную строку решает Analyze по всем блокам страницы).
        /// </summary>
        private static void AnalyzeRegion(List<Line> lines, double colLeft, double colRight,
            List<OcrParagraph> paragraphs, IndentStats stats)
        {
            if (lines.Count == 0)
                return;
            MarkScripts(lines); // пометить надстрочные/подстрочные слова

            // Геометрия блока: поля колонки, типичный кегль блока, признак выключки.
            double bodyLeft = colLeft, bodyRight = colRight;
            var heights = new List<double>(lines.Count);
            foreach (Line ln in lines)
                heights.Add(ln.Height);
            double em = MathUtil.Median(heights);
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

            // Центрированность строк (НА УРОВНЕ ПРОГОНА по общей оси центра, см. CenteredRuns) —
            // сигнал группировки: каждую центрированную строку выносим в СВОЙ абзац, сохраняя
            // исходную разбивку титула/шапки (двухстрочное название остаётся двумя строками — Word не
            // сольёт их в одну), и отделяем центрированный блок от обычного текста. При нулевой отбивке
            // абзацев соседние центрированные абзацы выглядят как исходные центрированные строки.
            bool[] centered = CenteredRuns(lines, bodyLeft, bodyRight, em, width, gapThreshold);

            // Разбить на группы строк-абзацев. Сигналы разрыва: большой зазор; центрированная строка
            // (выносится своим абзацем — и её граница с обычным текстом); красная строка и «короткий
            // хвост» justified; начало пункта списка («1.», «•») — иначе плотный одностроковый список
            // слипся бы в абзац с единственным маркером. groupCentered хранит однородный флаг
            // центрирования группы (центрированная строка изолирована, поэтому группа однородна).
            var groups = new List<List<Line>>();
            var groupCentered = new List<bool>();
            var current = new List<Line>();
            for (int i = 0; i < lines.Count; i++)
            {
                if (i > 0)
                {
                    bool gapBreak = lines[i - 1].MidY - lines[i].MidY > gapThreshold;
                    bool centeredBreak = centered[i] || centered[i - 1]; // центрированная строка — своим абзацем
                    bool indentBreak = lines[i].Left - bodyLeft > indentTol;
                    bool shortBreak = justified && lines[i - 1].Right < bodyRight - shortTol;
                    bool listBreak = StartsWithListMarker(lines[i]);
                    if (gapBreak || centeredBreak || indentBreak || shortBreak || listBreak)
                    {
                        groups.Add(current);
                        groupCentered.Add(centered[i - 1]);
                        current = new List<Line>();
                    }
                }
                current.Add(lines[i]);
            }
            groups.Add(current);
            groupCentered.Add(centered[lines.Count - 1]);

            // Абзацы: раны с форматом + выравнивание; сбор отступов первых строк (исключая центрирование).
            double maxIndent = MaxIndentFraction * width;
            for (int gi = 0; gi < groups.Count; gi++)
            {
                List<Line> g = groups[gi];
                double right = double.MinValue, bottom = double.MaxValue;
                foreach (Line ln in g)
                {
                    if (ln.Right > right) right = ln.Right;
                    if (ln.Bottom < bottom) bottom = ln.Bottom;
                }
                var para = new OcrParagraph
                {
                    Runs = BuildRuns(g),
                    // Центрированность решена на уровне прогона; иначе — выключка или левый край.
                    Alignment = groupCentered[gi] ? OcrAlignment.Center : DetectAlignment(g, bodyRight, em),
                    TopPt = g[0].Top,
                    LeftPt = g[0].Left,
                    RightPt = right,
                    BottomPt = bottom
                };
                if (para.Alignment != OcrAlignment.Center)
                {
                    ListMarker.Result m = ListMarker.Detect(para.Text);
                    para.ListKind = m.Kind;
                    para.ListNumber = m.Number;
                    para.ListContentStart = m.ContentStart;
                }
                paragraphs.Add(para);
                // Красную строку меряем ТОЛЬКО по нецентрированным абзацам: у центрированного блока
                // левый край — это центрирование, а не отступ, и он исказил бы медиану красной строки.
                if (para.Alignment == OcrAlignment.Justify)
                    stats.JustifiedGroups++;
                double ind = g[0].Left - bodyLeft;
                if (!groupCentered[gi] && ind > indentTol && ind <= maxIndent)
                {
                    stats.Indents.Add(ind);
                    if (para.Alignment == OcrAlignment.Justify)
                        stats.JustifiedIndents.Add(ind);
                }
            }
            stats.Groups += groups.Count;
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
                double domH = MathUtil.Median(heights);
                double domBottom = MathUtil.Median(bottoms);
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
        /// Пометить строки, входящие в центрированный прогон. Прогон — максимальная цепочка соседних
        /// (сверху вниз) строк с общей осью центра (|Δ midX| ≤ MidAxisTolEmFactor кегля) без разрыва
        /// абзаца (вертикальный зазор ≤ gapThreshold). Прогон считаем центрированным, если в нём есть
        /// хотя бы одна «плавающая» строка (<see cref="IsCentered"/>): тогда центрированы и его ШИРОКИЕ
        /// строки, дотягивающиеся до полей (внешне неотличимые от выключки) — общая ось с короткой
        /// центрированной строкой блока это доказывает. Обычный justified-абзац так не опознаётся: его
        /// короткий последний хвост прижат влево, а красная строка сдвигает ось полных строк — в обоих
        /// случаях ось сбивается и прогон обрывается на границе с телом. Чистая — под тест.
        /// </summary>
        private static bool[] CenteredRuns(List<Line> lines, double bodyLeft, double bodyRight, double em, double width, double gapThreshold)
        {
            int n = lines.Count;
            var centered = new bool[n];
            double midTol = MidAxisTolEmFactor * em;
            int start = 0;
            while (start < n)
            {
                int end = start; // расширяем прогон [start..end] по общей оси, пока нет разрыва абзаца
                while (end + 1 < n
                    && Math.Abs(MidX(lines[end + 1]) - MidX(lines[end])) <= midTol
                    && lines[end].MidY - lines[end + 1].MidY <= gapThreshold)
                    end++;
                bool anyFloating = false;
                for (int i = start; i <= end && !anyFloating; i++)
                    anyFloating = IsCentered(lines[i].Left - bodyLeft, bodyRight - lines[i].Right, em, width);
                if (anyFloating)
                    for (int i = start; i <= end; i++)
                        centered[i] = true;
                start = end + 1;
            }
            return centered;
        }

        private static double MidX(Line line) { return (line.Left + line.Right) / 2; }

        /// <summary>
        /// «Строка плавает по центру»: отступает от ОБОИХ полей больше порога (доля кегля ИЛИ доля
        /// ширины) и почти симметрична (левый и правый зазоры близки). Ловит и узкую (номер
        /// страницы), и ШИРОКУЮ (строка титула) центрированную строку. Отличие от красной строки —
        /// та упирается в правое поле (rightGap ≈ 0); от рваной левой — та у левого поля (leftGap ≈ 0).
        /// leftGap/rightGap — зазоры до полей тела (pt). Чистая — под тест.
        /// </summary>
        internal static bool IsCentered(double leftGap, double rightGap, double em, double width)
        {
            double minInset = Math.Max(CenterInsetEmFactor * em, CenterMinInsetFraction * width);
            double balance = CenterBalanceFraction * width;
            return leftGap > minInset && rightGap > minInset && Math.Abs(leftGap - rightGap) <= balance;
        }

        /// <summary>
        /// Выравнивание НЕцентрированного абзаца из геометрии (центрирование решено раньше — на
        /// уровне прогона по общей оси, см. <see cref="CenteredRuns"/>):
        ///  • Justify — многострочный абзац, где НЕпоследние строки достают до правого поля
        ///    (последняя строка justified-абзаца всегда рваная, её не учитываем);
        ///  • иначе Left (рваный справа / одиночная строка — визуально как по левому краю).
        /// </summary>
        private static OcrAlignment DetectAlignment(List<Line> group, double bodyRight, double em)
        {
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

        internal sealed class Line
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

        internal static List<Line> ToLines(IList<PdfWord> words)
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
        /// Типичный ПРОСВЕТ между рамками соседних строк (низ верхней − верх нижней), нижняя
        /// медиана положительных: перекрытия (строки соседних колонок на близких высотах) и
        /// слипшиеся рамки в счёт не идут. 0 — просветов нет (одна строка/сплошные перекрытия).
        /// </summary>
        private static double TypicalLineGap(List<Line> lines)
        {
            var gaps = new List<double>(lines.Count);
            for (int i = 1; i < lines.Count; i++)
            {
                double gap = lines[i - 1].Bottom - lines[i].Top;
                if (gap > 0)
                    gaps.Add(gap);
            }
            return MathUtil.Median(gaps);
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
