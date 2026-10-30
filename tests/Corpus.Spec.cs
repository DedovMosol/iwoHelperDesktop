using System;
using System.Collections.Generic;
using System.Text;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    // Корпус разметки (этап 0 плана распознавания). Документ описывается моделью ниже, из неё
    // рисуется настоящий PDF (Corpus.Render) и из неё же — то, что разбор обязан увидеть
    // (Corpus.Check). Одно описание на обе стороны: ожидание не может разойтись с рисунком.

    /// <summary>
    /// Выключка абзаца. Distribute — по ширине вместе с последней строкой (распределённая
    /// выключка Word): абзац кончается ПОЛНОЙ строкой, и границу абзаца выдаёт только отбивка.
    /// </summary>
    internal enum CAlign { Left, Right, Center, Justify, Distribute }

    /// <summary>Формат отрезка текста. Неизменяемый: варианты получают копии через With*.</summary>
    internal sealed class CStyle
    {
        public readonly string Font;
        public readonly double Size;
        public readonly bool Bold, Italic, Underline, Strike;
        public readonly int Color; // 0xRRGGBB

        public CStyle(string font, double size, bool bold = false, bool italic = false,
            bool underline = false, bool strike = false, int color = 0)
        {
            Font = font; Size = size; Bold = bold; Italic = italic;
            Underline = underline; Strike = strike; Color = color;
        }

        public CStyle WithBold(bool v = true) { return new CStyle(Font, Size, v, Italic, Underline, Strike, Color); }
        public CStyle WithItalic(bool v = true) { return new CStyle(Font, Size, Bold, v, Underline, Strike, Color); }
        public CStyle WithUnderline(bool v = true) { return new CStyle(Font, Size, Bold, Italic, v, Strike, Color); }
        public CStyle WithStrike(bool v = true) { return new CStyle(Font, Size, Bold, Italic, Underline, v, Color); }
        public CStyle WithColor(int c) { return new CStyle(Font, Size, Bold, Italic, Underline, Strike, c); }
        public CStyle WithSize(double s) { return new CStyle(Font, s, Bold, Italic, Underline, Strike, Color); }
        public CStyle WithFont(string f) { return new CStyle(f, Size, Bold, Italic, Underline, Strike, Color); }
    }

    internal sealed class CRun
    {
        public readonly string Text;
        public readonly CStyle Style;
        public CRun(string text, CStyle style) { Text = text; Style = style; }
    }

    internal abstract class CBlock
    {
        public double SpaceAfter;
    }

    /// <summary>Абзац или пункт списка (List != None).</summary>
    internal sealed class CParagraph : CBlock
    {
        public readonly List<CRun> Runs = new List<CRun>();
        public CAlign Align = CAlign.Justify;
        public double FirstIndent;   // красная строка (pt)
        public double LeftIndent;    // отступ всего абзаца слева (pt)
        public double LineSpacing = 1.15;
        public int Fill = -1;        // заливка под абзацем (0xRRGGBB), -1 — нет
        // Список: маркер рисуется на LeftIndent, текст и продолжения — на LeftIndent + Hang.
        public ListKind List = ListKind.None;
        public int Level;
        public int Number;
        public string Marker;
        public string MarkerText;    // как маркер видит читатель (знак Юникода значка Word); null — как Marker
        public CStyle MarkerStyle;   // null — стиль первого рана
        public double Hang = 18;
    }

    internal sealed class CCell
    {
        public readonly List<CRun> Runs = new List<CRun>();
        public int ColSpan = 1;
        public int RowSpan = 1;
    }

    /// <summary>Таблица: ячейки по строкам (накрытые объединением позиции — null).</summary>
    internal sealed class CTable : CBlock
    {
        public readonly List<CCell[]> Rows = new List<CCell[]>();
        public double[] ColWidths;
        public bool Ruled = true;
        public int HeaderFill = -1;
        public double LineWidth = 0.75;
    }

    /// <summary>Колонки газетной вёрстки: каждая колонка — свой список блоков.</summary>
    internal sealed class CColumns : CBlock
    {
        public readonly List<List<CBlock>> Columns = new List<List<CBlock>>();
        public double Gap = 24;
    }

    /// <summary>Надпись, повёрнутая на Angle (по часовой), на поле страницы — поток не занимает.</summary>
    internal sealed class CRotated : CBlock
    {
        public string Text;
        public CStyle Style;
        public double Angle = -90;
        public double X, Y; // точка начала (сверху вниз, как в PdfSharp)
    }

    /// <summary>Страница-скан: картинка с текстом и невидимый текстовый слой поверх (OCR).</summary>
    internal sealed class CScan : CBlock
    {
        public readonly List<string> Lines = new List<string>();
        public double FontSize = 14;
    }

    internal sealed class CPageBreak : CBlock { }

    internal sealed class CDoc
    {
        public string Name;
        public double PageW = 595, PageH = 842;
        public double MarginL = 70, MarginR = 56, MarginT = 64, MarginB = 64;
        public string Header;        // колонтитул сверху каждой страницы; null — нет
        public bool PageNumbers;     // номер страницы по центру снизу
        public readonly List<CBlock> Blocks = new List<CBlock>();
    }

    /// <summary>
    /// Детерминированный набор слов и фраз. Слова без цифр и маркеров в начале — иначе
    /// случайная фраза сама стала бы «пунктом списка».
    /// </summary>
    internal sealed class CText
    {
        private static readonly string[] Ru = (
            "документ договор сторона условие порядок оплата поставка товар срок качество работа " +
            "услуга заказчик исполнитель акт приёмка объём стоимость расчёт налог период отчёт проект " +
            "план задача решение вопрос предложение письмо ответ основание пункт раздел приложение " +
            "подпись печать дата номер адрес организация предприятие отдел сотрудник руководитель " +
            "специалист бухгалтерия склад продукция материал оборудование ремонт обслуживание контроль " +
            "проверка результат анализ данные таблица перечень требование стандарт норма правило " +
            "инструкция регламент процедура согласование утверждение изменение дополнение уведомление " +
            "обязательство ответственность гарантия претензия возврат замена доставка транспорт " +
            "маршрут график обязан выполнить предоставить обеспечить подготовить направить согласовать " +
            "принять оплатить проверить указанный настоящий соответствующий необходимый дополнительный " +
            "основной текущий годовой ежемесячный письменный электронный полный своевременный " +
            "надлежащий в на по для при от до из за с о к и или также если").Split(' ');
        private static readonly string[] En = (
            "document agreement party condition procedure payment delivery goods term quality work " +
            "service customer contractor statement acceptance volume price calculation tax period report " +
            "project plan task decision question proposal letter answer reason clause section annex " +
            "signature stamp date number address company enterprise department employee manager " +
            "specialist accounting warehouse product material equipment repair maintenance control " +
            "inspection result analysis data table list requirement standard norm rule instruction " +
            "regulation approval change addition notice obligation liability warranty claim return " +
            "replacement transport route schedule must perform provide ensure prepare send approve accept " +
            "pay verify specified present relevant necessary additional main current annual monthly " +
            "written electronic complete timely proper in on by for at from to of with and or also if").Split(' ');

        private readonly Random _rnd;
        private readonly string[] _bank;
        public readonly bool Russian;

        public CText(int seed, bool russian)
        {
            _rnd = new Random(seed);
            Russian = russian;
            _bank = russian ? Ru : En;
        }

        public int Next(int min, int maxExclusive) { return _rnd.Next(min, maxExclusive); }

        /// <summary>Слово из набора, не короче minLength знаков.</summary>
        public string Word(int minLength = 1)
        {
            for (int i = 0; i < 100; i++)
            {
                string w = _bank[_rnd.Next(_bank.Length)];
                if (w.Length >= minLength)
                    return w;
            }
            return Russian ? "документ" : "document";
        }

        /// <summary>Предложение из words слов: с заглавной, с точкой; изредка — запятая.</summary>
        public string Sentence(int words)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < words; i++)
            {
                string w = Word(i == 0 ? 3 : 1);
                if (i == 0)
                    w = char.ToUpperInvariant(w[0]) + w.Substring(1);
                if (i > 0)
                    sb.Append(' ');
                sb.Append(w);
                if (i > 1 && i < words - 2 && _rnd.Next(8) == 0)
                    sb.Append(',');
            }
            sb.Append('.');
            return sb.ToString();
        }

        /// <summary>Несколько предложений подряд — текст абзаца.</summary>
        public string Text(int sentences, int minWords = 6, int maxWords = 14)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < sentences; i++)
            {
                if (i > 0)
                    sb.Append(' ');
                sb.Append(Sentence(_rnd.Next(minWords, maxWords + 1)));
            }
            return sb.ToString();
        }

        /// <summary>Короткая фраза без точки (заголовок, ячейка, пункт).</summary>
        public string Phrase(int words)
        {
            string s = Sentence(words);
            return s.Substring(0, s.Length - 1).Replace(",", "");
        }
    }
}
