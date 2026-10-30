using System;
using System.Collections.Generic;
using ExcelMerger;

namespace ExcelMerger.Tests
{
    /// <summary>Параметры варианта: шрифт, кегль, язык, лист, поля, интервал.</summary>
    internal sealed class CVariant
    {
        public int Index;
        public int Seed;
        public string Font;
        public double Size;
        public bool Russian;
        public string Paper;
        public double PageW, PageH;
        public double Spacing;
        public override string ToString()
        {
            return Paper + ", " + Font + " " + Size + ", " + (Russian ? "ru" : "en") + ", ×" + Spacing;
        }
    }

    /// <summary>Документ корпуса: семейство, вариант и сама модель. Известные дефекты — в CorpusKnown.</summary>
    internal sealed class CCase
    {
        public int Id; // порядковый номер в корпусе: стабильное имя файла
        public string Family;
        public CVariant Variant;
        public CDoc Doc;
        public string Name { get { return Family + " #" + Variant.Index + " (" + Variant + ")"; } }
    }

    internal static class CorpusCases
    {
        // Шрифты, которые есть в любой Windows 8.1+ (и на раннере CI).
        private static readonly string[] Fonts = { "Times New Roman", "Arial", "Georgia", "Verdana", "Tahoma", "Courier New" };
        private static readonly double[] Sizes = { 10, 11, 12, 14, 10.5, 9 };
        private static readonly double[] Spacings = { 1.0, 1.15, 1.5 };

        private delegate void Builder(CCase c, CText text);

        private static readonly KeyValuePair<string, Builder>[] Families =
        {
            Family("абзацы по ширине с красной строкой", PlainJustified),
            Family("абзацы влево с отбивкой", PlainLeft),
            Family("заголовки и текст", Headings),
            Family("выключка: влево, по центру, вправо, по ширине", AlignmentMix),
            Family("маркированный список в два уровня", Bullets),
            Family("нумерованный список в два уровня", Numbered),
            Family("маркеры Word: Symbol, Courier «o», Wingdings", WordBullets),
            Family("списки с буквами, римскими и составными номерами", LetterLists),
            Family("таблица с линовкой", RuledTable),
            Family("таблица с объединёнными ячейками", SpannedTable),
            Family("таблица без линовки «метка — значение»", UnruledTable),
            Family("две колонки", TwoColumns),
            Family("три колонки", ThreeColumns),
            Family("полужирные, курсивные и цветные слова", Formatting),
            Family("подчёркнутые фразы", UnderlinedPhrases),
            Family("зачёркнутые фразы", StruckPhrases),
            Family("белый текст на тёмной плашке", WhiteOnDark),
            Family("белый текст на светлой цветной плашке", WhiteOnLight),
            Family("повёрнутая надпись на поле", Rotated),
            Family("скан с невидимым текстовым слоем", Scan),
            Family("колонтитулы и номера страниц", HeaderFooter),
            Family("разный кегль в абзаце", MixedSizes),
            Family("две колонки с разным интервалом", UnevenColumns),
            Family("абзацы с полной последней строкой и малой отбивкой", FullLastLine),
            // Новые семейства — только в конец: номер семейства входит в сид и шрифт варианта,
            // вставка в середину пересоздала бы все последующие документы и весь реестр.
            Family("таблица с переносом строк в ячейках", WrappedTable)
        };

        private static KeyValuePair<string, Builder> Family(string name, Builder b)
        {
            return new KeyValuePair<string, Builder>(name, b);
        }

        /// <summary>Все документы корпуса: perFamily вариантов каждого семейства.</summary>
        public static List<CCase> All(int perFamily)
        {
            var result = new List<CCase>();
            for (int f = 0; f < Families.Length; f++)
                for (int i = 0; i < perFamily; i++)
                {
                    CVariant v = Variant(f, i);
                    var c = new CCase { Id = result.Count, Family = Families[f].Key, Variant = v, Doc = NewDoc(v) };
                    c.Doc.Name = c.Name;
                    Families[f].Value(c, new CText(v.Seed, v.Russian));
                    result.Add(c);
                }
            return result;
        }

        private static CVariant Variant(int family, int i)
        {
            var v = new CVariant
            {
                Index = i,
                Seed = 7919 * (family + 1) + 104729 * i,
                Font = Fonts[(i + family) % Fonts.Length],
                Size = Sizes[(i * 5 + family) % Sizes.Length],
                Russian = (i + family) % 3 != 2,
                Spacing = Spacings[(i + 2 * family) % Spacings.Length]
            };
            switch ((i + family) % 4)
            {
                case 0: case 1: v.Paper = "A4"; v.PageW = 595; v.PageH = 842; break;
                case 2: v.Paper = "Letter"; v.PageW = 612; v.PageH = 792; break;
                default: v.Paper = "A5"; v.PageW = 420; v.PageH = 595; break;
            }
            return v;
        }

        private static CDoc NewDoc(CVariant v)
        {
            var d = new CDoc { PageW = v.PageW, PageH = v.PageH };
            if (v.Paper == "A5")
            {
                d.MarginL = 48; d.MarginR = 40; d.MarginT = 48; d.MarginB = 48;
            }
            return d;
        }

        private static CStyle Body(CCase c)
        {
            return new CStyle(c.Variant.Font, c.Variant.Size);
        }

        private static CParagraph Para(CCase c, string text, CAlign align, CStyle style = null)
        {
            var p = new CParagraph { Align = align, LineSpacing = c.Variant.Spacing };
            p.Runs.Add(new CRun(text, style ?? Body(c)));
            return p;
        }

        // ---------- семейства ----------

        private static void PlainJustified(CCase c, CText t)
        {
            int n = t.Next(3, 7);
            for (int i = 0; i < n; i++)
            {
                CParagraph p = Para(c, t.Text(t.Next(2, 5)), CAlign.Justify);
                p.FirstIndent = c.Variant.Size * 3;
                c.Doc.Blocks.Add(p);
            }
        }

        private static void PlainLeft(CCase c, CText t)
        {
            int n = t.Next(3, 7);
            for (int i = 0; i < n; i++)
            {
                CParagraph p = Para(c, t.Text(t.Next(1, 4)), CAlign.Left);
                p.SpaceAfter = c.Variant.Size * 0.8;
                c.Doc.Blocks.Add(p);
            }
        }

        /// <summary>
        /// Абзацы без красной строки, последняя строка полная (распределённая выключка), отбивка
        /// небольшая: ни отступ, ни короткий хвост, ни умышленный перевод строки границу не
        /// выдают — только межабзацный зазор. Проверяет порог зазора абзаца сам по себе.
        /// </summary>
        private static void FullLastLine(CCase c, CText t)
        {
            double[] after = { 0.4, 0.55, 0.7 };
            int n = t.Next(4, 8);
            for (int i = 0; i < n; i++)
            {
                CParagraph p = Para(c, t.Text(t.Next(1, 4)), CAlign.Distribute);
                p.SpaceAfter = c.Variant.Size * after[c.Variant.Index % after.Length];
                c.Doc.Blocks.Add(p);
            }
        }

        private static void Headings(CCase c, CText t)
        {
            int sections = t.Next(2, 4);
            for (int s = 0; s < sections; s++)
            {
                CParagraph h = Para(c, t.Phrase(t.Next(2, 6)), s % 2 == 0 ? CAlign.Left : CAlign.Center,
                    Body(c).WithBold().WithSize(c.Variant.Size + 4));
                h.SpaceAfter = c.Variant.Size;
                c.Doc.Blocks.Add(h);
                for (int i = 0; i < t.Next(1, 3); i++)
                {
                    CParagraph p = Para(c, t.Text(t.Next(2, 4)), CAlign.Justify);
                    p.SpaceAfter = c.Variant.Size * 0.8;
                    c.Doc.Blocks.Add(p);
                }
            }
        }

        private static void AlignmentMix(CCase c, CText t)
        {
            CAlign[] order = { CAlign.Left, CAlign.Center, CAlign.Right, CAlign.Justify };
            for (int i = 0; i < 4; i++)
            {
                CAlign a = order[(i + c.Variant.Index) % 4];
                string text = a == CAlign.Center || a == CAlign.Right ? t.Phrase(t.Next(3, 7)) : t.Text(t.Next(2, 4));
                CParagraph p = Para(c, text, a);
                p.SpaceAfter = c.Variant.Size;
                c.Doc.Blocks.Add(p);
            }
        }

        private static CParagraph Item(CCase c, CText t, ListKind kind, int level, int number, string marker,
            CStyle markerStyle = null)
        {
            CParagraph p = Para(c, t.Text(t.Next(1, 3), 4, 12), CAlign.Left);
            p.List = kind;
            p.Level = level;
            p.Number = number;
            p.Marker = marker;
            p.MarkerStyle = markerStyle;
            p.LeftIndent = 18 + level * 24;
            p.Hang = 18;
            p.SpaceAfter = c.Variant.Size * 0.3;
            return p;
        }

        private static void Bullets(CCase c, CText t)
        {
            c.Doc.Blocks.Add(Para(c, t.Text(1), CAlign.Left));
            for (int i = 0; i < t.Next(3, 7); i++)
            {
                c.Doc.Blocks.Add(Item(c, t, ListKind.Bulleted, 0, 0, "•"));
                if (i % 2 == 1)
                    c.Doc.Blocks.Add(Item(c, t, ListKind.Bulleted, 1, 0, "–"));
            }
        }

        private static void Numbered(CCase c, CText t)
        {
            c.Doc.Blocks.Add(Para(c, t.Text(1), CAlign.Left));
            int n = t.Next(3, 7);
            for (int i = 1; i <= n; i++)
            {
                c.Doc.Blocks.Add(Item(c, t, ListKind.Numbered, 0, i, i + "."));
                if (i % 2 == 0)
                    for (int k = 1; k <= 2; k++)
                        c.Doc.Blocks.Add(Item(c, t, ListKind.Numbered, 1, k, k + ")"));
            }
        }

        private static void WordBullets(CCase c, CText t)
        {
            CStyle symbol = Body(c).WithFont("Symbol");
            CStyle courier = Body(c).WithFont("Courier New");
            CStyle wingdings = Body(c).WithFont("Wingdings");
            for (int i = 0; i < t.Next(2, 4); i++)
            {
                c.Doc.Blocks.Add(Item(c, t, ListKind.Bulleted, 0, 0, "\uF0B7", symbol));
                c.Doc.Blocks.Add(Item(c, t, ListKind.Bulleted, 1, 0, "o", courier));
                c.Doc.Blocks.Add(Item(c, t, ListKind.Bulleted, 2, 0, "\uF0A7", wingdings));
            }
        }

        private static void LetterLists(CCase c, CText t)
        {
            string[] letters = c.Variant.Russian ? new[] { "а)", "б)", "в)", "г)" } : new[] { "a)", "b)", "c)", "d)" };
            string[] roman = { "I.", "II.", "III.", "IV." };
            for (int i = 0; i < 3; i++)
                c.Doc.Blocks.Add(Item(c, t, ListKind.Numbered, 0, i + 1, letters[i]));
            for (int i = 0; i < 3; i++)
                c.Doc.Blocks.Add(Item(c, t, ListKind.Numbered, 0, i + 1, roman[i]));
            for (int i = 0; i < 3; i++)
                c.Doc.Blocks.Add(Item(c, t, ListKind.Numbered, 1, i + 1, "1." + (i + 1) + "."));
        }

        private static CCell Cell(string text, CStyle style)
        {
            var cell = new CCell();
            cell.Runs.Add(new CRun(text, style));
            return cell;
        }

        /// <summary>Сколько колонок помещается на лист при этом кегле (длинное слово — ~10 em).</summary>
        private static int MaxColumns(CCase c)
        {
            double content = c.Doc.PageW - c.Doc.MarginL - c.Doc.MarginR;
            return Math.Max(2, Math.Min(5, (int)(content / (c.Variant.Size * 10))));
        }

        private static CTable Table(CCase c, CText t, int rows, int cols, bool ruled)
        {
            cols = Math.Min(cols, MaxColumns(c));
            var table = new CTable { Ruled = ruled, SpaceAfter = c.Variant.Size };
            table.ColWidths = new double[cols];
            for (int k = 0; k < cols; k++)
                table.ColWidths[k] = 60 + t.Next(0, 80);
            for (int r = 0; r < rows; r++)
            {
                var row = new CCell[cols];
                for (int k = 0; k < cols; k++)
                    row[k] = Cell(r == 0 ? t.Phrase(t.Next(1, 3)) : t.Phrase(t.Next(1, 5)),
                        r == 0 ? Body(c).WithBold() : Body(c));
                table.Rows.Add(row);
            }
            return table;
        }

        /// <summary>
        /// Отбивка после абзаца: у половины вариантов её нет (стиль «Обычный» по ГОСТ —
        /// интервалы 0), у другой — как в шаблоне Word (8 pt при 11 pt).
        /// </summary>
        private static double After(CCase c)
        {
            return c.Variant.Index % 2 == 0 ? 0 : c.Variant.Size * 0.75;
        }

        private static CParagraph Spaced(CCase c, string text, CAlign align)
        {
            CParagraph p = Para(c, text, align);
            p.SpaceAfter = After(c);
            return p;
        }

        private static void RuledTable(CCase c, CText t)
        {
            c.Doc.Blocks.Add(Spaced(c, t.Text(1), CAlign.Left));
            CTable table = Table(c, t, t.Next(3, 8), t.Next(2, 6), true);
            if (c.Variant.Index % 3 == 1)
                table.HeaderFill = 0xD9E2F3;
            c.Doc.Blocks.Add(table);
            c.Doc.Blocks.Add(Para(c, t.Text(1), CAlign.Left));
        }

        /// <summary>
        /// Таблица с переносом текста в ячейках (описание, условие, примечание — обычная таблица
        /// договора): несколько строк в одной ячейке должны остаться одной ячейкой, а на слайде —
        /// разными строками одной ячейки, а не одним склеенным словом на стыке строк.
        /// </summary>
        private static void WrappedTable(CCase c, CText t)
        {
            int cols = Math.Min(3, MaxColumns(c));
            var table = new CTable { Ruled = true, SpaceAfter = c.Variant.Size, ColWidths = new double[cols] };
            for (int k = 0; k < cols; k++)
                table.ColWidths[k] = 100;
            for (int r = 0; r < t.Next(3, 5); r++)
            {
                var row = new CCell[cols];
                for (int k = 0; k < cols; k++)
                    row[k] = Cell(r == 0 ? t.Phrase(t.Next(1, 3)) : t.Phrase(t.Next(6, 13)),
                        r == 0 ? Body(c).WithBold() : Body(c));
                table.Rows.Add(row);
            }
            c.Doc.Blocks.Add(table);
        }

        private static void SpannedTable(CCase c, CText t)
        {
            CTable table = Table(c, t, 4, 4, true);
            // Шапка: первые две колонки объединены; первая колонка — на две строки.
            table.Rows[0][0].ColSpan = 2;
            table.Rows[0][1] = null;
            table.Rows[1][0].RowSpan = 2;
            table.Rows[2][0] = null;
            c.Doc.Blocks.Add(table);
        }

        private static void UnruledTable(CCase c, CText t)
        {
            c.Doc.Blocks.Add(Spaced(c, t.Text(1), CAlign.Left));
            var table = new CTable { Ruled = false, ColWidths = new double[] { 150, 200 }, SpaceAfter = c.Variant.Size };
            for (int r = 0; r < t.Next(4, 8); r++)
                table.Rows.Add(new[]
                {
                    Cell(t.Phrase(t.Next(1, 3)) + ":", Body(c)),
                    Cell(t.Phrase(t.Next(1, 3)), Body(c))
                });
            c.Doc.Blocks.Add(table);
        }

        private static void Columns(CCase c, CText t, int count) { Columns(c, t, count, false); }

        /// <summary>
        /// Колонки газеты или бюллетеня. uneven — у второй колонки свой кегль и интервал, так что
        /// базовые линии колонок не совпадают: строки двух колонок уже не лежат на одной высоте,
        /// и разделить их может только вертикальный разрез страницы, а не разрыв внутри строки.
        /// </summary>
        private static void Columns(CCase c, CText t, int count, bool uneven)
        {
            CParagraph title = Para(c, t.Phrase(t.Next(3, 6)), CAlign.Center, Body(c).WithBold().WithSize(c.Variant.Size + 3));
            title.SpaceAfter = c.Variant.Size;
            c.Doc.Blocks.Add(title);
            // Промежуток: у чётных — как у Word по умолчанию (1,25 см), у нечётных — узкий,
            // как columnsep в LaTeX (10 pt) и в двухколоночных статьях.
            var cols = new CColumns { Gap = c.Variant.Index % 2 == 0 ? 35.4 : 12 };
            bool small = c.Variant.Paper == "A5" || c.Variant.Size >= 12;
            double colW = (c.Doc.PageW - c.Doc.MarginL - c.Doc.MarginR - cols.Gap * (count - 1)) / count;
            CStyle style = Body(c).WithSize(Math.Min(c.Variant.Size, Math.Floor(colW / 11 * 2) / 2));
            for (int k = 0; k < count; k++)
            {
                var col = new List<CBlock>();
                bool other = uneven && k % 2 == 1;
                CStyle colStyle = other ? style.WithSize(style.Size - 1) : style;
                for (int i = 0; i < t.Next(2, small ? 3 : 4); i++)
                {
                    CParagraph p = Para(c, t.Text(t.Next(1, small ? 3 : 4)), CAlign.Justify, colStyle);
                    p.SpaceAfter = c.Variant.Size * 0.6;
                    if (other)
                        p.LineSpacing = c.Variant.Spacing + 0.35;
                    col.Add(p);
                }
                cols.Columns.Add(col);
            }
            c.Doc.Blocks.Add(cols);
        }

        private static void TwoColumns(CCase c, CText t) { Columns(c, t, 2); }

        private static void ThreeColumns(CCase c, CText t) { Columns(c, t, 3); }

        private static void UnevenColumns(CCase c, CText t) { Columns(c, t, 2, true); }

        /// <summary>Абзац, где отдельные слова получают свой формат.</summary>
        private static CParagraph Styled(CCase c, CText t, Func<int, CStyle, CStyle> styleOf)
        {
            var p = new CParagraph { Align = CAlign.Justify, LineSpacing = c.Variant.Spacing, SpaceAfter = c.Variant.Size };
            string[] words = t.Text(t.Next(2, 4)).Split(' ');
            for (int i = 0; i < words.Length; i++)
                p.Runs.Add(new CRun((i > 0 ? " " : "") + words[i], styleOf(i, Body(c))));
            return p;
        }

        private static void Formatting(CCase c, CText t)
        {
            c.Doc.Blocks.Add(Styled(c, t, (i, s) => i % 5 == 2 ? s.WithBold() : s));
            c.Doc.Blocks.Add(Styled(c, t, (i, s) => i % 4 == 1 ? s.WithItalic() : s));
            c.Doc.Blocks.Add(Styled(c, t, (i, s) => i % 6 == 3 ? s.WithColor(0xC00000) : (i % 7 == 5 ? s.WithColor(0x2F5597) : s)));
            c.Doc.Blocks.Add(Styled(c, t, (i, s) => i % 6 == 4 ? s.WithUnderline() : s));
        }

        private static void UnderlinedPhrases(CCase c, CText t)
        {
            for (int k = 0; k < 3; k++)
                c.Doc.Blocks.Add(Styled(c, t, (i, s) => i % 9 >= 3 && i % 9 <= 5 ? s.WithUnderline() : s));
        }

        private static void StruckPhrases(CCase c, CText t)
        {
            for (int k = 0; k < 3; k++)
                c.Doc.Blocks.Add(Styled(c, t, (i, s) => i % 8 >= 2 && i % 8 <= 4 ? s.WithStrike() : s));
        }

        private static void Shaded(CCase c, CText t, int fill)
        {
            CParagraph intro = Para(c, t.Text(1), CAlign.Left);
            intro.SpaceAfter = c.Variant.Size;
            c.Doc.Blocks.Add(intro);
            for (int i = 0; i < 2; i++)
            {
                CParagraph p = Para(c, t.Phrase(t.Next(3, 7)), CAlign.Left, Body(c).WithBold().WithColor(0xFFFFFF));
                p.Fill = fill;
                p.SpaceAfter = c.Variant.Size;
                c.Doc.Blocks.Add(p);
                CParagraph body = Para(c, t.Text(1), CAlign.Left);
                body.SpaceAfter = c.Variant.Size;
                c.Doc.Blocks.Add(body);
            }
        }

        private static void WhiteOnDark(CCase c, CText t) { Shaded(c, t, c.Variant.Index % 2 == 0 ? 0x1F3864 : 0x404040); }

        private static void WhiteOnLight(CCase c, CText t)
        {
            // Акцентные цвета шапок Word (синий, оранжевый, серый): сумма RGB больше 1,5 —
            // для разбора это «светлая» подложка.
            int[] fills = { 0x5B9BD5, 0xED7D31, 0xA5A5A5 };
            Shaded(c, t, fills[c.Variant.Index % fills.Length]);
        }

        private static void Rotated(CCase c, CText t)
        {
            for (int i = 0; i < 3; i++)
            {
                CParagraph p = Para(c, t.Text(t.Next(2, 4)), CAlign.Justify);
                p.FirstIndent = c.Variant.Size * 3;
                c.Doc.Blocks.Add(p);
            }
            c.Doc.Blocks.Add(new CRotated
            {
                Text = t.Phrase(3),
                Style = Body(c).WithSize(8),
                Angle = -90,
                X = c.Doc.MarginL * 0.5,
                Y = c.Doc.PageH * 0.6
            });
        }

        private static void Scan(CCase c, CText t)
        {
            var scan = new CScan { FontSize = 12 + c.Variant.Index % 3 * 2 };
            for (int i = 0; i < t.Next(3, 6); i++)
                scan.Lines.Add(t.Phrase(t.Next(3, 6)));
            c.Doc.Blocks.Add(scan);
        }

        private static void HeaderFooter(CCase c, CText t)
        {
            c.Doc.Header = t.Phrase(4);
            c.Doc.PageNumbers = true;
            for (int i = 0; i < t.Next(8, 14); i++)
            {
                CParagraph p = Para(c, t.Text(t.Next(2, 5)), CAlign.Justify);
                p.FirstIndent = c.Variant.Size * 3;
                c.Doc.Blocks.Add(p);
            }
        }

        private static void MixedSizes(CCase c, CText t)
        {
            for (int k = 0; k < 3; k++)
                c.Doc.Blocks.Add(Styled(c, t, (i, s) => i % 7 == 3 ? s.WithSize(s.Size + 2) : s));
        }
    }
}
