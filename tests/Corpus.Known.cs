using System;
using System.Collections.Generic;
namespace ExcelMerger.Tests
{
    /// <summary>Известный дефект: в каком выводе, у какого семейства, какой вид и сколько документов.</summary>
    internal sealed class CKnown
    {
        public readonly string Output, Family, Aspect, Id;
        /// <summary>Документов семейства с этим видом и всего случаев в них (см. CFinding.Count).</summary>
        public readonly int Docs, Cases;
        public CKnown(string output, string family, string aspect, string id, int docs, int cases)
        {
            Output = output; Family = family; Aspect = aspect; Id = id; Docs = docs; Cases = cases;
        }
    }
    /// <summary>
    /// Реестр известных дефектов корпуса (25 семейств × 40 вариантов = 1000 PDF).
    /// Сверка строгая в обе стороны.
    /// </summary>
    internal static class CorpusKnown
    {
        internal const string Docx = "docx", Pptx = "pptx", Word = "word";
        private static CKnown K(string output, string family, string aspect, string id, int docs, int cases)
        {
            return new CKnown(output, family, aspect, id, docs, cases);
        }
        internal static readonly CKnown[] All =
        {
            K(Docx, "абзацы влево с отбивкой", "align-left", "RC-31", 17, 20),
            K(Docx, "абзацы влево с отбивкой", "merge", "RC-30", 2, 5),
            K(Docx, "абзацы влево с отбивкой", "order", "RC-14", 1, 6),
            K(Docx, "абзацы влево с отбивкой", "split", "RC-32", 2, 2),
            K(Docx, "абзацы по ширине с красной строкой", "split", "RC-32", 2, 2),
            K(Docx, "абзацы с полной последней строкой и малой отбивкой", "merge", "RC-30", 4, 13),
            K(Docx, "абзацы с полной последней строкой и малой отбивкой", "order", "RC-14", 13, 36),
            K(Docx, "абзацы с полной последней строкой и малой отбивкой", "split", "RC-32", 6, 6),
            K(Docx, "белый текст на светлой цветной плашке", "align-left", "RC-31", 24, 27),
            K(Docx, "белый текст на светлой цветной плашке", "merge", "RC-30", 20, 61),
            K(Docx, "белый текст на тёмной плашке", "align-left", "RC-31", 19, 21),
            K(Docx, "белый текст на тёмной плашке", "merge", "RC-30", 27, 74),
            K(Docx, "белый текст на тёмной плашке", "order", "RC-14", 4, 16),
            K(Docx, "белый текст на тёмной плашке", "split", "RC-32", 5, 5),
            K(Docx, "выключка: влево, по центру, вправо, по ширине", "align-right", "RC-15", 2, 2),
            K(Docx, "выключка: влево, по центру, вправо, по ширине", "merge", "RC-30", 1, 2),
            K(Docx, "выключка: влево, по центру, вправо, по ширине", "split", "RC-32", 5, 8),
            K(Docx, "две колонки", "order", "RC-14", 27, 456),
            K(Docx, "две колонки", "split", "RC-14", 26, 55),
            K(Docx, "две колонки с разным интервалом", "align-center", "RC-8", 1, 1),
            K(Docx, "две колонки с разным интервалом", "into-table", "RC-8", 1, 45),
            K(Docx, "две колонки с разным интервалом", "order", "RC-14", 27, 775),
            K(Docx, "две колонки с разным интервалом", "split", "RC-14", 34, 76),
            K(Docx, "две колонки с разным интервалом", "table-extra", "RC-8", 1, 1),
            K(Docx, "заголовки и текст", "into-table", "RC-8", 1, 17),
            K(Docx, "заголовки и текст", "order", "RC-14", 7, 25),
            K(Docx, "заголовки и текст", "split", "RC-32", 7, 7),
            K(Docx, "заголовки и текст", "table-extra", "RC-8", 1, 1),
            K(Docx, "зачёркнутые фразы", "order", "RC-14", 5, 11),
            K(Docx, "зачёркнутые фразы", "split", "RC-14", 2, 2),
            K(Docx, "колонтитулы и номера страниц", "lost", "RC-2", 1, 1),
            K(Docx, "колонтитулы и номера страниц", "split", "RC-32", 2, 2),
            K(Docx, "маркеры Word: Symbol, Courier «o», Wingdings", "list-centered", "RC-3", 28, 55),
            K(Docx, "маркеры Word: Symbol, Courier «o», Wingdings", "list-kind", "RC-4", 5, 8),
            K(Docx, "маркеры Word: Symbol, Courier «o», Wingdings", "list-level", "RC-9", 15, 28),
            K(Docx, "маркеры Word: Symbol, Courier «o», Wingdings", "split", "RC-3", 7, 14),
            K(Docx, "маркированный список в два уровня", "align-left", "RC-31", 4, 4),
            K(Docx, "маркированный список в два уровня", "list-centered", "RC-3", 29, 34),
            K(Docx, "маркированный список в два уровня", "list-level", "RC-9", 12, 15),
            K(Docx, "нумерованный список в два уровня", "align-left", "RC-31", 5, 5),
            K(Docx, "нумерованный список в два уровня", "list-centered", "RC-3", 24, 32),
            K(Docx, "нумерованный список в два уровня", "list-level", "RC-9", 10, 26),
            K(Docx, "нумерованный список в два уровня", "marker-lost", "RC-3", 1, 4),
            K(Docx, "нумерованный список в два уровня", "merge", "RC-30", 1, 2),
            K(Docx, "нумерованный список в два уровня", "split", "RC-32", 1, 1),
            K(Docx, "повёрнутая надпись на поле", "split", "RC-32", 1, 2),
            K(Docx, "подчёркнутые фразы", "order", "RC-14", 6, 19),
            K(Docx, "подчёркнутые фразы", "split", "RC-32", 1, 1),
            K(Docx, "полужирные, курсивные и цветные слова", "order", "RC-14", 2, 6),
            K(Docx, "полужирные, курсивные и цветные слова", "split", "RC-32", 2, 2),
            K(Docx, "разный кегль в абзаце", "order", "RC-14", 1, 2),
            K(Docx, "скан с невидимым текстовым слоем", "lost", "RC-2", 1, 1),
            K(Docx, "скан с невидимым текстовым слоем", "scan-duplicate", "RC-2", 40, 40),
            K(Docx, "списки с буквами, римскими и составными номерами", "list-centered", "RC-3", 22, 30),
            K(Docx, "списки с буквами, римскими и составными номерами", "list-level", "RC-9", 2, 3),
            K(Docx, "списки с буквами, римскими и составными номерами", "marker-lost", "RC-3", 6, 22),
            K(Docx, "списки с буквами, римскими и составными номерами", "order", "RC-14", 5, 16),
            K(Docx, "списки с буквами, римскими и составными номерами", "split", "RC-3", 4, 4),
            K(Docx, "таблица без линовки «метка — значение»", "align-left", "RC-31", 15, 15),
            K(Docx, "таблица без линовки «метка — значение»", "order", "RC-14", 4, 8),
            K(Docx, "таблица без линовки «метка — значение»", "split", "RC-32", 3, 3),
            K(Docx, "таблица без линовки «метка — значение»", "table-shape", "RC-8", 5, 5),
            K(Docx, "таблица с линовкой", "align-left", "RC-31", 25, 25),
            K(Docx, "таблица с линовкой", "merge", "RC-30", 12, 24),
            K(Docx, "три колонки", "align-justify", "RC-31", 3, 3),
            K(Docx, "три колонки", "extra", "RC-14", 1, 21),
            K(Docx, "три колонки", "lost", "RC-2", 1, 22),
            K(Docx, "три колонки", "order", "RC-14", 23, 769),
            K(Docx, "три колонки", "split", "RC-14", 39, 118),
            K(Pptx, "две колонки с разным интервалом", "pptx-table-extra", "RC-8", 1, 1),
            K(Pptx, "заголовки и текст", "pptx-table-extra", "RC-8", 1, 1),
            K(Pptx, "маркеры Word: Symbol, Courier «o», Wingdings", "pptx-marker-lost", "RC-4", 6, 6),
            K(Pptx, "скан с невидимым текстовым слоем", "pptx-lost", "WP-16", 1, 1),
            K(Pptx, "скан с невидимым текстовым слоем", "pptx-ocr-visible", "WP-16", 40, 40),
            K(Pptx, "таблица без линовки «метка — значение»", "pptx-table-shape", "RC-8", 5, 5),
        };
    }
}