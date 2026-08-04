using System;
using System.Collections.Generic;
using PdfSharp.Pdf;

namespace ExcelMerger
{
    /// <summary>
    /// Одна закладка оглавления: заголовок, страница-цель (индекс В СВОЁМ документе) и
    /// глубина вложенности (0 — верхний уровень). Плоский список с уровнями вместо дерева
    /// выбран намеренно: перенос закладок — это переназначение страниц, а дерево при этом
    /// только мешает (выпавший родитель не должен утаскивать за собой уцелевших детей).
    /// </summary>
    public sealed class PdfBookmark
    {
        public string Title;
        public int PageIndex;
        public int Level;
        public bool Opened;
    }

    /// <summary>
    /// Перенос оглавления из исходных документов в собранный. До 1.18.3 его не было вовсе:
    /// страницы копировались по одной, а закладки, заголовок и остальная «обвязка» документа
    /// оставались в источнике. Замер на файле с четырьмя страницами и тремя закладками —
    /// после объединения ноль закладок, после разделения ноль. Для документа с оглавлением
    /// это тихая потеря данных: файл открывается, выглядит целым, а навигация исчезла.
    ///
    /// Слой знает только про закладки и ничего — про то, ради чего собирают документ,
    /// поэтому им пользуются одинаково и объединение, и все четыре режима разделения.
    /// </summary>
    internal static class PdfBookmarks
    {
        /// <summary>
        /// Закладки документа плоским списком в порядке обхода в глубину. Закладка, чья цель
        /// не разрешается в страницу этого документа (битая ссылка, цель во внешнем файле),
        /// пропускается вместе со своей ветвью — вести читателя в никуда хуже, чем не вести.
        /// </summary>
        public static List<PdfBookmark> Read(PdfDocument doc)
        {
            var result = new List<PdfBookmark>();
            if (doc == null)
                return result;
            // Индекс страницы по идентификатору объекта: ссылочное равенство обёрток PdfPage
            // ненадёжно, тот же приём уже используется при разделении по закладкам.
            var indexByObject = new Dictionary<PdfObjectID, int>();
            for (int i = 0; i < doc.PageCount; i++)
            {
                PdfPage page = doc.Pages[i];
                if (page != null && page.Reference != null)
                    indexByObject[page.Reference.ObjectID] = i;
            }
            Collect(doc.Outlines, 0, indexByObject, result);
            return result;
        }

        private static void Collect(PdfOutlineCollection outlines, int level,
            Dictionary<PdfObjectID, int> indexByObject, List<PdfBookmark> into)
        {
            if (outlines == null)
                return;
            foreach (PdfOutline outline in outlines)
            {
                PdfPage dest = outline.DestinationPage;
                int index;
                if (dest == null || dest.Reference == null ||
                    !indexByObject.TryGetValue(dest.Reference.ObjectID, out index))
                    continue;
                into.Add(new PdfBookmark
                {
                    Title = outline.Title ?? string.Empty,
                    PageIndex = index,
                    Level = level,
                    Opened = outline.Opened
                });
                Collect(outline.Outlines, level + 1, indexByObject, into);
            }
        }

        /// <summary>
        /// Пересчёт закладок под новый состав документа. newIndexBySource — куда переехала
        /// каждая исходная страница (страницы, которой в результате нет, в карте нет).
        ///
        /// Закладка, чья страница не попала в результат, ОТБРАСЫВАЕТСЯ, а её дети
        /// поднимаются на освободившийся уровень: иначе половина оглавления исчезала бы
        /// вместе с одним удалённым разделом. Уровни после этого нормализуются — вложенность
        /// не может прыгнуть больше чем на ступень, иначе запись построит рваное дерево.
        /// Порядок обхода сохраняется: он и есть порядок оглавления, а не порядок страниц
        /// (при перемешивании страниц «содержание первого файла, затем второго» понятнее,
        /// чем список, скачущий по документу). Чистая — под тест.
        /// </summary>
        public static List<PdfBookmark> Remap(IList<PdfBookmark> marks, IDictionary<int, int> newIndexBySource)
        {
            var result = new List<PdfBookmark>();
            if (marks == null || newIndexBySource == null)
                return result;
            // Сколько уровней «съедено» выпавшими предками: пока мы внутри ветви выпавшей
            // закладки, её потомки поднимаются ровно на её уровень.
            var droppedAt = new List<int>();
            foreach (PdfBookmark mark in marks)
            {
                if (mark == null)
                    continue;
                // Вышли из ветви — забываем выпавших предков глубже текущего уровня.
                while (droppedAt.Count > 0 && droppedAt[droppedAt.Count - 1] >= mark.Level)
                    droppedAt.RemoveAt(droppedAt.Count - 1);
                int moved;
                if (!newIndexBySource.TryGetValue(mark.PageIndex, out moved))
                {
                    droppedAt.Add(mark.Level);
                    continue;
                }
                int level = mark.Level - droppedAt.Count;
                if (level < 0)
                    level = 0;
                // Ступенька за раз: после подъёма сирот уровень мог оторваться от предыдущего.
                int previous = result.Count == 0 ? -1 : result[result.Count - 1].Level;
                if (level > previous + 1)
                    level = previous + 1;
                result.Add(new PdfBookmark
                {
                    Title = mark.Title,
                    PageIndex = moved,
                    Level = level,
                    Opened = mark.Opened
                });
            }
            return result;
        }

        /// <summary>
        /// Записать плоский список с уровнями в документ. Закладка, указывающая за пределы
        /// документа, пропускается — записать её значило бы получить файл, который читатель
        /// объявит повреждённым.
        /// </summary>
        public static void Write(PdfDocument doc, IList<PdfBookmark> marks)
        {
            if (doc == null || marks == null || marks.Count == 0)
                return;
            // Стек коллекций по уровням: на уровне 0 пишем в оглавление документа, глубже —
            // в коллекцию последней закладки предыдущего уровня.
            var parents = new List<PdfOutlineCollection> { doc.Outlines };
            foreach (PdfBookmark mark in marks)
            {
                if (mark == null || mark.PageIndex < 0 || mark.PageIndex >= doc.PageCount)
                    continue;
                int level = mark.Level < 0 ? 0 : mark.Level;
                if (level >= parents.Count)
                    level = parents.Count - 1;   // рваный уровень прижимаем к достижимому
                PdfOutline added = parents[level].Add(mark.Title ?? string.Empty,
                    doc.Pages[mark.PageIndex], mark.Opened);
                parents.RemoveRange(level + 1, parents.Count - level - 1);
                parents.Add(added.Outlines);
            }
        }

        /// <summary>
        /// Карта «исходная страница → её место в результате» для подряд идущего диапазона.
        /// Чистая — под тест.
        /// </summary>
        public static Dictionary<int, int> RangeMap(int start, int end)
        {
            var map = new Dictionary<int, int>();
            for (int i = start; i <= end; i++)
                map[i] = i - start;
            return map;
        }
    }
}
