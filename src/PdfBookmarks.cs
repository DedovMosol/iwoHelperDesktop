using System;
using System.Collections.Generic;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;

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
        /// не разрешается в страницу этого документа (битая ссылка, цель во внешнем файле,
        /// «папка» без цели вовсе), пропускается — вести читателя в никуда хуже, чем не вести.
        /// Но только она сама: её дети поднимаются на её уровень, по тому же правилу, что и в
        /// <see cref="Remap"/>. Раньше вместе с ней молча исчезала вся ветвь — а «папка» без
        /// цели над главами — обычное устройство оглавления, и терялось почти всё.
        ///
        /// Дерево читается из сырых словарей, а не через <c>PdfDocument.Outlines</c>: PDFsharp
        /// 1.50 строит его целиком и БРОСАЕТ на любой цели, кроме массива прямо на месте, —
        /// на именованном назначении <c>/A &lt;&lt;/S /GoTo /D (имя)&gt;&gt;</c>, которое пишет
        /// LaTeX (hyperref), и на номере страницы числом. Падало тогда всё объединение или
        /// разделение, а закладки <c>/Dest (имя)</c> пропадали молча. Здесь разрешаются все виды
        /// цели из ISO 32000-1 §12.3.2: массив, ссылка на него, имя (каталог <c>/Dests</c>) и
        /// строка (дерево имён <c>/Names /Dests</c>), в том числе словарь назначения с <c>/D</c>.
        /// Оглавление — обвязка документа: испорченная закладка теряется сама, но никогда не
        /// роняет операцию над страницами.
        ///
        /// Читается документ, открытый из файла (так его открывают объединение и разделение).
        /// У оглавления, построенного в памяти через doc.Outlines, сырых /First и /Next нет до
        /// сохранения — их пишет запись, — и такой документ читается пустым.
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
            PdfDictionary catalog = doc.Internals.Catalog;
            PdfDictionary root = Deref(catalog.Elements["/Outlines"]) as PdfDictionary;
            if (root == null)
                return result;
            var reader = new OutlineReader(catalog, indexByObject, doc.PageCount);
            reader.Collect(Deref(root.Elements["/First"]) as PdfDictionary, result);
            return result;
        }

        /// <summary>
        /// Обход сырого дерева закладок. Все обращения — к элементам словарей напрямую
        /// (<c>Elements[key]</c>), без <c>GetValue</c>/<c>GetDictionary</c>: те по метаданным
        /// PDFsharp превращают словарь закладки в <c>PdfOutline</c>, а его конструктор и есть
        /// то место, что бросает. Цикл ссылок в испорченном файле, бесконечная цепочка
        /// <c>/Next</c> и сверхглубокая вложенность ограничены, чтобы обход всегда кончался.
        /// </summary>
        private sealed class OutlineReader
        {
            private const int MaxItems = 100000;   // больше закладок реальный документ не несёт
            private const int MaxDepth = 64;       // и глубже не вкладывает (глубина рекурсии, см. Collect)
            private const int MaxNameNodes = 100000;

            private readonly PdfDictionary _catalog;
            private readonly Dictionary<PdfObjectID, int> _indexByObject;
            private readonly int _pageCount;
            private readonly HashSet<object> _visited = new HashSet<object>();
            private Dictionary<string, PdfItem> _namedDests; // дерево имён /Names /Dests, лениво

            public OutlineReader(PdfDictionary catalog, Dictionary<PdfObjectID, int> indexByObject, int pageCount)
            {
                _catalog = catalog;
                _indexByObject = indexByObject;
                _pageCount = pageCount;
            }

            public void Collect(PdfDictionary first, List<PdfBookmark> into)
            {
                Collect(first, 0, 0, into);
            }

            /// <summary>
            /// level — уровень закладки в результате, depth — глубина рекурсии. Они расходятся:
            /// дети выпавшей закладки встают на её уровень, а стек при этом всё равно растёт.
            /// Ограничение по level цепочку вложенных «папок» без цели не остановило бы, и
            /// испорченный файл в несколько тысяч уровней переполнил бы стек потока —
            /// StackOverflowException не перехватывается и завершает процесс.
            /// </summary>
            private void Collect(PdfDictionary item, int level, int depth, List<PdfBookmark> into)
            {
                if (depth >= MaxDepth)
                    return;
                for (; item != null; item = Deref(item.Elements["/Next"]) as PdfDictionary)
                {
                    // Цикл /Next или /First на уже пройденную закладку — конец этой цепочки.
                    if (_visited.Count >= MaxItems || !_visited.Add(item))
                        return;
                    PdfDictionary firstChild = Deref(item.Elements["/First"]) as PdfDictionary;
                    int index = TargetPage(item);
                    if (index < 0)
                    {
                        // Сама закладка выпадает, дети встают на её место — тем же уровнем.
                        Collect(firstChild, level, depth + 1, into);
                        continue;
                    }
                    into.Add(new PdfBookmark
                    {
                        Title = StringOf(item.Elements["/Title"]),
                        PageIndex = index,
                        Level = level,
                        // ISO 32000-1 §12.3.3: положительный /Count — ветвь раскрыта.
                        Opened = IntegerOf(item.Elements["/Count"]) > 0
                    });
                    Collect(firstChild, level + 1, depth + 1, into);
                }
            }

            /// <summary>Индекс страницы цели закладки в этом документе; -1 — цели нет.</summary>
            private int TargetPage(PdfDictionary item)
            {
                try
                {
                    PdfItem dest = item.Elements["/Dest"];
                    if (dest == null)
                    {
                        PdfDictionary action = Deref(item.Elements["/A"]) as PdfDictionary;
                        // Только переход внутри документа: /GoToR ведёт в другой файл.
                        if (action == null || NameOf(action.Elements["/S"]) != "/GoTo")
                            return -1;
                        dest = action.Elements["/D"];
                    }
                    return PageOfDestination(dest, 0);
                }
                catch (Exception ex) when (MergeException.ShouldWrap(ex))
                {
                    return -1; // испорченная цель — выпадает только эта закладка
                }
            }

            private int PageOfDestination(PdfItem dest, int hops)
            {
                // Имя ведёт к назначению, назначение может оказаться словарём с /D: две ступени.
                if (hops > 2)
                    return -1;
                PdfItem value = Deref(dest);
                PdfArray array = value as PdfArray;
                if (array != null)
                    return PageOfArray(array);
                PdfDictionary dictionary = value as PdfDictionary;
                if (dictionary != null)
                    return PageOfDestination(dictionary.Elements["/D"], hops + 1);
                string name = NameOf(value);
                if (name != null)
                    return PageOfDestination(LookupName(name.TrimStart('/')), hops + 1);
                string text = StringOrNull(value);
                if (text != null)
                    return PageOfDestination(LookupName(text), hops + 1);
                return -1;
            }

            private int PageOfArray(PdfArray array)
            {
                if (array.Elements.Count == 0)
                    return -1;
                PdfItem page = array.Elements[0];
                PdfReference reference = page as PdfReference;
                int index;
                if (reference != null)
                    return _indexByObject.TryGetValue(reference.ObjectID, out index) ? index : -1;
                // Внутри документа цель обязана быть ссылкой на страницу, но генераторы пишут и
                // номер числом; как и pdf.js, понимаем его как индекс с нуля.
                if (!IsInteger(page))
                    return -1;
                int number = IntegerOf(page);
                return number >= 0 && number < _pageCount ? number : -1;
            }

            /// <summary>
            /// Именованное назначение: сначала дерево имён <c>/Names /Dests</c> (PDF 1.2+, ключи —
            /// строки), затем словарь каталога <c>/Dests</c> (PDF 1.1, ключи — имена). Генераторы
            /// путают строку и имя, поэтому имя любого вида ищется в обоих.
            /// </summary>
            private PdfItem LookupName(string name)
            {
                if (_namedDests == null)
                    _namedDests = BuildNameTree();
                PdfItem found;
                if (_namedDests.TryGetValue(name, out found))
                    return found;
                PdfDictionary dests = Deref(_catalog.Elements["/Dests"]) as PdfDictionary;
                return dests == null ? null : dests.Elements["/" + name];
            }

            private Dictionary<string, PdfItem> BuildNameTree()
            {
                var map = new Dictionary<string, PdfItem>(StringComparer.Ordinal);
                PdfDictionary names = Deref(_catalog.Elements["/Names"]) as PdfDictionary;
                PdfDictionary root = names == null ? null : Deref(names.Elements["/Dests"]) as PdfDictionary;
                if (root == null)
                    return map;
                var seen = new HashSet<object>();
                var pending = new Stack<PdfDictionary>();
                pending.Push(root);
                while (pending.Count > 0 && seen.Count < MaxNameNodes)
                {
                    PdfDictionary node = pending.Pop();
                    if (!seen.Add(node))
                        continue;
                    PdfArray pairs = Deref(node.Elements["/Names"]) as PdfArray;
                    if (pairs != null)
                        for (int i = 0; i + 1 < pairs.Elements.Count; i += 2)
                        {
                            string key = StringOrNull(Deref(pairs.Elements[i]));
                            if (key != null && !map.ContainsKey(key))
                                map.Add(key, pairs.Elements[i + 1]);
                        }
                    PdfArray kids = Deref(node.Elements["/Kids"]) as PdfArray;
                    if (kids != null)
                        for (int i = kids.Elements.Count - 1; i >= 0; i--)
                        {
                            PdfDictionary kid = Deref(kids.Elements[i]) as PdfDictionary;
                            if (kid != null)
                                pending.Push(kid);
                        }
                }
                return map;
            }
        }

        private static PdfItem Deref(PdfItem item)
        {
            PdfReference reference = item as PdfReference;
            return reference != null ? reference.Value : item;
        }

        private static string StringOrNull(PdfItem item)
        {
            item = Deref(item);
            PdfString text = item as PdfString;
            if (text != null)
                return text.Value;
            PdfStringObject textObject = item as PdfStringObject;
            return textObject != null ? textObject.Value : null;
        }

        private static string StringOf(PdfItem item)
        {
            return StringOrNull(item) ?? string.Empty;
        }

        private static string NameOf(PdfItem item)
        {
            // Только проверка типа через is: PdfName и PdfNameObject перегружают
            // operator !=(…, string), и «name != null» сравнивает name.Value со строкой —
            // NullReferenceException на любом не-имени (строке, массиве).
            item = Deref(item);
            if (item is PdfName)
                return ((PdfName)item).Value;
            if (item is PdfNameObject)
                return ((PdfNameObject)item).Value;
            return null;
        }

        private static bool IsInteger(PdfItem item)
        {
            item = Deref(item);
            return item is PdfInteger || item is PdfIntegerObject;
        }

        private static int IntegerOf(PdfItem item)
        {
            item = Deref(item);
            PdfInteger number = item as PdfInteger;
            if (number != null)
                return number.Value;
            PdfIntegerObject numberObject = item as PdfIntegerObject;
            return numberObject != null ? numberObject.Value : 0;
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
                // Добавляем закрытой и раскрываем потом: Add с opened=true увеличивает у предков
                // счётчик PDFsharp, и при сохранении тот затирает /Count своим значением (см. ниже).
                PdfOutline added = parents[level].Add(mark.Title ?? string.Empty,
                    doc.Pages[mark.PageIndex], false);
                added.Opened = mark.Opened;
                parents.RemoveRange(level + 1, parents.Count - level - 1);
                parents.Add(added.Outlines);
            }
            if (doc.Outlines.Count > 0)
                doc.Outlines[0].Parent.Elements["/Count"] = new PdfInteger(WriteCounts(doc.Outlines));
        }

        /// <summary>
        /// Дописать закладку верхнего уровня в конец оглавления документа, открытого для
        /// правки (так разделение по закладкам подписывает часть, чья глава выпала из сетки).
        /// Закладка добавляется свёрнутой по той же причине, что и в <see cref="Write"/>: иначе
        /// PDFsharp при сохранении затёр бы /Count корня своим счётчиком. /Count корня
        /// пересчитывается по сырым значениям: у загруженных закладок PDFsharp раскрытость не
        /// читает (Opened всегда false), а их /Count в словаре остаётся как был записан.
        /// </summary>
        public static void AppendTopLevel(PdfDocument doc, string title, int pageIndex)
        {
            if (doc == null || pageIndex < 0 || pageIndex >= doc.PageCount)
                return;
            PdfOutline added = doc.Outlines.Add(title ?? string.Empty, doc.Pages[pageIndex], false);
            int visible = 0;
            foreach (PdfOutline item in doc.Outlines)
            {
                visible++;
                int below = IntegerOf(item.Elements["/Count"]);
                if (below > 0)
                    visible += below; // раскрытая: под ней видно ровно столько
            }
            added.Parent.Elements["/Count"] = new PdfInteger(visible);
        }

        /// <summary>
        /// /Count по ISO 32000-1 §12.3.3 (таблицы 152, 153): у раскрытой закладки — сколько
        /// потомков видно под ней, у свёрнутой — минус столько, сколько стало бы видно, раскрой
        /// её; у корня — сколько закладок видно всего. Возвращает число видимых закладок
        /// коллекции, когда её владелец раскрыт.
        ///
        /// PDFsharp 1.50 считает иначе — число РАСКРЫТЫХ потомков — и пишет /Count только если оно
        /// больше нуля. Глава с листьями-разделами (листья раскрытыми не бывают) уходила
        /// в файл без /Count, то есть свёрнутой. Поэтому закладки добавляются с opened=false:
        /// счётчик PDFsharp остаётся нулём, и записанное здесь значение доживает до файла.
        /// Глубина рекурсии — глубина оглавления, а её при чтении ограничивает MaxDepth.
        /// </summary>
        private static int WriteCounts(PdfOutlineCollection items)
        {
            int visible = 0;
            foreach (PdfOutline item in items)
            {
                visible++;
                if (!item.HasChildren)
                    continue;
                int below = WriteCounts(item.Outlines);
                item.Elements["/Count"] = new PdfInteger(item.Opened ? below : -below);
                if (item.Opened)
                    visible += below;
            }
            return visible;
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
