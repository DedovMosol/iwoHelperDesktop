using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace ExcelMerger
{
    /// <summary>
    /// Разметка структуры страницы: какому логическому блоку (абзацу, пункту списка, ячейке)
    /// принадлежит каждая буква.
    ///
    /// Размеченный PDF — часть стандарта (ISO 32000, раздел 14.7; для доступности — PDF/UA):
    /// содержимое страницы обёрнуто в скобки «BDC … EMC» с номером элемента, а дерево
    /// /StructTreeRoot связывает эти номера с ролями (абзац, заголовок, пункт списка, ячейка).
    /// Word, PowerPoint и Acrobat пишут её всегда. Это значит, что во многих документах
    /// границы абзацев УЖЕ УКАЗАНЫ автором, и восстанавливать их по зазорам между строками —
    /// значит спорить с источником и ошибаться там, где ошибаться не обязательно.
    ///
    /// Здесь берётся только НОМЕР блока: он отвечает на вопрос «эти две строки — один абзац
    /// или разные», ради которого всё и затевалось. Роли (список, ячейка, заголовок) лежат в
    /// дереве структуры, которое разбирать пока не нужно.
    ///
    /// Чего этот номер НЕ решает: где строка стоит. Разметка описывает логику, а не геометрию,
    /// поэтому координаты по-прежнему берутся у букв, а номер блока только уточняет границы.
    ///
    /// Пометка «оформление» (Artifact) намеренно НЕ используется для отбрасывания текста:
    /// на реальных файлах ею оказывается помечен видимый текст (проверено — совпадений с
    /// дублями двойной отрисовки нет ни одного), и доверие к ней стирало бы содержимое.
    /// Буквы оформления просто остаются без номера и группируются по-старому.
    /// </summary>
    internal static class StructureBlocks
    {
        /// <summary>
        /// Карта «буква → номер блока» для страницы. Пустая, если разметки нет: тогда всё
        /// работает по-прежнему. Ключ — САМА буква: PdfPig отдаёт в словах те же объекты,
        /// что и в разметке, поэтому сравнение по ссылке точно и не требует поиска по координатам.
        /// </summary>
        public static Dictionary<UglyToad.PdfPig.Content.Letter, int> Map(UglyToad.PdfPig.Content.Page page)
        {
            var map = new Dictionary<UglyToad.PdfPig.Content.Letter, int>(ByReference.Instance);
            if (page == null)
                return map;
            IReadOnlyList<UglyToad.PdfPig.Content.MarkedContentElement> roots;
            try { roots = page.GetMarkedContents(); }
            catch { return map; } // разметка — подсказка, а не условие работы
            if (roots == null)
                return map;
            var stack = new Stack<UglyToad.PdfPig.Content.MarkedContentElement>();
            for (int i = roots.Count - 1; i >= 0; i--)
                stack.Push(roots[i]);
            while (stack.Count > 0)
            {
                UglyToad.PdfPig.Content.MarkedContentElement element = stack.Pop();
                if (element.Children != null)
                    foreach (UglyToad.PdfPig.Content.MarkedContentElement child in element.Children)
                        stack.Push(child);
                int id = element.MarkedContentIdentifier;
                if (id < 0 || element.Letters == null)
                    continue;
                foreach (UglyToad.PdfPig.Content.Letter letter in element.Letters)
                    map[letter] = id; // вложенная скобка уточняет внешнюю, поэтому перезапись верна
            }
            return map;
        }

        /// <summary>Номер блока слова — тот, что у большинства его букв; -1, если букв нет.</summary>
        public static int Of(Dictionary<UglyToad.PdfPig.Content.Letter, int> map,
            IReadOnlyList<UglyToad.PdfPig.Content.Letter> letters)
        {
            if (map == null || map.Count == 0 || letters == null || letters.Count == 0)
                return -1;
            int best = -1, bestCount = 0;
            var counts = new Dictionary<int, int>(4);
            foreach (UglyToad.PdfPig.Content.Letter letter in letters)
            {
                int id;
                if (!map.TryGetValue(letter, out id))
                    continue;
                int n;
                counts.TryGetValue(id, out n);
                n++;
                counts[id] = n;
                // Строго «больше»: при равенстве остаётся тот, что встретился раньше, — иначе
                // номер слова зависел бы от порядка перебора словаря.
                if (n > bestCount) { bestCount = n; best = id; }
            }
            return best;
        }

        /// <summary>Сравнение по ССЫЛКЕ: у букв нет равенства по значению, да оно и не нужно.</summary>
        private sealed class ByReference : IEqualityComparer<UglyToad.PdfPig.Content.Letter>
        {
            public static readonly ByReference Instance = new ByReference();

            public bool Equals(UglyToad.PdfPig.Content.Letter a, UglyToad.PdfPig.Content.Letter b)
            {
                return ReferenceEquals(a, b);
            }

            public int GetHashCode(UglyToad.PdfPig.Content.Letter letter)
            {
                return RuntimeHelpers.GetHashCode(letter);
            }
        }
    }
}
