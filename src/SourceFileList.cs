using System.Collections.Generic;
using System.IO;

namespace ExcelMerger
{
    /// <summary>Файл-источник свода: путь и признак включения в слияние.</summary>
    public class SourceFile
    {
        public string Path;
        public bool Include = true;

        public string FileName
        {
            get { return System.IO.Path.GetFileName(Path); }
        }
    }

    /// <summary>
    /// Список файлов свода до слияния: порядок (перестановка/перенос) и
    /// включение/исключение. Чистая модель без UI — покрыта юнит-тестами.
    /// </summary>
    public class SourceFileList
    {
        private readonly List<SourceFile> _items = new List<SourceFile>();

        public int Count
        {
            get { return _items.Count; }
        }

        public SourceFile this[int index]
        {
            get { return _items[index]; }
        }

        /// <summary>Заполнить список путями (в заданном порядке), все включены.</summary>
        public void SetFiles(IEnumerable<string> paths)
        {
            _items.Clear();
            foreach (string p in paths)
            {
                var f = new SourceFile();
                f.Path = p;
                _items.Add(f);
            }
        }

        public int MoveUp(int index)
        {
            return ListReorder.MoveUp(_items, index);
        }

        public int MoveDown(int index)
        {
            return ListReorder.MoveDown(_items, index);
        }

        public void Move(int from, int to)
        {
            ListReorder.Move(_items, from, to);
        }

        public void SetAllIncluded(bool included)
        {
            foreach (SourceFile f in _items)
                f.Include = included;
        }

        /// <summary>Сортировка по имени файла как в Проводнике («Отчёт 2» раньше «Отчёт 10»).</summary>
        public void SortByName()
        {
            _items.Sort(delegate(SourceFile a, SourceFile b)
            {
                return NaturalStringComparer.Instance.Compare(a.FileName, b.FileName);
            });
        }

        public int IncludedCount
        {
            get
            {
                int n = 0;
                foreach (SourceFile f in _items)
                    if (f.Include)
                        n++;
                return n;
            }
        }

        /// <summary>Включённые файлы в текущем порядке — то, что уйдёт в слияние.</summary>
        public List<string> IncludedInOrder()
        {
            var list = new List<string>();
            foreach (SourceFile f in _items)
                if (f.Include)
                    list.Add(f.Path);
            return list;
        }
    }
}
