using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Порядок страниц будущего PDF: добавление документов, перестановка
    /// и удаление. Чистая модель без UI и PDFsharp — покрыта юнит-тестами.
    /// Перестановка делегируется общему <see cref="ListReorder"/>.
    /// </summary>
    public class PdfPageOrder
    {
        private const int UndoLimit = 50; // хватает на сессию правок и не копит память

        /// <summary>Снимок жеста: порядок ссылок И их повороты (повороты мутируются в общих ссылках).</summary>
        private sealed class Snapshot
        {
            public List<PdfPageRef> Items;
            public List<int> Rotations;
        }

        private readonly List<PdfPageRef> _items = new List<PdfPageRef>();
        private readonly List<Snapshot> _undo = new List<Snapshot>();
        private readonly List<Snapshot> _redo = new List<Snapshot>();

        public int Count
        {
            get { return _items.Count; }
        }

        public PdfPageRef this[int index]
        {
            get { return _items[index]; }
        }

        /// <summary>Есть что откатывать (Ctrl+Z).</summary>
        public bool CanUndo
        {
            get { return _undo.Count > 0; }
        }

        /// <summary>Есть что возвращать (Ctrl+Y после отката).</summary>
        public bool CanRedo
        {
            get { return _redo.Count > 0; }
        }

        private Snapshot Capture()
        {
            var s = new Snapshot { Items = new List<PdfPageRef>(_items), Rotations = new List<int>(_items.Count) };
            foreach (PdfPageRef r in _items)
                s.Rotations.Add(r.Rotation);
            return s;
        }

        private void Restore(Snapshot s)
        {
            _items.Clear();
            _items.AddRange(s.Items);
            for (int i = 0; i < s.Items.Count; i++)
                s.Items[i].Rotation = s.Rotations[i]; // повороты возвращаются в те же общие ссылки
        }

        private static void Push(List<Snapshot> stack, Snapshot s)
        {
            stack.Add(s);
            if (stack.Count > UndoLimit)
                stack.RemoveAt(0);
        }

        /// <summary>
        /// Снимок ПЕРЕД пользовательским жестом (один жест — один снимок: формы зовут перед
        /// добавлением, переносом, удалением, вставкой и поворотом). Новый жест обнуляет
        /// ветку возврата (Ctrl+Y), как в любом редакторе. Стеки ограничены.
        /// </summary>
        public void Checkpoint()
        {
            Push(_undo, Capture());
            _redo.Clear();
        }

        /// <summary>Откатить последний жест (Ctrl+Z): порядок, состав И повороты. false — нечего.</summary>
        public bool Undo()
        {
            if (_undo.Count == 0)
                return false;
            Push(_redo, Capture());
            Snapshot s = _undo[_undo.Count - 1];
            _undo.RemoveAt(_undo.Count - 1);
            Restore(s);
            return true;
        }

        /// <summary>Вернуть откаченный жест (Ctrl+Y). false — нечего.</summary>
        public bool Redo()
        {
            if (_redo.Count == 0)
                return false;
            Push(_undo, Capture());
            Snapshot s = _redo[_redo.Count - 1];
            _redo.RemoveAt(_redo.Count - 1);
            Restore(s);
            return true;
        }

        /// <summary>Очистить список (например, при открытии другого документа). Обе истории тоже очищаются.</summary>
        public void Clear()
        {
            _items.Clear();
            _undo.Clear();
            _redo.Clear();
        }

        /// <summary>Все страницы документа добавляются в конец списка по порядку.</summary>
        public void AddDocument(string path, int pageCount)
        {
            InsertDocument(Count, path, pageCount);
        }

        /// <summary>
        /// Страницы документа вставляются подряд ПЕРЕД позицией insertAt (за пределами
        /// списка — прижимается к краю). Возвращает позицию первой вставленной страницы.
        /// </summary>
        public int InsertDocument(int insertAt, string path, int pageCount)
        {
            if (insertAt < 0) insertAt = 0;
            if (insertAt > _items.Count) insertAt = _items.Count;
            for (int i = 0; i < pageCount; i++)
            {
                var page = new PdfPageRef();
                page.SourcePath = path;
                page.PageIndex = i;
                _items.Insert(insertAt + i, page);
            }
            return insertAt;
        }

        /// <summary>
        /// Вставка готовых страниц (вставка из буфера) ПЕРЕД позицией insertAt.
        /// Возвращает позицию первой вставленной страницы.
        /// </summary>
        public int InsertAt(int insertAt, IList<PdfPageRef> pages)
        {
            if (insertAt < 0) insertAt = 0;
            if (insertAt > _items.Count) insertAt = _items.Count;
            if (pages != null)
                for (int i = 0; i < pages.Count; i++)
                    _items.Insert(insertAt + i, pages[i]);
            return insertAt;
        }

        /// <summary>
        /// Перенос набора страниц (вырезать → вставить) ПЕРЕД позицией insertAt в исходной
        /// нумерации. Возвращает позицию первой перенесённой страницы (-1 — пустой набор).
        /// </summary>
        public int MoveRange(IList<int> indices, int insertAt)
        {
            return ListReorder.MoveRange(_items, indices, insertAt);
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

        public void RemoveAt(IList<int> indices)
        {
            ListReorder.RemoveAt(_items, indices);
        }

        public List<PdfPageRef> ToList()
        {
            return new List<PdfPageRef>(_items);
        }

        /// <summary>
        /// Заменить порядок готовой ПЕРЕСТАНОВКОЙ тех же страниц (чередование пачек).
        /// В отличие от <see cref="Clear"/> не трогает стопки отмены: снимок для Ctrl+Z
        /// вызывающий делает сам через <see cref="Checkpoint"/> перед вызовом.
        /// </summary>
        public void SetOrder(IList<PdfPageRef> pages)
        {
            _items.Clear();
            if (pages != null)
                _items.AddRange(pages);
        }
    }
}
