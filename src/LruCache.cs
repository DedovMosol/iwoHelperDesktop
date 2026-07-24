using System;
using System.Collections.Generic;

namespace ExcelMerger
{
    /// <summary>
    /// Ограниченный по размеру кэш «наименее недавно использованных» (LRU) со
    /// строковыми ключами. При переполнении вытесняется самый несвежий элемент, и для
    /// него вызывается <c>onEvict</c> (освобождение ресурса). Любое обращение
    /// (<see cref="TryGet"/> при попадании или <see cref="Add"/>) делает ключ самым
    /// свежим. Ключи сравниваются без учёта регистра (рассчитан на пути файлов).
    ///
    /// НЕ потокобезопасен: вызывающий обеспечивает доступ из одного потока
    /// (в приложении — только фоновый поток рендера миниатюр). Логика вытеснения
    /// чистая и покрыта юнит-тестами.
    /// </summary>
    internal sealed class LruCache<TValue>
    {
        private readonly int _capacity;
        private readonly Action<TValue> _onEvict;
        private readonly Dictionary<string, TValue> _map =
            new Dictionary<string, TValue>(StringComparer.OrdinalIgnoreCase);
        // Голова — самый свежий ключ, хвост — кандидат на вытеснение.
        private readonly LinkedList<string> _order = new LinkedList<string>();

        public LruCache(int capacity, Action<TValue> onEvict)
        {
            if (capacity < 1)
                throw new ArgumentOutOfRangeException("capacity", "ёмкость LRU должна быть не меньше 1");
            _capacity = capacity;
            _onEvict = onEvict;
        }

        /// <summary>Число хранимых элементов.</summary>
        public int Count { get { return _map.Count; } }

        /// <summary>Значение по ключу; при попадании ключ становится самым свежим.</summary>
        public bool TryGet(string key, out TValue value)
        {
            if (_map.TryGetValue(key, out value))
            {
                Touch(key);
                return true;
            }
            return false;
        }

        /// <summary>Значение по ключу БЕЗ освежения (порядок вытеснения не меняется).</summary>
        public bool TryPeek(string key, out TValue value)
        {
            return _map.TryGetValue(key, out value);
        }

        /// <summary>
        /// Изъять элемент БЕЗ вызова onEvict (освобождение — на вызывающем: он же
        /// выполняет и сопутствующую очистку). false — ключа нет.
        /// </summary>
        public bool Remove(string key, out TValue value)
        {
            if (!_map.TryGetValue(key, out value))
                return false;
            _map.Remove(key);
            LinkedListNode<string> node = Find(key);
            if (node != null)
                _order.Remove(node);
            return true;
        }

        /// <summary>Снимок ключей (для перебора с изменением кэша по ходу).</summary>
        public List<string> KeysSnapshot()
        {
            return new List<string>(_map.Keys);
        }

        /// <summary>
        /// Кладёт значение. Существующий ключ — замена значения и подъём в свежие
        /// (прежнее значение не освобождается: замена в приложении не используется).
        /// Новый ключ сверх ёмкости вытесняет самый несвежий элемент через onEvict.
        /// </summary>
        public void Add(string key, TValue value)
        {
            if (_map.ContainsKey(key))
            {
                _map[key] = value;
                Touch(key);
                return;
            }
            _map[key] = value;
            _order.AddFirst(key);
            if (_map.Count > _capacity)
                EvictOldest();
        }

        /// <summary>Освобождает все элементы (onEvict) и очищает кэш.</summary>
        public void Clear()
        {
            if (_onEvict != null)
                foreach (TValue value in _map.Values)
                    _onEvict(value);
            _map.Clear();
            _order.Clear();
        }

        private void EvictOldest()
        {
            LinkedListNode<string> oldest = _order.Last;
            if (oldest == null)
                return;
            _order.RemoveLast();
            TValue value;
            if (_map.TryGetValue(oldest.Value, out value))
            {
                _map.Remove(oldest.Value);
                if (_onEvict != null)
                    _onEvict(value);
            }
        }

        private void Touch(string key)
        {
            // Список не длиннее ёмкости (мал) — линейный поиск дешевле накладных
            // расходов на второй словарь узлов.
            LinkedListNode<string> node = Find(key);
            if (node != null && node != _order.First)
            {
                _order.Remove(node);
                _order.AddFirst(node);
            }
        }

        private LinkedListNode<string> Find(string key)
        {
            for (LinkedListNode<string> n = _order.First; n != null; n = n.Next)
                if (string.Equals(n.Value, key, StringComparison.OrdinalIgnoreCase))
                    return n;
            return null;
        }
    }
}
