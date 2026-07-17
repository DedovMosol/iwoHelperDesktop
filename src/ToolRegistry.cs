using System.Collections.Generic;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>
    /// Реестр открытых окон-инструментов по ключу. Позволяет держать несколько
    /// разных инструментов открытыми одновременно, но не допускать двух копий
    /// одного. Закрытые (Disposed) окна считаются отсутствующими. Только UI-поток.
    /// </summary>
    public sealed class ToolRegistry
    {
        private readonly Dictionary<string, Form> _open =
            new Dictionary<string, Form>();

        /// <summary>Открыт ли инструмент с этим ключом (живое, не закрытое окно).</summary>
        public bool TryGetOpen(string key, out Form form)
        {
            if (_open.TryGetValue(key, out form))
            {
                if (!form.IsDisposed)
                    return true;
                _open.Remove(key); // окно уже закрыто — забываем
            }
            form = null;
            return false;
        }

        public void Add(string key, Form form)
        {
            _open[key] = form;
        }

        public void Remove(string key)
        {
            _open.Remove(key);
        }

        /// <summary>Живые открытые окна (для корректного закрытия при выходе).</summary>
        public List<Form> OpenForms()
        {
            var list = new List<Form>();
            foreach (Form f in _open.Values)
                if (f != null && !f.IsDisposed)
                    list.Add(f);
            return list;
        }
    }
}
