using System;
using System.IO;
using System.Reflection;

namespace ExcelMerger
{
    /// <summary>
    /// Загрузка сборок, вшитых в exe ресурсами (PdfSharp.dll): наружу
    /// по-прежнему поставляется один файл. Ensure() обязан выполняться до
    /// JIT-компиляции любого метода, чьё тело ссылается на типы PdfSharp, —
    /// поэтому публичные методы PdfMergeService не содержат таких типов
    /// и вызывают [NoInlining]-ядра только после Ensure().
    /// </summary>
    internal static class EmbeddedAssemblies
    {
        private static readonly object Sync = new object();
        private static bool _registered;
        private static Assembly _pdfSharp; // Assembly.Load(byte[]) кэшируем сами: CLR не дедуплицирует

        public static void Ensure()
        {
            lock (Sync)
            {
                if (_registered)
                    return;
                AppDomain.CurrentDomain.AssemblyResolve += OnResolve;
                _registered = true;
            }
        }

        private static Assembly OnResolve(object sender, ResolveEventArgs args)
        {
            if (!string.Equals(new AssemblyName(args.Name).Name, "PdfSharp", StringComparison.OrdinalIgnoreCase))
                return null;
            lock (Sync)
            {
                if (_pdfSharp != null)
                    return _pdfSharp;
                using (Stream stream = typeof(EmbeddedAssemblies).Assembly.GetManifestResourceStream("PdfSharp.dll"))
                {
                    if (stream == null)
                        return null;
                    var bytes = new byte[stream.Length];
                    int done = 0;
                    while (done < bytes.Length)
                    {
                        int read = stream.Read(bytes, done, bytes.Length - done);
                        if (read <= 0)
                            break;
                        done += read;
                    }
                    _pdfSharp = Assembly.Load(bytes);
                    return _pdfSharp;
                }
            }
        }
    }
}
