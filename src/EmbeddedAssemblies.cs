using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace ExcelMerger
{
    /// <summary>
    /// Загрузка сборок, вшитых в exe ресурсами (PdfSharp, PdfPig и его зависимости):
    /// наружу по-прежнему поставляется один файл. Резолв по простому имени —
    /// версионно-независим, поэтому заодно снимает нужду в binding-redirect для
    /// net48-полифиллов (System.Memory / System.Runtime.CompilerServices.Unsafe и др.).
    /// Ensure() обязан выполняться до JIT-компиляции любого метода, чьё тело ссылается
    /// на вшитые типы, — поэтому публичные методы сервисов не содержат таких типов и
    /// вызывают [NoInlining]-ядра только после Ensure().
    /// </summary>
    internal static class EmbeddedAssemblies
    {
        private static readonly object Sync = new object();
        private static bool _registered;
        // Кэшируем Assembly.Load(byte[]) сами: CLR не дедуплицирует загрузку из байтов.
        private static readonly Dictionary<string, Assembly> _loaded =
            new Dictionary<string, Assembly>(StringComparer.OrdinalIgnoreCase);

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
            string name = new AssemblyName(args.Name).Name;
            lock (Sync)
            {
                Assembly cached;
                if (_loaded.TryGetValue(name, out cached))
                    return cached;
                // Ресурс называется «<простое имя>.dll»; нет такого — сборка не наша.
                using (Stream stream = typeof(EmbeddedAssemblies).Assembly.GetManifestResourceStream(name + ".dll"))
                {
                    if (stream == null)
                        return null;
                    Assembly asm = Assembly.Load(ReadAll(stream));
                    _loaded[name] = asm;
                    return asm;
                }
            }
        }

        private static byte[] ReadAll(Stream stream)
        {
            var bytes = new byte[stream.Length];
            int done = 0;
            while (done < bytes.Length)
            {
                int read = stream.Read(bytes, done, bytes.Length - done);
                if (read <= 0)
                    // Недочитанный ресурс — повреждённый exe: лучше явная ошибка сразу,
                    // чем Assembly.Load обрезанных байтов с невнятным BadImageFormat.
                    throw new EndOfStreamException(
                        "Вшитый ресурс прочитан не полностью: " + done + " из " + bytes.Length + " байт.");
                done += read;
            }
            return bytes;
        }
    }
}
