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
            Assembly cached;
            lock (Sync)
                if (_loaded.TryGetValue(name, out cached))
                    return cached;

            // Ресурс называется «<простое имя>.dll»; нет такого — сборка не наша.
            byte[] bytes;
            using (Stream stream = typeof(EmbeddedAssemblies).Assembly.GetManifestResourceStream(name + ".dll"))
            {
                if (stream == null)
                    return null;
                bytes = ReadAll(stream);
            }

            // Assembly.Load зовём БЕЗ своего лока: обработчик AssemblyResolve выполняется
            // внутри загрузчика CLR, и удерживать поверх него собственную блокировку — это
            // инверсия блокировок (два воркера, впервые резолвящие РАЗНЫЕ вшитые сборки —
            // PdfSharp в одном окне и PdfPig в другом, — могли встать намертво).
            Assembly loaded = Assembly.Load(bytes);
            lock (Sync)
            {
                // Гонку выиграл другой поток — отдаём ЕГО сборку: два Assembly.Load одних и тех
                // же байтов дают РАЗНЫЕ типы, и вернуть разным вызывающим разные сборки значит
                // получить InvalidCastException между ними. Лишняя копия останется без ссылок.
                if (_loaded.TryGetValue(name, out cached))
                    return cached;
                _loaded[name] = loaded;
                return loaded;
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
