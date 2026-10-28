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
        private static readonly Cache Resources = new Cache(delegate(string name)
        {
            // Ресурс называется «<простое имя>.dll»; нет такого — сборка не наша.
            return typeof(EmbeddedAssemblies).Assembly.GetManifestResourceStream(name + ".dll");
        });

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
            return Resources.Resolve(new AssemblyName(args.Name).Name);
        }

        /// <summary>
        /// Сборки по простому имени из источника байтов; каждая загружается один раз. Отделён от
        /// подписки на AssemblyResolve, чтобы правила кэша и гонки проверялись тестом на своём
        /// источнике, не трогая загрузчик процесса.
        /// </summary>
        internal sealed class Cache
        {
            private readonly Func<string, Stream> _open;
            private readonly object _sync = new object();
            // Кэшируем Assembly.Load(byte[]) сами: CLR не дедуплицирует загрузку из байтов.
            private readonly Dictionary<string, Assembly> _loaded =
                new Dictionary<string, Assembly>(StringComparer.OrdinalIgnoreCase);

            /// <summary>open — поток байтов сборки по простому имени; null — сборка не наша.</summary>
            internal Cache(Func<string, Stream> open)
            {
                if (open == null)
                    throw new ArgumentNullException("open");
                _open = open;
            }

            internal Assembly Resolve(string name)
            {
                Assembly cached;
                lock (_sync)
                    if (_loaded.TryGetValue(name, out cached))
                        return cached;

                byte[] bytes;
                using (Stream stream = _open(name))
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
                lock (_sync)
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
        }

        internal static byte[] ReadAll(Stream stream)
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
