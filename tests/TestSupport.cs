using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// Общие помощники тестов окон (TS-13): ответ на модальные сообщения и текст окна. Раньше
    /// каждый файл тестов заводил свои копии.
    /// </summary>
    internal static partial class UnitTests
    {
        /// <summary>
        /// Отвечать result на модальные окна сообщений (MessageForm), пока их не наберётся max;
        /// seen получает весь текст окна. Возвращает таймер — его Dispose снимает ответчика.
        /// Без ответчика окно ошибки, открытое кодом под тестом, держало бы прогон до сторожа.
        /// </summary>
        private static Timer AnswerMessages(DialogResult result, int max, Action<string> seen)
        {
            int answered = 0;
            var timer = new Timer { Interval = 50 };
            timer.Tick += delegate
            {
                foreach (Form form in Application.OpenForms)
                {
                    if (!(form is MessageForm) || !form.Visible || form.DialogResult != DialogResult.None)
                        continue;
                    if (seen != null)
                        seen(ControlText(form));
                    form.DialogResult = result;
                    if (++answered >= max)
                        timer.Stop();
                    return; // коллекция окон меняется при закрытии — следующее на следующем тике
                }
            };
            timer.Start();
            return timer;
        }

        /// <summary>
        /// Принтер для живых тестов печати — одна политика на весь набор (TS-7): виртуальный, чтобы
        /// ничего не уехало на офисный принтер, иначе любой доступный; нет ни одного — честный
        /// пропуск с зависимостью «printer». На системный принтер по умолчанию полагаться нельзя:
        /// политика «Разрешить Windows управлять принтером по умолчанию» (стандарт с Win10)
        /// оставляет сеанс без него, и PrinterSettings без имени бросает InvalidPrinterException
        /// даже на перечислении листов. Поэтому имя задаётся всегда, пригодность — по IsValid.
        /// </summary>
        private static System.Drawing.Printing.PrinterSettings RequirePrinter()
        {
            string[] virtuals = { "Microsoft Print to PDF", "Microsoft XPS Document Writer", "Adobe PDF" };
            var names = new List<string>(virtuals);
            foreach (string installed in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                names.Add(installed);
            foreach (string name in names)
            {
                var candidate = new System.Drawing.Printing.PrinterSettings { PrinterName = name };
                if (candidate.IsValid)
                    return candidate;
            }
            SkipMissing("printer", "в системе нет ни одного доступного принтера — печать проверить нечем");
            return null; // недостижимо: SkipMissing бросает
        }

        // ---------- Построители PDF (TS-13) ----------

        /// <summary>Born-digital PDF: на каждой странице одна строка текста (Times New Roman 14).</summary>
        private static void WriteTextPdf(string path, params string[] pageTexts)
        {
            using (var doc = new PdfSharp.Pdf.PdfDocument())
            {
                var font = new PdfSharp.Drawing.XFont("Times New Roman", 14);
                foreach (string text in pageTexts)
                {
                    PdfSharp.Pdf.PdfPage page = doc.AddPage();
                    using (PdfSharp.Drawing.XGraphics g = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
                        g.DrawString(text, font, PdfSharp.Drawing.XBrushes.Black, new PdfSharp.Drawing.XPoint(50, 100));
                }
                doc.Save(path);
            }
        }

        /// <summary>PDF из заданного числа пустых страниц по точному пути.</summary>
        private static void MakeEmptyPagesPdf(string path, int pages)
        {
            using (var doc = new PdfSharp.Pdf.PdfDocument())
            {
                for (int i = 0; i < pages; i++)
                    doc.AddPage();
                doc.Save(path);
            }
        }

        // ---------- Доступ к внутреннему состоянию (TS-11, TS-13) ----------
        // Одна реализация вместо четырёх копий (ReviewPrivateField, AuditDField, SetField, AuditDCall).
        // Поле ищется по всей цепочке базовых типов; нет поля — понятная ошибка с именем типа:
        // переименование в продукте ломает тест сразу и с причиной, а не молча.

        private static System.Reflection.FieldInfo InternalFieldInfo(object target, string name)
        {
            for (Type t = target.GetType(); t != null; t = t.BaseType)
            {
                System.Reflection.FieldInfo field = t.GetField(name, System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.DeclaredOnly);
                if (field != null)
                    return field;
            }
            throw new Exception("нет поля " + name + " у " + target.GetType().Name + " — поле переименовано или удалено");
        }

        /// <summary>Записать внутреннее поле объекта (поиск по базовым типам).</summary>
        private static void SetInternal(object target, string name, object value)
        {
            InternalFieldInfo(target, name).SetValue(target, value);
        }

        /// <summary>Вызвать внутренний метод экземпляра; исключение метода выходит как есть.</summary>
        private static object CallInternal(object target, string name, params object[] args)
        {
            for (Type t = target.GetType(); t != null; t = t.BaseType)
                foreach (System.Reflection.MethodInfo method in t.GetMethods(System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.DeclaredOnly))
                    if (method.Name == name && method.GetParameters().Length == args.Length)
                    {
                        try { return method.Invoke(target, args); }
                        catch (System.Reflection.TargetInvocationException ex) when (ex.InnerException != null)
                        {
                            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                            throw;
                        }
                    }
            throw new Exception("нет метода " + name + "/" + args.Length + " у " + target.GetType().Name);
        }

        /// <summary>Текст окна и всех его элементов — для проверок «что увидел человек».</summary>
        private static string ControlText(Control root)
        {
            var text = new StringBuilder();
            var stack = new Stack<Control>();
            stack.Push(root);
            while (stack.Count > 0)
            {
                Control c = stack.Pop();
                if (!string.IsNullOrEmpty(c.Text))
                    text.Append(c.Text).Append('\n');
                foreach (Control child in c.Controls)
                    stack.Push(child);
            }
            return text.ToString();
        }

        // ---------- Детерминированные меры работы вместо секундомера (TS-5) ----------
        // Порог по времени зависит от машины и загрузки раннера; сколько прочитано байтов и
        // выделено памяти — нет. Проверка «не зависает» остаётся за сторожем раннера.

        /// <summary>
        /// Методы type (и вложенных в него типов), в IL которых есть вызов target. Правило «этот
        /// класс не зовёт X» проверяется по коду, а не по времени (TS-5): так закреплено решение,
        /// у которого нет иного наблюдаемого следа, например «пиксели — только через LockBits».
        /// IL разбирается по таблице OpCodes, поэтому байты операндов не принимаются за вызовы.
        /// </summary>
        private static List<string> CallersOf(Type type, System.Reflection.MethodBase target)
        {
            var callers = new List<string>();
            var types = new Stack<Type>();
            types.Push(type);
            while (types.Count > 0)
            {
                Type current = types.Pop();
                foreach (Type nested in current.GetNestedTypes(System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.NonPublic))
                    types.Push(nested);
                const System.Reflection.BindingFlags all = System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.DeclaredOnly;
                var methods = new List<System.Reflection.MethodBase>(current.GetMethods(all));
                methods.AddRange(current.GetConstructors(all));
                foreach (System.Reflection.MethodBase method in methods)
                    if (CallsMethod(method, target))
                        callers.Add(current.Name + "." + method.Name);
            }
            return callers;
        }

        private static readonly Dictionary<short, System.Reflection.Emit.OpCode> IlOpCodes = LoadIlOpCodes();

        private static Dictionary<short, System.Reflection.Emit.OpCode> LoadIlOpCodes()
        {
            var codes = new Dictionary<short, System.Reflection.Emit.OpCode>();
            foreach (System.Reflection.FieldInfo field in typeof(System.Reflection.Emit.OpCodes).GetFields(
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static))
            {
                var code = (System.Reflection.Emit.OpCode)field.GetValue(null);
                codes[code.Value] = code;
            }
            return codes;
        }

        private static bool CallsMethod(System.Reflection.MethodBase method, System.Reflection.MethodBase target)
        {
            System.Reflection.MethodBody body = method.GetMethodBody();
            if (body == null)
                return false;
            byte[] il = body.GetILAsByteArray();
            Type[] typeArgs = method.DeclaringType.IsGenericType ? method.DeclaringType.GetGenericArguments() : null;
            Type[] methodArgs = method.IsGenericMethod ? method.GetGenericArguments() : null;
            int i = 0;
            while (i < il.Length)
            {
                short value = il[i] == 0xFE && i + 1 < il.Length ? (short)(0xFE00 | il[i + 1]) : il[i];
                System.Reflection.Emit.OpCode code;
                if (!IlOpCodes.TryGetValue(value, out code))
                    throw new Exception("неизвестный опкод 0x" + value.ToString("X") + " в " + method.Name);
                i += code.Size;
                if (code.OperandType == System.Reflection.Emit.OperandType.InlineMethod)
                {
                    int token = BitConverter.ToInt32(il, i);
                    System.Reflection.MethodBase called = method.Module.ResolveMethod(token, typeArgs, methodArgs);
                    if (called == target)
                        return true;
                }
                i += IlOperandSize(code.OperandType, il, i);
            }
            return false;
        }

        private static int IlOperandSize(System.Reflection.Emit.OperandType type, byte[] il, int at)
        {
            switch (type)
            {
                case System.Reflection.Emit.OperandType.InlineNone: return 0;
                case System.Reflection.Emit.OperandType.ShortInlineBrTarget:
                case System.Reflection.Emit.OperandType.ShortInlineI:
                case System.Reflection.Emit.OperandType.ShortInlineVar: return 1;
                case System.Reflection.Emit.OperandType.InlineVar: return 2;
                case System.Reflection.Emit.OperandType.InlineI8:
                case System.Reflection.Emit.OperandType.InlineR: return 8;
                case System.Reflection.Emit.OperandType.InlineSwitch: return 4 + 4 * BitConverter.ToInt32(il, at);
                default: return 4; // токены, InlineI, InlineBrTarget, ShortInlineR
            }
        }

        /// <summary>
        /// Байты, выделенные в управляемой куче кодом action на этом потоке. Мера годна только для
        /// однопоточного кода: выделения других потоков сюда не попадают.
        /// </summary>
        private static long AllocatedBytes(Action action)
        {
            long before = GC.GetAllocatedBytesForCurrentThread();
            action();
            return GC.GetAllocatedBytesForCurrentThread() - before;
        }

        /// <summary>
        /// Поток только для чтения, считающий прочитанные байты. Дальше limit — исключение: код,
        /// потерявший предел чтения, падает сразу, а не зацикливается до сторожа прогона.
        /// </summary>
        private sealed class CountingReadStream : System.IO.Stream
        {
            private readonly System.IO.Stream _inner;
            private readonly long _limit;

            internal CountingReadStream(System.IO.Stream inner, long limit)
            {
                _inner = inner;
                _limit = limit;
            }

            internal long BytesRead { get; private set; }

            public override int Read(byte[] buffer, int offset, int count)
            {
                int read = _inner.Read(buffer, offset, count);
                BytesRead += read;
                if (BytesRead > _limit)
                    throw new InvalidOperationException("прочитано больше " + _limit + " байт — чтение не ограничено");
                return read;
            }

            public override bool CanRead { get { return true; } }
            public override bool CanSeek { get { return _inner.CanSeek; } }
            public override bool CanWrite { get { return false; } }
            public override long Length { get { return _inner.Length; } }
            public override long Position { get { return _inner.Position; } set { _inner.Position = value; } }
            public override void Flush() { }
            public override long Seek(long offset, System.IO.SeekOrigin origin) { return _inner.Seek(offset, origin); }
            public override void SetLength(long value) { throw new NotSupportedException(); }
            public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                    _inner.Dispose();
                base.Dispose(disposing);
            }
        }
    }
}

