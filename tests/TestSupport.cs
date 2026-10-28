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
    }
}

