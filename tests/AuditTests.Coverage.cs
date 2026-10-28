using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Xml;

namespace ExcelMerger.Tests
{
    /// <summary>
    /// TS-15 (audit.md): модули, на которые не ссылался ни один тест. Межпроцессная блокировка и
    /// файлы состояния — на настоящих файлах и с настоящим удержанием; PDF → PowerPoint — полный
    /// путь сервиса без Office; записка Word — живой Word.
    /// </summary>
    internal static partial class UnitTests
    {
        // ---------- AppDataLock ----------

        [Test("Аудит TS-15: AppDataLock — пока lock занят, действие не выполняется; освобождённый — берётся")]
        private static void AuditTs15LockWaitsAndGivesUp()
        {
            using (var dir = new TempDir("iwo_ts15_lock_"))
            {
                string state = dir.File("state.txt");
                int runs = 0;
                // Чужой держатель — тот же FileShare.None, что у другого процесса или сеанса.
                using (new FileStream(state + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
                {
                    var clock = System.Diagnostics.Stopwatch.StartNew();
                    bool ran = AppDataLock.TryRun(state, delegate { runs++; return true; });
                    AssertTrue(!ran && runs == 0, "занятый lock: действие не выполнено");
                    AssertTrue(clock.ElapsedMilliseconds >= 1500, "отказ — после ожидания, а не сразу: " + clock.ElapsedMilliseconds + " мс");
                }

                // Держатель отпускает во время ожидания — действие выполняется ровно один раз.
                var holder = new FileStream(state + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
                var release = new Thread(delegate() { Thread.Sleep(300); holder.Dispose(); });
                release.Start();
                bool afterRelease = AppDataLock.TryRun(state, delegate { runs++; return true; });
                release.Join();
                AssertTrue(afterRelease, "освобождённый lock взят");
                AssertEq(1, runs, "действие выполнено один раз");
            }
        }

        [Test("Аудит TS-15: AppDataLock — сбой действия не повторяет его и не держит lock")]
        private static void AuditTs15LockActionFailure()
        {
            using (var dir = new TempDir("iwo_ts15_lockfail_"))
            {
                string state = dir.File("state.txt");
                int runs = 0;
                // IOException «занято» из САМОГО действия (файл данных держит антивирус) — не повод
                // перезапускать действие, которое могло успеть что-то записать.
                var busy = new IOException("файл данных занят (тест)", unchecked((int)0x80070020));
                bool ran = AppDataLock.TryRun(state, delegate { runs++; throw busy; });
                AssertTrue(!ran, "сбой действия — false");
                AssertEq(1, runs, "действие не перезапущено");

                bool other = AppDataLock.TryRun(state, delegate { throw new InvalidOperationException("тест"); });
                AssertTrue(!other, "любой сбой действия — false, а не исключение наружу");
                AssertTrue(AppDataLock.TryRun(state, delegate { return true; }), "после сбоев lock свободен");

                AssertTrue(!AppDataLock.TryRun(null, delegate { return true; }), "без файла — false");
                AssertTrue(!AppDataLock.TryRun(state, null), "без действия — false");
                AssertTrue(!AppDataLock.TryRun(dir.Path + "\\bad|name.txt", delegate { return true; }),
                    "недопустимый путь — false, а не исключение");
            }
        }

        // ---------- AppStateFile ----------

        [Test("Аудит TS-15: AppStateFile — нет файла — пустой снимок; занятый — «не прочитан»; запись атомарна")]
        private static void AuditTs15StateFile()
        {
            using (var dir = new TempDir("iwo_ts15_state_"))
            {
                string[] lines;
                string path = Path.Combine(dir.Path, "вложенная", "state.txt");
                AssertTrue(AppStateFile.TryReadLines(path, out lines) && lines.Length == 0,
                    "нет файла и папки — достоверный пустой снимок");

                AppStateFile.WriteLines(path, new[] { "первая", "вторая" });
                AssertTrue(AppStateFile.TryReadLines(path, out lines), "записанный читается");
                AssertEq("первая|вторая", string.Join("|", lines), "строки на месте, папка создана");

                AppStateFile.WriteLines(path, new[] { "замена" });
                AssertTrue(AppStateFile.TryReadLines(path, out lines) && lines.Length == 1 && lines[0] == "замена",
                    "существующий файл заменён целиком");
                AssertEq(1, Directory.GetFiles(Path.GetDirectoryName(path)).Length, "временных файлов записи не осталось");

                using (new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    AssertTrue(!AppStateFile.TryReadLines(path, out lines), "файл есть, но не читается — не пустой снимок");
                    AssertEq(0, lines.Length, "при отказе строк нет");
                }
            }
        }

        // ---------- PDF → PowerPoint ----------


        [Test("Аудит TS-15: PDF → PowerPoint — слайд на страницу, текст источника, валидный пакет, порядок из набора")]
        private static void AuditTs15PdfToPptx()
        {
            using (var dir = new TempDir("iwo_ts15_pptx_"))
            {
                string pdf = dir.File("источник.pdf");
                WriteTextPdf(pdf, "Alpha first page", "Beta second page");
                string pptx = dir.File("итог.pptx");
                // Порядок задан набором, а не файлом: вторая страница — первым слайдом.
                var order = new List<PdfPageRef>
                {
                    new PdfPageRef { SourcePath = pdf, PageIndex = 1 },
                    new PdfPageRef { SourcePath = pdf, PageIndex = 0 }
                };
                ConvertResult result = PdfToPptxService.Convert(order, pptx);
                AssertEq(2, result.Pages, "две страницы");
                AssertEq(2, result.PagesWithText, "обе с текстом");
                List<string> problems = PptxValidationProblems(pptx);
                AssertEq(0, problems.Count, "пакет валиден: " + string.Join("; ", problems.ToArray()));
                AssertEq("Beta second page|Alpha first page", string.Join("|", Ts15SlideTexts(pptx).ToArray()),
                    "текст слайдов в порядке набора");
            }
        }

        [Test("Аудит TS-15: PDF → PowerPoint — отменённое преобразование не оставляет файла")]
        private static void AuditTs15PdfToPptxCancel()
        {
            using (var dir = new TempDir("iwo_ts15_pptxcancel_"))
            {
                string pdf = dir.File("источник.pdf");
                WriteTextPdf(pdf, "One", "Two", "Three");
                string pptx = dir.File("итог.pptx");
                var order = new List<PdfPageRef>();
                for (int i = 0; i < 3; i++)
                    order.Add(new PdfPageRef { SourcePath = pdf, PageIndex = i });
                AssertThrowsExactly<OperationCanceledException>("отмена до начала",
                    delegate { PdfToPptxService.Convert(order, pptx, null, delegate { return true; }); });
                AssertTrue(!File.Exists(pptx), "файла нет");
                AssertEq(1, Directory.GetFiles(dir.Path).Length, "временных файлов не осталось — только источник");
            }
        }

        /// <summary>Склеенный текст каждого слайда (a:t) в порядке слайдов.</summary>
        private static List<string> Ts15SlideTexts(string pptx)
        {
            var texts = new List<string>();
            using (ZipArchive zip = ZipFile.OpenRead(pptx))
            {
                for (int n = 1; ; n++)
                {
                    ZipArchiveEntry entry = zip.GetEntry("ppt/slides/slide" + n + ".xml");
                    if (entry == null)
                        break;
                    var xml = new XmlDocument();
                    using (Stream s = entry.Open())
                        xml.Load(s);
                    var ns = new XmlNamespaceManager(xml.NameTable);
                    ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    var text = new StringBuilder();
                    foreach (XmlNode t in xml.SelectNodes("//a:t", ns))
                        text.Append(t.InnerText);
                    texts.Add(text.ToString().Trim());
                }
            }
            return texts;
        }

        // ---------- EmbeddedAssemblies ----------

        /// <summary>Байты настоящей сборки для загрузки из памяти (небольшая зависимость из папки тестов).</summary>
        private static byte[] Ts15AssemblyBytes()
        {
            // PdfSharp — та самая вшиваемая сборка; копия в папке тестов, грузится из байтов отдельно.
            return File.ReadAllBytes(typeof(PdfSharp.Pdf.PdfDocument).Assembly.Location);
        }

        [Test("Аудит TS-15: EmbeddedAssemblies — своя сборка грузится один раз, чужое имя — null, источник открыт один раз")]
        private static void AuditTs15EmbeddedCache()
        {
            byte[] bytes = Ts15AssemblyBytes();
            int opened = 0;
            var cache = new EmbeddedAssemblies.Cache(delegate(string name)
            {
                if (name != "Зонд")
                    return null;
                Interlocked.Increment(ref opened);
                return new MemoryStream(bytes);
            });
            System.Reflection.Assembly first = cache.Resolve("Зонд");
            AssertTrue(first != null, "сборка загружена из байтов");
            AssertTrue(ReferenceEquals(first, cache.Resolve("ЗОНД")), "повтор (имя без учёта регистра) — тот же экземпляр");
            AssertEq(1, opened, "ресурс прочитан один раз");
            AssertTrue(cache.Resolve("Чужая") == null, "сборка не наша — null, загрузчик ищет дальше");
        }

        [Test("Аудит TS-15: EmbeddedAssemblies — первые одновременные запросы получают ОДИН экземпляр сборки")]
        private static void AuditTs15EmbeddedRace()
        {
            byte[] bytes = Ts15AssemblyBytes();
            var gate = new ManualResetEvent(false);
            var cache = new EmbeddedAssemblies.Cache(delegate(string name)
            {
                gate.WaitOne(); // все потоки дошли до чтения — гонка гарантирована
                return new MemoryStream(bytes);
            });
            const int threads = 8;
            var got = new System.Reflection.Assembly[threads];
            var workers = new Thread[threads];
            for (int i = 0; i < threads; i++)
            {
                int index = i;
                workers[i] = new Thread(delegate() { got[index] = cache.Resolve("Зонд"); });
                workers[i].Start();
            }
            Thread.Sleep(100);
            gate.Set();
            foreach (Thread t in workers)
                AssertTrue(t.Join(30000), "поток резолва завершился");
            gate.Close();
            for (int i = 1; i < threads; i++)
                AssertTrue(ReferenceEquals(got[0], got[i]),
                    "две загрузки одних байтов — разные типы и InvalidCastException; всем — один экземпляр");
        }

        /// <summary>Поток, который заявляет больше байтов, чем отдаёт, — как обрезанный ресурс.</summary>
        private sealed class Ts15ShortStream : MemoryStream
        {
            public Ts15ShortStream() : base(new byte[10]) { }
            public override long Length { get { return 100; } }
        }

        [Test("Аудит TS-15: EmbeddedAssemblies — недочитанный ресурс — понятная ошибка, а не Assembly.Load обрезка")]
        private static void AuditTs15EmbeddedTruncated()
        {
            var cache = new EmbeddedAssemblies.Cache(delegate { return new Ts15ShortStream(); });
            EndOfStreamException error = AssertThrowsExactly<EndOfStreamException>("обрезанный ресурс",
                delegate { cache.Resolve("Зонд"); });
            AssertTrue(error.Message.Contains("10 из 100"), "сколько прочитано из скольких: " + error.Message);
        }

        // ---------- FolderPicker ----------

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "FindWindowW", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern IntPtr Ts15FindWindow(string className, string title);

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "PostMessageW")]
        private static extern bool Ts15PostMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);

        /// <summary>
        /// Показать FolderPicker на STA-потоке и ответить на настоящий диалог: confirm — кнопка
        /// выбора папки (IDOK), иначе закрытие окна (отмена). Окно ищется по уникальному заголовку.
        /// </summary>
        private static string Ts15PickFolder(string initialDir, bool confirm, out bool dialogFound)
        {
            string title = "iwo-test " + Guid.NewGuid().ToString("N");
            string result = null;
            bool found = false;
            Exception error = null;
            var ui = new Thread(delegate()
            {
                try
                {
                    using (var owner = new System.Windows.Forms.Form { ShowInTaskbar = false })
                    {
                        owner.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                        owner.Location = new System.Drawing.Point(-32000, -32000);
                        owner.Show();
                        var answer = new Thread(delegate()
                        {
                            try
                            {
                                for (int i = 0; i < 300; i++)
                                {
                                    IntPtr dialog = Ts15FindWindow("#32770", title);
                                    if (dialog != IntPtr.Zero)
                                    {
                                        found = true;
                                        Thread.Sleep(300); // диалог дорисовался и принял папку
                                        if (confirm)
                                            Ts15PostMessage(dialog, 0x0111, (IntPtr)1, IntPtr.Zero); // WM_COMMAND, IDOK
                                        else
                                            Ts15PostMessage(dialog, 0x0010, IntPtr.Zero, IntPtr.Zero); // WM_CLOSE
                                        return;
                                    }
                                    Thread.Sleep(100);
                                }
                            }
                            catch (Exception ex) { error = ex; } // упавший фоновый поток уронил бы весь прогон
                        });
                        answer.IsBackground = true;
                        answer.Start();
                        result = FolderPicker.Show(owner, title, initialDir);
                        answer.Join(5000);
                    }
                }
                catch (Exception ex) { error = ex; }
            });
            ui.SetApartmentState(ApartmentState.STA);
            ui.IsBackground = true;
            ui.Start();
            if (!ui.Join(60000))
            {
                // Диалог не ответил: закрываем его, чтобы поток не остался висеть.
                IntPtr stuck = Ts15FindWindow("#32770", title);
                if (stuck != IntPtr.Zero)
                    Ts15PostMessage(stuck, 0x0010, IntPtr.Zero, IntPtr.Zero);
                ui.Join(10000);
                throw new Exception("диалог выбора папки не закрылся за 60 с");
            }
            if (error != null)
                throw new Exception(error.GetType().Name + ": " + error.Message);
            dialogFound = found;
            return result;
        }

        [Test("Аудит TS-15: FolderPicker — настоящий диалог: выбор возвращает папку, закрытие — null и без второго окна")]
        private static void AuditTs15FolderPicker()
        {
            using (var dir = new TempDir("iwo_ts15_folder_"))
            {
                bool found;
                string picked = Ts15PickFolder(dir.Path, true, out found);
                if (!found)
                    SkipMissing("desktop", "диалог выбора папки не появился — нет интерактивного рабочего стола");
                AssertEq(Path.GetFullPath(dir.Path).TrimEnd('\\'), (picked ?? "").TrimEnd('\\'), "выбрана начальная папка");

                string cancelled = Ts15PickFolder(dir.Path, false, out found);
                AssertTrue(found, "диалог появился и во второй раз");
                AssertTrue(cancelled == null, "закрытие диалога — отмена, а не путь: " + cancelled);
            }
        }

        // ---------- Записка Word ----------

        /// <summary>Текст документа по абзацам (включая ячейки таблиц), абзацы — через перевод строки.</summary>
        private static string Ts15DocxText(string docx)
        {
            XmlDocument xml = AuditBDocumentXml(docx);
            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            var text = new StringBuilder();
            foreach (XmlNode p in xml.SelectNodes("//w:p", ns))
            {
                foreach (XmlNode t in p.SelectNodes(".//w:t", ns))
                    text.Append(t.InnerText);
                text.Append('\n');
            }
            return text.ToString();
        }

        [Test("Аудит TS-15: записка Word — заголовок, итоги и таблица пропусков в .docx (живой Word)")]
        private static void AuditTs15WordNoteLive()
        {
            AuditBRequireWord();
            using (var dir = new TempDir("iwo_ts15_note_"))
            {
                var result = new MergeResult { OutputPath = dir.File("Свод.xlsx") };
                result.Files.Add(new FileResult { FileName = "a.xlsx", FullPath = dir.File("a.xlsx"), Ok = true, SheetName = "a" });
                result.Files.Add(new FileResult { FileName = "b.xlsx", FullPath = dir.File("b.xlsx"), Note = "файл защищён паролем" });
                result.OkCount = 1;
                result.SkipCount = 1;
                NoteContent note = NoteText.Build(result, dir.Path, new MergeOptions(), new DateTime(2026, 10, 27, 9, 30, 0));
                string docx = dir.File("Записка.docx");
                AuditBOfficeSta(120000, delegate { WordNoteWriter.Write(note, docx); });

                string all = Ts15DocxText(docx);
                AssertTrue(all.Contains(note.Title), "заголовок «" + note.Title + "» в тексте: " + all.Substring(0, Math.Min(400, all.Length)));
                AssertTrue(all.Contains(note.Subtitle), "подзаголовок");
                foreach (string paragraph in note.Body)
                    AssertTrue(all.Contains(paragraph), "абзац итогов: " + paragraph);
                AssertTrue(all.Contains("b.xlsx") && all.Contains("файл защищён паролем"), "пропуск и его причина в таблице");
            }
        }
    }
}
