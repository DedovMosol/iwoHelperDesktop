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

        /// <summary>Born-digital PDF: на каждой странице одна строка текста.</summary>
        private static void Ts15TextPdf(string path, params string[] pageTexts)
        {
            using (var doc = new PdfSharp.Pdf.PdfDocument())
            {
                var font = new PdfSharp.Drawing.XFont("Arial", 16);
                foreach (string text in pageTexts)
                {
                    PdfSharp.Pdf.PdfPage page = doc.AddPage();
                    using (PdfSharp.Drawing.XGraphics g = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
                        g.DrawString(text, font, PdfSharp.Drawing.XBrushes.Black, new PdfSharp.Drawing.XPoint(72, 120));
                }
                doc.Save(path);
            }
        }

        [Test("Аудит TS-15: PDF → PowerPoint — слайд на страницу, текст источника, валидный пакет, порядок из набора")]
        private static void AuditTs15PdfToPptx()
        {
            using (var dir = new TempDir("iwo_ts15_pptx_"))
            {
                string pdf = dir.File("источник.pdf");
                Ts15TextPdf(pdf, "Alpha first page", "Beta second page");
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
                Ts15TextPdf(pdf, "One", "Two", "Three");
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
