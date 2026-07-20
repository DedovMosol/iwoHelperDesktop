using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using Microsoft.Win32;

namespace ExcelMerger
{
    /// <summary>
    /// Поиск и запуск Ghostscript — движка сжатия PDF (downsampling изображений с
    /// сохранением текста, уровень Acrobat «Reduce File Size»). Фича опциональна:
    /// если GS не найден, вызывающий предлагает ссылку на скачивание.
    ///
    /// Приоритет поиска: вшитый инсталлятором рядом с exe (&lt;app&gt;\gs\bin) →
    /// реестр (GPL/Artifex Ghostscript) → Program Files\gs\* → профиль пользователя
    /// (gs*) → PATH. Результат кэшируется потокобезопасно (Lazy) — оба PDF-инструмента
    /// зовут его из фоновых потоков.
    /// </summary>
    public static class Ghostscript
    {
        /// <summary>Официальная страница загрузки GPL Ghostscript (Windows 64-bit installer).</summary>
        public const string DownloadPage = "https://ghostscript.com/releases/gsdnld.html";

        private static readonly Lazy<string> _exe = new Lazy<string>(
            delegate { return PickFirstExisting(Candidates(), File.Exists); },
            LazyThreadSafetyMode.ExecutionAndPublication);

        /// <summary>Полный путь к gswin*c.exe или null, если Ghostscript не найден.</summary>
        public static string Exe { get { return _exe.Value; } }

        /// <summary>Доступно ли сжатие (найден ли Ghostscript).</summary>
        public static bool Available { get { return Exe != null; } }

        /// <summary>
        /// Корень вшитого GS (&lt;app&gt;\gs) — когда используется бандл рядом с exe;
        /// иначе null. Для бандла аргументы получают -I на его lib/Resource, чтобы
        /// не зависеть от эвристики относительных путей GS.
        /// </summary>
        public static string BundledRoot
        {
            get
            {
                string exe = Exe;
                if (exe == null)
                    return null;
                try
                {
                    string root = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "gs");
                    if (exe.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                        return root;
                }
                catch { }
                return null;
            }
        }

        /// <summary>Первый существующий путь из кандидатов (null — ни одного). Чистая — под тест.</summary>
        public static string PickFirstExisting(IEnumerable<string> candidates, Func<string, bool> exists)
        {
            if (candidates == null || exists == null)
                return null;
            foreach (string c in candidates)
            {
                if (string.IsNullOrEmpty(c))
                    continue;
                bool ok;
                try { ok = exists(c); }
                catch { ok = false; }
                if (ok)
                    return c;
            }
            return null;
        }

        /// <summary>
        /// Запуск Ghostscript. Возвращает код выхода (или -1 при сбое/таймауте),
        /// stderr — через out. Потоки читаются асинхронно (без дедлока на буфере),
        /// процесс освобождается и добивается по таймауту.
        /// </summary>
        public static int Run(string argsLine, int timeoutMs, out string stderr)
        {
            stderr = string.Empty;
            string exe = Exe;
            if (exe == null)
                return -1;
            var err = new StringBuilder();
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = exe,
                    Arguments = argsLine,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                using (var p = Process.Start(psi))
                {
                    if (p == null)
                        return -1;
                    p.OutputDataReceived += delegate { }; // осушаем stdout, чтобы буфер не блокировал
                    p.ErrorDataReceived += delegate(object s, DataReceivedEventArgs e)
                    {
                        if (e.Data != null)
                            err.AppendLine(e.Data);
                    };
                    p.BeginOutputReadLine();
                    p.BeginErrorReadLine();
                    if (!p.WaitForExit(timeoutMs))
                    {
                        try { p.Kill(); } catch { }
                        try { p.WaitForExit(); } catch { } // дать ОС высвободить процесс
                        stderr = "timeout";
                        return -1;
                    }
                    p.WaitForExit(); // дождаться флаша асинхронных обработчиков
                    stderr = err.ToString();
                    return p.ExitCode;
                }
            }
            catch (Exception ex)
            {
                stderr = ex.Message;
                return -1;
            }
        }

        // ---- Кандидаты на путь к gswin*c.exe, в порядке приоритета ----

        private static List<string> Candidates()
        {
            var list = new List<string>();

            // 1. Вшитый инсталлятором рядом с exe: <app>\gs\bin\gswin64c.exe
            try
            {
                string appGsBin = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "gs", "bin");
                list.Add(Path.Combine(appGsBin, "gswin64c.exe"));
                list.Add(Path.Combine(appGsBin, "gswin32c.exe"));
            }
            catch { }

            // 2. Реестр (GPL/Artifex Ghostscript): из GS_DLL берём каталог bin.
            list.AddRange(RegistryCandidates());

            // 3. Program Files\gs\gs*\bin
            foreach (string pf in new[]
            {
                Environment.GetEnvironmentVariable("ProgramFiles"),
                Environment.GetEnvironmentVariable("ProgramFiles(x86)")
            })
            {
                if (!string.IsNullOrEmpty(pf))
                    AddGlobBin(list, Path.Combine(pf, "gs"));
            }

            // 4. Профиль пользователя: <profile>\gs*\bin (нестандартная пользовательская установка)
            try
            {
                AddGlobBin(list, Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
            }
            catch { }

            // 5. PATH
            try
            {
                string pathEnv = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
                foreach (string dir in pathEnv.Split(';'))
                {
                    string d = dir.Trim();
                    if (d.Length == 0)
                        continue;
                    try
                    {
                        list.Add(Path.Combine(d, "gswin64c.exe"));
                        list.Add(Path.Combine(d, "gswin32c.exe"));
                    }
                    catch { }
                }
            }
            catch { }

            return list;
        }

        /// <summary>Для каждого подкаталога parent\gs* добавляет bin\gswin64c.exe и gswin32c.exe.</summary>
        private static void AddGlobBin(List<string> list, string parent)
        {
            try
            {
                if (string.IsNullOrEmpty(parent) || !Directory.Exists(parent))
                    return;
                foreach (string dir in Directory.GetDirectories(parent, "gs*"))
                {
                    list.Add(Path.Combine(dir, "bin", "gswin64c.exe"));
                    list.Add(Path.Combine(dir, "bin", "gswin32c.exe"));
                }
            }
            catch { }
        }

        private static List<string> RegistryCandidates()
        {
            var result = new List<string>();
            string[] vendors = { @"SOFTWARE\GPL Ghostscript", @"SOFTWARE\Artifex Ghostscript" };
            RegistryHive[] hives = { RegistryHive.LocalMachine, RegistryHive.CurrentUser };
            RegistryView[] views = { RegistryView.Registry64, RegistryView.Registry32 };

            foreach (RegistryHive hive in hives)
                foreach (RegistryView view in views)
                {
                    try
                    {
                        using (RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, view))
                        {
                            foreach (string vendor in vendors)
                            {
                                using (RegistryKey vk = baseKey.OpenSubKey(vendor))
                                {
                                    if (vk == null)
                                        continue;
                                    foreach (string ver in vk.GetSubKeyNames())
                                    {
                                        string dll = ReadGsDll(vk, ver);
                                        if (string.IsNullOrEmpty(dll))
                                            continue;
                                        string dir = SafeDirName(dll);
                                        if (dir == null)
                                            continue;
                                        result.Add(Path.Combine(dir, "gswin64c.exe"));
                                        result.Add(Path.Combine(dir, "gswin32c.exe"));
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                }
            return result;
        }

        private static string ReadGsDll(RegistryKey vendorKey, string versionSubKey)
        {
            try
            {
                using (RegistryKey k = vendorKey.OpenSubKey(versionSubKey))
                    return k == null ? null : k.GetValue("GS_DLL") as string;
            }
            catch { return null; }
        }

        private static string SafeDirName(string path)
        {
            try { return Path.GetDirectoryName(path); }
            catch { return null; }
        }
    }
}
