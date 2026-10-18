using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

namespace ExcelMerger
{
    internal enum PdfReviewSourceError
    {
        None,
        Empty,
        InvalidPath,
        Missing,
        NotPdf,
        Unreadable
    }

    /// <summary>
    /// Результат проверки пути к одной стороне Review. Path задан только для существующего
    /// файла с расширением PDF; Probe дополнительно подтверждает, что PdfPig может его открыть.
    /// Защищённый файл остаётся допустимым источником — пароль спрашивает штатный Compare.
    /// </summary>
    internal sealed class PdfReviewSourceResult
    {
        public PdfReviewSourceError Error;
        public string Path;

        public bool IsValid { get { return Error == PdfReviewSourceError.None && Path != null; } }
    }

    internal enum PdfReviewDropTarget
    {
        Neutral,
        Left,
        Right
    }

    internal enum PdfReviewDropAction
    {
        None,
        AssignLeft,
        AssignRight,
        AssignBoth,
        NeedExplicitSide,
        TooMany
    }

    /// <summary>Чистое решение маршрута drop; WinForms-слой только исполняет его.</summary>
    internal sealed class PdfReviewDropPlan
    {
        public PdfReviewDropAction Action;
        public string LeftPath;
        public string RightPath;
    }

    /// <summary>
    /// Единая граница между строкой из поля/browse/drop и сервисом сравнения. Здесь только
    /// дешёвая пригодность источника; извлечение текста и лимиты остаются в PdfReviewService.
    /// </summary>
    internal static class PdfReviewInput
    {
        public static PdfReviewSourceResult Resolve(string text)
        {
            string value = text == null ? "" : text.Trim();
            if (value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"')
                value = value.Substring(1, value.Length - 2).Trim();
            if (value.Length == 0)
                return Error(PdfReviewSourceError.Empty);

            string full;
            try { full = Path.GetFullPath(value); }
            catch { return Error(PdfReviewSourceError.InvalidPath); }

            if (!string.Equals(Path.GetExtension(full), ".pdf", StringComparison.OrdinalIgnoreCase))
                return Error(PdfReviewSourceError.NotPdf);
            // Resolve зовётся на UI-потоке (Leave/Enter поля). Существование на локальном
            // диске проверяем сразу — это мгновенно и даёт ошибку без мигания «проверка…»;
            // сетевой/съёмный путь может висеть десятки секунд, его проверяет фоновая Probe.
            if (IsLocalDiskPath(full) && !File.Exists(full))
                return Error(PdfReviewSourceError.Missing);
            return new PdfReviewSourceResult { Path = full };
        }

        /// <summary>
        /// Путь на локальном жёстком/RAM/USB-диске. UNC (\\сервер\…, \\?\…), сетевые буквы,
        /// CD и неизвестные тома — не локальные: обращение к ним на UI-потоке недопустимо.
        /// Тип тома берётся по корню (GetDriveType), без чтения файла.
        /// </summary>
        internal static bool IsLocalDiskPath(string fullPath)
        {
            if (string.IsNullOrEmpty(fullPath) || fullPath.StartsWith(@"\\", StringComparison.Ordinal))
                return false;
            try
            {
                string root = Path.GetPathRoot(fullPath);
                if (string.IsNullOrEmpty(root))
                    return false;
                DriveType type = new DriveInfo(root).DriveType;
                return type == DriveType.Fixed || type == DriveType.Ram ||
                    type == DriveType.Removable;
            }
            catch { return false; }
        }

        public static PdfReviewSourceResult Probe(PdfReviewSourceResult source)
        {
            if (source == null || !source.IsValid)
                return source ?? Error(PdfReviewSourceError.InvalidPath);
            try
            {
                // Фоновый поток: здесь существование проверяется для любого пути, включая
                // сетевой, который Resolve намеренно не трогал.
                if (!File.Exists(source.Path))
                    return Error(PdfReviewSourceError.Missing);
                EmbeddedAssemblies.Ensure();
                return ProbeCore(source.Path);
            }
            catch
            {
                return Error(PdfReviewSourceError.Unreadable);
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static PdfReviewSourceResult ProbeCore(string path)
        {
            try
            {
                using (UglyToad.PdfPig.PdfDocument pdf = PdfPageProbe.OpenPig(path))
                    if (pdf.NumberOfPages > 0)
                        return new PdfReviewSourceResult { Path = path };
            }
            catch (Exception ex)
            {
                // PasswordRequired — валидный выбор. Compare повторит открытие и покажет
                // существующий диалог пароля, не превращая защищённый PDF в «битый».
                if (PdfPasswords.LooksPasswordProtected(ex))
                    return new PdfReviewSourceResult { Path = path };
            }
            return Error(PdfReviewSourceError.Unreadable);
        }

        public static PdfReviewDropPlan PlanDrop(IList<string> paths,
            PdfReviewDropTarget target, bool hasLeft, bool hasRight)
        {
            var plan = new PdfReviewDropPlan();
            int count = paths == null ? 0 : paths.Count;
            if (count == 0)
                return plan;
            if (count > 2)
            {
                plan.Action = PdfReviewDropAction.TooMany;
                return plan;
            }
            if (count == 2)
            {
                plan.Action = PdfReviewDropAction.AssignBoth;
                plan.LeftPath = paths[0];
                plan.RightPath = paths[1];
                return plan;
            }

            if (target == PdfReviewDropTarget.Left ||
                (target == PdfReviewDropTarget.Neutral && !hasLeft))
            {
                plan.Action = PdfReviewDropAction.AssignLeft;
                plan.LeftPath = paths[0];
            }
            else if (target == PdfReviewDropTarget.Right ||
                (target == PdfReviewDropTarget.Neutral && !hasRight))
            {
                plan.Action = PdfReviewDropAction.AssignRight;
                plan.RightPath = paths[0];
            }
            else
                plan.Action = PdfReviewDropAction.NeedExplicitSide;
            return plan;
        }

        private static PdfReviewSourceResult Error(PdfReviewSourceError error)
        {
            return new PdfReviewSourceResult { Error = error };
        }
    }
}
