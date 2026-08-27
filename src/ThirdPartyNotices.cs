using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelMerger
{
    /// <summary>Лицензии embedded-библиотек доступны и в portable single-file сборке.</summary>
    internal static class ThirdPartyNotices
    {
        private const string ResourceName = "thirdparty.notices.txt";
        private const string FileName = "THIRD-PARTY-NOTICES.txt";

        internal static bool IsPacked()
        {
            using (Stream stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream(ResourceName))
                return stream != null && stream.Length > 1024;
        }

        internal static void Open(IWin32Window owner, string title)
        {
            string path = Path.Combine(AppPaths.Root, FileName);
            try
            {
                Extract(path);
            }
            catch (Exception ex)
            {
                if (!File.Exists(path))
                {
                    Dialogs.Error(owner, title, Loc.T("common.err.openFailed"), ex.Message);
                    return;
                }
            }
            Ui.OpenPathOrWarn(owner, title, path);
        }

        private static void Extract(string path)
        {
            using (Stream source = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream(ResourceName))
            {
                if (source == null)
                    throw new FileNotFoundException(ResourceName);
                var existing = new FileInfo(path);
                if (existing.Exists && existing.Length == source.Length)
                    return;
                Directory.CreateDirectory(AppPaths.Root);
                using (var target = new FileStream(path, FileMode.Create,
                    FileAccess.Write, FileShare.None))
                    source.CopyTo(target);
            }
        }
    }
}
