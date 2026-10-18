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
            EmbeddedFile.Open(owner, title, ResourceName, Path.Combine(AppPaths.Root, FileName));
        }
    }
}
