using System.Runtime.CompilerServices;
using PdfSharp.Pdf;

namespace ExcelMerger
{
    /// <summary>Служебное создание простого PDF (для самопроверок --thumbcheck).</summary>
    internal static class PdfProbe
    {
        public static void WriteOnePagePdf(string path)
        {
            EmbeddedAssemblies.Ensure();
            WriteCore(path);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void WriteCore(string path)
        {
            using (var doc = new PdfDocument())
            {
                PdfPage page = doc.AddPage();
                page.Width = 420;   // A5, пункты
                page.Height = 595;
                doc.Save(path);
            }
        }
    }
}
