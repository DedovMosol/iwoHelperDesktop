using System;

namespace ExcelMerger
{
    /// <summary>
    /// Инструмент «PDF → PowerPoint»: переносит страницы цифровых PDF в .pptx так, что текст
    /// остаётся ТЕКСТОМ — его можно править, искать и копировать, а не разглядывать картинку.
    /// Всё, что текстом не является (фон, рамки, диаграммы, логотипы), ложится подложкой
    /// страницы, поэтому слайд выглядит как исходник.
    ///
    /// PowerPoint для конвертации не нужен: файл собирается своим кодом. Само окно — общее для
    /// конвертеров (<see cref="PdfConvertFormBase"/>).
    /// </summary>
    public class PptxForm : PdfConvertFormBase
    {
        public PptxForm() : this(null) { }

        public PptxForm(Action showHub) : base(showHub, Spec()) { }

        private static ConvertToolSpec Spec()
        {
            return new ConvertToolSpec
            {
                Prefix = "pptx",
                NameKey = "hub.pptx.name",
                Theme = ExcelMerger.Theme.PowerPointBand,
                ThemeDark = ExcelMerger.Theme.PowerPointBandDark,
                Extension = ".pptx",
                RequiresSta = false, // ни COM, ни WinRT — обычный фоновый поток
                HistoryKey = "hist.op.pdftopptx",
                RecordUsage = UsageStats.RecordPdfToPptx,
                Convert = PdfToPptxService.Convert
            };
        }
    }
}
