using System;

namespace ExcelMerger
{
    /// <summary>
    /// Инструмент «PDF → Word»: извлекает текстовый слой одного или НЕСКОЛЬКИХ цифровых PDF
    /// (сохранённых из Word, «Microsoft Print to PDF» и т.п.) в один редактируемый .docx.
    /// Страницы всех добавленных файлов показаны единой сеткой и собираются в выбранном
    /// порядке. Отсканированные документы (без текстового слоя) в настоящее время недоступны —
    /// при попытке будет понятное сообщение (файл цел).
    ///
    /// Само окно — общее для всех конвертеров (<see cref="PdfConvertFormBase"/>); здесь только
    /// то, чем этот вывод отличается от других: строки, цвет, расширение и сам конвертер.
    /// </summary>
    public class OcrForm : PdfConvertFormBase
    {
        public OcrForm() : this(null) { }

        public OcrForm(Action showHub) : base(showHub, Spec()) { }

        private static ConvertToolSpec Spec()
        {
            return new ConvertToolSpec
            {
                Prefix = "ocr",
                NameKey = "hub.ocr.name",
                Theme = ExcelMerger.Theme.WordViolet,
                ThemeDark = ExcelMerger.Theme.WordVioletDark,
                Extension = ".docx",
                RequiresSta = true, // требование Word COM
                HistoryKey = "hist.op.pdftoword",
                RecordUsage = UsageStats.RecordPdfToWord,
                Convert = PdfToWordService.Convert
            };
        }
    }
}
