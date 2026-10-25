using System.IO;
using System.Text;

namespace ExcelMerger
{
    /// <summary>
    /// Общий хвост командной строки Ghostscript: ресурсы вшитого движка, файл вывода и вход.
    /// Одно место на все построители аргументов (сжатие, преобразования, растеризация):
    /// раньше каждый собирал этот хвост сам, и ошибка в нём повторялась трижды.
    ///
    /// Главная тонкость — знак «%» в -sOutputFile. Движок читает это значение как шаблон
    /// printf (так задаётся «файл на страницу»: page-%04d.png), поэтому обычный путь с
    /// процентом ломается: «Предоплата 50%.pdf» даёт ошибку формата, а «Contract%20draft.pdf»
    /// тихо превращается в шаблон и порождает по файлу на страницу. Буквальный «%» движок
    /// принимает только удвоенным — это и делает <see cref="LiteralOutput"/>. Вход и пути -I
    /// шаблонами не являются и не экранируются. Чистые функции — под тестами.
    /// </summary>
    internal static class GsArguments
    {
        /// <summary>
        /// pdfwrite по умолчанию сам «угадывает» ориентацию страницы по преобладающему тексту
        /// (AutoRotatePages: /PageByPage у /default и /screen, /All у /ebook — таблица
        /// VectorDevices). Альбомная таблица на книжном листе или чертёж с подписями сбоку
        /// после сжатия, серого или починки разворачивались, и ни одна проверка (страницы,
        /// размер, цвет) этого не видела. Пересобранная страница обязана стоять как стояла.
        /// Общий для всех построителей pdfwrite — правка в одном месте.
        /// </summary>
        internal const string KeepPageOrientation = " -dAutoRotatePages=/None";

        /// <summary>Путь в кавычках: пробелы в путях — норма, без кавычек путь распадётся на два аргумента.</summary>
        internal static string Quote(string s)
        {
            return "\"" + s + "\"";
        }

        /// <summary>Буквальный путь вывода в синтаксисе -sOutputFile: каждый «%» удваивается.</summary>
        internal static string LiteralOutput(string path)
        {
            return string.IsNullOrEmpty(path) ? path : path.Replace("%", "%%");
        }

        /// <summary>
        /// Шаблон вывода «файл на страницу»: папка — буквально (её «%» экранируются), имя —
        /// шаблон движка как есть (в нём «%04d» и должен остаться форматом).
        /// </summary>
        internal static string OutputPattern(string directory, string fileNameFormat)
        {
            return Path.Combine(LiteralOutput(directory), fileNameFormat);
        }

        /// <summary>
        /// Дописать -I на ресурсы вшитого движка (bundledGsRoot != null), -sOutputFile и вход.
        /// outputSpec — уже в синтаксисе движка: из <see cref="LiteralOutput"/> или
        /// <see cref="OutputPattern"/>, чтобы вызывающий не мог забыть про экранирование
        /// молча — сырой путь здесь не принимается по смыслу параметра.
        /// </summary>
        internal static void AppendTail(StringBuilder sb, string bundledGsRoot,
            string outputSpec, string input)
        {
            if (!string.IsNullOrEmpty(bundledGsRoot))
            {
                // Вшитый движок: явные пути к его lib/Resource, чтобы не зависеть от
                // эвристики, которой он ищет их относительно себя.
                sb.Append(" -I ").Append(Quote(Path.Combine(bundledGsRoot, "lib")));
                sb.Append(" -I ").Append(Quote(Path.Combine(bundledGsRoot, "Resource", "Init")));
            }
            sb.Append(" -sOutputFile=").Append(Quote(outputSpec));
            sb.Append(' ').Append(Quote(input));
        }
    }
}
