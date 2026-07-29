using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
// В ExcelMerger уже есть свой CompressionLevel — уровни сжатия PDF (см. PdfCompression.cs),
// и внутри пространства имён он перекрывает системный. Псевдоним разводит эти два смысла:
// «насколько ужимать страницы» и «жать ли запись архива».
using ZipLevel = System.IO.Compression.CompressionLevel;

namespace ExcelMerger
{
    /// <summary>
    /// Сборка документа Open XML: набор частей → ZIP-архив по правилам OPC. Слой знает про
    /// упаковку и НИЧЕГО не знает про презентации — из тех же кирпичей собирается любой
    /// формат Office.
    ///
    /// Три правила упаковки, нарушение каждого из которых даёт «файл повреждён» без единой
    /// подробности от читателя:
    /// 1. `[Content_Types].xml` идёт ПЕРВОЙ записью архива и описывает КАЖДУЮ часть — либо
    ///    правилом по расширению (Default), либо поимённо (Override).
    /// 2. Имя части в архиве пишется БЕЗ ведущего слэша, а в `[Content_Types].xml` — С ним.
    ///    Это единственное расхождение во всём формате, и оно же самая частая ошибка.
    /// 3. XML-части пишутся в UTF-8 БЕЗ метки порядка байтов (BOM): читатель ждёт объявление
    ///    &lt;?xml с первого байта.
    ///
    /// Архив детерминирован (фиксированная дата записей): один и тот же вход даёт побайтово
    /// один и тот же файл — иначе тесты сравнивали бы движущуюся цель.
    /// </summary>
    internal sealed class OoxmlPackage
    {
        /// <summary>Дата записей архива: минимальная, какую формат ZIP умеет хранить.</summary>
        private static readonly DateTimeOffset ZipEpoch = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

        private static readonly UTF8Encoding Utf8NoBom = new UTF8Encoding(false);

        private sealed class Part
        {
            public string Name;         // «ppt/slides/slide1.xml» — без ведущего слэша
            public byte[] Data;
            public string ContentType;  // null — тип берётся из правила по расширению
            public bool Compress;
        }

        private readonly List<Part> _parts = new List<Part>();
        private readonly Dictionary<string, string> _defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public OoxmlPackage()
        {
            // Связи и разметка есть в любом документе Open XML.
            _defaults["rels"] = "application/vnd.openxmlformats-package.relationships+xml";
            _defaults["xml"] = "application/xml";
        }

        /// <summary>Число частей (без `[Content_Types].xml`, который добавляется при записи).</summary>
        public int Count { get { return _parts.Count; } }

        /// <summary>
        /// XML-часть. contentType — поимённый тип части (Override); null для файлов связей
        /// (.rels), тип которых задан правилом по расширению. Текст кодируется в UTF-8 без BOM.
        /// </summary>
        public void AddXml(string name, string xml, string contentType = null)
        {
            Add(name, Utf8NoBom.GetBytes(xml ?? string.Empty), contentType, true);
        }

        /// <summary>
        /// Двоичная часть (картинка). Тип задаётся правилом по расширению — так одна запись
        /// в `[Content_Types].xml` покрывает сотню картинок вместо сотни строк Override.
        /// Сжимать уже сжатое (PNG, JPEG) незачем: время тратится, размер не меняется.
        /// </summary>
        public void AddBinary(string name, byte[] data, string contentType)
        {
            if (string.IsNullOrEmpty(contentType))
                throw new ArgumentException("тип содержимого обязателен", "contentType");
            string ext = Extension(name);
            if (ext.Length == 0)
                throw new ArgumentException("двоичной части нужно расширение: " + name, "name");
            string known;
            if (_defaults.TryGetValue(ext, out known) && known != contentType)
                throw new ArgumentException("расширение «" + ext + "» уже отдано типу " + known, "contentType");
            _defaults[ext] = contentType;
            Add(name, data ?? new byte[0], null, false);
        }

        private void Add(string name, byte[] data, string contentType, bool compress)
        {
            ValidateName(name);
            if (!_names.Add(name))
                throw new ArgumentException("часть уже добавлена: " + name, "name");
            _parts.Add(new Part { Name = name, Data = data, ContentType = contentType, Compress = compress });
        }

        /// <summary>
        /// Имя части: без ведущего слэша и обратных слэшей, без пустых сегментов, только ASCII.
        /// Не-ASCII в имени заставляет ZIP выставлять флаг кодировки, а читатели Office
        /// расходятся в его трактовке — имена частей мы задаём сами, поэтому просто запрещаем.
        /// </summary>
        private static void ValidateName(string name)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("пустое имя части", "name");
            if (name[0] == '/' || name.IndexOf('\\') >= 0)
                throw new ArgumentException("имя части — без ведущего слэша и с прямыми слэшами: " + name, "name");
            if (name.IndexOf("//", StringComparison.Ordinal) >= 0)
                throw new ArgumentException("пустой сегмент в имени части: " + name, "name");
            for (int i = 0; i < name.Length; i++)
                if (name[i] > 0x7E || name[i] < 0x20)
                    throw new ArgumentException("имя части — только ASCII: " + name, "name");
        }

        private static string Extension(string name)
        {
            int dot = name.LastIndexOf('.');
            int slash = name.LastIndexOf('/');
            return dot > slash && dot >= 0 && dot + 1 < name.Length ? name.Substring(dot + 1) : string.Empty;
        }

        /// <summary>
        /// `[Content_Types].xml`: правила по расширению + поимённые типы частей. Порядок
        /// детерминирован (расширения по алфавиту, части — в порядке добавления). Чистая — под тест.
        /// </summary>
        internal string BuildContentTypes()
        {
            var exts = new List<string>(_defaults.Keys);
            exts.Sort(StringComparer.OrdinalIgnoreCase);
            var sb = new StringBuilder(512);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
            sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            foreach (string ext in exts)
                sb.Append("<Default Extension=\"").Append(ext)
                  .Append("\" ContentType=\"").Append(_defaults[ext]).Append("\"/>");
            foreach (Part p in _parts)
                if (p.ContentType != null)
                    sb.Append("<Override PartName=\"/").Append(p.Name)   // здесь слэш ОБЯЗАТЕЛЕН
                      .Append("\" ContentType=\"").Append(p.ContentType).Append("\"/>");
            sb.Append("</Types>");
            return sb.ToString();
        }

        /// <summary>Записать пакет в файл. Поток файловый — большой набор подложек не держим в памяти дважды.</summary>
        public void Save(string path)
        {
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                Write(fs);
        }

        /// <summary>Пакет целиком в памяти — для тестов и проверок.</summary>
        public byte[] ToArray()
        {
            using (var ms = new MemoryStream())
            {
                Write(ms);
                return ms.ToArray();
            }
        }

        private void Write(Stream target)
        {
            // leaveOpen: поток закрывает вызывающий (using выше), архив — только свои структуры.
            using (var zip = new ZipArchive(target, ZipArchiveMode.Create, true))
            {
                WriteEntry(zip, "[Content_Types].xml", Utf8NoBom.GetBytes(BuildContentTypes()), true);
                foreach (Part p in _parts)
                    WriteEntry(zip, p.Name, p.Data, p.Compress);
            }
        }

        private static void WriteEntry(ZipArchive zip, string name, byte[] data, bool compress)
        {
            ZipArchiveEntry entry = zip.CreateEntry(name,
                compress ? ZipLevel.Optimal : ZipLevel.NoCompression);
            entry.LastWriteTime = ZipEpoch;
            using (Stream s = entry.Open())
                s.Write(data, 0, data.Length);
        }
    }
}
