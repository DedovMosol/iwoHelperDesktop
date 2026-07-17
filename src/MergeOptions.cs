namespace ExcelMerger
{
    /// <summary>Параметры объединения. По умолчанию все опции выключены (поведение как раньше).</summary>
    public class MergeOptions
    {
        /// <summary>Добавить первым листом «Содержание»: оглавление с гиперссылками и статусами файлов.</summary>
        public bool AddToc;

        /// <summary>Заменить формулы вычисленными значениями — свод не будет зависеть от исходных файлов.</summary>
        public bool ValuesOnly;

        /// <summary>Брать все видимые листы каждого файла (иначе — только первый видимый).</summary>
        public bool AllSheets;
    }
}
