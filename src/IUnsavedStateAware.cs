namespace ExcelMerger
{
    /// <summary>
    /// Окно с незаписанным рабочим состоянием. При смене языка ShellContext не
    /// пересоздаёт его: сохранность страниц/списков важнее немедленного перевода.
    /// </summary>
    internal interface IUnsavedStateAware
    {
        bool HasUncommittedState { get; }
    }
}
