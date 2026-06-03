namespace CommonLibrary.Models
{
    public enum OfficeType
    {
        Excel,
        Word,
        PowerPoint
    }

    public enum RuleType
    {
        CellValue,
        Formula,
        Formatting,
        ObjectExistence,
        RangeComparison
    }

    public enum ValidationStatus
    {
        NotValidated,
        Valid,
        HasWarnings,
        Invalid
    }
}
