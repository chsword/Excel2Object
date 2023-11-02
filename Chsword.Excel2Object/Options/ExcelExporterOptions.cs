namespace Chsword.Excel2Object.Options;

public class ExcelExporterOptions
{
    /// <summary>
    ///     Excel file type default:xlsx
    /// </summary>
    public ExcelType ExcelType { get; set; } = ExcelType.Xlsx;

    public FormulaColumnsCollection FormulaColumns { get; set; } = new();

    /// <summary>
    ///     Sheet Title default:null
    /// </summary>
    public string? SheetTitle { get; set; }

    /// <summary>
    ///     Use when append export
    /// </summary>
    public byte[]? SourceExcelBytes { get; set; }

    public Func<string,Type,string>? MappingColumnAction { get; set; }
}