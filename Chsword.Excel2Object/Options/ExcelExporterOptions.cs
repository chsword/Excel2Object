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

    public Func<string, Type, string>? MappingColumnAction { get; set; }

    /// <summary>
    ///     Enable auto column width adjustment based on content
    /// </summary>
    public bool AutoColumnWidth { get; set; } = false;

    /// <summary>
    ///     Minimum column width in characters (default: 8)
    /// </summary>
    public int MinColumnWidth { get; set; } = 8;

    /// <summary>
    ///     Maximum column width in characters (default: 50)
    /// </summary>
    public int MaxColumnWidth { get; set; } = 50;

    /// <summary>
    ///     Default column width in characters when AutoColumnWidth is false (default: 16)
    /// </summary>
    public int DefaultColumnWidth { get; set; } = 16;
}