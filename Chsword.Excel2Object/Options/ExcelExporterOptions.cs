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

    /// <summary>
    ///     Write <see cref="DateTime" /> columns as text, the way versions before 2.4.0 did, instead of as
    ///     real date cells (default: false). A text date cannot be sorted, filtered or calculated with in
    ///     Excel, but comes out character for character as <c>[ExcelColumn(Format = ...)]</c> renders it.
    /// </summary>
    public bool DateTimeAsText { get; set; }

    /// <summary>
    ///     Freeze the header row, so it stays in view while the sheet is scrolled (default: false).
    /// </summary>
    public bool FreezeHeader { get; set; }

    /// <summary>
    ///     The values a column's cells are limited to, keyed by column title - the runtime counterpart of
    ///     <see cref="ExcelColumnAttribute.Dropdown" />, which it takes precedence over. Excel shows them
    ///     as a dropdown and rejects anything else.
    /// </summary>
    public IDictionary<string, string[]> Dropdowns { get; set; } = new Dictionary<string, string[]>();

    /// <summary>
    ///     Put Excel's filter dropdowns on the header row, over the data written in this export
    ///     (default: false).
    /// </summary>
    public bool AutoFilter { get; set; }
}