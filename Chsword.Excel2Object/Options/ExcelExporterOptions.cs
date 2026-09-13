using Chsword.Excel2Object.Styles;

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
    ///     How the sheet looks, written by what each style applies to - the header, every cell, one column,
    ///     every other row - instead of repeated on every <c>[ExcelColumn]</c>.
    /// </summary>
    public ExcelStyleSheet Styles { get; set; } = new();

    /// <summary>
    ///     Rules that restyle cells whose value meets a condition - Excel re-evaluates them as the sheet is
    ///     edited, so the colours keep following the data.
    /// </summary>
    public IList<ConditionalFormat> ConditionalFormats { get; set; } = new List<ConditionalFormat>();

    /// <summary>
    ///     这些列中连续相同的值会并成一格，按列标题指定。分组报表常用，例如同一城市的若干行只显示一次
    ///     城市名。
    /// </summary>
    /// <remarks>
    ///     仅相邻且相等的行参与，空单元格不参与；各列彼此独立判断。被并入的单元格仍保留各自的值，
    ///     Excel 只显示左上角那一个，因此导回对象时每一行的数据依然完整。
    ///     需要注意：Excel 中含合并单元格的区域无法排序，若同时启用 <see cref="AutoFilter" />，
    ///     筛选可用而排序会被 Excel 拒绝。
    /// </remarks>
    public IList<string> MergeRepeatedColumns { get; set; } = new List<string>();

    /// <summary>
    ///     另行指定要合并的区域，用于上面那条规则覆盖不到的情形。列以标题指定、行以数据行序号指定，
    ///     无需知道单元格地址：
    ///     <code>
    /// options.MergedRegions.Add(new MergedRegion("省份", "城市"));                    // 表头行，跨两列
    /// options.MergedRegions.Add(new MergedRegion("备注") {FirstRow = 1, LastRow = 3}); // 前三行数据
    /// options.MergedRegions.Add("A1:C1");                                             // 已知布局时
    ///     </code>
    /// </summary>
    public IList<MergedRegion> MergedRegions { get; set; } = new List<MergedRegion>();

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