namespace Chsword.Excel2Object.Options;

/// <summary>
///     一张工作表的表头：表名与各列的标题，按列的先后。
/// </summary>
public class ExcelSheetHeader
{
    public ExcelSheetHeader(string? sheetTitle, IReadOnlyList<string> columns)
    {
        SheetTitle = sheetTitle;
        Columns = columns;
    }

    /// <summary>实际读的那张表的名字：未指定表名时即第一张表。</summary>
    public string? SheetTitle { get; }

    /// <summary>表头上的各列标题，按列的先后；表中没有任何行时为空。</summary>
    public IReadOnlyList<string> Columns { get; }
}
