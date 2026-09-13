using NPOI.SS.Util;

namespace Chsword.Excel2Object.Options;

/// <summary>
///     某个单元格未能读取时的说明，通过 <see cref="ExcelImporterOptions.OnCellError" /> 报出。
/// </summary>
/// <remarks>
///     读取失败时该单元格取默认值、其余数据照常读取，可在回调中抛出异常使其中止；值无法转换为目标
///     类型时则在上报之后照旧抛出并中止导入。详见
///     <see cref="ExcelImporterOptions.OnCellError" />。
/// </remarks>
public class ExcelImportError
{
    internal ExcelImportError(string? sheetTitle, int rowIndex, int columnIndex, Exception exception)
    {
        SheetTitle = sheetTitle;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Exception = exception;
    }

    /// <summary>所在工作表的名称。</summary>
    public string? SheetTitle { get; }

    /// <summary>
    ///     所在行在工作表中的序号，自 0 计起。标题行的位置取决于
    ///     <see cref="ExcelImporterOptions.TitleSkipLine" />，并不必然是第 0 行。
    /// </summary>
    public int RowIndex { get; }

    /// <summary>所在列的序号，自 0 计起。</summary>
    public int ColumnIndex { get; }

    /// <summary>单元格地址，例如 <c>B3</c>。</summary>
    public string CellReference => new CellReference(RowIndex, ColumnIndex).FormatAsString();

    /// <summary>读取该单元格时抛出的异常。</summary>
    public Exception Exception { get; }

    public override string ToString()
    {
        return $"[{SheetTitle}!{CellReference}] {Exception.GetType().Name}: {Exception.Message}";
    }
}
