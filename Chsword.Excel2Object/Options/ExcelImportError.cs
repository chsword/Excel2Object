using NPOI.SS.Util;

namespace Chsword.Excel2Object.Options;

/// <summary>
///     某个单元格未能读取时的说明，通过 <see cref="ExcelImporterOptions.OnCellError" /> 报出。
/// </summary>
/// <remarks>
///     导入不会因单个单元格失败而中止：该单元格取默认值，其余数据照常读取。若希望失败即中止，
///     可在回调中抛出异常，例如 <c>options.OnCellError = e =&gt; throw e.Exception;</c>。
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

    /// <summary>所在行的序号，自 0 计起（0 为标题行）。</summary>
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
