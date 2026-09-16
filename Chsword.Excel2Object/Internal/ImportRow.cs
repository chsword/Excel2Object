using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     一行数据，与它从哪里来无关：整份读入内存的工作簿，或流式读取解出的一行。
/// </summary>
internal interface IImportRow
{
    string? SheetTitle { get; }

    int RowIndex { get; }

    /// <summary>该行存在的各格，按列序号给出。</summary>
    IEnumerable<KeyValuePair<int, CellData>> Cells { get; }

    /// <summary>该行有几格。整行为空的行不参与导入，判断这一点不必把每格都取出来。</summary>
    int CellCount { get; }

    /// <summary>某一列上的那一格；不存在时为 <see cref="CellData.None" />。</summary>
    CellData Cell(int columnIndex);
}

/// <summary>整份读入内存时的一行，取值经由 NPOI 的单元格。</summary>
internal sealed class NpoiRow : IImportRow
{
    private readonly ImportContext _context;
    private readonly bool _date1904;
    private readonly IRow _row;

    public NpoiRow(IRow row, ImportContext context, bool date1904)
    {
        _row = row;
        _context = context;
        _date1904 = date1904;
    }

    public string? SheetTitle => _row.Sheet?.SheetName;

    public int RowIndex => _row.RowNum;

    public int CellCount => _row.Cells.Count;

    public IEnumerable<KeyValuePair<int, CellData>> Cells
    {
        get
        {
            foreach (var cell in _row.Cells)
                yield return new KeyValuePair<int, CellData>(cell.ColumnIndex, Of(cell, _context, _date1904));
        }
    }

    public CellData Cell(int columnIndex)
    {
        var cell = _row.GetCell(columnIndex);
        return cell == null ? CellData.None : Of(cell, _context, _date1904);
    }

    /// <summary>
    ///     把 NPOI 的单元格取成与来源无关的一格。公式在此求值：求值器按工作簿创建一次，其结果的类型
    ///     与取值即为这一格的类型与取值。
    /// </summary>
    private static CellData Of(ICell cell, ImportContext context, bool date1904)
    {
        try
        {
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    var style = cell.CellStyle;
                    return CellData.OfNumber(cell.NumericCellValue, style?.DataFormat ?? 0,
                        style?.GetDataFormatString(), date1904);
                case CellType.String:
                    return CellData.OfText(cell.StringCellValue);
                case CellType.Boolean:
                    return CellData.OfBoolean(cell.BooleanCellValue);
                case CellType.Blank:
                    return CellData.Blank();
                case CellType.Formula:
                    return Of(context.Evaluator(cell.Sheet.Workbook).EvaluateInCell(cell), context, date1904);
                default:
                    return CellData.OfError(cell.ToString() ?? string.Empty);
            }
        }
        catch (Exception e)
        {
            context.Report(cell.Sheet?.SheetName, cell.RowIndex, cell.ColumnIndex, e);
            return CellData.Failure;
        }
    }
}
