using Chsword.Excel2Object.Options;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     一次导入过程中跨单元格共享的状态：读取失败的上报方式，以及公式求值器。
/// </summary>
internal sealed class ImportContext
{
    private readonly ExcelImporterOptions _options;
    private IFormulaEvaluator? _evaluator;
    private IWorkbook? _evaluatorWorkbook;

    public ImportContext(ExcelImporterOptions options)
    {
        _options = options;
    }

    /// <summary>
    ///     工作簿的公式求值器。求值器按工作簿创建一次并在整个导入过程中复用：其构造会读取工作簿的
    ///     公式环境，逐单元格创建会使含公式的表格导入变慢数倍。
    /// </summary>
    public IFormulaEvaluator Evaluator(IWorkbook workbook)
    {
        if (_evaluator != null && ReferenceEquals(_evaluatorWorkbook, workbook)) return _evaluator;

        _evaluator = WorkbookFactory.CreateFormulaEvaluator(workbook);
        _evaluatorWorkbook = workbook;
        return _evaluator;
    }

    /// <summary>
    ///     上报某个单元格的读取失败。未设置回调时不作任何输出——类库不应向标准输出写日志。
    /// </summary>
    public void Report(ICell? cell, Exception exception)
    {
        if (_options.OnCellError == null) return;
        _options.OnCellError(new ExcelImportError(cell?.Sheet?.SheetName, cell?.RowIndex ?? -1,
            cell?.ColumnIndex ?? -1, exception));
    }

    public void Report(IRow? row, int columnIndex, Exception exception)
    {
        if (_options.OnCellError == null) return;
        _options.OnCellError(new ExcelImportError(row?.Sheet?.SheetName, row?.RowNum ?? -1, columnIndex,
            exception));
    }
}
