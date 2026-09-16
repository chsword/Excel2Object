using System.Runtime.ExceptionServices;
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

    /// <summary>回调自身抛出的异常，它在向外传播途中不应被再次当作读取失败上报。</summary>
    private Exception? _abortingWith;

    /// <summary>
    ///     上报某个单元格的读取失败。未设置回调时不作任何输出——类库不应向标准输出写日志。
    /// </summary>
    /// <remarks>
    ///     读取路径有嵌套：读取日期单元格会先取其文本，取文本又可能递归求值公式。回调若抛出异常
    ///     （即调用方选择中止导入），该异常会途经外层的 catch，若不加区分将被再次上报，回调也就被
    ///     调用了两次。因此此处记下回调抛出的异常，再次经过时原样抛出，既不重复上报，也不会被外层
    ///     的 catch 吞掉。该记录保存在本次导入的上下文中，不改动调用方的异常对象。
    /// </remarks>
    public void Report(ICell? cell, Exception exception)
    {
        Report(cell?.Sheet?.SheetName, cell?.RowIndex ?? -1, cell?.ColumnIndex ?? -1, exception);
    }

    /// <summary>单元格可能并不存在（空单元格不占位），此时以其所在行与列号定位。</summary>
    public void Report(IRow? row, int columnIndex, Exception exception)
    {
        var cell = row?.GetCell(columnIndex);
        if (cell != null)
        {
            Report(cell, exception);
            return;
        }

        Report(row?.Sheet?.SheetName, row?.RowNum ?? -1, columnIndex, exception);
    }

    /// <summary>位置由调用方给出：流式读取没有单元格对象，只有行列号。</summary>
    public void Report(string? sheetTitle, int rowIndex, int columnIndex, Exception exception)
    {
        if (_options.OnCellError == null) return;

        if (ReferenceEquals(exception, _abortingWith))
            ExceptionDispatchInfo.Capture(exception).Throw();

        try
        {
            _options.OnCellError(new ExcelImportError(sheetTitle, rowIndex, columnIndex, exception));
        }
        catch (Exception fromCallback)
        {
            _abortingWith = fromCallback;
            throw;
        }
    }
}
