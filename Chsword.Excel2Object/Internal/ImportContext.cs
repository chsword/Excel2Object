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

    /// <summary>标记回调自身抛出的异常，使其在向外传播途中不被再次当作读取失败上报。</summary>
    private const string AbortKey = "Chsword.Excel2Object.AbortingFromCallback";

    /// <summary>
    ///     上报某个单元格的读取失败。未设置回调时不作任何输出——类库不应向标准输出写日志。
    /// </summary>
    /// <remarks>
    ///     读取路径有嵌套：读取日期单元格会先取其文本，取文本又可能递归求值公式。回调若抛出异常
    ///     （即调用方选择中止导入），该异常会途经外层的 catch，若不加区分将被再次上报，回调也就被
    ///     调用了两次。因此回调抛出的异常在此标记，再次经过时原样抛出，既不重复上报，也不会被外层
    ///     的 catch 吞掉。
    /// </remarks>
    public void Report(ICell? cell, Exception exception)
    {
        Report(cell?.Sheet?.SheetName, cell?.RowIndex ?? -1, cell?.ColumnIndex ?? -1, exception);
    }

    private void Report(string? sheetTitle, int rowIndex, int columnIndex, Exception exception)
    {
        if (_options.OnCellError == null) return;

        if (exception.Data.Contains(AbortKey))
            ExceptionDispatchInfo.Capture(exception).Throw();

        try
        {
            _options.OnCellError(new ExcelImportError(sheetTitle, rowIndex, columnIndex, exception));
        }
        catch (Exception fromCallback)
        {
            if (!fromCallback.Data.IsReadOnly) fromCallback.Data[AbortKey] = true;
            throw;
        }
    }
}
