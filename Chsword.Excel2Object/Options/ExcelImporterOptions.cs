namespace Chsword.Excel2Object.Options;

public class ExcelImporterOptions
{
    public string? SheetTitle { get; set; }
    public int TitleSkipLine { get; set; }

    /// <summary>
    ///     某个单元格未能读取时的回调。默认为 <c>null</c>，此时该单元格取默认值并继续导入。
    ///     在回调中抛出异常即可使导入中止：<c>options.OnCellError = e =&gt; throw e.Exception;</c>。
    /// </summary>
    public Action<ExcelImportError>? OnCellError { get; set; }
}
