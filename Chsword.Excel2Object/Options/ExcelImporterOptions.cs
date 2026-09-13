namespace Chsword.Excel2Object.Options;

public class ExcelImporterOptions
{
    public string? SheetTitle { get; set; }
    public int TitleSkipLine { get; set; }

    /// <summary>
    ///     某个单元格未能读取或未能转换为目标类型时的回调，用于得知是哪一个单元格、因何失败。
    /// </summary>
    /// <remarks>
    ///     两类失败的后续处理并不相同，回调只负责告知：
    ///     <list type="bullet">
    ///         <item>
    ///             <description>
    ///                 <b>读取失败</b>（例如日期序列号超出 Excel 日历）：该单元格取默认值，导入继续。
    ///             </description>
    ///         </item>
    ///         <item>
    ///             <description>
    ///                 <b>转换失败</b>（例如文本 <c>abc</c> 写入 <c>int</c> 属性）：上报之后照旧抛出，
    ///                 导入中止。这一行为与既有版本一致。
    ///             </description>
    ///         </item>
    ///     </list>
    ///     在回调中抛出异常，可使读取失败也中止导入：
    ///     <c>options.OnCellError = e =&gt; throw e.Exception;</c>。
    /// </remarks>
    public Action<ExcelImportError>? OnCellError { get; set; }
}
