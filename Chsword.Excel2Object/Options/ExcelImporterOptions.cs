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

    /// <summary>
    ///     模型上写着的某个列标题在表头中找不到时的回调，用于得知哪一列没有对上。默认不设回调，
    ///     此时该属性保持其默认值，导入照常进行——与既有版本一致。
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         标题差一个空格、多一个单位（「金额」与「金额（元）」），导入并不会报错，只是那一列
    ///         悄悄全为空。这是导入类问题中最常见的一种，<see cref="OnCellError" /> 只覆盖到单元格
    ///         层面，对此无能为力。
    ///     </para>
    ///     <para>回调在读过表头之后、取第一行数据之前调用，每个对不上的标题调用一次。</para>
    ///     <example>
    ///         <code>
    /// // 只是记下来
    /// options.OnMissingColumn = missing =&gt; logger.Warn(missing.ToString());
    ///
    /// // 或者干脆不接受这样的文件
    /// options.OnMissingColumn = missing =&gt; throw new Excel2ObjectException(missing.ToString());
    ///         </code>
    ///     </example>
    /// </remarks>
    public Action<ExcelColumnMissing>? OnMissingColumn { get; set; }
}
