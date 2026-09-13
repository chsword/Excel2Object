namespace Chsword.Excel2Object.Options;

/// <summary>
///     一处要合并的区域。列以标题指定，行以数据行的序号指定，均无需知道单元格地址。
/// </summary>
/// <remarks>
///     列在工作表中的位置由 <c>Order</c>、特性的声明顺序以及公式列的插入位置共同决定，行数则取决于
///     数据条数，二者在编写导出代码时都无从得知，因此此处不以 <c>A1</c> 这样的地址表达。
///     确已知道最终布局时，仍可直接写地址：<c>options.MergedRegions.Add("A1:C1")</c> 会隐式转换为
///     一处区域。
/// </remarks>
/// <example>
///     <code>
/// // 表头行上，把「省份」到「城市」两列并成一格
/// options.MergedRegions.Add(new MergedRegion("省份", "城市"));
///
/// // 「备注」列的第 1 至第 3 行数据并成一格
/// options.MergedRegions.Add(new MergedRegion("备注") { FirstRow = 1, LastRow = 3 });
/// </code>
/// </example>
public class MergedRegion
{
    public MergedRegion()
    {
    }

    public MergedRegion(string column, string? lastColumn = null)
    {
        Column = column;
        LastColumn = lastColumn;
    }

    /// <summary>起始列的标题。</summary>
    public string? Column { get; set; }

    /// <summary>结束列的标题；省略则只占 <see cref="Column" /> 一列。</summary>
    public string? LastColumn { get; set; }

    /// <summary>
    ///     起始的数据行序号，自 1 计起（1 即表头下方的第一行数据）。省略则指表头行本身。
    /// </summary>
    public int? FirstRow { get; set; }

    /// <summary>结束的数据行序号；省略则只占 <see cref="FirstRow" /> 一行。</summary>
    public int? LastRow { get; set; }

    /// <summary>直接写定的单元格地址，如 <c>A1:C1</c>；由字符串隐式转换而来。</summary>
    internal string? Address { get; private set; }

    /// <summary>
    ///     已知最终布局时，可直接以 <c>A1:C1</c> 这样的地址书写。
    /// </summary>
    public static implicit operator MergedRegion(string address)
    {
        return new MergedRegion {Address = address};
    }

    public override string ToString()
    {
        if (Address != null) return Address;
        var columns = LastColumn == null ? Column : $"{Column}..{LastColumn}";
        if (FirstRow == null) return $"表头 [{columns}]";
        return LastRow == null ? $"第 {FirstRow} 行 [{columns}]" : $"第 {FirstRow}–{LastRow} 行 [{columns}]";
    }
}
