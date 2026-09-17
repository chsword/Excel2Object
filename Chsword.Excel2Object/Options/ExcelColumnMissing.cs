namespace Chsword.Excel2Object.Options;

/// <summary>
///     模型上写着的某个列标题，在表头中找不到。该属性因而不会被填上，整列都是默认值。
/// </summary>
/// <remarks>
///     这是导入类问题中最常见的一种：标题差了一个空格、多了个单位（「金额」与「金额（元）」），
///     导入不会报错，只是那一列悄悄全为空。<see cref="ExcelImporterOptions.OnMissingColumn" /> 即为
///     此而设。
/// </remarks>
public class ExcelColumnMissing
{
    public ExcelColumnMissing(string title, string propertyName, string? sheetTitle,
        IReadOnlyList<string> headerTitles, IReadOnlyList<string> similarTitles)
    {
        Title = title;
        PropertyName = propertyName;
        SheetTitle = sheetTitle;
        HeaderTitles = headerTitles;
        SimilarTitles = similarTitles;
    }

    /// <summary>模型上写着的标题。</summary>
    public string Title { get; }

    /// <summary>要它的那个属性。</summary>
    public string PropertyName { get; }

    public string? SheetTitle { get; }

    /// <summary>表头上实际有的标题，按列的先后。</summary>
    public IReadOnlyList<string> HeaderTitles { get; }

    /// <summary>
    ///     表头上与之相近的标题：一方包含另一方即算（「金额」与「金额（元）」），不作更多猜测。
    /// </summary>
    public IReadOnlyList<string> SimilarTitles { get; }

    public override string ToString()
    {
        var where = SheetTitle == null ? "" : $"工作表 [{SheetTitle}] 的";
        var similar = SimilarTitles.Count == 0
            ? ""
            : $"，表头上与之相近的是 [{string.Join("]、[", SimilarTitles)}]";
        return $"{where}表头中没有 [{Title}] 这一列（属性 {PropertyName} 因而不会被填上）{similar}。";
    }
}
