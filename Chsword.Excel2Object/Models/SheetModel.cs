namespace Chsword.Excel2Object;

internal class SheetModel
{
    public List<ExcelColumn> Columns { get; set; } = new();
    public int Index { get; set; }
    /// <summary>
    ///     数据行。声明为惰性序列而非列表：流式导出要边取边写，一次只在内存里留一行。
    ///     该序列只会被遍历一次。
    /// </summary>
    public IEnumerable<Dictionary<string, object>> Rows { get; set; } = Enumerable.Empty<Dictionary<string, object>>();
    public string Title { get; private set; } = null!;

    public static SheetModel Create(string? title)
    {
        return new SheetModel
        {
            Title = title ?? "Sheet1"
        };
    }
}