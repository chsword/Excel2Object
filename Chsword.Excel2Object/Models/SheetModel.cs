namespace Chsword.Excel2Object;

internal class SheetModel
{
    public List<ExcelColumn> Columns { get; set; } = new();
    public int Index { get; set; }
    public List<Dictionary<string, object>> Rows { get; set; } = new();
    public string Title { get; private set; } = null!;
    public static SheetModel Create(string? title)
    {
        return new SheetModel
        {
            Title = title ?? "Sheet1",
        };
    }
}