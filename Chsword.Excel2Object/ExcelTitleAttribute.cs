namespace Chsword.Excel2Object;

/// <summary>
///     On property ,it will be a column title
///     On class ,it will be a sheet title
/// </summary>
public class ExcelTitleAttribute : Attribute
{
    public ExcelTitleAttribute(string title)
    {
        Title = title;
    }

    public int Order { get; set; }
    public string Title { get; set; }
}