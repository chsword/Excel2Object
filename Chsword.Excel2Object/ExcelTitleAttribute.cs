namespace Chsword.Excel2Object;

/// <summary>
///     On property, it will be a column title.
///     On class, it will be a sheet title.
/// </summary>
public class ExcelTitleAttribute : Attribute
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelTitleAttribute"/> class.
    /// </summary>
    /// <param name="title">The title of the column or sheet.</param>
    public ExcelTitleAttribute(string title)
    {
        Title = title;
    }

    /// <summary>
    ///     Gets or sets the order of the column or sheet.
    /// </summary>
    public int Order { get; set; }

    /// <summary>
    ///     Gets or sets the title of the column or sheet.
    /// </summary>
    public string Title { get; set; }
}