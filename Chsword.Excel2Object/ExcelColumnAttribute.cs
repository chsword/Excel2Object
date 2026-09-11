using Chsword.Excel2Object.Styles;

namespace Chsword.Excel2Object;

/// <summary>
/// Attribute to define properties for Excel columns, including styles for cells and headers.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelColumnAttribute : ExcelTitleAttribute, IExcelHeaderStyle, IExcelCellStyle
{
    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelColumnAttribute"/> class.
    /// </summary>
    /// <param name="title">The title of the column.</param>
    public ExcelColumnAttribute(string title) : base(title)
    {
    }

    // Cell

    /// <summary>
    /// Gets or sets the horizontal alignment of the cell.
    /// </summary>
    public HorizontalAlignment CellAlignment { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the cell text is bold.
    /// </summary>
    public bool CellBold { get; set; }

    /// <summary>
    /// Gets or sets the font color of the cell.
    /// </summary>
    public ExcelStyleColor CellFontColor { get; set; }

    /// <summary>
    /// Gets or sets the font family of the cell.
    /// </summary>
    public string? CellFontFamily { get; set; }

    /// <summary>
    /// Gets or sets the font height of the cell.
    /// </summary>
    public double CellFontHeight { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the cell text is italic.
    /// </summary>
    public bool CellItalic { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the cell text has a strikeout.
    /// </summary>
    public bool CellStrikeout { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the cell text is underlined.
    /// </summary>
    public bool CellUnderline { get; set; }

    /// <summary>
    /// Gets or sets the format of the cell. It only has an effect on <see cref="System.DateTime"/>
    /// columns, where it is a .NET format string the value is rendered with before being written as
    /// text - for example "yyyy-MM-dd HH:mm:ss" turns the cell into the text "2026-09-11 14:30:45".
    /// <para>
    /// Two things to know, both long-standing:
    /// </para>
    /// <list type="bullet">
    /// <item>
    /// A format that happens to be one of Excel's builtin format names, such as "m/d/yy", is silently
    /// ignored - the cell comes out exactly as if no format had been given.
    /// </item>
    /// <item>
    /// When a formula column takes over such a column (they are matched by title), the value is
    /// written with this format and the formula is dropped.
    /// </item>
    /// </list>
    /// On every other column type - string, numeric, boolean, Uri - it is ignored.
    /// </summary>
    public string? Format { get; set; }

    // Header

    /// <summary>
    /// Gets or sets the horizontal alignment of the header.
    /// </summary>
    public HorizontalAlignment HeaderAlignment { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the header text is bold.
    /// </summary>
    public bool HeaderBold { get; set; }

    /// <summary>
    /// Gets or sets the font color of the header.
    /// </summary>
    public ExcelStyleColor HeaderFontColor { get; set; }

    /// <summary>
    /// Gets or sets the font family of the header.
    /// </summary>
    public string? HeaderFontFamily { get; set; }

    /// <summary>
    /// Gets or sets the font height of the header.
    /// </summary>
    public double HeaderFontHeight { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the header text is italic.
    /// </summary>
    public bool HeaderItalic { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the header text has a strikeout.
    /// </summary>
    public bool HeaderStrikeout { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the header text is underlined.
    /// </summary>
    public bool HeaderUnderline { get; set; }
}
