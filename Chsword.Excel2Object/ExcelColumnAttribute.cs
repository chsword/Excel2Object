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
    /// Gets or sets the display format of a <see cref="System.DateTime"/> column. The column is written
    /// as real date cells, and this .NET format string is translated into the Excel number format that
    /// shows the same thing - "yyyy-MM-dd HH:mm:ss" becomes the cell format <c>yyyy-mm-dd hh:mm:ss</c>,
    /// "yyyy年MM月dd日" becomes <c>yyyy"年"mm"月"dd"日"</c>. Without it the column shows
    /// <c>yyyy-mm-dd hh:mm:ss</c>. An Excel format such as "m/d/yy", "[$-409]d-mmm-yy" or "yyyy/m/d;@" is
    /// taken as is.
    /// <para>
    /// Parts Excel cannot show - time zone offsets (<c>zzz</c>, <c>K</c>) and eras (<c>g</c>) - are
    /// dropped, and <c>hh</c> without <c>tt</c> shows the 24-hour clock, Excel having no 12-hour clock
    /// without an AM/PM marker. Excel also reads a minute as a month unless it sits next to the hour or the
    /// seconds, so write "HH:mm" rather than "HH时mm分" - the literal in between would make Excel show the
    /// month. A date before 1900, which Excel's calendar does not reach, is written as
    /// text rendered with this format instead; so is every date when
    /// <see cref="Options.ExcelExporterOptions.DateTimeAsText"/> is set.
    /// </para>
    /// A formula column that takes over such a column (they are matched by title) keeps the format,
    /// unless its <see cref="Options.FormulaColumn.FormulaResultType"/> says the formula yields something
    /// other than a date. On every other column type - string, numeric, boolean, Uri - it is ignored.
    /// </summary>
    public string? Format { get; set; }

    /// <summary>
    /// Gets or sets the values this column's cells are limited to. Excel shows them as a dropdown and
    /// rejects anything else: <c>[ExcelColumn("状态", Dropdown = new[] {"启用", "停用"})]</c>.
    /// <para>
    /// A list Excel cannot hold inline - over 255 characters in total, or a value carrying a comma or a
    /// quote - is written to a hidden sheet the dropdown then reads from, which works the same way in
    /// Excel. For values only known at runtime use <see cref="Options.ExcelExporterOptions.Dropdowns"/>,
    /// which takes precedence over this one.
    /// </para>
    /// </summary>
    public string[]? Dropdown { get; set; }

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
