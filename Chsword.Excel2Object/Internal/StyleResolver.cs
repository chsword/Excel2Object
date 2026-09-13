using Chsword.Excel2Object.Styles;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     Works out how one cell looks, by laying the styles an export declares over each other from the
///     widest to the narrowest.
/// </summary>
/// <remarks>
///     For a data cell: every cell, then the row's stripe, then the column's <c>[ExcelColumn]</c>
///     attribute, and last the style written for that column in the stylesheet. For a header cell: the
///     stylesheet's header style, then the attribute's <c>Header…</c> properties. Only the properties a
///     style sets take part, so the layers add up rather than replace one another.
/// </remarks>
internal static class StyleResolver
{
    /// <summary>
    ///     The size a header has had since before there was a stylesheet: any column declaring an
    ///     <c>[ExcelColumn]</c> got one whether or not it asked. It sits under everything, so a size set
    ///     anywhere - on the attribute or in the stylesheet - is the one that shows.
    /// </summary>
    private static readonly ExcelStyle LegacyHeaderSize =
        new ExcelStyle().FontSize(ExcelConstants.DefaultFontHeightInPoints);

    public static ExcelStyle? Header(ExcelColumn column, ExcelStyleSheet styles)
    {
        var resolved = Merge(FromHeader(column.HeaderStyle), styles.HeaderStyle);
        // the legacy size applies to the columns it always applied to, and steps aside as soon as the
        // export writes a header style of its own - one sheet should not mix sizes row by column
        return column.HeaderStyle == null || styles.HeaderStyle != null ? resolved : Merge(resolved, LegacyHeaderSize);
    }

    /// <summary>
    ///     How a column's cells look on one kind of row, settled once for the whole sheet.
    /// </summary>
    /// <param name="rowIndex">Any row of that kind; only its odd- or evenness is read.</param>
    public static ResolvedColumnStyle Resolve(ExcelColumn column, ExcelStyleSheet styles, int rowIndex)
    {
        var style = Cell(column, styles, rowIndex);
        var written = ForThisColumn(column, styles);
        var fallback = Broad(styles, rowIndex);

        // a date column takes a date format, a value column takes anything else; what was written for
        // the column itself is used either way
        var forColumn = ColumnStyleFormat(column, styles);
        var date = written ?? (IsDateFormat(fallback) ? fallback : null);
        var value = forColumn ?? (IsDateFormat(fallback) ? null : fallback);

        return new ResolvedColumnStyle(style, date, value, forColumn);
    }

    /// <summary>
    ///     The format written for this column and no wider: the stylesheet's <c>Column(title)</c>, else
    ///     the <c>[ExcelColumn]</c>'s own - which has always meant the format of a date column.
    /// </summary>
    private static string? ForThisColumn(ExcelColumn column, ExcelStyleSheet styles)
    {
        return ColumnStyleFormat(column, styles) ?? column.CellStyle?.Format;
    }

    private static string? ColumnStyleFormat(ExcelColumn column, ExcelStyleSheet styles)
    {
        return column.Title != null && styles.ColumnStyles.TryGetValue(column.Title, out var byTitle)
            ? byTitle.NumberFormat
            : null;
    }

    /// <summary>The format written for every cell, or for the stripe this row falls on.</summary>
    private static string? Broad(ExcelStyleSheet styles, int rowIndex)
    {
        var stripe = Stripe(styles, rowIndex);
        return stripe?.NumberFormat ?? styles.CellsStyle?.NumberFormat;
    }

    private static bool IsDateFormat(string? format)
    {
        return format != null &&
               ExcelDateFormat.PartsShown(ExcelDateFormat.ToExcel(format)) != ExcelDateFormat.Parts.None;
    }

    private static ExcelStyle? Stripe(ExcelStyleSheet styles, int rowIndex)
    {
        return (rowIndex - ExcelConstants.DefaultDataStartRowIndex) % 2 == 0
            ? styles.OddRowsStyle
            : styles.EvenRowsStyle;
    }

    public static ExcelStyle? Cell(ExcelColumn column, ExcelStyleSheet styles, int rowIndex)
    {
        var resolved = Merge(Stripe(styles, rowIndex), styles.CellsStyle);
        resolved = Merge(FromCell(column.CellStyle), resolved);
        if (column.Title != null && styles.ColumnStyles.TryGetValue(column.Title, out var byTitle))
            resolved = Merge(byTitle, resolved);

        return resolved;
    }

    /// <summary>
    ///     The cell half of an <c>[ExcelColumn]</c>, as a style. Its properties say nothing when they are
    ///     left at their default, which is what the attribute has always meant by them.
    /// </summary>
    private static ExcelStyle? FromCell(IExcelCellStyle? attribute)
    {
        if (attribute == null) return null;
        var style = new ExcelStyle();

        if (!string.IsNullOrWhiteSpace(attribute.CellFontFamily)) style.FontFamily(attribute.CellFontFamily!);
        if (attribute.CellFontHeight > 0) style.FontSize(attribute.CellFontHeight);
        if (attribute.CellFontColor > 0) style.Color(attribute.CellFontColor);
        if (attribute.CellBold) style.Bold();
        if (attribute.CellItalic) style.Italic();
        if (attribute.CellStrikeout) style.Strikeout();
        if (attribute.CellUnderline) style.Underline();
        if (attribute.CellAlignment != HorizontalAlignment.General) style.Align(attribute.CellAlignment);
        // Format is not copied: on the attribute it has always meant the format of a date column only,
        // while a style's Format applies to whatever column it is written for

        return style.IsEmpty ? null : style;
    }

    private static ExcelStyle? FromHeader(IExcelHeaderStyle? attribute)
    {
        if (attribute == null) return null;
        var style = new ExcelStyle();

        if (!string.IsNullOrWhiteSpace(attribute.HeaderFontFamily)) style.FontFamily(attribute.HeaderFontFamily!);
        if (attribute.HeaderFontHeight > 0) style.FontSize(attribute.HeaderFontHeight);
        if (attribute.HeaderFontColor > 0) style.Color(attribute.HeaderFontColor);
        if (attribute.HeaderBold) style.Bold();
        if (attribute.HeaderItalic) style.Italic();
        if (attribute.HeaderStrikeout) style.Strikeout();
        if (attribute.HeaderUnderline) style.Underline();
        if (attribute.HeaderAlignment != HorizontalAlignment.General) style.Align(attribute.HeaderAlignment);

        return style.IsEmpty ? null : style;
    }

    /// <summary>One style laid over another, either of which may be nothing at all.</summary>
    private static ExcelStyle? Merge(ExcelStyle? over, ExcelStyle? under)
    {
        if (over == null) return under;
        return under == null ? over : over.Over(under);
    }
}
