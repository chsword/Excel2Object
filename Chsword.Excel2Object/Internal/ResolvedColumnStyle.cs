using Chsword.Excel2Object.Styles;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     How one column's cells look on one kind of row - odd or even - worked out once for the sheet
///     rather than once per cell.
/// </summary>
/// <remarks>
///     The number format is settled here too, and separately from the rest of the look, because where it
///     was written decides where it applies. A format written for this column - in the stylesheet's
///     <c>Column(title)</c> or on the <c>[ExcelColumn]</c> - is a deliberate statement about this column
///     and is used as it is. A format written for every cell, or for the row stripes, is a default: it
///     reaches the columns it has something to say to, so <c>#,##0.00</c> leaves the dates alone and a
///     date format leaves the numbers alone.
/// </remarks>
internal readonly struct ResolvedColumnStyle
{
    /// <summary>Everything but the number format: font, fill, alignment, borders.</summary>
    public ExcelStyle? Style { get; }

    /// <summary>The format a date cell of this column shows, if any.</summary>
    public string? DateFormat { get; }

    /// <summary>The format any other cell of this column shows, if any.</summary>
    public string? ValueFormat { get; }

    /// <summary>
    ///     The format a text cell of this column shows. Text is kept verbatim under Excel's text format
    ///     unless the export asked for something else for this very column: a number format that fell on
    ///     the whole sheet would quietly drop the protection a leading zero relies on.
    /// </summary>
    public string? TextFormat { get; }

    public ResolvedColumnStyle(ExcelStyle? style, string? dateFormat, string? valueFormat, string? textFormat)
    {
        Style = style;
        DateFormat = dateFormat;
        ValueFormat = valueFormat;
        TextFormat = textFormat;
    }
}
