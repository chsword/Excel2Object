namespace Chsword.Excel2Object.Styles;

/// <summary>
///     The styles of an export, written by what they apply to rather than repeated on every column: the
///     header, every data cell, one column, or every other row.
/// </summary>
/// <remarks>
///     Styles layer from the widest to the narrowest, and only the properties a style sets take part:
///     <see cref="Cells" />, then <see cref="OddRows" /> / <see cref="EvenRows" />, then the column's own
///     <c>[ExcelColumn]</c> attribute, and last <see cref="Column" /> - the one written for that column
///     here, which is the most deliberate. <see cref="Header" /> layers under the attribute's
///     <c>Header…</c> properties the same way.
/// </remarks>
public class ExcelStyleSheet
{
    internal ExcelStyle? HeaderStyle { get; private set; }
    internal ExcelStyle? CellsStyle { get; private set; }
    internal ExcelStyle? OddRowsStyle { get; private set; }
    internal ExcelStyle? EvenRowsStyle { get; private set; }
    internal Dictionary<string, ExcelStyle> ColumnStyles { get; } = new(StringComparer.Ordinal);

    /// <summary>True when the export declares no styles here at all.</summary>
    internal bool IsEmpty => HeaderStyle == null && CellsStyle == null && OddRowsStyle == null &&
                             EvenRowsStyle == null && ColumnStyles.Count == 0;

    /// <summary>How the header row looks.</summary>
    public ExcelStyleSheet Header(Action<ExcelStyle> style)
    {
        HeaderStyle = Build(HeaderStyle, style);
        return this;
    }

    /// <summary>How every data cell looks, under everything more specific.</summary>
    public ExcelStyleSheet Cells(Action<ExcelStyle> style)
    {
        CellsStyle = Build(CellsStyle, style);
        return this;
    }

    /// <summary>How the cells of one column look, found by its title.</summary>
    public ExcelStyleSheet Column(string title, Action<ExcelStyle> style)
    {
        ColumnStyles[title] = Build(ColumnStyles.TryGetValue(title, out var existing) ? existing : null, style);
        return this;
    }

    /// <summary>How the first, third, fifth … data row looks.</summary>
    public ExcelStyleSheet OddRows(Action<ExcelStyle> style)
    {
        OddRowsStyle = Build(OddRowsStyle, style);
        return this;
    }

    /// <summary>How the second, fourth, sixth … data row looks - the other half of a striped table.</summary>
    public ExcelStyleSheet EvenRows(Action<ExcelStyle> style)
    {
        EvenRowsStyle = Build(EvenRowsStyle, style);
        return this;
    }

    private static ExcelStyle Build(ExcelStyle? existing, Action<ExcelStyle> style)
    {
        if (style == null) throw new Excel2ObjectException("A style needs something to set.");
        var built = new ExcelStyle();
        style(built);
        return built.Over(existing);
    }
}
