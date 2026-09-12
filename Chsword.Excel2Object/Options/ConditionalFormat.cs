using Chsword.Excel2Object.Styles;

namespace Chsword.Excel2Object.Options;

/// <summary>How a cell's value is compared with <see cref="ConditionalFormat.Value" />.</summary>
public enum ConditionalOperator
{
    /// <summary>No comparison: the rule is the <see cref="ConditionalFormat.Formula" />.</summary>
    None = 0,
    Between,
    NotBetween,
    Equal,
    NotEqual,
    GreaterThan,
    LessThan,
    GreaterThanOrEqual,
    LessThanOrEqual
}

/// <summary>
///     A rule that restyles the cells of one column when their value meets a condition - Excel's
///     conditional formatting, so the colours follow the data as it is edited.
/// </summary>
/// <example>
///     <code>
/// options.ConditionalFormats.Add(new ConditionalFormat("金额")
/// {
///     Operator = ConditionalOperator.GreaterThan,
///     Value = 10000,
///     FontColor = ExcelStyleColor.Red,
///     Bold = true
/// });
/// </code>
/// </example>
public class ConditionalFormat
{
    public ConditionalFormat()
    {
    }

    public ConditionalFormat(string title)
    {
        Title = title;
    }

    /// <summary>The title of the column the rule watches.</summary>
    public string? Title { get; set; }

    /// <summary>
    ///     How the cell is compared with <see cref="Value" />. Leave it unset to write the rule as a
    ///     <see cref="Formula" /> instead.
    /// </summary>
    public ConditionalOperator Operator { get; set; }

    /// <summary>
    ///     What the cell is compared with. A number, a <see cref="DateTime" /> or a bool is written as
    ///     Excel's literal for it, and a string is quoted, so <c>Value = "急件"</c> compares with that text.
    /// </summary>
    public object? Value { get; set; }

    /// <summary>The other end of <see cref="ConditionalOperator.Between" /> and its negation.</summary>
    public object? Value2 { get; set; }

    /// <summary>
    ///     An Excel condition of your own, true when the rule should apply, e.g. <c>$C2&gt;$D2</c>. It is
    ///     written against the first data row (row 2) and Excel moves it down the column from there, so
    ///     the row part stays relative while the column part is usually anchored with <c>$</c>.
    /// </summary>
    public string? Formula { get; set; }

    /// <summary>
    ///     Restyle the whole row rather than the cell that met the condition (default: false).
    /// </summary>
    public bool WholeRow { get; set; }

    /// <summary>The colour the text takes when the rule applies.</summary>
    public ExcelStyleColor FontColor { get; set; }

    /// <summary>The colour the cells are filled with when the rule applies.</summary>
    public ExcelStyleColor BackgroundColor { get; set; }

    public bool Bold { get; set; }

    public bool Italic { get; set; }
}
