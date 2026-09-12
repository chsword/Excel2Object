using System.Globalization;
using Chsword.Excel2Object.Options;
using Chsword.Excel2Object.Styles;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     Writes the conditional formatting rules an export declares: Excel re-evaluates them as the sheet
///     is edited, so the colours keep following the data.
/// </summary>
internal static class ConditionalFormatting
{
    public static void Apply(ISheet sheet, ExcelColumn[] columns, int lastDataRowIndex,
        IList<ConditionalFormat> formats)
    {
        if (formats.Count == 0) return;

        // an export with no rows still gets its rules, so a template keeps them once it is filled in
        var lastRow = Math.Max(lastDataRowIndex, ExcelConstants.DefaultDataStartRowIndex);
        var sheetFormatting = sheet.SheetConditionalFormatting;

        foreach (var format in formats)
        {
            var column = IndexOf(columns, format.Title);
            Validate(format);
            var rule = CreateRule(sheetFormatting, format, column);

            var font = rule.CreateFontFormatting();
            if (format.FontColor > 0) font.FontColorIndex = (short) format.FontColor;
            if (format.Bold || format.Italic) font.SetFontStyle(format.Italic, format.Bold);

            if (format.BackgroundColor > 0)
            {
                var pattern = rule.CreatePatternFormatting();
                pattern.FillBackgroundColor = (short) format.BackgroundColor;
                pattern.FillPattern = FillPattern.SolidForeground;
            }

            var lastColumn = format.WholeRow ? columns.Length - 1 : column;
            var range = new CellRangeAddress(ExcelConstants.DefaultDataStartRowIndex, lastRow,
                format.WholeRow ? 0 : column, lastColumn);
            sheetFormatting.AddConditionalFormatting(new[] {range}, rule);
        }
    }

    private static int IndexOf(ExcelColumn[] columns, string? title)
    {
        for (var i = 0; i < columns.Length; i++)
            if (columns[i].Title == title)
                return i;

        throw new Excel2ObjectException(
            $"Conditional format refers to column [{title}], which is not in the sheet.");
    }

    /// <summary>
    ///     Says what a rule is missing before NPOI is handed an incomplete one, naming the column and the
    ///     operator so the rule at fault is the one you go and look at.
    /// </summary>
    private static void Validate(ConditionalFormat format)
    {
        if (format.Operator == ConditionalOperator.None)
        {
            if (string.IsNullOrEmpty(format.Formula))
                throw new Excel2ObjectException(
                    $"{Describe(format)} has neither an Operator to compare with nor a Formula of its own.");
            return;
        }

        if (format.Value == null)
            throw new Excel2ObjectException($"{Describe(format)} needs a Value to compare with.");

        if (format.Value2 == null &&
            format.Operator is ConditionalOperator.Between or ConditionalOperator.NotBetween)
            throw new Excel2ObjectException(
                $"{Describe(format)} needs a Value2 as well - {format.Operator} takes both ends of the range.");
    }

    private static string Describe(ConditionalFormat format)
    {
        return format.Operator == ConditionalOperator.None
            ? $"The conditional format on column [{format.Title}]"
            : $"The conditional format on column [{format.Title}] ({format.Operator})";
    }

    private static IConditionalFormattingRule CreateRule(ISheetConditionalFormatting sheetFormatting,
        ConditionalFormat format, int column)
    {
        // a rule over the whole row can only be an expression - Excel compares the cell it is written on
        if (format.Operator == ConditionalOperator.None || format.WholeRow)
            return sheetFormatting.CreateConditionalFormattingRule(
                format.Formula ?? ComparisonFormula(format, column));

        return sheetFormatting.CreateConditionalFormattingRule(Operator(format.Operator),
            Literal(format.Value), format.Value2 == null ? null : Literal(format.Value2));
    }

    /// <summary>The comparison written as an expression, for a rule that colours more than its own cell.</summary>
    private static string ComparisonFormula(ConditionalFormat format, int column)
    {
        // the column is anchored and the row is not, so Excel moves the rule down the rows
        var cell = $"${CellReference.ConvertNumToColString(column)}{ExcelConstants.DefaultDataStartRowIndex + 1}";
        var value = Literal(format.Value);

        return format.Operator switch
        {
            ConditionalOperator.Between => $"AND({cell}>={value},{cell}<={Literal(format.Value2)})",
            ConditionalOperator.NotBetween => $"OR({cell}<{value},{cell}>{Literal(format.Value2)})",
            ConditionalOperator.Equal => $"{cell}={value}",
            ConditionalOperator.NotEqual => $"{cell}<>{value}",
            ConditionalOperator.GreaterThan => $"{cell}>{value}",
            ConditionalOperator.LessThan => $"{cell}<{value}",
            ConditionalOperator.GreaterThanOrEqual => $"{cell}>={value}",
            ConditionalOperator.LessThanOrEqual => $"{cell}<={value}",
            _ => throw new Excel2ObjectException($"Unsupported conditional operator {format.Operator}.")
        };
    }

    private static ComparisonOperator Operator(ConditionalOperator value)
    {
        return value switch
        {
            ConditionalOperator.Between => ComparisonOperator.Between,
            ConditionalOperator.NotBetween => ComparisonOperator.NotBetween,
            ConditionalOperator.Equal => ComparisonOperator.Equal,
            ConditionalOperator.NotEqual => ComparisonOperator.NotEqual,
            ConditionalOperator.GreaterThan => ComparisonOperator.GreaterThan,
            ConditionalOperator.LessThan => ComparisonOperator.LessThan,
            ConditionalOperator.GreaterThanOrEqual => ComparisonOperator.GreaterThanOrEqual,
            ConditionalOperator.LessThanOrEqual => ComparisonOperator.LessThanOrEqual,
            _ => throw new Excel2ObjectException($"Unsupported conditional operator {value}.")
        };
    }

    /// <summary>What the value looks like inside an Excel formula.</summary>
    private static string Literal(object? value)
    {
        return value switch
        {
            null => throw new Excel2ObjectException("A conditional format has no value to compare with."),
            DateTime date => $"DATE({date.Year},{date.Month},{date.Day})",
            bool flag => flag ? "TRUE" : "FALSE",
            string text => "\"" + text.Replace("\"", "\"\"") + "\"",
            IFormattable number when value is not char => number.ToString(null, CultureInfo.InvariantCulture),
            _ => "\"" + value.ToString()!.Replace("\"", "\"\"") + "\""
        };
    }
}
