using Chsword.Excel2Object.Functions;

namespace Chsword.Excel2Object;

/// <summary>
///     The Excel functions usable inside a formula expression, by category.
/// </summary>
/// <remarks>
///     None of these are ever executed: a formula lambda is only ever read as an expression tree and
///     translated to formula text, so <c>ExcelFunctions.Math.Abs(c["Age"])</c> becomes <c>ABS(B4)</c>.
/// </remarks>
public static class ExcelFunctions
{
    /// <summary>Every category at once, for formulas that mix them.</summary>
    public static IAllFunction All { get; set; } = null!;

    /// <summary>Logical functions: IF, AND, OR, IFERROR, SWITCH, ...</summary>
    public static IConditionFunction Condition { get; set; } = null!;

    /// <summary>Date and time functions: DATE, EOMONTH, NETWORKDAYS, ...</summary>
    public static IDateTimeFunction DateAndTime { get; set; } = null!;

    /// <summary>Database functions: DSUM, DGET, DCOUNT, ...</summary>
    public static IDatabaseFunction Database { get; set; } = null!;

    /// <summary>Engineering functions: CONVERT, DEC2BIN, BITAND, ...</summary>
    public static IEngineeringFunction Engineering { get; set; } = null!;

    /// <summary>Financial functions: PMT, NPV, IRR, ...</summary>
    public static IFinancialFunction Financial { get; set; } = null!;

    /// <summary>Information functions: ISBLANK, ISNUMBER, NA, ...</summary>
    public static IInformationFunction Information { get; set; } = null!;

    /// <summary>Math and trigonometry functions: ABS, MOD, POWER, SUMIF, ...</summary>
    public static IMathFunction Math { get; set; } = null!;

    /// <summary>Lookup and reference functions: VLOOKUP, INDEX, MATCH, ...</summary>
    public static IReferenceFunction Reference { get; set; } = null!;

    /// <summary>Statistical functions: SUM, AVERAGE, COUNTIF, MAX, ...</summary>
    public static IStatisticsFunction Statistics { get; set; } = null!;

    /// <summary>Text functions: LEFT, MID, LEN, TEXT, TEXTJOIN, ...</summary>
    public static ITextFunction Text { get; set; } = null!;
}
