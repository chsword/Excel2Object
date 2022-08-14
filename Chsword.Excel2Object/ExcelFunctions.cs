using Chsword.Excel2Object.Functions;

namespace Chsword.Excel2Object;

public static class ExcelFunctions
{
    public static IAllFunction All { get; set; } = null!;
    public static IConditionFunction Condition { get; set; } = null!;
    public static IDateTimeFunction DateAndTime { get; set; } = null!;

    public static IMathFunction Math { get; set; } = null!;
    public static IReferenceFunction Reference { get; set; } = null!;
    public static IStatisticsFunction Statistics { get; set; } = null!;
    public static ITextFunction Text { get; set; } = null!;
}