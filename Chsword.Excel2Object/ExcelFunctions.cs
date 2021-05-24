using Chsword.Excel2Object.Functions;

namespace Chsword.Excel2Object
{
    public static class ExcelFunctions
    {
        public static IAllFunction All { get; set;}

        public static IMathFunction Math { get; set; }
        public static IStatisticsFunction Statistics { get; set; }
        public static IConditionFunction Condition { get; set; }
        public static IReferenceFunction Reference { get; set; }
        public static IDateTimeFunction DateAndTime { get; set; }
        public static ITextFunction  Text { get; set; }

    }
}