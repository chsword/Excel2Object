namespace Chsword.Excel2Object.Internal;

internal static class ExcelConstants
{
    public const int DefaultColumnWidthMultiplier = 256;
    public const int DefaultHeaderRowIndex = 0;
    public const int DefaultDataStartRowIndex = 1;
    public const short DefaultFontHeightInPoints = 10;
    
    public static class CellTypes
    {
        public const string Text = "text";
        public const string DateTime = "datetime";
        public const string Number = "number";
    }
    
    public static class BooleanValues
    {
        public static readonly string[] TrueValues = { "1", "是", "yes", "true" };
        public static readonly string[] FalseValues = { "0", "否", "no", "false" };
    }
    
    public static class DateFormats
    {
        public const string YearSuffix = "年";
        public const string MonthSuffix = "月";
        public const string DaySuffix = "日";
        public const string DefaultYearMonthSuffix = "-01-01";
        public const string DefaultDaySuffix = "-01";
    }
}
