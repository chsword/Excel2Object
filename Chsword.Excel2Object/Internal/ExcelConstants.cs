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
        public const string Boolean = "boolean";
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
        
        /// <summary>
        /// Provides a comprehensive list of common date and time format patterns used for parsing date/time strings from Excel cells.
        /// </summary>
        /// <remarks>
        /// The formats are organized into groups, including:
        /// <list type="bullet">
        /// <item>ISO 8601 formats (e.g., "yyyy-MM-ddTHH:mm:ss")</item>
        /// <item>Date with various separators (e.g., "yyyy/MM/dd", "dd-MM-yyyy")</item>
        /// <item>Date with time (12-hour and 24-hour formats, with or without AM/PM)</item>
        /// <item>Short date formats (single digit month/day)</item>
        /// <item>Time only formats (24-hour and 12-hour with AM/PM)</item>
        /// </list>
        /// This array is intended for use when attempting to parse date/time values from Excel cells that may be formatted in a variety of ways.
        /// </remarks>
        public static readonly string[] CommonDateTimeFormats = 
        [
            // ISO 8601 formats
            "yyyy-MM-ddTHH:mm:ss",
            "yyyy-MM-ddTHH:mm:ssZ",
            "yyyy-MM-ddTHH:mm:ss.fff",
            "yyyy-MM-ddTHH:mm:ss.fffZ",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-dd HH:mm",
            
            // Date with various separators
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "yyyy.MM.dd",
            "dd-MM-yyyy",
            "dd/MM/yyyy",
            "dd.MM.yyyy",
            "MM-dd-yyyy",
            "MM/dd/yyyy",
            "MM.dd.yyyy",
            
            // Date with time (12-hour format)
            "yyyy-MM-dd hh:mm:ss tt",
            "yyyy-MM-dd hh:mm tt",
            "dd-MM-yyyy hh:mm:ss tt",
            "dd/MM/yyyy hh:mm:ss tt",
            "MM-dd-yyyy hh:mm:ss tt",
            "MM/dd/yyyy hh:mm:ss tt",
            
            // Date with time (24-hour format)
            "dd-MM-yyyy HH:mm:ss",
            "dd/MM/yyyy HH:mm:ss",
            "MM-dd-yyyy HH:mm:ss",
            "MM/dd/yyyy HH:mm:ss",
            "dd-MM-yyyy HH:mm",
            "dd/MM/yyyy HH:mm",
            "MM-dd-yyyy HH:mm",
            "MM/dd/yyyy HH:mm",
            
            // Short date formats (with single digit month/day)
            "yyyy/M/d",
            "yyyy-M-d",
            "d/M/yyyy",
            "d-M-yyyy",
            "M/d/yyyy",
            "M-d-yyyy",
            
            // Time only formats (24-hour)
            "HH:mm:ss",
            "HH:mm",
            "H:mm:ss",
            "H:mm",
            
            // Time only formats (12-hour with AM/PM)
            "hh:mm:ss tt",
            "hh:mm tt",
            "h:mm:ss tt",
            "h:mm tt"
        ];
    }
}
