namespace Chsword.Excel2Object.Functions;

/// <summary>
///     Excel date and time functions. See <see cref="IExcelFunction" /> for how these are translated.
/// </summary>
public interface IDateTimeFunction : IExcelFunction
{
    // ---- building dates and times ----

    /// <summary>DATE - the date of the given year, month and day.</summary>
    ColumnValue Date(ColumnValue year, ColumnValue month, ColumnValue day);

    /// <summary>DATEVALUE - converts a date written as text to a date.</summary>
    ColumnValue DateValue(ColumnValue date);

    /// <summary>TIME - the time of the given hour, minute and second.</summary>
    ColumnValue Time(ColumnValue hour, ColumnValue minute, ColumnValue second);

    /// <summary>TIMEVALUE - converts a time written as text to a time.</summary>
    ColumnValue TimeValue(ColumnValue time);

    /// <summary>NOW - the current date and time.</summary>
    ColumnValue Now();

    /// <summary>TODAY - the current date.</summary>
    ColumnValue Today();

    // ---- parts of a date ----

    /// <summary>YEAR - the year of a date.</summary>
    ColumnValue Year(ColumnValue date);

    /// <summary>MONTH - the month of a date.</summary>
    ColumnValue Month(ColumnValue date);

    /// <summary>DAY - the day of a date.</summary>
    ColumnValue Day(ColumnValue date);

    /// <summary>HOUR - the hour of a time.</summary>
    ColumnValue Hour(ColumnValue time);

    /// <summary>MINUTE - the minute of a time.</summary>
    ColumnValue Minute(ColumnValue time);

    /// <summary>SECOND - the second of a time.</summary>
    ColumnValue Second(ColumnValue time);

    /// <summary>WEEKDAY - the day of the week of a date, 1 for Sunday.</summary>
    ColumnValue Weekday(ColumnValue date);

    /// <summary>WEEKDAY - the day of the week of a date, numbered by <paramref name="firstDay" />.</summary>
    ColumnValue Weekday(ColumnValue date, ColumnValue firstDay);

    /// <summary>WEEKNUM - the week of the year a date falls in.</summary>
    [ExcelFunctionName("WEEKNUM")]
    ColumnValue WeekNum(ColumnValue date);

    /// <summary>WEEKNUM - the week of the year, with the week numbering scheme of <paramref name="returnType" />.</summary>
    [ExcelFunctionName("WEEKNUM")]
    ColumnValue WeekNum(ColumnValue date, ColumnValue returnType);

    /// <summary>ISOWEEKNUM - the ISO 8601 week of the year a date falls in.</summary>
    [ExcelFunctionName("ISOWEEKNUM", Future = true)]
    ColumnValue IsoWeekNum(ColumnValue date);

    // ---- arithmetic on dates ----

    /// <summary>EDATE - the date the given number of months before or after a date.</summary>
    [ExcelFunctionName("EDATE")]
    ColumnValue EDate(ColumnValue date, ColumnValue months);

    /// <summary>EOMONTH - the last day of the month, the given number of months away.</summary>
    [ExcelFunctionName("EOMONTH")]
    ColumnValue EoMonth(ColumnValue date, ColumnValue months);

    /// <summary>DAYS - the number of days between two dates.</summary>
    [ExcelFunctionName("DAYS", Future = true)]
    ColumnValue Days(ColumnValue start, ColumnValue end);

    /// <summary>DAYS360 - the number of days between two dates on a 360 day year.</summary>
    [ExcelFunctionName("DAYS360")]
    ColumnValue Days360(ColumnValue start, ColumnValue end);

    /// <summary>DATEDIF - the difference of two dates in the given unit ("Y", "M", "D", "MD", "YM", "YD").</summary>
    [ExcelFunctionName("DATEDIF")]
    ColumnValue DateDif(ColumnValue start, ColumnValue end, string unit);

    /// <summary>YEARFRAC - the fraction of a year between two dates.</summary>
    [ExcelFunctionName("YEARFRAC")]
    ColumnValue YearFrac(ColumnValue start, ColumnValue end);

    /// <summary>YEARFRAC - the fraction of a year between two dates, on the given day count basis.</summary>
    [ExcelFunctionName("YEARFRAC")]
    ColumnValue YearFrac(ColumnValue start, ColumnValue end, string unit);

    /// <summary>YEARFRAC - the fraction of a year between two dates, on the given day count basis.</summary>
    [ExcelFunctionName("YEARFRAC")]
    ColumnValue YearFrac(ColumnValue start, ColumnValue end, ColumnValue basis);

    /// <summary>NETWORKDAYS - the number of working days between two dates.</summary>
    [ExcelFunctionName("NETWORKDAYS")]
    ColumnValue NetworkDays(ColumnValue start, ColumnValue end);

    /// <summary>NETWORKDAYS - the number of working days between two dates, excluding holidays.</summary>
    [ExcelFunctionName("NETWORKDAYS")]
    ColumnValue NetworkDays(ColumnValue start, ColumnValue end, ColumnMatrix holidays);

    /// <summary>NETWORKDAYS.INTL - working days between two dates, with a custom weekend.</summary>
    [ExcelFunctionName("NETWORKDAYS.INTL", Future = true)]
    ColumnValue NetworkDaysIntl(ColumnValue start, ColumnValue end, ColumnValue weekend);

    /// <summary>NETWORKDAYS.INTL - as above, excluding holidays.</summary>
    [ExcelFunctionName("NETWORKDAYS.INTL", Future = true)]
    ColumnValue NetworkDaysIntl(ColumnValue start, ColumnValue end, ColumnValue weekend, ColumnMatrix holidays);

    /// <summary>WORKDAY - the date the given number of working days away.</summary>
    [ExcelFunctionName("WORKDAY")]
    ColumnValue WorkDay(ColumnValue start, ColumnValue days);

    /// <summary>WORKDAY - the date the given number of working days away, excluding holidays.</summary>
    [ExcelFunctionName("WORKDAY")]
    ColumnValue WorkDay(ColumnValue start, ColumnValue days, ColumnMatrix holidays);

    /// <summary>WORKDAY.INTL - as WORKDAY, with a custom weekend.</summary>
    [ExcelFunctionName("WORKDAY.INTL", Future = true)]
    ColumnValue WorkDayIntl(ColumnValue start, ColumnValue days, ColumnValue weekend);

    /// <summary>WORKDAY.INTL - as above, excluding holidays.</summary>
    [ExcelFunctionName("WORKDAY.INTL", Future = true)]
    ColumnValue WorkDayIntl(ColumnValue start, ColumnValue days, ColumnValue weekend, ColumnMatrix holidays);
}
