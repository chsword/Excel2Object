namespace Chsword.Excel2Object.Functions;

/// <summary>
/// Interface for date and time functions used in Excel to object conversion.
/// </summary>
public interface IDateTimeFunction
{
    /// <summary>
    /// Returns the sequential serial number that represents a particular date.
    /// </summary>
    ColumnValue Date(ColumnValue year, ColumnValue month, ColumnValue day);

    /// <summary>
    /// Calculates the difference between two dates.
    /// </summary>
    ColumnValue DateDif(ColumnValue startDate, ColumnValue endDate, ColumnValue unit);

    /// <summary>
    /// Returns the number of days between two dates.
    /// </summary>
    ColumnValue Days(ColumnValue endDate, ColumnValue startDate);

    /// <summary>
    /// Converts a date in the form of text to a serial number.
    /// </summary>
    ColumnValue DateValue(ColumnValue dateText);

    /// <summary>
    /// Returns the current date and time.
    /// </summary>
    ColumnValue Now();

    /// <summary>
    /// Returns the serial number for a particular time.
    /// </summary>
    ColumnValue Time(ColumnValue hour, ColumnValue minute, ColumnValue second);

    /// <summary>
    /// Converts a time in the form of text to a serial number.
    /// </summary>
    ColumnValue TimeValue(ColumnValue timeText);

    /// <summary>
    /// Returns the current date.
    /// </summary>
    ColumnValue Today();

    /// <summary>
    /// Returns the day of the week corresponding to a date.
    /// </summary>
    ColumnValue Weekday(ColumnValue serialNumber, ColumnValue returnType);

    /// <summary>
    /// Returns the day of the week corresponding to a date (default return type).
    /// </summary>
    ColumnValue Weekday(ColumnValue serialNumber);

    /// <summary>
    /// Returns the year of a date.
    /// </summary>
    ColumnValue Year(ColumnValue serialNumber);

    /// <summary>
    /// Calculates the fraction of the year represented by the number of whole days between two dates.
    /// </summary>
    ColumnValue YearFrac(ColumnValue startDate, ColumnValue endDate, ColumnValue basis);

    /// <summary>
    /// Calculates the fraction of the year represented by the number of whole days between two dates (default basis).
    /// </summary>
    ColumnValue YearFrac(ColumnValue startDate, ColumnValue endDate);

    /// <summary>
    /// Returns the hour of a time value.
    /// </summary>
    ColumnValue Hour(ColumnValue serialNumber);

    /// <summary>
    /// Returns the minute of a time value.
    /// </summary>
    ColumnValue Minute(ColumnValue serialNumber);

    /// <summary>
    /// Returns the second of a time value.
    /// </summary>
    ColumnValue Second(ColumnValue serialNumber);

    /// <summary>
    /// Returns the month of a date.
    /// </summary>
    ColumnValue Month(ColumnValue serialNumber);

    /// <summary>
    /// Returns the day of a date.
    /// </summary>
    ColumnValue Day(ColumnValue serialNumber);

    /// <summary>
    /// Returns the day of the year.
    /// </summary>
    ColumnValue DayOfYear(ColumnValue serialNumber);

    /// <summary>
    /// Returns the week number of the year.
    /// </summary>
    ColumnValue WeekNum(ColumnValue serialNumber, ColumnValue returnType);

    /// <summary>
    /// Returns the week number of the year (default return type).
    /// </summary>
    ColumnValue WeekNum(ColumnValue serialNumber);

    /// <summary>
    /// Returns a date that is a specified number of months before or after a specified date.
    /// </summary>
    ColumnValue EDate(ColumnValue startDate, ColumnValue months);

    /// <summary>
    /// Returns the last day of the month that is a specified number of months before or after a specified date.
    /// </summary>
    ColumnValue EoMonth(ColumnValue startDate, ColumnValue months);

    /// <summary>
    /// Converts a date to a number representing where that date falls within a 1900 date system.
    /// </summary>
    ColumnValue Days360(ColumnValue startDate, ColumnValue endDate, ColumnValue method);

    /// <summary>
    /// Converts a date to a number representing where that date falls within a 1900 date system (default method).
    /// </summary>
    ColumnValue Days360(ColumnValue startDate, ColumnValue endDate);

    /// <summary>
    /// Returns the number of working days between two dates.
    /// </summary>
    ColumnValue NetworkDays(ColumnValue startDate, ColumnValue endDate, ColumnMatrix holidays);

    /// <summary>
    /// Returns the number of working days between two dates (no holidays).
    /// </summary>
    ColumnValue NetworkDays(ColumnValue startDate, ColumnValue endDate);

    /// <summary>
    /// Returns a date that is a specified number of working days before or after a date.
    /// </summary>
    ColumnValue WorkDay(ColumnValue startDate, ColumnValue days, ColumnMatrix holidays);

    /// <summary>
    /// Returns a date that is a specified number of working days before or after a date (no holidays).
    /// </summary>
    ColumnValue WorkDay(ColumnValue startDate, ColumnValue days);
}