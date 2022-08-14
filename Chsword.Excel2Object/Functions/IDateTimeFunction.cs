namespace Chsword.Excel2Object.Functions;

public interface IDateTimeFunction
{
    ColumnValue Date(ColumnValue year, ColumnValue month, ColumnValue day);
    ColumnValue DateDif(ColumnValue start, ColumnValue end, string unit);
    ColumnValue Days(ColumnValue start, ColumnValue end);
    ColumnValue DateValue(ColumnValue date);
    ColumnValue Now();
    ColumnValue Time(ColumnValue hour, ColumnValue minute, ColumnValue second);
    ColumnValue TimeValue(ColumnValue time);
    ColumnValue Today();
    ColumnValue Weekday(ColumnValue date, ColumnValue firstDay);
    ColumnValue Year(ColumnValue date);
    ColumnValue YearFrac(ColumnValue start, ColumnValue end, string unit);
    ColumnValue Hour(ColumnValue time);
    ColumnValue Minute(ColumnValue time);
    ColumnValue Second(ColumnValue time);
    ColumnValue Month(ColumnValue date);
    ColumnValue Day(ColumnValue date);
}