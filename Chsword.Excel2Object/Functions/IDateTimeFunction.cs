namespace Chsword.Excel2Object.Functions
{
    public interface IDateTimeFunction
    {
        ColumnValue Date(ColumnValue year, ColumnValue month, ColumnValue day);
        ColumnValue DateDif(ColumnValue start, ColumnValue end, string unit);
        ColumnValue Days(ColumnValue start, ColumnValue end);
    }
}