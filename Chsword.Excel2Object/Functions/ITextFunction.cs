namespace Chsword.Excel2Object.Functions
{
    public interface ITextFunction
    {
        ColumnValue Asc(string str);
        ColumnValue Find(ColumnValue findText, ColumnValue withinText, int startNum);
        ColumnValue Find(ColumnValue findText, ColumnValue withinText);
    }
}