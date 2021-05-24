namespace Chsword.Excel2Object.Functions
{
    public interface ITextFunction
    {
        ColumnValue Find(ColumnValue findText, ColumnValue withinText, int startNum);
        ColumnValue Find(ColumnValue findText, ColumnValue withinText);
        ColumnValue Asc(string str);
    }
}