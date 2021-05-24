namespace Chsword.Excel2Object.Functions
{
    public interface IReferenceFunction
    {
        ColumnValue Lookup(ColumnValue val, ColumnMatrix lookupVector, ColumnMatrix resultVector);

        ColumnValue VLookup(ColumnValue val, ColumnMatrix tableArray, ColumnValue colIndexNum,
            bool rangeLookup = false);

        ColumnValue Match(ColumnValue val, ColumnMatrix tableArray, int matchType);
        ColumnValue Choose(ColumnValue indexNum, params ColumnValue[] values);
        ColumnValue Index(ColumnMatrix array, ColumnValue rowNum, ColumnValue columnNum);
    }
}