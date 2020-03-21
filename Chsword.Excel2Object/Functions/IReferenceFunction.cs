namespace Chsword.Excel2Object.Functions
{
    public interface IReferenceFunction
    {
        ColumnValue Lookup(ColumnValue val, ColumnMatrix lookupVector, ColumnMatrix resultVector);

        ColumnValue VLookup(ColumnValue val, ColumnMatrix tableArray, ColumnValue colIndexNum,
            bool rangeLookup = false);
    }
}