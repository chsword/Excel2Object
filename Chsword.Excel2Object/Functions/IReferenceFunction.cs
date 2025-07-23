namespace Chsword.Excel2Object.Functions;

/// <summary>
/// Interface for reference functions used in Excel to object conversion.
/// </summary>
public interface IReferenceFunction
{
    /// <summary>
    /// Uses an index number to return a value from a list of values.
    /// </summary>
    ColumnValue Choose(ColumnValue indexNum, params ColumnValue[] values);

    /// <summary>
    /// Returns a value or reference to a value from within a table or range.
    /// </summary>
    ColumnValue Index(ColumnMatrix array, ColumnValue rowNum, ColumnValue columnNum);

    /// <summary>
    /// Returns a value or reference to a value from within a table or range (single parameter).
    /// </summary>
    ColumnValue Index(ColumnMatrix array, ColumnValue rowNum);

    /// <summary>
    /// Searches for a value in a vector and returns a corresponding value in another vector.
    /// </summary>
    ColumnValue Lookup(ColumnValue lookupValue, ColumnMatrix lookupVector, ColumnMatrix resultVector);

    /// <summary>
    /// Searches for a value in a vector and returns the value (approximate match).
    /// </summary>
    ColumnValue Lookup(ColumnValue lookupValue, ColumnMatrix lookupVector);

    /// <summary>
    /// Returns the relative position of an item in an array.
    /// </summary>
    ColumnValue Match(ColumnValue lookupValue, ColumnMatrix lookupArray, ColumnValue matchType);

    /// <summary>
    /// Searches for a value in the leftmost column of a table and returns a value in the same row from a specified column.
    /// </summary>
    ColumnValue VLookup(ColumnValue lookupValue, ColumnMatrix tableArray, ColumnValue colIndexNum, ColumnValue rangeLookup);

    /// <summary>
    /// Searches for a value in the leftmost column of a table and returns a value in the same row from a specified column (exact match).
    /// </summary>
    ColumnValue VLookup(ColumnValue lookupValue, ColumnMatrix tableArray, ColumnValue colIndexNum);

    /// <summary>
    /// Searches for a value in the top row of a table and returns a value in the same column from a specified row.
    /// </summary>
    ColumnValue HLookup(ColumnValue lookupValue, ColumnMatrix tableArray, ColumnValue rowIndexNum, ColumnValue rangeLookup);

    /// <summary>
    /// Searches for a value in the top row of a table and returns a value in the same column from a specified row (exact match).
    /// </summary>
    ColumnValue HLookup(ColumnValue lookupValue, ColumnMatrix tableArray, ColumnValue rowIndexNum);

    /// <summary>
    /// Returns the reference specified by a text string.
    /// </summary>
    ColumnValue Indirect(ColumnValue refText, ColumnValue a1);

    /// <summary>
    /// Returns the reference specified by a text string (A1 style).
    /// </summary>
    ColumnValue Indirect(ColumnValue refText);

    /// <summary>
    /// Returns a reference offset from a given reference.
    /// </summary>
    ColumnValue Offset(ColumnMatrix reference, ColumnValue rows, ColumnValue cols, ColumnValue height, ColumnValue width);

    /// <summary>
    /// Returns a reference offset from a given reference (without height and width).
    /// </summary>
    ColumnValue Offset(ColumnMatrix reference, ColumnValue rows, ColumnValue cols);

    /// <summary>
    /// Returns the row number of a reference.
    /// </summary>
    ColumnValue Row(ColumnMatrix reference);

    /// <summary>
    /// Returns the row number of the current cell.
    /// </summary>
    ColumnValue Row();

    /// <summary>
    /// Returns the column number of a reference.
    /// </summary>
    ColumnValue Column(ColumnMatrix reference);

    /// <summary>
    /// Returns the column number of the current cell.
    /// </summary>
    ColumnValue Column();

    /// <summary>
    /// Returns the number of rows in a reference.
    /// </summary>
    ColumnValue Rows(ColumnMatrix array);

    /// <summary>
    /// Returns the number of columns in a reference.
    /// </summary>
    ColumnValue Columns(ColumnMatrix array);

    /// <summary>
    /// Returns information about the formatting, location, or contents of a cell.
    /// </summary>
    ColumnValue Cell(ColumnValue infoType, ColumnMatrix reference);

    /// <summary>
    /// Returns information about the formatting, location, or contents of the current cell.
    /// </summary>
    ColumnValue Cell(ColumnValue infoType);

    /// <summary>
    /// Returns the address of a cell as text.
    /// </summary>
    ColumnValue Address(ColumnValue rowNum, ColumnValue columnNum, ColumnValue absNum, ColumnValue a1, ColumnValue sheetText);

    /// <summary>
    /// Returns the address of a cell as text (simplified).
    /// </summary>
    ColumnValue Address(ColumnValue rowNum, ColumnValue columnNum);
}