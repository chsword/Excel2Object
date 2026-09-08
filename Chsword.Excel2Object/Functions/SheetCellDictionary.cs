namespace Chsword.Excel2Object.Functions;

/// <summary>
///     Refers to cells on another sheet of the same workbook from inside a formula expression,
///     e.g. <c>c.Sheet("Products").Columns("Name", "Price")</c> becomes <c>'Products'!A:B</c>.
///     Obtained from <see cref="ColumnCellDictionary.Sheet" />.
/// </summary>
/// <remarks>
///     Like <see cref="ColumnCellDictionary" />, this type is only ever captured in an expression tree
///     and translated to formula text; its members are never executed.
/// </remarks>
public class SheetCellDictionary
{
    /// <summary>
    ///     The cell in the named column on the same row as the formula cell.
    /// </summary>
    public ColumnValue this[string columnName] => throw new NotImplementedException();

    /// <summary>
    ///     The cell in the named column at the given 1-based row number.
    /// </summary>
    public ColumnValue this[string columnName, int rowNumber] => throw new NotImplementedException();

    /// <summary>
    ///     A rectangular range, e.g. <c>'Sheet'!A2:C20</c>.
    /// </summary>
    public ColumnMatrix Matrix(string keyA, int rowA, string keyB, int rowB)
    {
        throw new NotImplementedException();
    }

    /// <summary>
    ///     A whole-column range, e.g. <c>'Sheet'!A:F</c>.
    /// </summary>
    public ColumnMatrix Columns(string keyA, string keyB)
    {
        throw new NotImplementedException();
    }
}
