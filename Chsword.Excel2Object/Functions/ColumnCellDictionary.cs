namespace Chsword.Excel2Object.Functions;

/// <summary>
/// Represents a dictionary of column cells.
/// </summary>
public class ColumnCellDictionary : Dictionary<string, ColumnValue>
{
    /// <summary>
    /// Gets the <see cref="ColumnValue"/> at the specified column name and row number.
    /// </summary>
    /// <param name="columnName">The name of the column.</param>
    /// <param name="rowNumber">The number of the row.</param>
    /// <returns>The <see cref="ColumnValue"/> at the specified column name and row number.</returns>
    public ColumnValue this[string columnName, int rowNumber] => throw new NotImplementedException();

    /// <summary>
    /// Gets a matrix of column values between the specified keys and row numbers.
    /// </summary>
    /// <param name="keyA">The first key.</param>
    /// <param name="rowA">The row number for the first key.</param>
    /// <param name="keyB">The second key.</param>
    /// <param name="rowB">The row number for the second key.</param>
    /// <returns>A <see cref="ColumnMatrix"/> representing the matrix of column values.</returns>
    public ColumnMatrix Matrix(string keyA, int rowA, string keyB, int rowB)
    {
        throw new NotImplementedException();
    }

    /// <summary>
    /// Gets a whole-column range between two columns, e.g. <c>A:F</c>.
    /// </summary>
    /// <param name="keyA">The first column title.</param>
    /// <param name="keyB">The last column title.</param>
    public ColumnMatrix Columns(string keyA, string keyB)
    {
        throw new NotImplementedException();
    }

    /// <summary>
    /// Refers to another sheet of the same workbook so its cells can be used in this formula.
    /// The sheet must already exist in the workbook when the formula is written.
    /// </summary>
    /// <param name="sheetTitle">The title of the other sheet.</param>
    public SheetCellDictionary Sheet(string sheetTitle)
    {
        throw new NotImplementedException();
    }
}