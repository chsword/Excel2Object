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
        return new ColumnMatrix(keyA, rowA, keyB, rowB);
    }
}