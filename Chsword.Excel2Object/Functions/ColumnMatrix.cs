namespace Chsword.Excel2Object.Functions;

/// <summary>
/// Represents a matrix of column values.
/// </summary>
public class ColumnMatrix
{
    public string StartColumn { get; set; }
    public int StartRow { get; set; }
    public string EndColumn { get; set; }
    public int EndRow { get; set; }

    public ColumnMatrix(string startColumn, int startRow, string endColumn, int endRow)
    {
        StartColumn = startColumn;
        StartRow = startRow;
        EndColumn = endColumn;
        EndRow = endRow;
    }

    public override string ToString()
    {
        if (StartColumn == EndColumn && StartRow == EndRow)
        {
            return $"{StartColumn}{StartRow}";
        }
        return $"{StartColumn}{StartRow}:{EndColumn}{EndRow}";
    }
}