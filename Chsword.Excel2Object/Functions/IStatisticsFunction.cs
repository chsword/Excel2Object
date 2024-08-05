namespace Chsword.Excel2Object.Functions;

/// <summary>
/// Interface for statistical functions used in Excel to object conversion.
/// </summary>
public interface IStatisticsFunction
{
    /// <summary>
    /// Calculates the sum of the specified matrices.
    /// </summary>
    /// <param name="matrix">The matrices to sum.</param>
    /// <returns>A <see cref="ColumnValue"/> representing the sum of the matrices.</returns>
    ColumnValue Sum(params ColumnMatrix[] matrix);
}