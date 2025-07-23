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

    /// <summary>
    /// Calculates the average of the specified matrices.
    /// </summary>
    ColumnValue Average(params ColumnMatrix[] matrix);

    /// <summary>
    /// Counts the number of cells that contain numbers.
    /// </summary>
    ColumnValue Count(params ColumnMatrix[] matrix);

    /// <summary>
    /// Counts the number of non-empty cells.
    /// </summary>
    ColumnValue CountA(params ColumnMatrix[] matrix);

    /// <summary>
    /// Counts the number of empty cells.
    /// </summary>
    ColumnValue CountBlank(ColumnMatrix matrix);

    /// <summary>
    /// Counts the number of cells that meet a criteria.
    /// </summary>
    ColumnValue CountIf(ColumnMatrix range, ColumnValue criteria);

    /// <summary>
    /// Counts the number of cells that meet multiple criteria.
    /// </summary>
    ColumnValue CountIfs(ColumnMatrix criteriaRange1, ColumnValue criteria1, ColumnMatrix criteriaRange2, ColumnValue criteria2);

    /// <summary>
    /// Returns the largest value among the specified matrices.
    /// </summary>
    ColumnValue Max(params ColumnMatrix[] matrix);

    /// <summary>
    /// Returns the smallest value among the specified matrices.
    /// </summary>
    ColumnValue Min(params ColumnMatrix[] matrix);

    /// <summary>
    /// Returns the median of the given numbers.
    /// </summary>
    ColumnValue Median(params ColumnMatrix[] matrix);

    /// <summary>
    /// Returns the most frequently occurring value.
    /// </summary>
    ColumnValue Mode(params ColumnMatrix[] matrix);

    /// <summary>
    /// Calculates the standard deviation of the specified matrices.
    /// </summary>
    ColumnValue StDev(params ColumnMatrix[] matrix);

    /// <summary>
    /// Calculates the variance of the specified matrices.
    /// </summary>
    ColumnValue Var(params ColumnMatrix[] matrix);

    /// <summary>
    /// Returns the k-th largest value in a dataset.
    /// </summary>
    ColumnValue Large(ColumnMatrix array, ColumnValue k);

    /// <summary>
    /// Returns the k-th smallest value in a dataset.
    /// </summary>
    ColumnValue Small(ColumnMatrix array, ColumnValue k);

    /// <summary>
    /// Returns the rank of a number in a list of numbers.
    /// </summary>
    ColumnValue Rank(ColumnValue number, ColumnMatrix array, ColumnValue order);

    /// <summary>
    /// Returns the percentile of values in a range.
    /// </summary>
    ColumnValue Percentile(ColumnMatrix array, ColumnValue k);

    /// <summary>
    /// Returns the quartile of a dataset.
    /// </summary>
    ColumnValue Quartile(ColumnMatrix array, ColumnValue quart);

    /// <summary>
    /// Sums the values of cells that meet a criteria.
    /// </summary>
    ColumnValue SumIf(ColumnMatrix range, ColumnValue criteria, ColumnMatrix sumRange);

    /// <summary>
    /// Sums the values of cells that meet multiple criteria.
    /// </summary>
    ColumnValue SumIfs(ColumnMatrix sumRange, ColumnMatrix criteriaRange1, ColumnValue criteria1);

    /// <summary>
    /// Averages the values of cells that meet a criteria.
    /// </summary>
    ColumnValue AverageIf(ColumnMatrix range, ColumnValue criteria, ColumnMatrix averageRange);

    /// <summary>
    /// Averages the values of cells that meet multiple criteria.
    /// </summary>
    ColumnValue AverageIfs(ColumnMatrix averageRange, ColumnMatrix criteriaRange1, ColumnValue criteria1);
}