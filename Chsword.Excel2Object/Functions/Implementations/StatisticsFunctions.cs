using System;
using System.Linq;

namespace Chsword.Excel2Object.Functions.Implementations;

/// <summary>
/// Implementation of statistical functions for Excel formulas
/// </summary>
public class StatisticsFunctions : IStatisticsFunction
{
    /// <summary>
    /// Calculates the sum of the specified matrices
    /// </summary>
    public ColumnValue Sum(params ColumnMatrix[] matrix)
    {
        if (matrix == null || matrix.Length == 0)
            return new ColumnValue();

        double total = 0;
        
        foreach (var mat in matrix)
        {
            if (mat?.Values != null)
            {
                total += mat.Values.Sum(ConvertToDouble);
            }
        }

        return new ColumnValue { Value = total };
    }

    /// <summary>
    /// Calculates the average of the specified matrices
    /// </summary>
    public ColumnValue Average(params ColumnMatrix[] matrix)
    {
        if (matrix == null || matrix.Length == 0)
            return new ColumnValue();

        var values = matrix.Where(m => m?.Values != null)
                          .SelectMany(m => m.Values)
                          .Where(v => v != null)
                          .ToList();

        if (values.Count == 0) return new ColumnValue();

        double total = values.Sum(ConvertToDouble);
        return new ColumnValue { Value = total / values.Count };
    }

    /// <summary>
    /// Counts the number of cells that contain numbers
    /// </summary>
    public ColumnValue Count(params ColumnMatrix[] matrix)
    {
        if (matrix == null || matrix.Length == 0)
            return new ColumnValue();

        int count = matrix.Where(m => m?.Values != null)
                         .SelectMany(m => m.Values)
                         .Count(v => v != null && IsNumeric(v));

        return new ColumnValue { Value = count };
    }

    /// <summary>
    /// Counts the number of non-empty cells
    /// </summary>
    public ColumnValue CountA(params ColumnMatrix[] matrix)
    {
        if (matrix == null || matrix.Length == 0)
            return new ColumnValue();

        int count = matrix.Where(m => m?.Values != null)
                         .SelectMany(m => m.Values)
                         .Count(v => v != null && !string.IsNullOrEmpty(v.ToString()));

        return new ColumnValue { Value = count };
    }

    /// <summary>
    /// Returns the largest value among the specified matrices
    /// </summary>
    public ColumnValue Max(params ColumnMatrix[] matrix)
    {
        if (matrix == null || matrix.Length == 0)
            return new ColumnValue();

        var values = matrix.Where(m => m?.Values != null)
                          .SelectMany(m => m.Values)
                          .Where(v => v != null && IsNumeric(v))
                          .Select(ConvertToDouble)
                          .ToList();

        if (values.Count == 0) return new ColumnValue();

        return new ColumnValue { Value = values.Max() };
    }

    /// <summary>
    /// Returns the smallest value among the specified matrices
    /// </summary>
    public ColumnValue Min(params ColumnMatrix[] matrix)
    {
        if (matrix == null || matrix.Length == 0)
            return new ColumnValue();

        var values = matrix.Where(m => m?.Values != null)
                          .SelectMany(m => m.Values)
                          .Where(v => v != null && IsNumeric(v))
                          .Select(ConvertToDouble)
                          .ToList();

        if (values.Count == 0) return new ColumnValue();

        return new ColumnValue { Value = values.Min() };
    }

    /// <summary>
    /// Calculates the standard deviation of the specified matrices
    /// </summary>
    public ColumnValue StdDev(params ColumnMatrix[] matrix)
    {
        if (matrix == null || matrix.Length == 0)
            return new ColumnValue();

        var values = matrix.Where(m => m?.Values != null)
                          .SelectMany(m => m.Values)
                          .Where(v => v != null && IsNumeric(v))
                          .Select(ConvertToDouble)
                          .ToList();

        if (values.Count <= 1) return new ColumnValue();

        double avg = values.Average();
        double sumOfSquares = values.Sum(v => Math.Pow(v - avg, 2));
        double variance = sumOfSquares / (values.Count - 1);

        return new ColumnValue { Value = Math.Sqrt(variance) };
    }

    /// <summary>
    /// Calculates the variance of the specified matrices
    /// </summary>
    public ColumnValue Var(params ColumnMatrix[] matrix)
    {
        if (matrix == null || matrix.Length == 0)
            return new ColumnValue();

        var values = matrix.Where(m => m?.Values != null)
                          .SelectMany(m => m.Values)
                          .Where(v => v != null && IsNumeric(v))
                          .Select(ConvertToDouble)
                          .ToList();

        if (values.Count <= 1) return new ColumnValue();

        double avg = values.Average();
        double sumOfSquares = values.Sum(v => Math.Pow(v - avg, 2));

        return new ColumnValue { Value = sumOfSquares / (values.Count - 1) };
    }

    // Helper methods
    private static bool IsNumeric(object value)
    {
        return value switch
        {
            int or double or float or decimal => true,
            string str => double.TryParse(str, out _),
            _ => false
        };
    }

    private static double ConvertToDouble(object val)
    {
        if (val == null) return 0;
        
        return val switch
        {
            double d => d,
            int i => i,
            float f => f,
            decimal dec => (double)dec,
            string str when double.TryParse(str, out var result) => result,
            _ => 0
        };
    }
}
