using System;

namespace Chsword.Excel2Object.Functions.Implementations;

/// <summary>
/// Implementation of mathematical functions for Excel formulas
/// </summary>
public class MathFunctions : IMathFunction
{
    /// <summary>
    /// Returns the absolute value of a number
    /// </summary>
    public int Abs(object val)
    {
        if (val == null) return 0;
        
        return val switch
        {
            int i => Math.Abs(i),
            double d => (int)Math.Abs(d),
            float f => (int)Math.Abs(f),
            decimal dec => (int)Math.Abs(dec),
            string str when int.TryParse(str, out var result) => Math.Abs(result),
            string str when double.TryParse(str, out var result) => (int)Math.Abs(result),
            _ => throw new ArgumentException($"Cannot convert {val} to a numeric value for ABS function")
        };
    }

    /// <summary>
    /// Returns a number rounded up to the nearest even integer
    /// </summary>
    public int Even(object val)
    {
        var num = ConvertToDouble(val);
        var result = Math.Ceiling(Math.Abs(num) / 2) * 2;
        return num >= 0 ? (int)result : -(int)result;
    }

    /// <summary>
    /// Returns the factorial of a number
    /// </summary>
    public int Fact(object val)
    {
        var num = ConvertToInt(val);
        if (num < 0) throw new ArgumentException("FACT function requires a non-negative integer");
        if (num > 170) throw new ArgumentException("FACT function input too large");
        
        long result = 1;
        for (int i = 2; i <= num; i++)
        {
            result *= i;
            if (result > int.MaxValue) throw new OverflowException("Factorial result too large");
        }
        return (int)result;
    }

    /// <summary>
    /// Rounds a number down to the nearest integer
    /// </summary>
    public int Int(object val)
    {
        var num = ConvertToDouble(val);
        return (int)Math.Floor(num);
    }

    /// <summary>
    /// Returns the value of π (pi)
    /// </summary>
    public double PI()
    {
        return Math.PI;
    }

    /// <summary>
    /// Returns a random number between 0 and 1
    /// </summary>
    public double Rand()
    {
        return Random.Shared.NextDouble();
    }

    /// <summary>
    /// Rounds a number to a specified number of digits
    /// </summary>
    public int Round(object val, object digits)
    {
        var num = ConvertToDouble(val);
        var digitCount = ConvertToInt(digits);
        return (int)Math.Round(num, digitCount);
    }

    /// <summary>
    /// Rounds a number down to a specified number of digits
    /// </summary>
    public int RoundDown(object val, object digits)
    {
        var num = ConvertToDouble(val);
        var digitCount = ConvertToInt(digits);
        var multiplier = Math.Pow(10, digitCount);
        return (int)(Math.Floor(num * multiplier) / multiplier);
    }

    /// <summary>
    /// Rounds a number up to a specified number of digits
    /// </summary>
    public int RoundUp(object val, object digits)
    {
        var num = ConvertToDouble(val);
        var digitCount = ConvertToInt(digits);
        var multiplier = Math.Pow(10, digitCount);
        return (int)(Math.Ceiling(num * multiplier) / multiplier);
    }

    /// <summary>
    /// Returns the square root of a number
    /// </summary>
    public int Sqrt(object val)
    {
        var num = ConvertToDouble(val);
        if (num < 0) throw new ArgumentException("SQRT function requires a non-negative number");
        return (int)Math.Sqrt(num);
    }

    // Helper methods
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
            _ => throw new ArgumentException($"Cannot convert {val} to a numeric value")
        };
    }

    private static int ConvertToInt(object val)
    {
        if (val == null) return 0;
        
        return val switch
        {
            int i => i,
            double d => (int)d,
            float f => (int)f,
            decimal dec => (int)dec,
            string str when int.TryParse(str, out var result) => result,
            string str when double.TryParse(str, out var result) => (int)result,
            _ => throw new ArgumentException($"Cannot convert {val} to an integer value")
        };
    }
}
