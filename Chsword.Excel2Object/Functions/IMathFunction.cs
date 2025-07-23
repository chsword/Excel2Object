namespace Chsword.Excel2Object.Functions;

public interface IMathFunction
{
    /// <summary>
    /// Returns the absolute value of a number
    /// </summary>
    ColumnValue Abs(ColumnValue val);

    /// <summary>
    /// Returns number rounded up to the nearest even integer
    /// </summary>
    ColumnValue Even(ColumnValue val);

    /// <summary>
    /// Returns the factorial of a number
    /// </summary>
    ColumnValue Fact(ColumnValue val);

    /// <summary>
    /// Rounds a number down to the nearest integer
    /// </summary>
    ColumnValue Int(ColumnValue val);

    /// <summary>
    /// Returns the value of π (pi)
    /// </summary>
    ColumnValue PI();

    /// <summary>
    /// Returns a random number between 0 and 1
    /// </summary>
    ColumnValue Rand();

    /// <summary>
    /// Rounds a number to a specified number of digits
    /// </summary>
    ColumnValue Round(ColumnValue val, ColumnValue digits);

    /// <summary>
    /// Rounds a number down to a specified number of digits
    /// </summary>
    ColumnValue RoundDown(ColumnValue val, ColumnValue digits);

    /// <summary>
    /// Rounds a number up to a specified number of digits
    /// </summary>
    ColumnValue RoundUp(ColumnValue val, ColumnValue digits);

    /// <summary>
    /// Returns the square root of a number
    /// </summary>
    ColumnValue Sqrt(ColumnValue val);

    /// <summary>
    /// Returns number rounded up to the nearest odd integer
    /// </summary>
    ColumnValue Odd(ColumnValue val);

    /// <summary>
    /// Returns a number raised to a power
    /// </summary>
    ColumnValue Power(ColumnValue number, ColumnValue power);

    /// <summary>
    /// Returns the exponential value of e raised to the power of number
    /// </summary>
    ColumnValue Exp(ColumnValue number);

    /// <summary>
    /// Returns the natural logarithm of a number
    /// </summary>
    ColumnValue Ln(ColumnValue number);

    /// <summary>
    /// Returns the logarithm of a number to a specified base
    /// </summary>
    ColumnValue Log(ColumnValue number, ColumnValue logBase);

    /// <summary>
    /// Returns the base-10 logarithm of a number
    /// </summary>
    ColumnValue Log10(ColumnValue number);

    /// <summary>
    /// Returns the sine of an angle
    /// </summary>
    ColumnValue Sin(ColumnValue number);

    /// <summary>
    /// Returns the cosine of an angle
    /// </summary>
    ColumnValue Cos(ColumnValue number);

    /// <summary>
    /// Returns the tangent of an angle
    /// </summary>
    ColumnValue Tan(ColumnValue number);

    /// <summary>
    /// Returns the greatest common divisor
    /// </summary>
    ColumnValue Gcd(ColumnValue number1, ColumnValue number2);

    /// <summary>
    /// Returns the least common multiple
    /// </summary>
    ColumnValue Lcm(ColumnValue number1, ColumnValue number2);

    /// <summary>
    /// Returns the remainder from division
    /// </summary>
    ColumnValue Mod(ColumnValue number, ColumnValue divisor);

    /// <summary>
    /// Returns a random integer between the numbers you specify
    /// </summary>
    ColumnValue RandBetween(ColumnValue bottom, ColumnValue top);

    /// <summary>
    /// Returns the sign of a number
    /// </summary>
    ColumnValue Sign(ColumnValue number);

    /// <summary>
    /// Truncates a number to an integer
    /// </summary>
    ColumnValue Trunc(ColumnValue number, ColumnValue numDigits);
}