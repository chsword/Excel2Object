namespace Chsword.Excel2Object.Functions;

/// <summary>
///     Excel math and trigonometry functions. See <see cref="IExcelFunction" /> for how these are translated.
/// </summary>
public interface IMathFunction : IExcelFunction
{
    // ---- rounding and sign ----

    /// <summary>ABS - absolute value.</summary>
    ColumnValue Abs(ColumnValue val);

    /// <summary>SIGN - 1, 0 or -1 according to the sign of the number.</summary>
    ColumnValue Sign(ColumnValue val);

    /// <summary>EVEN - rounds away from zero to the nearest even integer.</summary>
    ColumnValue Even(ColumnValue val);

    /// <summary>ODD - rounds away from zero to the nearest odd integer.</summary>
    ColumnValue Odd(ColumnValue val);

    /// <summary>INT - rounds down to the nearest integer.</summary>
    ColumnValue Int(ColumnValue val);

    /// <summary>TRUNC - truncates to an integer.</summary>
    ColumnValue Trunc(ColumnValue val);

    /// <summary>TRUNC - truncates to the given number of digits.</summary>
    ColumnValue Trunc(ColumnValue val, ColumnValue digits);

    /// <summary>ROUND - rounds to the given number of digits.</summary>
    ColumnValue Round(ColumnValue val, ColumnValue digits);

    /// <summary>ROUNDDOWN - rounds toward zero.</summary>
    ColumnValue RoundDown(ColumnValue val, ColumnValue digits);

    /// <summary>ROUNDUP - rounds away from zero.</summary>
    ColumnValue RoundUp(ColumnValue val, ColumnValue digits);

    /// <summary>MROUND - rounds to the nearest multiple of <paramref name="multiple" />.</summary>
    ColumnValue MRound(ColumnValue val, ColumnValue multiple);

    /// <summary>CEILING - rounds up to the nearest multiple of <paramref name="significance" />.</summary>
    ColumnValue Ceiling(ColumnValue val, ColumnValue significance);

    /// <summary>CEILING.MATH - rounds up to the nearest integer or multiple of significance.</summary>
    [ExcelFunctionName("CEILING.MATH", Future = true)]
    ColumnValue CeilingMath(ColumnValue val);

    /// <summary>CEILING.MATH - rounds up to the nearest multiple of significance.</summary>
    [ExcelFunctionName("CEILING.MATH", Future = true)]
    ColumnValue CeilingMath(ColumnValue val, ColumnValue significance);

    /// <summary>CEILING.MATH - rounds up, with control over how negative numbers are handled.</summary>
    [ExcelFunctionName("CEILING.MATH", Future = true)]
    ColumnValue CeilingMath(ColumnValue val, ColumnValue significance, ColumnValue mode);

    /// <summary>FLOOR - rounds down to the nearest multiple of <paramref name="significance" />.</summary>
    ColumnValue Floor(ColumnValue val, ColumnValue significance);

    /// <summary>FLOOR.MATH - rounds down to the nearest integer or multiple of significance.</summary>
    [ExcelFunctionName("FLOOR.MATH", Future = true)]
    ColumnValue FloorMath(ColumnValue val);

    /// <summary>FLOOR.MATH - rounds down to the nearest multiple of significance.</summary>
    [ExcelFunctionName("FLOOR.MATH", Future = true)]
    ColumnValue FloorMath(ColumnValue val, ColumnValue significance);

    /// <summary>FLOOR.MATH - rounds down, with control over how negative numbers are handled.</summary>
    [ExcelFunctionName("FLOOR.MATH", Future = true)]
    ColumnValue FloorMath(ColumnValue val, ColumnValue significance, ColumnValue mode);

    // ---- arithmetic ----

    /// <summary>MOD - remainder of a division.</summary>
    ColumnValue Mod(ColumnValue number, ColumnValue divisor);

    /// <summary>QUOTIENT - integer part of a division.</summary>
    ColumnValue Quotient(ColumnValue numerator, ColumnValue denominator);

    /// <summary>POWER - raises a number to a power.</summary>
    ColumnValue Power(ColumnValue number, ColumnValue power);

    /// <summary>EXP - e raised to a power.</summary>
    ColumnValue Exp(ColumnValue val);

    /// <summary>LN - natural logarithm.</summary>
    ColumnValue Ln(ColumnValue val);

    /// <summary>LOG - base 10 logarithm.</summary>
    ColumnValue Log(ColumnValue val);

    /// <summary>LOG - logarithm to the given base.</summary>
    ColumnValue Log(ColumnValue val, ColumnValue baseValue);

    /// <summary>LOG10 - base 10 logarithm.</summary>
    ColumnValue Log10(ColumnValue val);

    /// <summary>SQRT - square root.</summary>
    ColumnValue Sqrt(ColumnValue val);

    /// <summary>SQRTPI - square root of (number * pi).</summary>
    [ExcelFunctionName("SQRTPI")]
    ColumnValue SqrtPi(ColumnValue val);

    /// <summary>PI - the constant pi.</summary>
    [ExcelFunctionName("PI")]
    ColumnValue PI();

    /// <summary>GCD - greatest common divisor.</summary>
    ColumnValue Gcd(params ColumnValue[] values);

    /// <summary>LCM - least common multiple.</summary>
    ColumnValue Lcm(params ColumnValue[] values);

    /// <summary>FACT - factorial.</summary>
    ColumnValue Fact(ColumnValue val);

    /// <summary>FACTDOUBLE - double factorial.</summary>
    ColumnValue FactDouble(ColumnValue val);

    /// <summary>MULTINOMIAL - multinomial coefficient of a set of numbers.</summary>
    ColumnValue MultiNomial(params ColumnValue[] values);

    /// <summary>COMBIN - number of combinations without repetition.</summary>
    ColumnValue Combin(ColumnValue number, ColumnValue numberChosen);

    /// <summary>COMBINA - number of combinations with repetition.</summary>
    [ExcelFunctionName("COMBINA", Future = true)]
    ColumnValue CombinA(ColumnValue number, ColumnValue numberChosen);

    // ---- aggregation ----

    /// <summary>PRODUCT - multiplies its arguments.</summary>
    ColumnValue Product(params ColumnValue[] values);

    /// <summary>SUMIF - sums the cells of a range that meet a criteria.</summary>
    ColumnValue SumIf(ColumnMatrix range, ColumnValue criteria);

    /// <summary>SUMIF - sums <paramref name="sumRange" /> where <paramref name="range" /> meets a criteria.</summary>
    ColumnValue SumIf(ColumnMatrix range, ColumnValue criteria, ColumnMatrix sumRange);

    /// <summary>SUMIFS - sums the cells that meet several criteria; pass criteria as range, criteria pairs.</summary>
    ColumnValue SumIfs(ColumnMatrix sumRange, params ColumnValue[] criteria);

    /// <summary>SUMPRODUCT - sum of the products of corresponding cells.</summary>
    ColumnValue SumProduct(params ColumnValue[] values);

    /// <summary>SUMSQ - sum of the squares of its arguments.</summary>
    [ExcelFunctionName("SUMSQ")]
    ColumnValue SumSq(params ColumnValue[] values);

    /// <summary>SERIESSUM - sum of a power series.</summary>
    [ExcelFunctionName("SERIESSUM")]
    ColumnValue SeriesSum(ColumnValue x, ColumnValue n, ColumnValue m, ColumnMatrix coefficients);

    /// <summary>SUBTOTAL - subtotal of a list, by function number.</summary>
    ColumnValue Subtotal(ColumnValue functionNum, params ColumnValue[] values);

    /// <summary>AGGREGATE - aggregate of a list, with control over which values to ignore.</summary>
    [ExcelFunctionName("AGGREGATE", Future = true)]
    ColumnValue Aggregate(ColumnValue functionNum, ColumnValue options, params ColumnValue[] values);

    // ---- random ----

    /// <summary>RAND - a random number in [0, 1).</summary>
    ColumnValue Rand();

    /// <summary>RANDBETWEEN - a random integer between two values.</summary>
    ColumnValue RandBetween(ColumnValue bottom, ColumnValue top);

    /// <summary>RANDARRAY - an array of random numbers.</summary>
    [ExcelFunctionName("RANDARRAY", Future = true)]
    ColumnValue RandArray(ColumnValue rows, ColumnValue columns);

    // ---- number bases ----

    /// <summary>BASE - converts a number to text in the given radix.</summary>
    [ExcelFunctionName("BASE", Future = true)]
    ColumnValue Base(ColumnValue number, ColumnValue radix);

    /// <summary>BASE - converts a number to text in the given radix, padded to a minimum length.</summary>
    [ExcelFunctionName("BASE", Future = true)]
    ColumnValue Base(ColumnValue number, ColumnValue radix, ColumnValue minLength);

    /// <summary>DECIMAL - converts text in the given radix to a number.</summary>
    [ExcelFunctionName("DECIMAL", Future = true)]
    ColumnValue Decimal(ColumnValue text, ColumnValue radix);

    /// <summary>ROMAN - converts a number to roman numerals.</summary>
    ColumnValue Roman(ColumnValue number);

    /// <summary>ARABIC - converts roman numerals to a number.</summary>
    [ExcelFunctionName("ARABIC", Future = true)]
    ColumnValue Arabic(ColumnValue text);

    // ---- matrices ----

    /// <summary>MDETERM - determinant of a matrix.</summary>
    [ExcelFunctionName("MDETERM")]
    ColumnValue MDeterm(ColumnMatrix array);

    /// <summary>MINVERSE - inverse of a matrix.</summary>
    [ExcelFunctionName("MINVERSE")]
    ColumnValue MInverse(ColumnMatrix array);

    /// <summary>MMULT - matrix product of two arrays.</summary>
    [ExcelFunctionName("MMULT")]
    ColumnValue MMult(ColumnMatrix array1, ColumnMatrix array2);

    /// <summary>MUNIT - the identity matrix of the given dimension.</summary>
    [ExcelFunctionName("MUNIT", Future = true)]
    ColumnValue MUnit(ColumnValue dimension);

    // ---- trigonometry ----

    /// <summary>SIN - sine.</summary>
    ColumnValue Sin(ColumnValue val);

    /// <summary>COS - cosine.</summary>
    ColumnValue Cos(ColumnValue val);

    /// <summary>TAN - tangent.</summary>
    ColumnValue Tan(ColumnValue val);

    /// <summary>COT - cotangent.</summary>
    [ExcelFunctionName("COT", Future = true)]
    ColumnValue Cot(ColumnValue val);

    /// <summary>SEC - secant.</summary>
    [ExcelFunctionName("SEC", Future = true)]
    ColumnValue Sec(ColumnValue val);

    /// <summary>CSC - cosecant.</summary>
    [ExcelFunctionName("CSC", Future = true)]
    ColumnValue Csc(ColumnValue val);

    /// <summary>ASIN - arcsine.</summary>
    [ExcelFunctionName("ASIN")]
    ColumnValue Asin(ColumnValue val);

    /// <summary>ACOS - arccosine.</summary>
    [ExcelFunctionName("ACOS")]
    ColumnValue Acos(ColumnValue val);

    /// <summary>ATAN - arctangent.</summary>
    [ExcelFunctionName("ATAN")]
    ColumnValue Atan(ColumnValue val);

    /// <summary>ATAN2 - arctangent of the given x and y coordinates.</summary>
    [ExcelFunctionName("ATAN2")]
    ColumnValue Atan2(ColumnValue x, ColumnValue y);

    /// <summary>ACOT - arccotangent.</summary>
    [ExcelFunctionName("ACOT", Future = true)]
    ColumnValue Acot(ColumnValue val);

    /// <summary>SINH - hyperbolic sine.</summary>
    ColumnValue Sinh(ColumnValue val);

    /// <summary>COSH - hyperbolic cosine.</summary>
    ColumnValue Cosh(ColumnValue val);

    /// <summary>TANH - hyperbolic tangent.</summary>
    ColumnValue Tanh(ColumnValue val);

    /// <summary>COTH - hyperbolic cotangent.</summary>
    [ExcelFunctionName("COTH", Future = true)]
    ColumnValue Coth(ColumnValue val);

    /// <summary>SECH - hyperbolic secant.</summary>
    [ExcelFunctionName("SECH", Future = true)]
    ColumnValue Sech(ColumnValue val);

    /// <summary>CSCH - hyperbolic cosecant.</summary>
    [ExcelFunctionName("CSCH", Future = true)]
    ColumnValue Csch(ColumnValue val);

    /// <summary>ASINH - inverse hyperbolic sine.</summary>
    [ExcelFunctionName("ASINH")]
    ColumnValue Asinh(ColumnValue val);

    /// <summary>ACOSH - inverse hyperbolic cosine.</summary>
    [ExcelFunctionName("ACOSH")]
    ColumnValue Acosh(ColumnValue val);

    /// <summary>ATANH - inverse hyperbolic tangent.</summary>
    [ExcelFunctionName("ATANH")]
    ColumnValue Atanh(ColumnValue val);

    /// <summary>ACOTH - inverse hyperbolic cotangent.</summary>
    [ExcelFunctionName("ACOTH", Future = true)]
    ColumnValue Acoth(ColumnValue val);

    /// <summary>DEGREES - converts radians to degrees.</summary>
    ColumnValue Degrees(ColumnValue angle);

    /// <summary>RADIANS - converts degrees to radians.</summary>
    ColumnValue Radians(ColumnValue angle);
}
