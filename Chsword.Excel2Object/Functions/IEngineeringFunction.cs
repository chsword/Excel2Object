namespace Chsword.Excel2Object.Functions;

/// <summary>
///     Excel engineering functions. See <see cref="IExcelFunction" /> for how these are translated.
/// </summary>
public interface IEngineeringFunction : IExcelFunction
{
    /// <summary>CONVERT - converts a number from one unit of measurement to another.</summary>
    ColumnValue Convert(ColumnValue number, ColumnValue fromUnit, ColumnValue toUnit);

    /// <summary>DEC2BIN - converts a decimal number to binary.</summary>
    [ExcelFunctionName("DEC2BIN")]
    ColumnValue Dec2Bin(ColumnValue number);

    /// <summary>DEC2OCT - converts a decimal number to octal.</summary>
    [ExcelFunctionName("DEC2OCT")]
    ColumnValue Dec2Oct(ColumnValue number);

    /// <summary>DEC2HEX - converts a decimal number to hexadecimal.</summary>
    [ExcelFunctionName("DEC2HEX")]
    ColumnValue Dec2Hex(ColumnValue number);

    /// <summary>BIN2DEC - converts a binary number to decimal.</summary>
    [ExcelFunctionName("BIN2DEC")]
    ColumnValue Bin2Dec(ColumnValue number);

    /// <summary>BIN2OCT - converts a binary number to octal.</summary>
    [ExcelFunctionName("BIN2OCT")]
    ColumnValue Bin2Oct(ColumnValue number);

    /// <summary>BIN2HEX - converts a binary number to hexadecimal.</summary>
    [ExcelFunctionName("BIN2HEX")]
    ColumnValue Bin2Hex(ColumnValue number);

    /// <summary>OCT2DEC - converts an octal number to decimal.</summary>
    [ExcelFunctionName("OCT2DEC")]
    ColumnValue Oct2Dec(ColumnValue number);

    /// <summary>OCT2BIN - converts an octal number to binary.</summary>
    [ExcelFunctionName("OCT2BIN")]
    ColumnValue Oct2Bin(ColumnValue number);

    /// <summary>OCT2HEX - converts an octal number to hexadecimal.</summary>
    [ExcelFunctionName("OCT2HEX")]
    ColumnValue Oct2Hex(ColumnValue number);

    /// <summary>HEX2DEC - converts a hexadecimal number to decimal.</summary>
    [ExcelFunctionName("HEX2DEC")]
    ColumnValue Hex2Dec(ColumnValue number);

    /// <summary>HEX2BIN - converts a hexadecimal number to binary.</summary>
    [ExcelFunctionName("HEX2BIN")]
    ColumnValue Hex2Bin(ColumnValue number);

    /// <summary>HEX2OCT - converts a hexadecimal number to octal.</summary>
    [ExcelFunctionName("HEX2OCT")]
    ColumnValue Hex2Oct(ColumnValue number);

    /// <summary>BITAND - the bitwise AND of two numbers.</summary>
    [ExcelFunctionName("BITAND", Future = true)]
    ColumnValue BitAnd(ColumnValue number1, ColumnValue number2);

    /// <summary>BITOR - the bitwise OR of two numbers.</summary>
    [ExcelFunctionName("BITOR", Future = true)]
    ColumnValue BitOr(ColumnValue number1, ColumnValue number2);

    /// <summary>BITXOR - the bitwise XOR of two numbers.</summary>
    [ExcelFunctionName("BITXOR", Future = true)]
    ColumnValue BitXor(ColumnValue number1, ColumnValue number2);

    /// <summary>BITLSHIFT - a number shifted left by the given number of bits.</summary>
    [ExcelFunctionName("BITLSHIFT", Future = true)]
    ColumnValue BitLShift(ColumnValue number, ColumnValue shiftAmount);

    /// <summary>BITRSHIFT - a number shifted right by the given number of bits.</summary>
    [ExcelFunctionName("BITRSHIFT", Future = true)]
    ColumnValue BitRShift(ColumnValue number, ColumnValue shiftAmount);

    /// <summary>DELTA - 1 when two numbers are equal, 0 otherwise.</summary>
    ColumnValue Delta(ColumnValue number1, ColumnValue number2);

    /// <summary>GESTEP - 1 when a number is at least the step value, 0 otherwise.</summary>
    [ExcelFunctionName("GESTEP")]
    ColumnValue GeStep(ColumnValue number, ColumnValue step);
}
