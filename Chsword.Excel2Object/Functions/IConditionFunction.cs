namespace Chsword.Excel2Object.Functions;

/// <summary>
///     Excel logical functions. See <see cref="IExcelFunction" /> for how these are translated.
/// </summary>
/// <remarks>
///     C# operators are translated too, so the conditions themselves are written naturally:
///     <c>m.Qty &gt; 5</c>, <c>a &amp;&amp; b</c> (AND), <c>a || b</c> (OR), <c>!a</c> (NOT) and the
///     conditional operator <c>cond ? x : y</c> (IF).
/// </remarks>
public interface IConditionFunction : IExcelFunction
{
    /// <summary>IF - one value when the condition holds, another when it does not.</summary>
    ColumnValue If(ColumnValue condition, ColumnValue value1, ColumnValue value2);

    /// <summary>IF - a value when the condition holds.</summary>
    ColumnValue If(ColumnValue condition, ColumnValue value1);

    /// <summary>IFS - the value of the first condition that holds; pass condition, value pairs.</summary>
    [ExcelFunctionName("IFS", Future = true)]
    ColumnValue Ifs(params ColumnValue[] conditionsAndValues);

    /// <summary>SWITCH - matches a value against a list of results; pass match, result pairs.</summary>
    [ExcelFunctionName("SWITCH", Future = true)]
    ColumnValue Switch(ColumnValue expression, params ColumnValue[] valuesAndResults);

    /// <summary>IFERROR - a fallback value when the expression is an error.</summary>
    [ExcelFunctionName("IFERROR")]
    ColumnValue IfError(ColumnValue val, ColumnValue valueIfError);

    /// <summary>IFNA - a fallback value when the expression is #N/A.</summary>
    [ExcelFunctionName("IFNA", Future = true)]
    ColumnValue IfNa(ColumnValue val, ColumnValue valueIfNa);

    /// <summary>AND - TRUE when every argument is true.</summary>
    ColumnValue And(params ColumnValue[] conditions);

    /// <summary>OR - TRUE when any argument is true.</summary>
    ColumnValue Or(params ColumnValue[] conditions);

    /// <summary>XOR - TRUE when an odd number of arguments are true.</summary>
    [ExcelFunctionName("XOR", Future = true)]
    ColumnValue Xor(params ColumnValue[] conditions);

    /// <summary>NOT - reverses the logic of its argument.</summary>
    ColumnValue Not(ColumnValue condition);

    /// <summary>TRUE - the logical value TRUE.</summary>
    [ExcelFunctionName("TRUE")]
    ColumnValue True();

    /// <summary>FALSE - the logical value FALSE.</summary>
    [ExcelFunctionName("FALSE")]
    ColumnValue False();
}
