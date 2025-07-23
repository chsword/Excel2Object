namespace Chsword.Excel2Object.Functions;

/// <summary>
/// Interface for conditional functions used in Excel to object conversion.
/// </summary>
public interface IConditionFunction
{
    /// <summary>
    /// Returns one value if a condition is true and another value if it's false.
    /// </summary>
    ColumnValue If(ColumnValue condition, ColumnValue valueIfTrue, ColumnValue valueIfFalse);

    /// <summary>
    /// Returns a value you specify if a formula evaluates to an error; otherwise, returns the result of the formula.
    /// </summary>
    ColumnValue IfError(ColumnValue value, ColumnValue valueIfError);

    /// <summary>
    /// Returns a value you specify if the expression resolves to #N/A; otherwise returns the result of the expression.
    /// </summary>
    ColumnValue IfNa(ColumnValue value, ColumnValue valueIfNa);

    /// <summary>
    /// Checks whether one or more conditions are met and returns a value that corresponds to the first TRUE condition.
    /// </summary>
    ColumnValue Ifs(params ColumnValue[] conditionsAndValues);

    /// <summary>
    /// Returns TRUE if any argument is TRUE.
    /// </summary>
    ColumnValue Or(params ColumnValue[] logicalValues);

    /// <summary>
    /// Returns TRUE if all arguments are TRUE.
    /// </summary>
    ColumnValue And(params ColumnValue[] logicalValues);

    /// <summary>
    /// Reverses the logic of its argument.
    /// </summary>
    ColumnValue Not(ColumnValue logical);

    /// <summary>
    /// Checks whether a value is an error.
    /// </summary>
    ColumnValue IsError(ColumnValue value);

    /// <summary>
    /// Checks whether a value is the #N/A error.
    /// </summary>
    ColumnValue IsNa(ColumnValue value);

    /// <summary>
    /// Checks whether a value is a number.
    /// </summary>
    ColumnValue IsNumber(ColumnValue value);

    /// <summary>
    /// Checks whether a value is text.
    /// </summary>
    ColumnValue IsText(ColumnValue value);

    /// <summary>
    /// Checks whether a value is blank.
    /// </summary>
    ColumnValue IsBlank(ColumnValue value);

    /// <summary>
    /// Checks whether a value is a logical value.
    /// </summary>
    ColumnValue IsLogical(ColumnValue value);

    /// <summary>
    /// Checks whether a reference is to a cell containing a formula.
    /// </summary>
    ColumnValue IsFormula(ColumnValue reference);

    /// <summary>
    /// Checks whether a value is even.
    /// </summary>
    ColumnValue IsEven(ColumnValue number);

    /// <summary>
    /// Checks whether a value is odd.
    /// </summary>
    ColumnValue IsOdd(ColumnValue number);
}