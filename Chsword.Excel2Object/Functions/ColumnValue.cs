// ReSharper disable UnusedParameter.Global

using System.Runtime.CompilerServices;

namespace Chsword.Excel2Object.Functions;

/// <summary>
///     A value inside a formula expression: a cell, a literal or the result of a function.
/// </summary>
/// <remarks>
///     This type only ever appears in an expression tree, so none of its members run. The operators
///     exist to be captured and translated: <c>+ - * /</c> and the comparisons map to the same Excel
///     operators, <c>&amp;</c> is concatenation, <c>%</c> becomes <c>MOD</c>, <c>^</c> becomes Excel's
///     power operator (not a bitwise xor), and <c>!</c>, <c>&amp;&amp;</c> and <c>||</c> become
///     <c>NOT</c>, <c>AND</c> and <c>OR</c>.
/// </remarks>
public class ColumnValue
{
    public static ColumnValue operator +(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    public static ColumnValue operator &(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    public static ColumnValue operator /(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    public static ColumnValue operator ==(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    public static explicit operator DateTime(ColumnValue operand)
    {
        throw new NotImplementedException();
    }

    public static explicit operator int(ColumnValue operand)
    {
        throw new NotImplementedException();
    }

    public static ColumnValue operator >(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    public static ColumnValue operator >=(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    public static implicit operator ColumnValue(int operand)
    {
        throw new NotImplementedException();
    }

    public static implicit operator ColumnValue(float operand)
    {
        throw new NotImplementedException();
    }

    public static implicit operator ColumnValue(double operand)
    {
        throw new NotImplementedException();
    }

    public static implicit operator ColumnValue(string operand)
    {
        throw new NotImplementedException();
    }

    public static implicit operator ColumnValue(long operand)
    {
        throw new NotImplementedException();
    }

    public static implicit operator ColumnValue(decimal operand)
    {
        throw new NotImplementedException();
    }

    public static implicit operator ColumnValue(bool operand)
    {
        throw new NotImplementedException();
    }

    public static implicit operator ColumnValue(DateTime operand)
    {
        throw new NotImplementedException();
    }

    public static ColumnValue operator !=(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    public static ColumnValue operator <(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    public static ColumnValue operator <=(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    public static ColumnValue operator *(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    public static ColumnValue operator -(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    public static ColumnValue operator -(ColumnValue a)
    {
        throw new NotImplementedException();
    }

    /// <summary>Remainder of a division, written as <c>MOD(a,b)</c>.</summary>
    public static ColumnValue operator %(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    /// <summary>Raises to a power, written as Excel's <c>^</c> operator.</summary>
    public static ColumnValue operator ^(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    /// <summary>Reverses a condition, written as <c>NOT(x)</c>.</summary>
    public static ColumnValue operator !(ColumnValue a)
    {
        throw new NotImplementedException();
    }

    /// <summary>Either condition, written as <c>OR(a,b)</c>.</summary>
    public static ColumnValue operator |(ColumnValue a, ColumnValue b)
    {
        throw new NotImplementedException();
    }

    // && and || are only allowed on a type that defines operator true / operator false alongside
    // & and |, so these exist to let cell conditions be written as c["A"] > 1 && c["B"] < 2. Like
    // every other member here they are only ever captured in an expression tree, never run.
    public static bool operator true(ColumnValue a)
    {
        throw new NotImplementedException();
    }

    public static bool operator false(ColumnValue a)
    {
        throw new NotImplementedException();
    }

    /// <summary>A range used where a single value is expected, e.g. <c>SUM(A1:B2)</c>.</summary>
    public static implicit operator ColumnValue(ColumnMatrix operand)
    {
        throw new NotImplementedException();
    }

    public static explicit operator double(ColumnValue operand)
    {
        throw new NotImplementedException();
    }

    public static explicit operator float(ColumnValue operand)
    {
        throw new NotImplementedException();
    }

    public static explicit operator decimal(ColumnValue operand)
    {
        throw new NotImplementedException();
    }

    public static explicit operator long(ColumnValue operand)
    {
        throw new NotImplementedException();
    }

    public static explicit operator bool(ColumnValue operand)
    {
        throw new NotImplementedException();
    }

    public static explicit operator string(ColumnValue operand)
    {
        throw new NotImplementedException();
    }

    // The == / != operators above exist only to be captured in expression trees,
    // so equality here falls back to reference identity rather than throwing.
    public override bool Equals(object? obj)
    {
        return ReferenceEquals(this, obj);
    }

    public override int GetHashCode()
    {
        return RuntimeHelpers.GetHashCode(this);
    }
}