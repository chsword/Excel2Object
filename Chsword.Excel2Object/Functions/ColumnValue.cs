// ReSharper disable UnusedParameter.Global

using System.Runtime.CompilerServices;

namespace Chsword.Excel2Object.Functions;

// this type only used in the Expression
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