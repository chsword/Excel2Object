using System;

namespace Chsword.Excel2Object.Functions
{
    public class ColumnValue
    {
        #region Compare

        public static ColumnValue operator ==(ColumnValue a, ColumnValue b)
        {
            throw new NotImplementedException();
        }
        public static ColumnValue operator !=(ColumnValue a, ColumnValue b)
        {
            throw new NotImplementedException();
        }
        public static ColumnValue operator >(ColumnValue a, ColumnValue b)
        {
            throw new NotImplementedException();
        }
        public static ColumnValue operator <(ColumnValue a, ColumnValue b)
        {
            throw new NotImplementedException();
        }
        public static ColumnValue operator >=(ColumnValue a, ColumnValue b)
        {
            throw new NotImplementedException();
        }
        public static ColumnValue operator <=(ColumnValue a, ColumnValue b)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region Binary


        public static ColumnValue operator +(ColumnValue a, ColumnValue b)
        {
            throw new NotImplementedException();
        }
        public static ColumnValue operator -(ColumnValue a, ColumnValue b)
        {
            throw new NotImplementedException();
        }
        public static ColumnValue operator *(ColumnValue a, ColumnValue b)
        {
            throw new NotImplementedException();
        }
        public static ColumnValue operator /(ColumnValue a, ColumnValue b)
        {
            throw new NotImplementedException();
        }

        public static ColumnValue operator &(ColumnValue a, ColumnValue b)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region Convert

        public static implicit operator ColumnValue(int operand)
        {
            throw new NotImplementedException();
        }
        public static implicit operator ColumnValue(string operand)
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
        #endregion
        public static ColumnValue operator -(ColumnValue a)
        {
            throw new NotImplementedException();
        }
    }
}