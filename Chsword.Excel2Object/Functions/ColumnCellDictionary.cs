﻿using System;
using System.Collections.Generic;

namespace Chsword.Excel2Object.Functions
{
    public class ColumnCellDictionary : Dictionary<string, ColumnValue>
    {
        public ColumnValue this[string columnName, int rowNumber] => throw new NotImplementedException();

        public ColumnMatrix Matrix(string keyA, int rowA, string keyB, int rowB)
        {
            throw new NotImplementedException();
        }

        public dynamic Model => throw new NotImplementedException();
    }
}