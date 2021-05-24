using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using Chsword.Excel2Object.Functions;
using Chsword.Excel2Object.Internal;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests
{
    public class BaseFunctionTest
    {
        protected void TestFunction(Expression<Func<ColumnCellDictionary, object>> exp, string expected)
        {
            var convert = new ExpressionConvert(new[] { "One","Two","Three","Four","Five","Six" }, 3);
            var ret = convert.Convert(exp);
            Assert.AreEqual(expected, ret);
        }
    }
}