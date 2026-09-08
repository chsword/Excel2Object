using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using Chsword.Excel2Object.Functions;
using Chsword.Excel2Object.Internal;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

public class BaseFunctionTest
{
    /// <summary>
    ///     Column titles of the other sheets a formula may refer to; unknown sheets resolve to null.
    /// </summary>
    private static readonly Dictionary<string, string[]> OtherSheets = new()
    {
        ["客户端详细信息"] = new[] {"客户", "地区", "电话", "邮箱", "负责人", "备注"},
        ["It's Rates"] = new[] {"Code", "Rate"}
    };

    protected void TestFunction(Expression<Func<ColumnCellDictionary, object>> exp, string expected)
    {
        var convert = new ExpressionConvert(new[] {"One", "Two", "Three", "Four", "Five", "Six"}, 3,
            title => OtherSheets.TryGetValue(title, out var columns) ? columns : null);
        var ret = convert.Convert(exp);
        Assert.AreEqual(expected, ret);
    }
}