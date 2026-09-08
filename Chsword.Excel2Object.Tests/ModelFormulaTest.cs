using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Linq.Expressions;
using Chsword.Excel2Object.Functions;
using Chsword.Excel2Object.Internal;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     Formula columns added with FormulaColumns.Add&lt;TModel&gt; refer to columns through model properties
///     instead of title strings (issue #22).
/// </summary>
[TestClass]
public class ModelFormulaTest : BaseExcelTest
{
    public class OrderLine
    {
        [ExcelTitle("Product")] public string Product { get; set; } = "";
        [ExcelTitle("Price")] public decimal Price { get; set; }
        [ExcelTitle("Qty")] public int Qty { get; set; }
        [Display(Name = "Ordered")] public DateTime Ordered { get; set; }
        [ExcelTitle("Discount")] public decimal? Discount { get; set; }
        public string Note { get; set; } = "";
    }

    private static readonly string[] Columns = {"Product", "Price", "Qty", "Ordered", "Discount"};

    private static string Convert(Expression<Func<ColumnCellDictionary, OrderLine, object>> formula)
    {
        return new ExpressionConvert(Columns, 1).Convert(formula);
    }

    private static void Test(Expression<Func<ColumnCellDictionary, OrderLine, object>> formula, string expected)
    {
        Assert.AreEqual(expected, Convert(formula));
    }

    [TestMethod]
    public void PropertiesMapToCellsOnTheCurrentRow()
    {
        Test((c, m) => m.Price * m.Qty, "B2*C2");
        Test((c, m) => m.Qty, "C2");
        Test((c, m) => m.Ordered.Year, "YEAR(D2)");
        Test((c, m) => m.Price * (1 - m.Discount.Value), "B2*(1-E2)");
    }

    [TestMethod]
    public void StringConcatenationBecomesAmpersand()
    {
        Test((c, m) => m.Product + "-" + m.Qty, "A2&\"-\"&C2");
    }

    [TestMethod]
    public void PropertiesCanBePassedToFunctions()
    {
        Test((c, m) => ExcelFunctions.Condition.If(m.Qty > 5, "bulk", "single"), "IF(C2>5,\"bulk\",\"single\")");
        Test((c, m) => ExcelFunctions.Math.Round(m.Price * m.Qty, 2), "ROUND(B2*C2,2)");
    }

    [TestMethod]
    public void ModelAndTitleAccessCanBeMixed()
    {
        Test((c, m) => m.Price * c["Qty", 2], "B2*C2");
        Test((c, m) => ExcelFunctions.Statistics.Sum(c.Matrix("Price", 2, "Price", 10)) - m.Price,
            "SUM(B2:B10)-B2");
    }

    [TestMethod]
    public void UnexportedPropertyIsRejected()
    {
        var ex = Assert.ThrowsException<Excel2ObjectException>(() => Convert((c, m) => m.Note));
        StringAssert.Contains(ex.Message, "[Note]");
        StringAssert.Contains(ex.Message, nameof(OrderLine));
    }

    [TestMethod]
    public void ParenthesesFollowExcelPrecedence()
    {
        Test((c, m) => (m.Price + m.Qty) * m.Qty, "(B2+C2)*C2");
        Test((c, m) => m.Price + m.Qty * m.Qty, "B2+C2*C2");
        Test((c, m) => m.Price - (m.Qty - m.Price), "B2-(C2-B2)");
        Test((c, m) => m.Price - m.Qty - m.Price, "B2-C2-B2");
        Test((c, m) => -(m.Price + m.Qty), "-(B2+C2)");
        Test((c, m) => m.Price * m.Qty > m.Price + 1, "B2*C2>B2+1");
    }

    [TestMethod]
    public void ExportedFormulaEvaluates()
    {
        var lines = new List<OrderLine>
        {
            new() {Product = "Apple", Price = 3, Qty = 4, Ordered = new DateTime(2026, 9, 8), Discount = 0.5m},
            new() {Product = "Pear", Price = 5, Qty = 2, Ordered = new DateTime(2026, 9, 9), Discount = null}
        };
        var bytes = new ExcelExporter().ObjectToExcelBytes(lines, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.FormulaColumns.Add<OrderLine>("Total", (c, m) => m.Price * m.Qty);
            options.FormulaColumns.Add<OrderLine>("Label", (c, m) => m.Product + "-" + m.Qty);
            options.FormulaColumns.Add<OrderLine>("Net",
                (c, m) => ExcelFunctions.Math.Round(m.Price * m.Qty * (1 - m.Discount.Value), 2));
        });
        Assert.IsNotNull(bytes);

        var result = ExcelHelper.ExcelToObject<Dictionary<string, object>>(bytes).ToList();
        Assert.AreEqual("12", result[0]["Total"]);
        Assert.AreEqual("Apple-4", result[0]["Label"]);
        Assert.AreEqual("6", result[0]["Net"]);
        Assert.AreEqual("10", result[1]["Total"]);
        Assert.AreEqual("10", result[1]["Net"]); // null discount is a blank cell, which Excel treats as 0
    }
}
