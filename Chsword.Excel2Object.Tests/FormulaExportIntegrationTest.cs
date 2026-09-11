using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using Chsword.Excel2Object.Functions;
using Chsword.Excel2Object.Options;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     The formulas produced for the newly supported functions have to survive NPOI's formula parser,
///     which is what actually writes them into the workbook.
/// </summary>
[TestClass]
public class FormulaExportIntegrationTest : BaseExcelTest
{
    private static readonly List<Dictionary<string, object>> Rows = new()
    {
        new() {["Name"] = "a", ["Qty"] = 4, ["Price"] = 2.5},
        new() {["Name"] = "b", ["Qty"] = 9, ["Price"] = 1.5}
    };

    private static byte[] Export(string title, Expression<Func<ColumnCellDictionary, object>> formula)
    {
        return new ExcelExporter().ObjectToExcelBytes(Rows, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.FormulaColumns.Add(new FormulaColumn
            {
                Title = title,
                Formula = formula,
                AfterColumnTitle = "Price"
            });
        });
    }

    [TestMethod]
    public void NewFunctionsProduceParsableFormulas()
    {
        var cases = new List<Expression<Func<ColumnCellDictionary, object>>>
        {
            c => ExcelFunctions.Statistics.Average(c.Matrix("Qty", 2, "Qty", 3)),
            c => ExcelFunctions.Statistics.CountIf(c.Matrix("Qty", 2, "Qty", 3), ">3"),
            c => ExcelFunctions.Math.SumIf(c.Matrix("Qty", 2, "Qty", 3), ">3", c.Matrix("Price", 2, "Price", 3)),
            c => ExcelFunctions.Math.Mod(c["Qty"], 2),
            c => ExcelFunctions.Text.Upper(c["Name"]),
            c => ExcelFunctions.Text.Text(c["Price"], "0.00"),
            c => ExcelFunctions.Condition.IfError(c["Price"] / c["Qty"], 0),
            c => ExcelFunctions.Condition.And(c["Qty"] > 1, c["Price"] > 1),
            c => ExcelFunctions.Information.IsBlank(c["Name"]),
            c => ExcelFunctions.DateAndTime.EoMonth(DateTime.Today, 1),
            c => ExcelFunctions.Financial.Pmt(0.05, 12, c["Price"]),
            c => ExcelFunctions.Engineering.Dec2Hex(c["Qty"]),
            c => (int) c["Qty"] > 5 ? "bulk" : "single",
            c => c["Qty"] % 2,
            c => c["Price"] ^ 2,
            // functions Excel gained after 2007, which the sheet has to store with the _xlfn. prefix
            c => ExcelFunctions.Condition.Ifs(c["Qty"] > 5, "bulk", c["Qty"] > 0, "single"),
            c => ExcelFunctions.Condition.IfNa(c["Qty"], 0),
            c => ExcelFunctions.Statistics.StDevP(c.Matrix("Qty", 2, "Qty", 3)),
            c => ExcelFunctions.Statistics.MaxIfs(c.Matrix("Qty", 2, "Qty", 3),
                c.Matrix("Price", 2, "Price", 3), ">1"),
            c => ExcelFunctions.Reference.XLookup(c["Name"], c.Matrix("Name", 2, "Name", 3),
                c.Matrix("Qty", 2, "Qty", 3)),
            c => ExcelFunctions.Reference.Filter(c.Matrix("Name", 2, "Qty", 3), c["Qty"] > 1),
            c => ExcelFunctions.Reference.Sort(c.Matrix("Name", 2, "Qty", 3)),
            c => ExcelFunctions.Text.TextJoin("-", true, c["Name"], c["Qty"]),
            c => ExcelFunctions.Text.Concat(c["Name"], c["Qty"]),
            c => ExcelFunctions.DateAndTime.IsoWeekNum(DateTime.Today),
            c => ExcelFunctions.Information.IsFormula(c["Price"])
        };
        foreach (var formula in cases)
        {
            var bytes = Export("F", formula);
            Assert.IsNotNull(bytes, formula.ToString());
        }
    }

    [TestMethod]
    public void FutureFunctionsAreStoredWithTheirPrefix()
    {
        var bytes = Export("Level",
            c => ExcelFunctions.Condition.Ifs(c["Qty"] > 5, "bulk", c["Qty"] > 0, "single"));
        using var stream = new System.IO.MemoryStream(bytes);
        var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(stream);
        var sheet = workbook.GetSheetAt(0);
        var formula = sheet.GetRow(1).Cells.Last().CellFormula;
        StringAssert.StartsWith(formula, "_xlfn.IFS(");
    }

    [TestMethod]
    public void FutureFunctionsKeepTheirPrefixInXlsToo()
    {
        // The _xlfn. prefix is how a function outside the file format's own function table is named,
        // which the 97-2003 format needs at least as much as xlsx does, so it is not gated on type.
        var bytes = new ExcelExporter().ObjectToExcelBytes(Rows, options =>
        {
            options.ExcelType = ExcelType.Xls;
            options.FormulaColumns.Add(new FormulaColumn
            {
                Title = "Level",
                Formula = (Expression<Func<ColumnCellDictionary, object>>) (c =>
                    ExcelFunctions.Condition.Ifs(c["Qty"] > 5, "bulk", c["Qty"] > 0, "single")),
                AfterColumnTitle = "Price"
            });
        });
        using var stream = new System.IO.MemoryStream(bytes);
        var workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(stream);
        var formula = workbook.GetSheetAt(0).GetRow(1).Cells.Last().CellFormula;
        StringAssert.StartsWith(formula, "_xlfn.IFS(");
    }

    [TestMethod]
    public void ReadingBackASheetWithAFutureFunctionDoesNotThrow()
    {
        // NPOI cannot evaluate the functions it does not implement; reading such a sheet has to
        // degrade to an empty value rather than fail the whole import.
        var bytes = Export("Level",
            c => ExcelFunctions.Condition.Ifs(c["Qty"] > 5, "bulk", c["Qty"] > 0, "single"));
        var result = ExcelHelper.ExcelToObject<Dictionary<string, object>>(bytes).ToList();
        Assert.AreEqual(2, result.Count);
        Assert.AreEqual("a", result[0]["Name"]);
    }

    [TestMethod]
    public void ConditionalSumIsComputedByExcel()
    {
        var bytes = Export("Total", c => ExcelFunctions.Math.Round(c["Qty"] * c["Price"], 2));
        var result = ExcelHelper.ExcelToObject<Dictionary<string, object>>(bytes).ToList();
        Assert.AreEqual("10", result[0]["Total"]);
        Assert.AreEqual("13.5", result[1]["Total"]);
    }
}
