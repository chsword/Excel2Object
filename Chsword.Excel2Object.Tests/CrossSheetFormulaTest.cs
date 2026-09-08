using System.Collections.Generic;
using System.Linq;
using Chsword.Excel2Object.Options;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     Formula columns that refer to another sheet of the same workbook via c.Sheet("...").
/// </summary>
[TestClass]
public class CrossSheetFormulaTest : BaseExcelTest
{
    private static readonly List<Dictionary<string, object>> Products = new()
    {
        new() {["Name"] = "Apple", ["Price"] = "3"},
        new() {["Name"] = "Pear", ["Price"] = "5"},
        new() {["Name"] = "Plum", ["Price"] = "7"}
    };

    private static readonly List<Dictionary<string, object>> Orders = new()
    {
        new() {["Product"] = "Pear", ["Qty"] = "2"},
        new() {["Product"] = "Plum", ["Qty"] = "10"},
        new() {["Product"] = "Apple", ["Qty"] = "1"}
    };

    [TestMethod]
    public void VLookupIntoPreviouslyWrittenSheet()
    {
        var exporter = new ExcelExporter();
        var bytes = exporter.ObjectToExcelBytes(Products, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.SheetTitle = "Products";
        });
        Assert.IsNotNull(bytes);

        // Append a second sheet whose Total column looks the unit price up on the first sheet by title.
        bytes = exporter.ObjectToExcelBytes(Orders, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.SheetTitle = "Orders";
            options.SourceExcelBytes = bytes;
            options.FormulaColumns.Add("Total",
                c => ExcelFunctions.Reference.VLookup(c["Product"], c.Sheet("Products").Columns("Name", "Price"),
                    2, false) * c["Qty"]);
        });
        Assert.IsNotNull(bytes);

        // The importer evaluates formulas, so the cross-sheet lookup must produce real numbers.
        var result = ExcelHelper.ExcelToObject<Dictionary<string, object>>(bytes, "Orders").ToList();
        Assert.AreEqual(3, result.Count);
        Assert.AreEqual("10", result[0]["Total"]); // Pear 5 * 2
        Assert.AreEqual("70", result[1]["Total"]); // Plum 7 * 10
        Assert.AreEqual("3", result[2]["Total"]); // Apple 3 * 1
    }

    [TestMethod]
    public void SheetTitleWithSpaceIsQuoted()
    {
        var exporter = new ExcelExporter();
        var bytes = exporter.ObjectToExcelBytes(Products, options =>
        {
            options.ExcelType = ExcelType.Xls;
            options.SheetTitle = "Price List";
        });
        bytes = exporter.ObjectToExcelBytes(Orders, options =>
        {
            options.ExcelType = ExcelType.Xls;
            options.SheetTitle = "Orders";
            options.SourceExcelBytes = bytes;
            options.FormulaColumns.Add("Total",
                c => ExcelFunctions.Reference.VLookup(c["Product"], c.Sheet("Price List").Matrix("Name", 2, "Price", 4),
                    2, false) * c["Qty"]);
        });
        Assert.IsNotNull(bytes);

        var result = ExcelHelper.ExcelToObject<Dictionary<string, object>>(bytes, "Orders").ToList();
        Assert.AreEqual("10", result[0]["Total"]);
        Assert.AreEqual("70", result[1]["Total"]);
        Assert.AreEqual("3", result[2]["Total"]);
    }

    [TestMethod]
    public void MissingSheetIsReported()
    {
        var ex = Assert.ThrowsException<Excel2ObjectException>(() =>
            new ExcelExporter().ObjectToExcelBytes(Orders, options =>
            {
                options.ExcelType = ExcelType.Xlsx;
                options.SheetTitle = "Orders";
                options.FormulaColumns.Add("Total", c => c.Sheet("Nowhere")["B", 2] * c["Qty"]);
            }));
        StringAssert.Contains(ex.Message, "[Total]");
        StringAssert.Contains(ex.Message, "[Nowhere]");
    }

    [TestMethod]
    public void UnknownColumnTitleIsReported()
    {
        var exporter = new ExcelExporter();
        var bytes = exporter.ObjectToExcelBytes(Products, options => options.SheetTitle = "Products");
        var ex = Assert.ThrowsException<Excel2ObjectException>(() =>
            exporter.ObjectToExcelBytes(Orders, options =>
            {
                options.SheetTitle = "Orders";
                options.SourceExcelBytes = bytes;
                options.FormulaColumns.Add("Total", c => c.Sheet("Products")["Prise", 2] * c["Qty"]);
            }));
        StringAssert.Contains(ex.Message, "[Total]");
        StringAssert.Contains(ex.Message, "[Prise]");
    }
}
