using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertReferenceTests : BaseFunctionTest
{
    [TestMethod]
    public void Choose()
    {
        TestFunction(c => ExcelFunctions.Reference.Choose(2, c["One", 2], c["One", 3]
            , c["One", 4], c["One", 5]), "CHOOSE(2,A2,A3,A4,A5)");
    }

    [TestMethod]
    public void Index()
    {
        TestFunction(c => ExcelFunctions.Reference.Index(c.Matrix("One", 2, "Two", 6),
            2, 3), "INDEX(A2:B6,2,3)");
    }

    [TestMethod]
    public void Lookup()
    {
        TestFunction(c => ExcelFunctions.Reference.Lookup(4.19, c.Matrix("One", 2, "One", 6),
            c.Matrix("Two", 2, "Two", 6)), "LOOKUP(4.19,A2:A6,B2:B6)");
    }

    [TestMethod]
    public void Match()
    {
        TestFunction(c => ExcelFunctions.Reference.Match(39, c.Matrix("Two", 2, "Two", 5),
            1), "MATCH(39,B2:B5,1)");
    }

    [TestMethod]
    public void VLookup()
    {
        TestFunction(c => ExcelFunctions.Reference.VLookup(c["One"], c.Matrix("One", 10, "Three", 20),
            2, true), "VLOOKUP(A4,A10:C20,2,TRUE)");
        TestFunction(c => ExcelFunctions.Reference.VLookup("袁", c.Matrix("Two", 2, "Five", 7),
            2, false), "VLOOKUP(\"袁\",B2:E7,2,FALSE)");
    }

    [TestMethod]
    public void WholeColumns()
    {
        TestFunction(c => ExcelFunctions.Reference.VLookup(c["One"], c.Columns("Two", "Four"), 2, false),
            "VLOOKUP(A4,B:D,2,FALSE)");
    }

    [TestMethod]
    public void VLookupOverSheet()
    {
        // Titles of the other sheet resolve to its column letters.
        TestFunction(c => ExcelFunctions.Reference.VLookup(c["One"],
                c.Sheet("客户端详细信息").Columns("客户", "备注"), 3, false),
            "VLOOKUP(A4,'客户端详细信息'!A:F,3,FALSE)");
        TestFunction(c => ExcelFunctions.Reference.VLookup(c["One"],
                c.Sheet("客户端详细信息").Matrix("客户", 2, "电话", 100), 2, true),
            "VLOOKUP(A4,'客户端详细信息'!A2:C100,2,TRUE)");
    }

    [TestMethod]
    public void CellOverSheet()
    {
        // Same row as the formula cell, and an explicit row.
        TestFunction(c => c.Sheet("客户端详细信息")["电话"] & c["Two"], "'客户端详细信息'!C4&B4");
        TestFunction(c => c.Sheet("客户端详细信息")["电话", 2], "'客户端详细信息'!C2");
    }

    [TestMethod]
    public void SheetTitleFromVariableAndApostrophe()
    {
        var sheet = "It's Rates";
        TestFunction(c => c["One"] * c.Sheet(sheet)["Rate", 2], "A4*'It''s Rates'!B2");
    }

    [TestMethod]
    public void UnknownSheetFallsBackToColumnLetters()
    {
        // A sheet the exporter knows nothing about can still be referenced by column letters...
        TestFunction(c => ExcelFunctions.Reference.VLookup(c["One"], c.Sheet("Legacy").Columns("A", "F"), 3, false),
            "VLOOKUP(A4,'Legacy'!A:F,3,FALSE)");
        // ...but not by a title, since there is nothing to resolve it against.
        var ex = Assert.ThrowsException<Excel2ObjectException>(() => TestFunction(c => c.Sheet("Legacy")["Total"], ""));
        StringAssert.Contains(ex.Message, "[Total]");
        StringAssert.Contains(ex.Message, "[Legacy]");
    }

    [TestMethod]
    public void UnknownColumnOnCurrentSheetThrows()
    {
        var ex = Assert.ThrowsException<Excel2ObjectException>(() => TestFunction(c => c["Seven"], ""));
        StringAssert.Contains(ex.Message, "[Seven]");
    }
}