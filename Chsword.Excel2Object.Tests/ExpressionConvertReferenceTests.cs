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
            2, c["Two"]), "VLOOKUP(A4,A10:C20,2,B4)");
        TestFunction(c => ExcelFunctions.Reference.VLookup("袁", c.Matrix("Two", 2, "Five", 7),
            2), "VLOOKUP(\"袁\",B2:E7,2)");
        //todo seach over sheet !
        // = VLOOKUP （A2，"客户端详细信息"！A:F，3，FALSE）
    }

    [TestMethod]
    public void HLookupTest()
    {
        TestFunction(c => ExcelFunctions.Reference.HLookup(c["One"], c.Matrix("One", 1, "Five", 4), 2, c["Two"]), "HLOOKUP(A4,A1:E4,2,B4)");
        TestFunction(c => ExcelFunctions.Reference.HLookup("key", c.Matrix("Two", 1, "Four", 3), 3), "HLOOKUP(\"key\",B1:D3,3)");
    }

    [TestMethod]
    public void RowTest()
    {
        TestFunction(c => ExcelFunctions.Reference.Row(c.Matrix("A", 1, "A", 1)), "ROW(A1:A1)");
        TestFunction(c => ExcelFunctions.Reference.Row(), "ROW()");
    }

    [TestMethod]
    public void ColumnTest()
    {
        TestFunction(c => ExcelFunctions.Reference.Column(c.Matrix("A", 1, "A", 1)), "COLUMN(A1:A1)");
        TestFunction(c => ExcelFunctions.Reference.Column(), "COLUMN()");
    }

    [TestMethod]
    public void OffsetTest()
    {
        TestFunction(c => ExcelFunctions.Reference.Offset(c.Matrix("A", 1, "A", 1), 1, 1), "OFFSET(A1:A1,1,1)");
        TestFunction(c => ExcelFunctions.Reference.Offset(c.Matrix("B", 4, "B", 4), c["One"], c["Two"], c["Three"], c["Four"]), "OFFSET(B4:B4,A4,B4,C4,D4)");
    }

    [TestMethod]
    public void IndirectTest()
    {
        TestFunction(c => ExcelFunctions.Reference.Indirect(c["A1"]), "INDIRECT(A14)");
        TestFunction(c => ExcelFunctions.Reference.Indirect("B1"), "INDIRECT(\"B1\")");
    }

    [TestMethod]
    public void AddressTest()
    {
        TestFunction(c => ExcelFunctions.Reference.Address(1, 1), "ADDRESS(1,1)");
        TestFunction(c => ExcelFunctions.Reference.Address(c["One"], c["Two"], 1, c["Three"], "Sheet1"), "ADDRESS(A4,B4,1,C4,\"Sheet1\")");
    }

    [TestMethod]
    public void RowsTest()
    {
        TestFunction(c => ExcelFunctions.Reference.Rows(c.Matrix("A", 1, "A", 10)), "ROWS(A1:A10)");
    }

    [TestMethod]
    public void ColumnsTest()
    {
        TestFunction(c => ExcelFunctions.Reference.Columns(c.Matrix("A", 1, "Z", 1)), "COLUMNS(A1:Z1)");
    }

    [TestMethod]
    public void CellTest()
    {
        TestFunction(c => ExcelFunctions.Reference.Cell("address", c.Matrix("A", 1, "A", 1)), "CELL(\"address\",A1:A1)");
        TestFunction(c => ExcelFunctions.Reference.Cell("type"), "CELL(\"type\")");
    }
}