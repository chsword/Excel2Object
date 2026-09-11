using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertLookupTests : BaseFunctionTest
{
    [TestMethod]
    public void Lookups()
    {
        TestFunction(c => ExcelFunctions.Reference.HLookup(c["One"], c.Matrix("One", 1, "Three", 3), 2),
            "HLOOKUP(A4,A1:C3,2,FALSE)");
        TestFunction(
            c => ExcelFunctions.Reference.XLookup(c["One"], c.Matrix("Two", 2, "Two", 9),
                c.Matrix("Three", 2, "Three", 9)),
            "_xlfn.XLOOKUP(A4,B2:B9,C2:C9)");
        TestFunction(c => ExcelFunctions.Reference.Match(c["One"], c.Matrix("Two", 2, "Two", 9)),
            "MATCH(A4,B2:B9)");
    }

    [TestMethod]
    public void Addressing()
    {
        TestFunction(c => ExcelFunctions.Reference.Row(), "ROW()");
        TestFunction(c => ExcelFunctions.Reference.Row(c["One"]), "ROW(A4)");
        TestFunction(c => ExcelFunctions.Reference.Rows(c.Matrix("One", 2, "One", 9)), "ROWS(A2:A9)");
        TestFunction(c => ExcelFunctions.Reference.Indirect("A1"), "INDIRECT(\"A1\")");
        TestFunction(c => ExcelFunctions.Reference.Offset(c["One"], 1, 0), "OFFSET(A4,1,0)");
        TestFunction(c => ExcelFunctions.Reference.Hyperlink(c["One"], "open"), "HYPERLINK(A4,\"open\")");
    }

    [TestMethod]
    public void DynamicArrays()
    {
        TestFunction(c => ExcelFunctions.Reference.Unique(c.Matrix("One", 2, "One", 9)), "_xlfn.UNIQUE(A2:A9)");
        TestFunction(c => ExcelFunctions.Reference.Sort(c.Matrix("One", 2, "Two", 9)), "_xlfn._xlws.SORT(A2:B9)");
        TestFunction(c => ExcelFunctions.Reference.Filter(c.Matrix("One", 2, "Two", 9), c["Three"] > 1),
            "_xlfn._xlws.FILTER(A2:B9,C4>1)");
        TestFunction(c => ExcelFunctions.Reference.Sequence(10), "_xlfn.SEQUENCE(10)");
    }
}
