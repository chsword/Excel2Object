using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertTextFunctionsTests : BaseFunctionTest
{
    [TestMethod]
    public void Pieces()
    {
        TestFunction(c => ExcelFunctions.Text.Left(c["One"], 3), "LEFT(A4,3)");
        TestFunction(c => ExcelFunctions.Text.Right(c["One"], 3), "RIGHT(A4,3)");
        TestFunction(c => ExcelFunctions.Text.Mid(c["One"], 2, 3), "MID(A4,2,3)");
        TestFunction(c => ExcelFunctions.Text.Len(c["One"]), "LEN(A4)");
        TestFunction(c => ExcelFunctions.Text.LenB(c["One"]), "LENB(A4)");
    }

    [TestMethod]
    public void SearchAndEdit()
    {
        TestFunction(c => ExcelFunctions.Text.Search("a", c["One"]), "SEARCH(\"a\",A4)");
        TestFunction(c => ExcelFunctions.Text.Substitute(c["One"], "-", "/"), "SUBSTITUTE(A4,\"-\",\"/\")");
        TestFunction(c => ExcelFunctions.Text.Replace(c["One"], 1, 2, "xx"), "REPLACE(A4,1,2,\"xx\")");
        TestFunction(c => ExcelFunctions.Text.Trim(c["One"]), "TRIM(A4)");
        TestFunction(c => ExcelFunctions.Text.Rept("-", 5), "REPT(\"-\",5)");
    }

    [TestMethod]
    public void CaseAndJoin()
    {
        TestFunction(c => ExcelFunctions.Text.Upper(c["One"]), "UPPER(A4)");
        TestFunction(c => ExcelFunctions.Text.Proper(c["One"]), "PROPER(A4)");
        TestFunction(c => ExcelFunctions.Text.Concat(c["One"], c["Two"]), "_xlfn.CONCAT(A4,B4)");
        TestFunction(c => ExcelFunctions.Text.TextJoin("-", true, c["One"], c["Two"]),
            "_xlfn.TEXTJOIN(\"-\",TRUE,A4,B4)");
    }

    [TestMethod]
    public void Conversion()
    {
        TestFunction(c => ExcelFunctions.Text.Text(c["One"], "0.00"), "TEXT(A4,\"0.00\")");
        TestFunction(c => ExcelFunctions.Text.Value(c["One"]), "VALUE(A4)");
        TestFunction(c => ExcelFunctions.Text.Char(65), "CHAR(65)");
    }

    [TestMethod]
    public void QuotesInsideTextAreDoubled()
    {
        TestFunction(c => ExcelFunctions.Text.Substitute(c["One"], "\"", "'"), "SUBSTITUTE(A4,\"\"\"\",\"'\")");
    }
}
