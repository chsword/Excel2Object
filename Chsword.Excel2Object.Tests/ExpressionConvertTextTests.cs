using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertTextTests : BaseFunctionTest
{
    [TestMethod]
    public void FindTest()
    {
        TestFunction(c => ExcelFunctions.Text.Find("M", c["One", 2]), "FIND(\"M\",A2)");
        TestFunction(c => ExcelFunctions.Text.Find("M", c["One", 2], 2), "FIND(\"M\",A2,2)");
        TestFunction(c => ExcelFunctions.Text.Find(c["Two"], c["Three"]), "FIND(B4,C4)");
    }

    [TestMethod]
    public void SearchTest()
    {
        TestFunction(c => ExcelFunctions.Text.Search("text", c["One"]), "SEARCH(\"text\",A4)");
        TestFunction(c => ExcelFunctions.Text.Search("abc", c["Two", 5], 3), "SEARCH(\"abc\",B5,3)");
        TestFunction(c => ExcelFunctions.Text.Search(c["Three"], c["Four", 1]), "SEARCH(C4,D1)");
    }

    [TestMethod]
    public void LeftTest()
    {
        TestFunction(c => ExcelFunctions.Text.Left(c["One"], 5), "LEFT(A4,5)");
        TestFunction(c => ExcelFunctions.Text.Left(c["Two", 3], c["Three"]), "LEFT(B3,C4)");
    }

    [TestMethod]
    public void RightTest()
    {
        TestFunction(c => ExcelFunctions.Text.Right(c["One"], 3), "RIGHT(A4,3)");
        TestFunction(c => ExcelFunctions.Text.Right(c["Four", 2], c["Five"]), "RIGHT(D2,E4)");
    }

    [TestMethod]
    public void MidTest()
    {
        TestFunction(c => ExcelFunctions.Text.Mid(c["One"], 2, 4), "MID(A4,2,4)");
        TestFunction(c => ExcelFunctions.Text.Mid(c["Two", 1], c["Three"], c["Four", 5]), "MID(B1,C4,D5)");
    }

    [TestMethod]
    public void LenTest()
    {
        TestFunction(c => ExcelFunctions.Text.Len(c["One"]), "LEN(A4)");
        TestFunction(c => ExcelFunctions.Text.Len(c["Six", 7]), "LEN(F7)");
    }

    [TestMethod]
    public void UpperTest()
    {
        TestFunction(c => ExcelFunctions.Text.Upper(c["One"]), "UPPER(A4)");
        TestFunction(c => ExcelFunctions.Text.Upper(c["Two", 8]), "UPPER(B8)");
    }

    [TestMethod]
    public void LowerTest()
    {
        TestFunction(c => ExcelFunctions.Text.Lower(c["One"]), "LOWER(A4)");
        TestFunction(c => ExcelFunctions.Text.Lower(c["Three", 6]), "LOWER(C6)");
    }

    [TestMethod]
    public void ProperTest()
    {
        TestFunction(c => ExcelFunctions.Text.Proper(c["One"]), "PROPER(A4)");
        TestFunction(c => ExcelFunctions.Text.Proper(c["Four", 2]), "PROPER(D2)");
    }

    [TestMethod]
    public void TrimTest()
    {
        TestFunction(c => ExcelFunctions.Text.Trim(c["One"]), "TRIM(A4)");
        TestFunction(c => ExcelFunctions.Text.Trim(c["Five", 3]), "TRIM(E3)");
    }

    [TestMethod]
    public void SubstituteTest()
    {
        TestFunction(c => ExcelFunctions.Text.Substitute(c["One"], "old", "new"), "SUBSTITUTE(A4,\"old\",\"new\")");
        TestFunction(c => ExcelFunctions.Text.Substitute(c["Two"], c["Three"], c["Four"], 1), "SUBSTITUTE(B4,C4,D4,1)");
        TestFunction(c => ExcelFunctions.Text.Substitute(c["One", 5], "a", "b"), "SUBSTITUTE(A5,\"a\",\"b\")");
    }

    [TestMethod]
    public void ReplaceTest()
    {
        TestFunction(c => ExcelFunctions.Text.Replace(c["One"], 2, 3, "new"), "REPLACE(A4,2,3,\"new\")");
        TestFunction(c => ExcelFunctions.Text.Replace(c["Two", 1], c["Three"], c["Four"], c["Five", 2]), "REPLACE(B1,C4,D4,E2)");
    }

    [TestMethod]
    public void ReptTest()
    {
        TestFunction(c => ExcelFunctions.Text.Rept(c["One"], 3), "REPT(A4,3)");
        TestFunction(c => ExcelFunctions.Text.Rept("*", c["Two"]), "REPT(\"*\",B4)");
    }

    [TestMethod]
    public void ConcatenateTest()
    {
        TestFunction(c => ExcelFunctions.Text.Concatenate(c["One"], c["Two"]), "CONCATENATE(A4,B4)");
        TestFunction(c => ExcelFunctions.Text.Concatenate(c["One"], " ", c["Two"], "!"), "CONCATENATE(A4,\" \",B4,\"!\")");
        TestFunction(c => ExcelFunctions.Text.Concatenate(c["Three", 2], c["Four", 2], c["Five", 2]), "CONCATENATE(C2,D2,E2)");
    }

    [TestMethod]
    public void TextTest()
    {
        TestFunction(c => ExcelFunctions.Text.Text(c["One"], "0.00"), "TEXT(A4,\"0.00\")");
        TestFunction(c => ExcelFunctions.Text.Text(c["Two", 5], c["Three"]), "TEXT(B5,C4)");
    }

    [TestMethod]
    public void ValueTest()
    {
        TestFunction(c => ExcelFunctions.Text.Value(c["One"]), "VALUE(A4)");
        TestFunction(c => ExcelFunctions.Text.Value(c["Six", 1]), "VALUE(F1)");
    }

    [TestMethod]
    public void CodeTest()
    {
        TestFunction(c => ExcelFunctions.Text.Code(c["One"]), "CODE(A4)");
        TestFunction(c => ExcelFunctions.Text.Code(c["Two", 9]), "CODE(B9)");
    }

    [TestMethod]
    public void CharTest()
    {
        TestFunction(c => ExcelFunctions.Text.Char(65), "CHAR(65)");
        TestFunction(c => ExcelFunctions.Text.Char(c["One"]), "CHAR(A4)");
    }

    [TestMethod]
    public void ExactTest()
    {
        TestFunction(c => ExcelFunctions.Text.Exact(c["One"], c["Two"]), "EXACT(A4,B4)");
        TestFunction(c => ExcelFunctions.Text.Exact(c["Three", 1], "test"), "EXACT(C1,\"test\")");
    }

    [TestMethod]
    public void FixedTest()
    {
        TestFunction(c => ExcelFunctions.Text.Fixed(c["One"], 2), "FIXED(A4,2)");
        TestFunction(c => ExcelFunctions.Text.Fixed(c["Two"], c["Three"], c["Four"]), "FIXED(B4,C4,D4)");
        TestFunction(c => ExcelFunctions.Text.Fixed(1234.567, 1), "FIXED(1234.567,1)");
    }

    [TestMethod]
    public void AscTest()
    {
        TestFunction(c => ExcelFunctions.Text.Asc(c["One"]), "ASC(A4)");
        TestFunction(c => ExcelFunctions.Text.Asc(c["Four", 6]), "ASC(D6)");
    }
}