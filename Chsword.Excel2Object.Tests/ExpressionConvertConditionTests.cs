using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertConditionTests : BaseFunctionTest
{
    [TestMethod]
    public void If()
    {
        TestFunction(c => ExcelFunctions.Condition.If(c["One"] == "Yes", 1, 2), "IF(A4=\"Yes\",1,2)");
    }

    [TestMethod]
    public void IfErrorTest()
    {
        TestFunction(c => ExcelFunctions.Condition.IfError(c["One"], "Error"), "IFERROR(A4,\"Error\")");
        TestFunction(c => ExcelFunctions.Condition.IfError(c["One"] / c["Two"], 0), "IFERROR(A4/B4,0)");
    }

    [TestMethod]
    public void IfNaTest()
    {
        TestFunction(c => ExcelFunctions.Condition.IfNa(c["One"], "N/A"), "IFNA(A4,\"N/A\")");
        TestFunction(c => ExcelFunctions.Condition.IfNa(c["One"], c["Two"]), "IFNA(A4,B4)");
    }

    [TestMethod]
    public void IfsTest()
    {
        TestFunction(c => ExcelFunctions.Condition.Ifs(c["One"] > 90, "A", c["One"] > 80, "B", c["One"] > 70, "C"), "IFS(A4>90,\"A\",A4>80,\"B\",A4>70,\"C\")");
    }

    [TestMethod]
    public void IsBlankTest()
    {
        TestFunction(c => ExcelFunctions.Condition.IsBlank(c["One"]), "ISBLANK(A4)");
    }

    [TestMethod]
    public void IsErrorTest()
    {
        TestFunction(c => ExcelFunctions.Condition.IsError(c["One"]), "ISERROR(A4)");
        TestFunction(c => ExcelFunctions.Condition.IsError(c["One"] / c["Two"]), "ISERROR(A4/B4)");
    }

    [TestMethod]
    public void IsEvenTest()
    {
        TestFunction(c => ExcelFunctions.Condition.IsEven(c["One"]), "ISEVEN(A4)");
        TestFunction(c => ExcelFunctions.Condition.IsEven(42), "ISEVEN(42)");
    }

    [TestMethod]
    public void IsLogicalTest()
    {
        TestFunction(c => ExcelFunctions.Condition.IsLogical(c["One"]), "ISLOGICAL(A4)");
    }

    [TestMethod]
    public void IsNaTest()
    {
        TestFunction(c => ExcelFunctions.Condition.IsNa(c["One"]), "ISNA(A4)");
    }

    [TestMethod]
    public void IsNumberTest()
    {
        TestFunction(c => ExcelFunctions.Condition.IsNumber(c["One"]), "ISNUMBER(A4)");
    }

    [TestMethod]
    public void IsOddTest()
    {
        TestFunction(c => ExcelFunctions.Condition.IsOdd(c["One"]), "ISODD(A4)");
        TestFunction(c => ExcelFunctions.Condition.IsOdd(43), "ISODD(43)");
    }

    [TestMethod]
    public void IsTextTest()
    {
        TestFunction(c => ExcelFunctions.Condition.IsText(c["One"]), "ISTEXT(A4)");
    }

    [TestMethod]
    public void IsFormulaTest()
    {
        TestFunction(c => ExcelFunctions.Condition.IsFormula(c["One"]), "ISFORMULA(A4)");
    }

    [TestMethod]
    public void OrTest()
    {
        TestFunction(c => ExcelFunctions.Condition.Or(c["One"] == "Yes", c["Two"] == "No"), "OR(A4=\"Yes\",B4=\"No\")");
    }

    [TestMethod]
    public void AndTest()
    {
        TestFunction(c => ExcelFunctions.Condition.And(c["One"] > 0, c["Two"] < 100), "AND(A4>0,B4<100)");
    }

    [TestMethod]
    public void NotTest()
    {
        TestFunction(c => ExcelFunctions.Condition.Not(c["One"] == "Yes"), "NOT(A4=\"Yes\")");
    }
}