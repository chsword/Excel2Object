using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertLogicalTests : BaseFunctionTest
{
    [TestMethod]
    public void AndOrNotXor()
    {
        TestFunction(c => ExcelFunctions.Condition.And(c["One"] > 1, c["Two"] < 2), "AND(A4>1,B4<2)");
        TestFunction(c => ExcelFunctions.Condition.Or(c["One"] > 1, c["Two"] < 2), "OR(A4>1,B4<2)");
        TestFunction(c => ExcelFunctions.Condition.Xor(c["One"] > 1, c["Two"] < 2), "_xlfn.XOR(A4>1,B4<2)");
        TestFunction(c => ExcelFunctions.Condition.Not(c["One"] > 1), "NOT(A4>1)");
    }

    [TestMethod]
    public void TrueAndFalse()
    {
        TestFunction(c => ExcelFunctions.Condition.True(), "TRUE()");
        TestFunction(c => ExcelFunctions.Condition.False(), "FALSE()");
    }

    [TestMethod]
    public void IfVariants()
    {
        TestFunction(c => ExcelFunctions.Condition.If(c["One"] > 1, "y", "n"), "IF(A4>1,\"y\",\"n\")");
        TestFunction(c => ExcelFunctions.Condition.If(c["One"] > 1, "y"), "IF(A4>1,\"y\")");
        TestFunction(c => ExcelFunctions.Condition.Ifs(c["One"] > 1, "y", c["One"] > 0, "m"),
            "_xlfn.IFS(A4>1,\"y\",A4>0,\"m\")");
        TestFunction(c => ExcelFunctions.Condition.Switch(c["One"], 1, "one", 2, "two"),
            "_xlfn.SWITCH(A4,1,\"one\",2,\"two\")");
    }

    [TestMethod]
    public void ErrorHandling()
    {
        TestFunction(c => ExcelFunctions.Condition.IfError(c["One"] / c["Two"], 0), "IFERROR(A4/B4,0)");
        TestFunction(c => ExcelFunctions.Condition.IfNa(c["One"], ""), "_xlfn.IFNA(A4,\"\")");
    }

    [TestMethod]
    public void OperatorsBecomeLogicalFunctions()
    {
        TestFunction(c => (int) c["One"] > 1 && (int) c["Two"] < 2, "AND(A4>1,B4<2)");
        TestFunction(c => (int) c["One"] > 1 || (int) c["Two"] < 2, "OR(A4>1,B4<2)");
        TestFunction(c => !((int) c["One"] > 1), "NOT(A4>1)");
    }

    [TestMethod]
    public void CellConditionsCanBeCombinedWithTheLogicalOperators()
    {
        // a cell comparison is a ColumnValue, so && and || only compile because ColumnValue defines
        // operator true / operator false alongside & and |
        TestFunction(c => c["One"] > 1 && c["Two"] < 2, "AND(A4>1,B4<2)");
        TestFunction(c => c["One"] > 1 || c["Two"] < 2, "OR(A4>1,B4<2)");
        TestFunction(c => c["One"] > 1 | c["Two"] < 2, "OR(A4>1,B4<2)");
        TestFunction(c => !(c["One"] > 1), "NOT(A4>1)");
        TestFunction(c => (c["One"] > 1 && c["Two"] < 2) || c["Three"] == 0,
            "OR(AND(A4>1,B4<2),C4=0)");
    }

    [TestMethod]
    public void ConditionalOperatorBecomesIf()
    {
        TestFunction(c => (int) c["One"] > 5 ? "bulk" : "single", "IF(A4>5,\"bulk\",\"single\")");
    }
}
