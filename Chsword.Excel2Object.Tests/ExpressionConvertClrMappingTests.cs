using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     Plain C# written inside a formula lambda - System.Math, string members, captured variables -
///     is translated to the Excel function that does the same thing.
/// </summary>
[TestClass]
public class ExpressionConvertClrMappingTests : BaseFunctionTest
{
    [TestMethod]
    public void SystemMathBecomesExcelMath()
    {
        TestFunction(c => Math.Abs((int) c["One"]), "ABS(A4)");
        TestFunction(c => Math.Pow((double) c["One"], 2), "POWER(A4,2)");
        TestFunction(c => Math.Max((int) c["One"], (int) c["Two"]), "MAX(A4,B4)");
        TestFunction(c => Math.Log((double) c["One"]), "LN(A4)");
        // .NET takes (y, x), Excel takes (x_num, y_num)
        TestFunction(c => Math.Atan2((double) c["One"], (double) c["Two"]), "ATAN2(B4,A4)");
        TestFunction(c => Math.Ceiling((double) c["One"]), "_xlfn.CEILING.MATH(A4)");
    }

    [TestMethod]
    public void StringMembersBecomeTextFunctions()
    {
        TestFunction(c => ((string) c["One"]).ToUpper(), "UPPER(A4)");
        TestFunction(c => ((string) c["One"]).Length, "LEN(A4)");
        TestFunction(c => ((string) c["One"]).Substring(2), "MID(A4,3,LEN(A4))");
        TestFunction(c => ((string) c["One"]).Substring(2, 3), "MID(A4,3,3)");
        TestFunction(c => ((string) c["One"]).Replace("-", "/"), "SUBSTITUTE(A4,\"-\",\"/\")");
        // FIND and EXACT, because .NET's Contains/StartsWith are case sensitive while Excel's
        // SEARCH and = are not
        TestFunction(c => ((string) c["One"]).Contains("xy"), "ISNUMBER(FIND(\"xy\",A4))");
        TestFunction(c => ((string) c["One"]).StartsWith("xy"), "EXACT(LEFT(A4,LEN(\"xy\")),\"xy\")");
        TestFunction(c => ((string) c["One"]).EndsWith("xy"), "EXACT(RIGHT(A4,LEN(\"xy\")),\"xy\")");
        // FIND errors when the text is absent, where .NET returns -1
        TestFunction(c => ((string) c["One"]).IndexOf("xy"), "IFERROR(FIND(\"xy\",A4)-1,-1)");
        // the char overloads of the same methods translate the same way
        TestFunction(c => ((string) c["One"]).Contains('x'), "ISNUMBER(FIND(\"x\",A4))");
        TestFunction(c => ((string) c["One"]).StartsWith('x'), "EXACT(LEFT(A4,LEN(\"x\")),\"x\")");
    }

    [TestMethod]
    public void OverloadsExcelCannotExpressAreNotTranslated()
    {
        // silently dropping startIndex or a StringComparison would change what the formula means
        StringAssert.Contains(Convert(c => ((string) c["One"]).IndexOf("xy", 2)), "unspport");
        StringAssert.Contains(
            Convert(c => ((string) c["One"]).StartsWith("xy", StringComparison.OrdinalIgnoreCase)),
            "unspport");
        StringAssert.Contains(Convert(c => Math.Round((double) c["One"], 2, MidpointRounding.ToEven)),
            "unspport");
        // IsNullOrWhiteSpace counts tabs and the Unicode spaces, Excel's TRIM only the ASCII space
        StringAssert.Contains(Convert(c => string.IsNullOrWhiteSpace((string) c["One"])), "unspport");
        // no Excel function rounds halves to even the way Math.Round does...
        StringAssert.Contains(Convert(c => Math.Round((double) c["One"], 2)), "unspport");
        StringAssert.Contains(Convert(c => Math.Round((double) c["One"])), "unspport");
        // ... and Excel's TRIM also collapses runs of spaces inside the text, which Trim() does not
        StringAssert.Contains(Convert(c => ((string) c["One"]).Trim()), "unspport");
        // the Excel functions themselves are still reachable when that is what is wanted
        Assert.AreEqual("ROUND(A4,2)", Convert(c => ExcelFunctions.Math.Round(c["One"], 2)));
        Assert.AreEqual("TRIM(A4)", Convert(c => ExcelFunctions.Text.Trim(c["One"])));
        // ... while the plain overloads still do translate
        Assert.AreEqual("ISNUMBER(FIND(\"xy\",A4))", Convert(c => ((string) c["One"]).Contains("xy")));
    }

    [TestMethod]
    public void DateLiteralsKeepSubSecondPrecision()
    {
        TestFunction(c => c["One"] > new DateTime(2024, 3, 1, 8, 30, 0),
            "A4>DATE(2024,3,1)+TIME(8,30,0)");
        TestFunction(c => c["One"] > new DateTime(2024, 3, 1, 8, 30, 0, 500),
            "A4>DATE(2024,3,1)+TIME(8,30,0)+5.787037037037037E-06");
    }

    [TestMethod]
    public void CapturedVariablesBecomeLiterals()
    {
        var rate = 0.25m;
        var label = "VIP";
        TestFunction(c => c["One"] * rate, "A4*0.25");
        TestFunction(c => c["One"] + label, "A4+\"VIP\"");
    }

    [TestMethod]
    public void NumbersAreWrittenWithAnInvariantDecimalPoint()
    {
        TestFunction(c => c["One"] * 1.5, "A4*1.5");
        TestFunction(c => c["One"] * 0.125m, "A4*0.125");
    }

    [TestMethod]
    public void StringJoinKeepsEmptyEntries()
    {
        TestFunction(c => string.Join("-", new[] {(string) c["One"], (string) c["Two"]}),
            "_xlfn.TEXTJOIN(\"-\",FALSE,A4,B4)");
    }

    [TestMethod]
    public void BitwiseOperatorsOnIntegersBecomeTheBitFunctions()
    {
        TestFunction(c => (int) c["One"] & 6, "_xlfn.BITAND(A4,6)");
        TestFunction(c => (int) c["One"] | 6, "_xlfn.BITOR(A4,6)");
        TestFunction(c => (int) c["One"] ^ 6, "_xlfn.BITXOR(A4,6)");
    }

    [TestMethod]
    public void XorOnBooleansIsAFutureFunction()
    {
        TestFunction(c => (int) c["One"] > 1 ^ (int) c["Two"] < 2, "_xlfn.XOR(A4>1,B4<2)");
    }

    [TestMethod]
    public void ModuloBecomesMod()
    {
        TestFunction(c => c["One"] % 2, "MOD(A4,2)");
    }

    [TestMethod]
    public void CaretBecomesThePowerOperator()
    {
        TestFunction(c => c["One"] ^ 2, "A4^2");
        TestFunction(c => (c["One"] ^ 2) * 3, "A4^2*3");
        TestFunction(c => c["One"] ^ (c["Two"] ^ 2), "A4^(B4^2)");
    }
}
