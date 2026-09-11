using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertMathFunctionsTests : BaseFunctionTest
{
    [TestMethod]
    public void AbsTest()
    {
        TestFunction(c => ExcelFunctions.Math.Abs(c["One"]), "ABS(A4)");
    }

    [TestMethod]
    public void PITest()
    {
        TestFunction(c => ExcelFunctions.Math.PI(), "PI()");
    }

    [TestMethod]
    public void Rounding()
    {
        TestFunction(c => ExcelFunctions.Math.Round(c["One"], 2), "ROUND(A4,2)");
        TestFunction(c => ExcelFunctions.Math.RoundUp(c["One"], 0), "ROUNDUP(A4,0)");
        TestFunction(c => ExcelFunctions.Math.Ceiling(c["One"], 5), "CEILING(A4,5)");
        TestFunction(c => ExcelFunctions.Math.CeilingMath(c["One"]), "_xlfn.CEILING.MATH(A4)");
        TestFunction(c => ExcelFunctions.Math.FloorMath(c["One"], 5), "_xlfn.FLOOR.MATH(A4,5)");
        TestFunction(c => ExcelFunctions.Math.MRound(c["One"], 5), "MROUND(A4,5)");
        TestFunction(c => ExcelFunctions.Math.Trunc(c["One"], 1), "TRUNC(A4,1)");
    }

    [TestMethod]
    public void Arithmetic()
    {
        TestFunction(c => ExcelFunctions.Math.Mod(c["One"], 3), "MOD(A4,3)");
        TestFunction(c => ExcelFunctions.Math.Power(c["One"], 2), "POWER(A4,2)");
        TestFunction(c => ExcelFunctions.Math.Ln(c["One"]), "LN(A4)");
        TestFunction(c => ExcelFunctions.Math.Log(c["One"], 2), "LOG(A4,2)");
        TestFunction(c => ExcelFunctions.Math.Quotient(c["One"], c["Two"]), "QUOTIENT(A4,B4)");
        TestFunction(c => ExcelFunctions.Math.Gcd(c["One"], c["Two"]), "GCD(A4,B4)");
    }

    [TestMethod]
    public void ConditionalSums()
    {
        TestFunction(c => ExcelFunctions.Math.SumIf(c.Matrix("One", 2, "One", 9), ">100"),
            "SUMIF(A2:A9,\">100\")");
        TestFunction(
            c => ExcelFunctions.Math.SumIf(c.Matrix("One", 2, "One", 9), ">100", c.Matrix("Two", 2, "Two", 9)),
            "SUMIF(A2:A9,\">100\",B2:B9)");
        TestFunction(
            c => ExcelFunctions.Math.SumIfs(c.Matrix("Two", 2, "Two", 9), c.Matrix("One", 2, "One", 9), ">100"),
            "SUMIFS(B2:B9,A2:A9,\">100\")");
        TestFunction(
            c => ExcelFunctions.Math.SumProduct(c.Matrix("One", 2, "One", 9), c.Matrix("Two", 2, "Two", 9)),
            "SUMPRODUCT(A2:A9,B2:B9)");
    }

    [TestMethod]
    public void Trigonometry()
    {
        TestFunction(c => ExcelFunctions.Math.Sin(c["One"]), "SIN(A4)");
        TestFunction(c => ExcelFunctions.Math.Atan2(c["One"], c["Two"]), "ATAN2(A4,B4)");
        TestFunction(c => ExcelFunctions.Math.Degrees(c["One"]), "DEGREES(A4)");
    }
}
