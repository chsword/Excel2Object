using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertMathFunctionsTests : BaseFunctionTest
{
    [TestMethod]
    public void AbsTest()
    {
        TestFunction(c => ExcelFunctions.Math.Abs(c["One"]), "ABS(A4)");
        TestFunction(c => ExcelFunctions.Math.Abs(c["Two", 5]), "ABS(B5)");
    }

    [TestMethod]
    public void PITest()
    {
        TestFunction(c => ExcelFunctions.Math.PI(), "PI()");
    }

    [TestMethod]
    public void EvenTest()
    {
        TestFunction(c => ExcelFunctions.Math.Even(c["One"]), "EVEN(A4)");
        TestFunction(c => ExcelFunctions.Math.Even(c["Three", 2]), "EVEN(C2)");
    }

    [TestMethod]
    public void OddTest()
    {
        TestFunction(c => ExcelFunctions.Math.Odd(c["One"]), "ODD(A4)");
        TestFunction(c => ExcelFunctions.Math.Odd(c["Two", 1]), "ODD(B1)");
    }

    [TestMethod]
    public void FactTest()
    {
        TestFunction(c => ExcelFunctions.Math.Fact(c["One"]), "FACT(A4)");
        TestFunction(c => ExcelFunctions.Math.Fact(c["Four", 6]), "FACT(D6)");
    }

    [TestMethod]
    public void IntTest()
    {
        TestFunction(c => ExcelFunctions.Math.Int(c["One"]), "INT(A4)");
        TestFunction(c => ExcelFunctions.Math.Int(c["Five", 3]), "INT(E3)");
    }

    [TestMethod]
    public void RandTest()
    {
        TestFunction(c => ExcelFunctions.Math.Rand(), "RAND()");
    }

    [TestMethod]
    public void RandBetweenTest()
    {
        TestFunction(c => ExcelFunctions.Math.RandBetween(c["One"], c["Two"]), "RANDBETWEEN(A4,B4)");
        TestFunction(c => ExcelFunctions.Math.RandBetween(1, 100), "RANDBETWEEN(1,100)");
    }

    [TestMethod]
    public void RoundTest()
    {
        TestFunction(c => ExcelFunctions.Math.Round(c["One"], c["Two"]), "ROUND(A4,B4)");
        TestFunction(c => ExcelFunctions.Math.Round(c["Three", 5], 2), "ROUND(C5,2)");
    }

    [TestMethod]
    public void RoundDownTest()
    {
        TestFunction(c => ExcelFunctions.Math.RoundDown(c["One"], c["Two"]), "ROUNDDOWN(A4,B4)");
        TestFunction(c => ExcelFunctions.Math.RoundDown(c["One"], 1), "ROUNDDOWN(A4,1)");
    }

    [TestMethod]
    public void RoundUpTest()
    {
        TestFunction(c => ExcelFunctions.Math.RoundUp(c["One"], c["Two"]), "ROUNDUP(A4,B4)");
        TestFunction(c => ExcelFunctions.Math.RoundUp(c["One"], 0), "ROUNDUP(A4,0)");
    }

    [TestMethod]
    public void SqrtTest()
    {
        TestFunction(c => ExcelFunctions.Math.Sqrt(c["One"]), "SQRT(A4)");
        TestFunction(c => ExcelFunctions.Math.Sqrt(c["Six", 2]), "SQRT(F2)");
    }

    [TestMethod]
    public void PowerTest()
    {
        TestFunction(c => ExcelFunctions.Math.Power(c["One"], c["Two"]), "POWER(A4,B4)");
        TestFunction(c => ExcelFunctions.Math.Power(c["One"], 2), "POWER(A4,2)");
    }

    [TestMethod]
    public void ExpTest()
    {
        TestFunction(c => ExcelFunctions.Math.Exp(c["One"]), "EXP(A4)");
        TestFunction(c => ExcelFunctions.Math.Exp(c["Two", 1]), "EXP(B1)");
    }

    [TestMethod]
    public void LnTest()
    {
        TestFunction(c => ExcelFunctions.Math.Ln(c["One"]), "LN(A4)");
        TestFunction(c => ExcelFunctions.Math.Ln(c["Three", 7]), "LN(C7)");
    }

    [TestMethod]
    public void LogTest()
    {
        TestFunction(c => ExcelFunctions.Math.Log(c["One"], c["Two"]), "LOG(A4,B4)");
        TestFunction(c => ExcelFunctions.Math.Log(c["One"], 10), "LOG(A4,10)");
    }

    [TestMethod]
    public void Log10Test()
    {
        TestFunction(c => ExcelFunctions.Math.Log10(c["One"]), "LOG10(A4)");
        TestFunction(c => ExcelFunctions.Math.Log10(c["Four", 3]), "LOG10(D3)");
    }

    [TestMethod]
    public void SinTest()
    {
        TestFunction(c => ExcelFunctions.Math.Sin(c["One"]), "SIN(A4)");
        TestFunction(c => ExcelFunctions.Math.Sin(c["Two", 8]), "SIN(B8)");
    }

    [TestMethod]
    public void CosTest()
    {
        TestFunction(c => ExcelFunctions.Math.Cos(c["One"]), "COS(A4)");
        TestFunction(c => ExcelFunctions.Math.Cos(c["Five", 1]), "COS(E1)");
    }

    [TestMethod]
    public void TanTest()
    {
        TestFunction(c => ExcelFunctions.Math.Tan(c["One"]), "TAN(A4)");
        TestFunction(c => ExcelFunctions.Math.Tan(c["Six", 9]), "TAN(F9)");
    }

    [TestMethod]
    public void ModTest()
    {
        TestFunction(c => ExcelFunctions.Math.Mod(c["One"], c["Two"]), "MOD(A4,B4)");
        TestFunction(c => ExcelFunctions.Math.Mod(c["One"], 3), "MOD(A4,3)");
    }

    [TestMethod]
    public void GcdTest()
    {
        TestFunction(c => ExcelFunctions.Math.Gcd(c["One"], c["Two"]), "GCD(A4,B4)");
        TestFunction(c => ExcelFunctions.Math.Gcd(24, 36), "GCD(24,36)");
    }

    [TestMethod]
    public void LcmTest()
    {
        TestFunction(c => ExcelFunctions.Math.Lcm(c["One"], c["Two"]), "LCM(A4,B4)");
        TestFunction(c => ExcelFunctions.Math.Lcm(4, 6), "LCM(4,6)");
    }

    [TestMethod]
    public void SignTest()
    {
        TestFunction(c => ExcelFunctions.Math.Sign(c["One"]), "SIGN(A4)");
        TestFunction(c => ExcelFunctions.Math.Sign(c["Three", 2]), "SIGN(C2)");
    }

    [TestMethod]
    public void TruncTest()
    {
        TestFunction(c => ExcelFunctions.Math.Trunc(c["One"], c["Two"]), "TRUNC(A4,B4)");
        TestFunction(c => ExcelFunctions.Math.Trunc(c["One"], 2), "TRUNC(A4,2)");
    }
}