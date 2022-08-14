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
}