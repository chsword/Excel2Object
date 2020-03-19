using Microsoft.VisualStudio.TestTools.UnitTesting;
using static Chsword.Excel2Object.ExcelFunctions;
namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExpressionConvertMathFunctionsTests :BaseFunctionTest
    {
        [TestMethod]
        public void AbsTest()
        {
            TestFunction(c => Math.Abs(c["One"]), "ABS(A4)");
        }

        [TestMethod]
        public void PITest()
        {
            TestFunction(c => Math.PI(), "PI()");
        }
    }
}