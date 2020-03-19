using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExpressionConvertStatisticsTests : BaseFunctionTest
    {
        [TestMethod]
        public void Sum()
        {
            TestFunction(c => ExcelFunctions.Statistics.Sum(c.Matrix("One", 1, "Two", 2)), "SUM(A1:B2)");
            TestFunction(c => ExcelFunctions.Statistics.Sum(c.Matrix("One", 1, "Two", 2), c.Matrix("Six", 11, "Five", 2)), "SUM(A1:B2,F11:E2)");

        }
    }
}