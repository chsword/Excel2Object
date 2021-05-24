using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExpressionConvertTextTests : BaseFunctionTest
    {
        [TestMethod]
        public void Find()
        {
            TestFunction(c => ExcelFunctions.Text.Find("M",c["One",2]), "FIND(\"M\",A2)");
            TestFunction(c => ExcelFunctions.Text.Find("M", c["One", 2],2), "FIND(\"M\",A2,2)");
        }
    }
}