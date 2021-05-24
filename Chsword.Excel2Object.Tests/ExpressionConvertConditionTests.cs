using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExpressionConvertConditionTests : BaseFunctionTest
    {
        [TestMethod]
        public void If()
        {
            TestFunction(c => ExcelFunctions.Condition.If(c["One"]=="Yes",1,2), "IF(A4=\"Yes\",1,2)");
        }
    }
}