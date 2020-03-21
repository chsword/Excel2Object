using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExpressionConvertReferenceTests : BaseFunctionTest
    {
        [TestMethod]
        public void Lookup()
        {
            TestFunction(c => ExcelFunctions.Reference.Lookup(4.19,c.Matrix("One", 2, "One", 6),
                c.Matrix("Two", 2, "Two", 6)), "LOOKUP(4.19,A2:A6,B2:B6)");
        }

        [TestMethod]
        public void VLookup()
        {
            TestFunction(c => ExcelFunctions.Reference.VLookup(c["One"], c.Matrix("One", 10, "Three", 20),
                2, true), "VLOOKUP(A4,A10:C20,2,TRUE)");
            TestFunction(c => ExcelFunctions.Reference.VLookup("袁", c.Matrix("Two", 2, "Five", 7),
                2, false), "VLOOKUP(\"袁\",B2:E7,2,FALSE)");
            //todo seach over sheet !
            // = VLOOKUP （A2，"客户端详细信息"！A:F，3，FALSE）
        }


    }
}