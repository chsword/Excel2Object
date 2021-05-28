using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExpressionConvertReferenceTests : BaseFunctionTest
    {
        [TestMethod]
        public void Choose()
        {
            TestFunction(c => ExcelFunctions.Reference.Choose(2, c["One", 2], c["One", 3]
                , c["One", 4], c["One", 5]), "CHOOSE(2,A2,A3,A4,A5)");
        }

        [TestMethod]
        public void Index()
        {
            TestFunction(c => ExcelFunctions.Reference.Index(c.Matrix("One", 2, "Two", 6),
                2, 3), "INDEX(A2:B6,2,3)");
        }

        [TestMethod]
        public void Lookup()
        {
            TestFunction(c => ExcelFunctions.Reference.Lookup(4.19, c.Matrix("One", 2, "One", 6),
                c.Matrix("Two", 2, "Two", 6)), "LOOKUP(4.19,A2:A6,B2:B6)");
        }

        [TestMethod]
        public void Match()
        {
            TestFunction(c => ExcelFunctions.Reference.Match(39, c.Matrix("Two", 2, "Two", 5),
                1), "MATCH(39,B2:B5,1)");
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