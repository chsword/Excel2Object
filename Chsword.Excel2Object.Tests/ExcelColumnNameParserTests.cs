using Chsword.Excel2Object.Internal;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExcelColumnNameParserTests
    {
        [TestMethod]
        public void Parse()
        {
            AssertHelper("A", 0);
            AssertHelper("B", 1);
            AssertHelper("Z", 25);
            AssertHelper("AA", 26);
            AssertHelper("DL", 115);
            AssertHelper("ACM", 766);
        }

        private void AssertHelper(string v1, int v2)
        {
            var ret = ExcelColumnNameParser.Parse(v2);
            Assert.AreEqual(v1, ret);
        }
    }
}