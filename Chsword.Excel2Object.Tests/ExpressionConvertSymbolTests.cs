using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExpressionConvertSymbolTests : BaseFunctionTest
    {
        [TestMethod]
        public void Addition()
        {
            TestFunction(c => c["One"] + c["Two"], "A4+B4");
            TestFunction(c => c["One"] + 1, "A4+1");
            TestFunction(c => 11 + c["Two"], "11+B4");
            TestFunction(c => 1 + 2, "3");
        }

        [TestMethod]
        public void Colon()
        {
            TestFunction(c => c.Matrix("One", 1, "Two", 2), "A1:B2");
        }

        [TestMethod]
        public void Division()
        {
            TestFunction(c => c["One"] / c["Two"], "A4/B4");
            TestFunction(c => c["One"] / 1, "A4/1");
            TestFunction(c => 11 / c["Two"], "11/B4");
            TestFunction(c => 1 / 2, "0");
        }

        [TestMethod]
        public void Equal()
        {
            TestFunction(c => c["One"] == c["Two"], "A4=B4");
            TestFunction(c => c["One"] == 1, "A4=1");
            TestFunction(c => 11 == c["Two"], "11=B4");
            // TestFunction(c => 1 == 2, "");
        }

        [TestMethod]
        public void GreaterThan()
        {
            TestFunction(c => c["One"] > c["Two"], "A4>B4");
            TestFunction(c => c["One"] > 1, "A4>1");
            TestFunction(c => 11 > c["Two"], "11>B4");
            // TestFunction(c => 1 == 2, "");
        }

        [TestMethod]
        public void GreaterThanOrEqual()
        {
            TestFunction(c => c["One"] >= c["Two"], "A4>=B4");
            TestFunction(c => c["One"] >= 1, "A4>=1");
            TestFunction(c => 11 >= c["Two"], "11>=B4");
            // TestFunction(c => 1 == 2, "");
        }

        [TestMethod]
        public void LessThan()
        {
            TestFunction(c => c["One"] < c["Two"], "A4<B4");
            TestFunction(c => c["One"] < 1, "A4<1");
            TestFunction(c => 11 < c["Two"], "11<B4");
            // TestFunction(c => 1 == 2, "");
        }

        [TestMethod]
        public void LessThanOrEqual()
        {
            TestFunction(c => c["One"] <= c["Two"], "A4<=B4");
            TestFunction(c => c["One"] <= 1, "A4<=1");
            TestFunction(c => 11 <= c["Two"], "11<=B4");
            // TestFunction(c => 1 == 2, "");
        }

        [TestMethod]
        public void Multiplication()
        {
            TestFunction(c => c["One"] * c["Two"], "A4*B4");
            TestFunction(c => c["One"] * 1, "A4*1");
            TestFunction(c => 11 * c["Two"], "11*B4");
            TestFunction(c => 1 * 2, "2");
        }

        [TestMethod]
        public void Negative()
        {
            TestFunction(c => -c["Two"], "-B4");
            TestFunction(c => -2, "-2");
        }

        [TestMethod]
        public void NotEqual()
        {
            TestFunction(c => c["One"] != c["Two"], "A4<>B4");
            TestFunction(c => c["One"] != 1, "A4<>1");
            TestFunction(c => 11 != c["Two"], "11<>B4");
            // TestFunction(c => 1 == 2, "");
        }

        [TestMethod]
        public void StringJoin()
        {
            TestFunction(c => c["One"] & c["Two"], "A4&B4");
            TestFunction(c => c["One"] & "1", "A4&\"1\"");
            TestFunction(c => "11" & c["Two"], "\"11\"&B4");
        }

        [TestMethod]
        public void Subtraction()
        {
            TestFunction(c => c["One"] - c["Two"], "A4-B4");
            TestFunction(c => c["One"] - 1, "A4-1");
            TestFunction(c => 11 - c["Two"], "11-B4");
            TestFunction(c => 1 - 2, "-1");
        }
    }
}