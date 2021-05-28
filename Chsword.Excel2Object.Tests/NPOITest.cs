using System;
using Chsword.Excel2Object.Styles;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class NPOITest
    {
        [TestMethod]
        public void CheckHorizontalAlignment()
        {
            var localNames = Enum.GetNames(typeof(HorizontalAlignment));
            var npoiNames = Enum.GetNames(typeof(NPOI.SS.UserModel.HorizontalAlignment));
            Assert.AreEqual(string.Join(",", localNames), string.Join(",", npoiNames));
        }
    }
}