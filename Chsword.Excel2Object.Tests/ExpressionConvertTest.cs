using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using Chsword.Excel2Object.Internal;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExpressionConvertTests
    {
        [TestMethod]
        public void Now()
        {
            Expression<Func<Dictionary<string, object>, object>> exp = c => DateTime.Now;
            var convert = new ExpressionConvert(new string[]{ }, 3);
            var ret = convert.Convert(exp);
            Assert.AreEqual("NOW()", ret);
        }
        [TestMethod]
        public void Year()
        {
            Expression<Func<Dictionary<string, object>, object>> exp = c => DateTime.Now.Year;
            var convert = new ExpressionConvert(new string[] { }, 3);
            var ret = convert.Convert(exp);
            Assert.AreEqual("YEAR(NOW())",ret);
        }
        [TestMethod]
        public void Month()
        {
            Expression<Func<Dictionary<string, object>, object>> exp = c => DateTime.Now.Month;
            var convert = new ExpressionConvert(new string[] { }, 3);
            var ret = convert.Convert(exp);
            Assert.AreEqual("MONTH(NOW())", ret);
        }
        [TestMethod]
        public void Day()
        {
            Expression<Func<Dictionary<string, object>, object>> exp = c => DateTime.Now.Day;
            var convert = new ExpressionConvert(new string[] { }, 3);
            var ret = convert.Convert(exp);
            Assert.AreEqual("DAY(NOW())", ret);
        }
        [TestMethod]
        public void Column()
        {
            Expression<Func<Dictionary<string, object>, object>> exp = c => c["One"];
            var convert = new ExpressionConvert(new string[] { "One" }, 3);
            var ret = convert.Convert(exp);
            Assert.AreEqual("A4", ret);
        }
    }
}