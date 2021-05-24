﻿using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using Chsword.Excel2Object.Functions;
using Chsword.Excel2Object.Internal;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static Chsword.Excel2Object.ExcelFunctions;
namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExpressionConvertTests:BaseFunctionTest
    {
        #region Basic Functions

        [TestMethod]
        public void Column()
        {
            TestFunction(c => c["One"],"A4");
        }
        [TestMethod]
        public void ColumnWithRow()
        {
            TestFunction(c => c["One",1], "A1");
        }
        #endregion

        #region Date Functions

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
        public void EDate()
        {
            Expression<Func<Dictionary<string, object>, object>> exp = c => DateTime.Now.AddMonths(3);
            var convert = new ExpressionConvert(new string[] { }, 3);
            var ret = convert.Convert(exp);
            Assert.AreEqual("EDATE(NOW(),3)", ret);
        }
        [TestMethod]
        public void EDateWithColumn()
        {
            Expression<Func<Dictionary<string, ColumnValue>, object>> exp = c => ((DateTime)c["Date"]).AddMonths((int)c["Month"]);
            var convert = new ExpressionConvert(new string[] {"Date", "Month" }, 3);
            var ret = convert.Convert(exp);
            Assert.AreEqual("EDATE(A4,B4)", ret);
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
        public void Date()
        {
            TestFunction(c=>ExcelFunctions.DateAndTime.Date(2020,2,2),"DATE(2020,2,2)");
        }
        [TestMethod]
        public void DateDif()
        {
            TestFunction(c => ExcelFunctions.DateAndTime.DateDif(c["One"],c["Two"],"YD"), "DATEDIF(A4,B4,\"YD\")");
        }

        [TestMethod]
        public void Days()
        {
            TestFunction(c => ExcelFunctions.DateAndTime.Days(c["One"], c["Two"]), "DAYS(A4,B4)");
        }
        
        #endregion

    }
}