using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertDateTimeTests : BaseFunctionTest
{
    [TestMethod]
    public void PartsOfADate()
    {
        TestFunction(c => ExcelFunctions.DateAndTime.Year(c["One"]), "YEAR(A4)");
        TestFunction(c => ExcelFunctions.DateAndTime.WeekNum(c["One"]), "WEEKNUM(A4)");
        TestFunction(c => ExcelFunctions.DateAndTime.IsoWeekNum(c["One"]), "_xlfn.ISOWEEKNUM(A4)");
        TestFunction(c => ExcelFunctions.DateAndTime.Weekday(c["One"]), "WEEKDAY(A4)");
    }

    [TestMethod]
    public void DateArithmetic()
    {
        TestFunction(c => ExcelFunctions.DateAndTime.EDate(c["One"], 3), "EDATE(A4,3)");
        TestFunction(c => ExcelFunctions.DateAndTime.EoMonth(c["One"], 0), "EOMONTH(A4,0)");
        TestFunction(c => ExcelFunctions.DateAndTime.Days360(c["One"], c["Two"]), "DAYS360(A4,B4)");
        TestFunction(c => ExcelFunctions.DateAndTime.DateDif(c["One"], c["Two"], "Y"), "DATEDIF(A4,B4,\"Y\")");
    }

    [TestMethod]
    public void WorkingDays()
    {
        TestFunction(c => ExcelFunctions.DateAndTime.NetworkDays(c["One"], c["Two"]), "NETWORKDAYS(A4,B4)");
        TestFunction(
            c => ExcelFunctions.DateAndTime.NetworkDays(c["One"], c["Two"], c.Matrix("Six", 1, "Six", 20)),
            "NETWORKDAYS(A4,B4,F1:F20)");
        TestFunction(c => ExcelFunctions.DateAndTime.NetworkDaysIntl(c["One"], c["Two"], 11),
            "_xlfn.NETWORKDAYS.INTL(A4,B4,11)");
        TestFunction(c => ExcelFunctions.DateAndTime.WorkDay(c["One"], 5), "WORKDAY(A4,5)");
    }

    [TestMethod]
    public void ClrDateTimeMembers()
    {
        TestFunction(c => DateTime.Today, "TODAY()");
        TestFunction(c => ((DateTime) c["One"]).AddDays(7), "(A4+7)");
        TestFunction(c => ((DateTime) c["One"]).AddMonths(1), "EDATE(A4,1)");
        TestFunction(c => ((DateTime) c["One"]).Hour, "HOUR(A4)");
    }

    [TestMethod]
    public void DateConstantsBecomeDateCalls()
    {
        TestFunction(c => c["One"] > new DateTime(2024, 3, 1), "A4>DATE(2024,3,1)");
    }
}
