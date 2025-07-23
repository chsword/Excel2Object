using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExpressionConvertDateTimeTests : BaseFunctionTest
    {
        [TestMethod]
        public void DateTest()
        {
            TestFunction(c => ExcelFunctions.DateAndTime.Date(2020, 2, 2), "DATE(2020,2,2)");
            TestFunction(c => ExcelFunctions.DateAndTime.Date(c["One"], c["Two"], c["Three"]), "DATE(A4,B4,C4)");
        }

        [TestMethod]
        public void TimeTest()
        {
            TestFunction(c => ExcelFunctions.DateAndTime.Time(12, 30, 45), "TIME(12,30,45)");
            TestFunction(c => ExcelFunctions.DateAndTime.Time(c["One"], c["Two"], c["Three"]), "TIME(A4,B4,C4)");
        }

        [TestMethod]
        public void DateValueTest()
        {
            TestFunction(c => ExcelFunctions.DateAndTime.DateValue(c["One"]), "DATEVALUE(A4)");
            TestFunction(c => ExcelFunctions.DateAndTime.DateValue("2020-12-31"), "DATEVALUE(\"2020-12-31\")");
        }

        [TestMethod]
        public void TimeValueTest()
        {
            TestFunction(c => ExcelFunctions.DateAndTime.TimeValue(c["One"]), "TIMEVALUE(A4)");
            TestFunction(c => ExcelFunctions.DateAndTime.TimeValue("14:30:45"), "TIMEVALUE(\"14:30:45\")");
        }

        [TestMethod]
        public void EDateTest()
        {
            TestFunction(c => ExcelFunctions.DateAndTime.EDate(c["One"], 6), "EDATE(A4,6)");
            TestFunction(c => ExcelFunctions.DateAndTime.EDate(c["One"], c["Two"]), "EDATE(A4,B4)");
        }

        [TestMethod]
        public void EoMonthTest()
        {
            TestFunction(c => ExcelFunctions.DateAndTime.EoMonth(c["One"], 0), "EOMONTH(A4,0)");
            TestFunction(c => ExcelFunctions.DateAndTime.EoMonth(c["One"], c["Two"]), "EOMONTH(A4,B4)");
        }

        [TestMethod]
        public void NetworkDaysTest()
        {
            TestFunction(c => ExcelFunctions.DateAndTime.NetworkDays(c["One"], c["Two"]), "NETWORKDAYS(A4,B4)");
            TestFunction(c => ExcelFunctions.DateAndTime.NetworkDays(c["One"], c["Two"], c.Matrix("C", 1, "C", 10)), "NETWORKDAYS(A4,B4,C1:C10)");
        }

        [TestMethod]
        public void WorkDayTest()
        {
            TestFunction(c => ExcelFunctions.DateAndTime.WorkDay(c["One"], 30), "WORKDAY(A4,30)");
            TestFunction(c => ExcelFunctions.DateAndTime.WorkDay(c["One"], c["Two"], c.Matrix("C", 1, "C", 10)), "WORKDAY(A4,B4,C1:C10)");
        }

        [TestMethod]
        public void YearFracTest()
        {
            TestFunction(c => ExcelFunctions.DateAndTime.YearFrac(c["One"], c["Two"]), "YEARFRAC(A4,B4)");
            TestFunction(c => ExcelFunctions.DateAndTime.YearFrac(c["One"], c["Two"], 1), "YEARFRAC(A4,B4,1)");
        }

        [TestMethod]
        public void DaysTest()
        {
            TestFunction(c => ExcelFunctions.DateAndTime.Days(c["One"], c["Two"]), "DAYS(A4,B4)");
            TestFunction(c => ExcelFunctions.DateAndTime.Days("2020-12-31", "2020-01-01"), "DAYS(\"2020-12-31\",\"2020-01-01\")");
        }

        [TestMethod]
        public void Days360Test()
        {
            TestFunction(c => ExcelFunctions.DateAndTime.Days360(c["One"], c["Two"]), "DAYS360(A4,B4)");
            TestFunction(c => ExcelFunctions.DateAndTime.Days360(c["One"], c["Two"], c["Three"]), "DAYS360(A4,B4,C4)");
        }
    }
}
