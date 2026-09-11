using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertDatabaseTests : BaseFunctionTest
{
    [TestMethod]
    public void DatabaseFunctionsTakeTableFieldAndCriteria()
    {
        TestFunction(
            c => ExcelFunctions.Database.DSum(c.Matrix("One", 1, "Three", 9), "Two",
                c.Matrix("Five", 1, "Six", 2)),
            "DSUM(A1:C9,\"Two\",E1:F2)");
        TestFunction(
            c => ExcelFunctions.Database.DGet(c.Matrix("One", 1, "Three", 9), "Two",
                c.Matrix("Five", 1, "Six", 2)),
            "DGET(A1:C9,\"Two\",E1:F2)");
    }
}
