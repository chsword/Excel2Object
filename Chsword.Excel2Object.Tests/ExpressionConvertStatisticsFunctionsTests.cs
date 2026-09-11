using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertStatisticsFunctionsTests : BaseFunctionTest
{
    [TestMethod]
    public void AveragesAndCounts()
    {
        TestFunction(c => ExcelFunctions.Statistics.Average(c.Matrix("One", 2, "One", 9)), "AVERAGE(A2:A9)");
        TestFunction(c => ExcelFunctions.Statistics.Count(c.Matrix("One", 2, "One", 9)), "COUNT(A2:A9)");
        TestFunction(c => ExcelFunctions.Statistics.CountA(c.Matrix("One", 2, "One", 9)), "COUNTA(A2:A9)");
        TestFunction(c => ExcelFunctions.Statistics.CountIf(c.Matrix("One", 2, "One", 9), ">3"),
            "COUNTIF(A2:A9,\">3\")");
        TestFunction(
            c => ExcelFunctions.Statistics.AverageIf(c.Matrix("One", 2, "One", 9), ">3",
                c.Matrix("Two", 2, "Two", 9)),
            "AVERAGEIF(A2:A9,\">3\",B2:B9)");
    }

    [TestMethod]
    public void ExtremesAndOrder()
    {
        TestFunction(c => ExcelFunctions.Statistics.Max(c.Matrix("One", 2, "One", 9)), "MAX(A2:A9)");
        TestFunction(c => ExcelFunctions.Statistics.Min(c["One"], c["Two"]), "MIN(A4,B4)");
        TestFunction(c => ExcelFunctions.Statistics.Large(c.Matrix("One", 2, "One", 9), 2), "LARGE(A2:A9,2)");
        TestFunction(c => ExcelFunctions.Statistics.Rank(c["One"], c.Matrix("One", 2, "One", 9)),
            "RANK(A4,A2:A9)");
        TestFunction(c => ExcelFunctions.Statistics.RankEq(c["One"], c.Matrix("One", 2, "One", 9), 0),
            "_xlfn.RANK.EQ(A4,A2:A9,0)");
    }

    [TestMethod]
    public void SpreadAndDistribution()
    {
        TestFunction(c => ExcelFunctions.Statistics.Median(c.Matrix("One", 2, "One", 9)), "MEDIAN(A2:A9)");
        TestFunction(c => ExcelFunctions.Statistics.StDevP(c.Matrix("One", 2, "One", 9)), "_xlfn.STDEV.P(A2:A9)");
        TestFunction(c => ExcelFunctions.Statistics.VarS(c.Matrix("One", 2, "One", 9)), "_xlfn.VAR.S(A2:A9)");
        TestFunction(
            c => ExcelFunctions.Statistics.PercentileInc(c.Matrix("One", 2, "One", 9), 0.9), 
            "_xlfn.PERCENTILE.INC(A2:A9,0.9)");
        TestFunction(c => ExcelFunctions.Statistics.NormSInv(0.95), "_xlfn.NORM.S.INV(0.95)");
    }

    [TestMethod]
    public void SumStillAcceptsRangesAndValues()
    {
        TestFunction(c => ExcelFunctions.Statistics.Sum(c.Matrix("One", 1, "Two", 2)), "SUM(A1:B2)");
        TestFunction(c => ExcelFunctions.Statistics.Sum(c["One"], c.Matrix("Two", 1, "Two", 9)),
            "SUM(A4,B1:B9)");
    }
}
