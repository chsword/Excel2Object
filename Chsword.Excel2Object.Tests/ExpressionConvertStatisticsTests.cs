using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertStatisticsTests : BaseFunctionTest
{
    [TestMethod]
    public void SumTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.Sum(c.Matrix("One", 1, "Two", 2)), "SUM(A1:B2)");
        TestFunction(
            c => ExcelFunctions.Statistics.Sum(c.Matrix("One", 1, "Two", 2), c.Matrix("Six", 11, "Five", 2)),
            "SUM(A1:B2,F11:E2)");
    }

    [TestMethod]
    public void AverageTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.Average(c.Matrix("One", 1, "Three", 5)), "AVERAGE(A1:C5)");
        TestFunction(
            c => ExcelFunctions.Statistics.Average(c.Matrix("Two", 2, "Four", 4), c.Matrix("Five", 1, "Six", 3)),
            "AVERAGE(B2:D4,E1:F3)");
    }

    [TestMethod]
    public void CountTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.Count(c.Matrix("One", 1, "Two", 10)), "COUNT(A1:B10)");
        TestFunction(c => ExcelFunctions.Statistics.Count(c.Matrix("Three", 5, "Five", 8)), "COUNT(C5:E8)");
    }

    [TestMethod]
    public void CountATest()
    {
        TestFunction(c => ExcelFunctions.Statistics.CountA(c.Matrix("One", 1, "Two", 10)), "COUNTA(A1:B10)");
        TestFunction(c => ExcelFunctions.Statistics.CountA(c.Matrix("Four", 2, "Six", 7)), "COUNTA(D2:F7)");
    }

    [TestMethod]
    public void CountBlankTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.CountBlank(c.Matrix("One", 1, "Three", 5)), "COUNTBLANK(A1:C5)");
        TestFunction(c => ExcelFunctions.Statistics.CountBlank(c.Matrix("Two", 3, "Four", 9)), "COUNTBLANK(B3:D9)");
    }

    [TestMethod]
    public void CountIfTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.CountIf(c.Matrix("One", 1, "One", 10), ">5"), "COUNTIF(A1:A10,\">5\")");
        TestFunction(c => ExcelFunctions.Statistics.CountIf(c.Matrix("Two", 2, "Two", 8), c["Three"]), "COUNTIF(B2:B8,C4)");
    }

    [TestMethod]
    public void CountIfsTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.CountIfs(
            c.Matrix("One", 1, "One", 10), ">5",
            c.Matrix("Two", 1, "Two", 10), "<100"), 
            "COUNTIFS(A1:A10,\">5\",B1:B10,\"<100\")");
    }

    [TestMethod]
    public void MaxTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.Max(c.Matrix("One", 1, "Three", 5)), "MAX(A1:C5)");
        TestFunction(
            c => ExcelFunctions.Statistics.Max(c.Matrix("Two", 2, "Four", 4), c.Matrix("Five", 1, "Six", 3)),
            "MAX(B2:D4,E1:F3)");
    }

    [TestMethod]
    public void MinTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.Min(c.Matrix("One", 1, "Three", 5)), "MIN(A1:C5)");
        TestFunction(
            c => ExcelFunctions.Statistics.Min(c.Matrix("Two", 2, "Four", 4), c.Matrix("Five", 1, "Six", 3)),
            "MIN(B2:D4,E1:F3)");
    }

    [TestMethod]
    public void MedianTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.Median(c.Matrix("One", 1, "Two", 10)), "MEDIAN(A1:B10)");
        TestFunction(c => ExcelFunctions.Statistics.Median(c.Matrix("Three", 2, "Five", 8)), "MEDIAN(C2:E8)");
    }

    [TestMethod]
    public void ModeTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.Mode(c.Matrix("One", 1, "Two", 10)), "MODE(A1:B10)");
        TestFunction(c => ExcelFunctions.Statistics.Mode(c.Matrix("Four", 3, "Six", 7)), "MODE(D3:F7)");
    }

    [TestMethod]
    public void StDevTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.StDev(c.Matrix("One", 1, "Three", 5)), "STDEV(A1:C5)");
        TestFunction(c => ExcelFunctions.Statistics.StDev(c.Matrix("Two", 2, "Four", 6)), "STDEV(B2:D6)");
    }

    [TestMethod]
    public void VarTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.Var(c.Matrix("One", 1, "Three", 5)), "VAR(A1:C5)");
        TestFunction(c => ExcelFunctions.Statistics.Var(c.Matrix("Five", 1, "Six", 4)), "VAR(E1:F4)");
    }

    [TestMethod]
    public void LargeTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.Large(c.Matrix("One", 1, "One", 10), 2), "LARGE(A1:A10,2)");
        TestFunction(c => ExcelFunctions.Statistics.Large(c.Matrix("Two", 3, "Two", 8), c["Three"]), "LARGE(B3:B8,C4)");
    }

    [TestMethod]
    public void SmallTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.Small(c.Matrix("One", 1, "One", 10), 3), "SMALL(A1:A10,3)");
        TestFunction(c => ExcelFunctions.Statistics.Small(c.Matrix("Three", 2, "Three", 9), c["Four"]), "SMALL(C2:C9,D4)");
    }

    [TestMethod]
    public void RankTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.Rank(c["One"], c.Matrix("Two", 1, "Two", 10), 1), "RANK(A4,B1:B10,1)");
        TestFunction(c => ExcelFunctions.Statistics.Rank(85, c.Matrix("One", 1, "One", 20), 0), "RANK(85,A1:A20,0)");
    }

    [TestMethod]
    public void PercentileTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.Percentile(c.Matrix("One", 1, "One", 20), 0.8), "PERCENTILE(A1:A20,0.8)");
        TestFunction(c => ExcelFunctions.Statistics.Percentile(c.Matrix("Two", 5, "Two", 15), c["Three"]), "PERCENTILE(B5:B15,C4)");
    }

    [TestMethod]
    public void QuartileTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.Quartile(c.Matrix("One", 1, "One", 12), 1), "QUARTILE(A1:A12,1)");
        TestFunction(c => ExcelFunctions.Statistics.Quartile(c.Matrix("Three", 2, "Three", 10), c["Four"]), "QUARTILE(C2:C10,D4)");
    }

    [TestMethod]
    public void SumIfTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.SumIf(
            c.Matrix("One", 1, "One", 10), ">5", 
            c.Matrix("Two", 1, "Two", 10)), 
            "SUMIF(A1:A10,\">5\",B1:B10)");
        TestFunction(c => ExcelFunctions.Statistics.SumIf(
            c.Matrix("Three", 2, "Three", 8), c["Four"], 
            c.Matrix("Five", 2, "Five", 8)), 
            "SUMIF(C2:C8,D4,E2:E8)");
    }

    [TestMethod]
    public void SumIfsTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.SumIfs(
            c.Matrix("One", 1, "One", 10),
            c.Matrix("Two", 1, "Two", 10), ">100"), 
            "SUMIFS(A1:A10,B1:B10,\">100\")");
    }

    [TestMethod]
    public void AverageIfTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.AverageIf(
            c.Matrix("One", 1, "One", 10), ">0", 
            c.Matrix("Two", 1, "Two", 10)), 
            "AVERAGEIF(A1:A10,\">0\",B1:B10)");
        TestFunction(c => ExcelFunctions.Statistics.AverageIf(
            c.Matrix("Three", 3, "Three", 12), c["Four"], 
            c.Matrix("Five", 3, "Five", 12)), 
            "AVERAGEIF(C3:C12,D4,E3:E12)");
    }

    [TestMethod]
    public void AverageIfsTest()
    {
        TestFunction(c => ExcelFunctions.Statistics.AverageIfs(
            c.Matrix("One", 1, "One", 10),
            c.Matrix("Two", 1, "Two", 10), ">50"), 
            "AVERAGEIFS(A1:A10,B1:B10,\">50\")");
    }
}