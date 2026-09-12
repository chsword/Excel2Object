using System.Collections.Generic;
using System.IO;
using Chsword.Excel2Object.Options;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     A long sheet is only usable when its header stays in view and can filter the rows below it.
/// </summary>
[TestClass]
public class HeaderViewTest : BaseExcelTest
{
    public class Model
    {
        [ExcelTitle("城市")] public string City { get; set; } = "";
        [ExcelTitle("人数")] public int Count { get; set; }
    }

    private static readonly List<Model> Rows = new()
    {
        new() {City = "北京", Count = 3},
        new() {City = "上海", Count = 5}
    };

    private static IWorkbook Export(ExcelType excelType, System.Action<ExcelExporterOptions> configure,
        List<Model>? rows = null)
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(rows ?? Rows, options =>
        {
            options.ExcelType = excelType;
            configure(options);
        });
        Assert.IsNotNull(bytes);
        return WorkbookFactory.Create(new MemoryStream(bytes));
    }

    [TestMethod]
    public void TheHeaderRowIsFrozenInBothFileFormats()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var pane = Export(excelType, options => options.FreezeHeader = true).GetSheetAt(0).PaneInformation;
            Assert.IsNotNull(pane, excelType.ToString());
            Assert.AreEqual(1, pane.HorizontalSplitPosition, excelType.ToString());
            Assert.AreEqual(1, pane.HorizontalSplitTopRow, excelType.ToString());
            Assert.AreEqual(0, pane.VerticalSplitPosition, excelType.ToString());
        }
    }

    /// <summary>The filter covers the header and every row written under it.</summary>
    [TestMethod]
    public void TheHeaderRowCarriesTheFilterOverTheData()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var workbook = Export(excelType, options => options.AutoFilter = true);
            Assert.AreEqual(1, workbook.NumberOfNames, excelType.ToString());
            Assert.AreEqual("Sheet1!$A$1:$B$3", workbook.GetNameAt(0).RefersToFormula, excelType.ToString());
        }
    }

    /// <summary>Without rows the dropdowns still belong on the header.</summary>
    [TestMethod]
    public void AnEmptySheetStillGetsItsFilter()
    {
        var workbook = Export(ExcelType.Xlsx, options => options.AutoFilter = true, new List<Model>());
        Assert.AreEqual("Sheet1!$A$1:$B$1", workbook.GetNameAt(0).RefersToFormula);
    }

    [TestMethod]
    public void NeitherIsOnByDefault()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var workbook = Export(excelType, _ => { });
            Assert.IsNull(workbook.GetSheetAt(0).PaneInformation, excelType.ToString());
            Assert.AreEqual(0, workbook.NumberOfNames, excelType.ToString());
        }
    }

    /// <summary>An appended sheet gets the same treatment, and the sheet already there is left alone.</summary>
    [TestMethod]
    public void AnAppendedSheetIsFrozenAndFilteredOnItsOwn()
    {
        var first = new ExcelExporter().ObjectToExcelBytes(Rows, options => options.ExcelType = ExcelType.Xlsx)!;
        var bytes = new ExcelExporter().ObjectToExcelBytes(Rows, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.SourceExcelBytes = first;
            options.SheetTitle = "第二页";
            options.FreezeHeader = true;
            options.AutoFilter = true;
        })!;

        var workbook = WorkbookFactory.Create(new MemoryStream(bytes));
        Assert.IsNotNull(workbook.GetSheet("第二页").PaneInformation);
        Assert.IsNull(workbook.GetSheetAt(0).PaneInformation);
        Assert.AreEqual("第二页!$A$1:$B$3", workbook.GetNameAt(0).RefersToFormula);
    }
}
