using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Options;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     A column can limit its cells to a list of values; Excel shows that list as a dropdown.
/// </summary>
[TestClass]
public class DropdownTest : BaseExcelTest
{
    private const string ListSheetName = "_excel2object_lists";

    public class Model
    {
        [ExcelTitle("城市")] public string City { get; set; } = "北京";

        [ExcelColumn("状态", Dropdown = new[] {"启用", "停用", "待审"})]
        public string Status { get; set; } = "启用";
    }

    private static readonly List<Model> Rows = new() {new(), new()};

    private static IWorkbook Export(ExcelType excelType, Action<ExcelExporterOptions>? configure = null,
        List<Model>? rows = null)
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(rows ?? Rows, options =>
        {
            options.ExcelType = excelType;
            configure?.Invoke(options);
        });
        Assert.IsNotNull(bytes);
        return WorkbookFactory.Create(new MemoryStream(bytes));
    }

    [TestMethod]
    public void TheAttributesValuesBecomeADropdownInBothFileFormats()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var sheet = Export(excelType).GetSheetAt(0);
            var validations = sheet.GetDataValidations();
            Assert.AreEqual(1, validations.Count, excelType.ToString());

            var validation = validations[0];
            CollectionAssert.AreEqual(new[] {"启用", "停用", "待审"},
                validation.ValidationConstraint.ExplicitListValues, excelType.ToString());
            // the header keeps its own value, the dropdown belongs to the rows under it
            Assert.AreEqual("B2:B3", validation.Regions.CellRangeAddresses[0].FormatAsString(),
                excelType.ToString());
        }
    }

    /// <summary>Values only known at runtime come through the options, and win over the attribute.</summary>
    [TestMethod]
    public void OptionsDropdownsOverrideTheAttribute()
    {
        var runtime = new[] {"甲", "乙"};
        var sheet = Export(ExcelType.Xlsx, options => options.Dropdowns["状态"] = runtime).GetSheetAt(0);

        var validation = sheet.GetDataValidations().Single();
        CollectionAssert.AreEqual(runtime, validation.ValidationConstraint.ExplicitListValues);
    }

    /// <summary>A column that declares none has no validation of its own.</summary>
    [TestMethod]
    public void AColumnWithoutValuesGetsNoDropdown()
    {
        var sheet = Export(ExcelType.Xlsx).GetSheetAt(0);
        Assert.AreEqual(1, sheet.GetDataValidations().Count);
        Assert.AreEqual(1, sheet.GetDataValidations()[0].Regions.CellRangeAddresses[0].FirstColumn);
    }

    /// <summary>An empty export is a template, and a template needs its dropdown on the first row.</summary>
    [TestMethod]
    public void AnEmptySheetStillGetsTheDropdown()
    {
        var sheet = Export(ExcelType.Xlsx, rows: new List<Model>()).GetSheetAt(0);
        Assert.AreEqual("B2", sheet.GetDataValidations().Single().Regions.CellRangeAddresses[0].FormatAsString());
    }

    /// <summary>
    ///     Excel holds at most 255 characters inline, and splits the list on commas. A longer list, or one
    ///     with a comma in it, goes to a hidden sheet the dropdown reads from - xls would otherwise throw.
    /// </summary>
    [TestMethod]
    public void ALongListMovesToAHiddenSheet()
    {
        var many = Enumerable.Range(0, 60).Select(i => $"取值{i:D3}").ToArray();
        Assert.IsTrue(string.Join(",", many).Length > 255);

        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var workbook = Export(excelType, options => options.Dropdowns["状态"] = many);
            var validation = workbook.GetSheetAt(0).GetDataValidations().Single();
            Assert.AreEqual($"{ListSheetName}!$A$1:$A$60", validation.ValidationConstraint.Formula1,
                excelType.ToString());

            var listSheet = workbook.GetSheet(ListSheetName);
            Assert.AreEqual(SheetVisibility.Hidden, workbook.GetSheetVisibility(workbook.GetSheetIndex(listSheet)),
                excelType.ToString());
            Assert.AreEqual("取值000", listSheet.GetRow(0).GetCell(0).StringCellValue, excelType.ToString());
            Assert.AreEqual("取值059", listSheet.GetRow(59).GetCell(0).StringCellValue, excelType.ToString());
        }
    }

    /// <summary>A value carrying a comma cannot go inline either - Excel would read it as two values.</summary>
    [TestMethod]
    public void AValueWithACommaMovesToAHiddenSheet()
    {
        var values = new[] {"甲, 乙", "丙"};
        var workbook = Export(ExcelType.Xlsx, options => options.Dropdowns["状态"] = values);

        Assert.AreEqual($"{ListSheetName}!$A$1:$A$2",
            workbook.GetSheetAt(0).GetDataValidations().Single().ValidationConstraint.Formula1);
        Assert.AreEqual("甲, 乙", workbook.GetSheet(ListSheetName).GetRow(0).GetCell(0).StringCellValue);
    }

    /// <summary>Two long lists share the hidden sheet, each in its own column.</summary>
    [TestMethod]
    public void EachLongListGetsItsOwnColumnOnTheHiddenSheet()
    {
        var first = Enumerable.Range(0, 60).Select(i => $"甲{i:D3}").ToArray();
        var second = Enumerable.Range(0, 70).Select(i => $"乙{i:D3}").ToArray();

        var workbook = Export(ExcelType.Xlsx, options =>
        {
            options.Dropdowns["城市"] = first;
            options.Dropdowns["状态"] = second;
        });

        var formulas = workbook.GetSheetAt(0).GetDataValidations()
            .Select(v => v.ValidationConstraint.Formula1).ToArray();
        CollectionAssert.AreEquivalent(
            new[] {$"{ListSheetName}!$A$1:$A$60", $"{ListSheetName}!$B$1:$B$70"}, formulas);

        var listSheet = workbook.GetSheet(ListSheetName);
        Assert.AreEqual("甲000", listSheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.AreEqual("乙000", listSheet.GetRow(0).GetCell(1).StringCellValue);
        // the shorter list ends where it ends; the longer one keeps going on its own
        Assert.IsNull(listSheet.GetRow(65).GetCell(0));
        Assert.AreEqual("乙065", listSheet.GetRow(65).GetCell(1).StringCellValue);
    }

    /// <summary>The values a dropdown allows still import as the values they are.</summary>
    [TestMethod]
    public void ADropdownColumnRoundTrips()
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(Rows, options => options.ExcelType = ExcelType.Xlsx)!;
        var back = ExcelHelper.ExcelToObject<Model>(bytes).ToArray();

        Assert.AreEqual(2, back.Length);
        Assert.AreEqual("启用", back[0].Status);
    }
}
