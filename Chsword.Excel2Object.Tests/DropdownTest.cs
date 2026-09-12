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
            Assert.AreEqual("_excel2object_list1", validation.ValidationConstraint.Formula1, excelType.ToString());
            Assert.AreEqual($"{ListSheetName}!$A$1:$A$60",
                workbook.GetName("_excel2object_list1").RefersToFormula, excelType.ToString());

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
            workbook.GetName(workbook.GetSheetAt(0).GetDataValidations().Single().ValidationConstraint.Formula1)
                .RefersToFormula);
        Assert.AreEqual("甲, 乙", workbook.GetSheet(ListSheetName).GetRow(0).GetCell(0).StringCellValue);
    }

    /// <summary>A quote cannot go inline either - Excel ends the list literal at it.</summary>
    [TestMethod]
    public void AValueWithAQuoteMovesToAHiddenSheet()
    {
        var values = new[] {"14\" 屏", "15 寸"};
        var workbook = Export(ExcelType.Xlsx, options => options.Dropdowns["状态"] = values);

        Assert.AreEqual($"{ListSheetName}!$A$1:$A$2",
            workbook.GetName(workbook.GetSheetAt(0).GetDataValidations().Single().ValidationConstraint.Formula1)
                .RefersToFormula);
        Assert.AreEqual("14\" 屏", workbook.GetSheet(ListSheetName).GetRow(0).GetCell(0).StringCellValue);
    }

    /// <summary>
    ///     A sheet of that name the workbook already owns is left alone - the values go to a free name
    ///     instead of into someone else's data.
    /// </summary>
    [TestMethod]
    public void AVisibleSheetOfTheSameNameIsNotWrittenInto()
    {
        var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
        var mine = workbook.CreateSheet(ListSheetName);
        mine.CreateRow(0).CreateCell(0).SetCellValue("我的数据");
        var stream = new MemoryStream();
        workbook.Write(stream, true);

        var many = Enumerable.Range(0, 60).Select(i => $"取值{i:D3}").ToArray();
        var bytes = ExcelHelper.AppendObjectToExcelBytes(stream.ToArray(), Rows, options =>
        {
            options.SheetTitle = "数据";
            options.Dropdowns["状态"] = many;
        })!;

        var back = WorkbookFactory.Create(new MemoryStream(bytes));
        Assert.AreEqual($"{ListSheetName}2!$A$1:$A$60",
            back.GetName(back.GetSheet("数据").GetDataValidations().Single().ValidationConstraint.Formula1)
                .RefersToFormula);
        Assert.AreEqual("我的数据", back.GetSheet(ListSheetName).GetRow(0).GetCell(0).StringCellValue);
        Assert.AreEqual("取值000", back.GetSheet(ListSheetName + "2").GetRow(0).GetCell(0).StringCellValue);
    }

    /// <summary>The freeze and the filter reach an appended sheet through the same options.</summary>
    [TestMethod]
    public void AppendTakesOptions()
    {
        var first = new ExcelExporter().ObjectToExcelBytes(Rows, options => options.ExcelType = ExcelType.Xlsx)!;
        var bytes = ExcelHelper.AppendObjectToExcelBytes(first, Rows, options =>
        {
            options.SheetTitle = "第二页";
            options.FreezeHeader = true;
        })!;

        var workbook = WorkbookFactory.Create(new MemoryStream(bytes));
        Assert.IsNotNull(workbook.GetSheet("第二页").PaneInformation);
        Assert.IsNull(workbook.GetSheetAt(0).PaneInformation);
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

        var ranges = workbook.GetSheetAt(0).GetDataValidations()
            .Select(v => workbook.GetName(v.ValidationConstraint.Formula1).RefersToFormula).ToArray();
        CollectionAssert.AreEquivalent(
            new[] {$"{ListSheetName}!$A$1:$A$60", $"{ListSheetName}!$B$1:$B$70"}, ranges);

        var listSheet = workbook.GetSheet(ListSheetName);
        Assert.AreEqual("甲000", listSheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.AreEqual("乙000", listSheet.GetRow(0).GetCell(1).StringCellValue);
        // the shorter list ends where it ends; the longer one keeps going on its own
        Assert.IsNull(listSheet.GetRow(65).GetCell(0));
        Assert.AreEqual("乙065", listSheet.GetRow(65).GetCell(1).StringCellValue);
    }

    /// <summary>
    ///     Excel counts the quotes it wraps the inline list in, so the last list that still fits is two
    ///     characters shorter than the limit itself.
    /// </summary>
    [DataTestMethod]
    [DataRow(253, true)]
    [DataRow(254, false)]
    public void TheInlineListStopsAtWhatExcelHolds(int joinedLength, bool inline)
    {
        // two values, so the join adds one comma
        var first = new string('a', joinedLength / 2);
        var second = new string('b', joinedLength - 1 - first.Length);
        var values = new[] {first, second};
        Assert.AreEqual(joinedLength, string.Join(",", values).Length);

        var validation = Export(ExcelType.Xlsx, options => options.Dropdowns["状态"] = values)
            .GetSheetAt(0).GetDataValidations().Single();

        if (inline)
            CollectionAssert.AreEqual(values, validation.ValidationConstraint.ExplicitListValues);
        else
            Assert.AreEqual("_excel2object_list1", validation.ValidationConstraint.Formula1);
    }

    /// <summary>A list Excel cannot show an empty entry in should say so, not throw a NullReference.</summary>
    [TestMethod]
    public void ANullValueIsRejectedByName()
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options => options.Dropdowns["状态"] = new[] {"启用", null!}));

        StringAssert.Contains(e.Message, "状态");
        StringAssert.Contains(e.Message, "index 1");
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
