using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Options;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExportDateFormatTest : BaseExcelTest
{
    /// <summary>
    ///     A [ExcelColumn] Format that Excel has no builtin number format for is applied by writing the
    ///     value as already-formatted text; an empty nullable date stays a blank cell.
    /// </summary>
    [TestMethod]
    public void ExportDateTest()
    {
        var birthday = new DateTime(2026, 9, 11, 14, 30, 45);
        var list = new List<TestModelDatePerson>
        {
            new() {Age = 18, Birthday = birthday, Birthday2 = birthday, Name = "test"},
            new() {Age = 20, Birthday = birthday, Name = "test2"}
        };
        var path = GetFilePath(DateTime.Now.Ticks + "test.xls");
        ExcelHelper.ObjectToExcel(list, path);
        Assert.IsTrue(File.Exists(path));

        using (var stream = File.OpenRead(path))
        {
            var sheet = WorkbookFactory.Create(stream).GetSheetAt(0);
            Assert.AreEqual("出生日期", sheet.GetRow(0).GetCell(2).StringCellValue);

            var firstRow = sheet.GetRow(1);
            Assert.AreEqual(CellType.Numeric, firstRow.GetCell(1).CellType);
            Assert.AreEqual(18d, firstRow.GetCell(1).NumericCellValue);
            Assert.AreEqual(CellType.String, firstRow.GetCell(2).CellType);
            Assert.AreEqual(birthday.ToString("yyyy-MM-dd HH:mm:ss"), firstRow.GetCell(2).StringCellValue);
            Assert.AreEqual(birthday.ToString("yyyy-MM-dd HH:mm:ss"), firstRow.GetCell(3).StringCellValue);

            Assert.AreEqual(CellType.Blank, sheet.GetRow(2).GetCell(3).CellType);
        }

        var result = ExcelHelper.ExcelToObject<TestModelDatePerson>(path)!.ToList();
        Assert.AreEqual(2, result.Count);
        Assert.AreEqual("test", result[0].Name);
        Assert.AreEqual(birthday, result[0].Birthday);
        Assert.AreEqual(birthday, result[0].Birthday2);
        Assert.AreEqual(20, result[1].Age);
        Assert.IsNull(result[1].Birthday2);

        File.Delete(path);
    }

    public class BuiltinFormatModel
    {
        [ExcelColumn("Builtin", Format = "m/d/yy")] public DateTime Builtin { get; set; }
        [ExcelColumn("None")] public DateTime None { get; set; }
    }

    /// <summary>
    ///     Format is applied by rendering the value into text, which only happens for a format Excel has
    ///     no builtin of its own for. One that collides with a builtin name is silently ignored - a
    ///     long-standing trap worth pinning, since "m/d/yy" looks like the most natural thing to write.
    /// </summary>
    [TestMethod]
    public void ABuiltinFormatNameOnADateColumnIsIgnored()
    {
        var when = new DateTime(2026, 9, 11, 14, 30, 45);
        var bytes = new ExcelExporter().ObjectToExcelBytes(
            new List<BuiltinFormatModel> {new() {Builtin = when, None = when}},
            options => options.ExcelType = ExcelType.Xlsx);
        Assert.IsNotNull(bytes);

        using var stream = new MemoryStream(bytes);
        var row = WorkbookFactory.Create(stream).GetSheetAt(0).GetRow(1);
        Assert.AreEqual(row.GetCell(1).StringCellValue, row.GetCell(0).StringCellValue);
        Assert.AreEqual(row.GetCell(1).CellStyle.DataFormat, row.GetCell(0).CellStyle.DataFormat);
    }

    /// <summary>
    ///     A formula column declared to return a DateTime must be written with a date number format,
    ///     otherwise Excel shows the raw serial number (e.g. 46282.6 instead of 2026-09-11).
    /// </summary>
    [TestMethod]
    public void FormulaColumnWithDateResultIsDateFormatted()
    {
        var list = new List<Dictionary<string, object>> {new() {["Name"] = "a"}};
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var bytes = new ExcelExporter().ObjectToExcelBytes(list, options =>
            {
                options.ExcelType = excelType;
                options.FormulaColumns.Add(new FormulaColumn
                {
                    Title = "Now",
                    Formula = c => DateTime.Now,
                    AfterColumnTitle = "Name",
                    FormulaResultType = typeof(DateTime)
                });
            });
            Assert.IsNotNull(bytes);

            using var stream = new MemoryStream(bytes);
            var cell = WorkbookFactory.Create(stream).GetSheetAt(0).GetRow(1).GetCell(1);
            Assert.AreEqual(CellType.Formula, cell.CellType, excelType.ToString());
            Assert.AreEqual(HSSFDataFormat.GetBuiltinFormat("m/d/yy"), cell.CellStyle.DataFormat,
                excelType.ToString());
        }
    }
}