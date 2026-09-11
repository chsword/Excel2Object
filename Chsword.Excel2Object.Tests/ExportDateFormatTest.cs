using System;
using System.Collections.Generic;
using System.IO;
using Chsword.Excel2Object.Options;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExportDateFormatTest : BaseExcelTest
{
    [TestMethod]
    public void ExportDateTest()
    {
        var list = new List<TestModelDatePerson>
        {
            new()
            {
                Age = 18,
                Birthday = DateTime.Now,
                Birthday2 = DateTime.Now,
                Name = "test"
            },
            new()
            {
                Age = 18,
                Birthday = DateTime.Now,

                Name = "test2"
            }
        };
        ExcelHelper.ObjectToExcel(list, GetFilePath(DateTime.Now.Ticks + "test.xls"));
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

    [TestMethod]
    public void MyTestMethod()
    {
        var list = HSSFDataFormat.GetBuiltinFormats();
        foreach (var item in list) Console.WriteLine(item);
    }
}