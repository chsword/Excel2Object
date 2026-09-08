using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     Typed model values are written with the matching Excel cell type: numbers as numeric cells, bools as
///     boolean cells, nulls as blank cells. Strings stay text so leading zeros survive.
/// </summary>
[TestClass]
public class TypedCellTest : BaseExcelTest
{
    public class Item
    {
        [ExcelTitle("Code")] public string Code { get; set; } = "";
        [ExcelTitle("Qty")] public int Qty { get; set; }
        [ExcelTitle("Price")] public decimal? Price { get; set; }
        [ExcelTitle("Active")] public bool Active { get; set; }
    }

    private static IRow FirstDataRow(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        return WorkbookFactory.Create(stream).GetSheetAt(0).GetRow(1);
    }

    [TestMethod]
    public void NumbersBoolsAndNullsGetTheirOwnCellTypes()
    {
        var items = new List<Item>
        {
            new() {Code = "007", Qty = 3, Price = 2.5m, Active = true},
            new() {Code = "008", Qty = 0, Price = null, Active = false}
        };
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var bytes = new ExcelExporter().ObjectToExcelBytes(items, o => o.ExcelType = excelType);
            Assert.IsNotNull(bytes);

            var row = FirstDataRow(bytes);
            Assert.AreEqual(CellType.String, row.GetCell(0).CellType, excelType.ToString());
            Assert.AreEqual("007", row.GetCell(0).StringCellValue);
            Assert.AreEqual(CellType.Numeric, row.GetCell(1).CellType);
            Assert.AreEqual(3d, row.GetCell(1).NumericCellValue);
            Assert.AreEqual(CellType.Numeric, row.GetCell(2).CellType);
            Assert.AreEqual(2.5d, row.GetCell(2).NumericCellValue);
            Assert.AreEqual(CellType.Boolean, row.GetCell(3).CellType);
            Assert.IsTrue(row.GetCell(3).BooleanCellValue);

            // and the round trip through the importer still yields the original values
            var result = ExcelHelper.ExcelToObject<Item>(bytes).ToList();
            Assert.AreEqual("007", result[0].Code);
            Assert.AreEqual(3, result[0].Qty);
            Assert.AreEqual(2.5m, result[0].Price);
            Assert.IsTrue(result[0].Active);
            Assert.IsNull(result[1].Price);
            Assert.IsFalse(result[1].Active);
        }
    }

    [TestMethod]
    public void NullNumberIsBlankNotEmptyText()
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(new List<Item> {new() {Price = null}});
        Assert.IsNotNull(bytes);
        Assert.AreEqual(CellType.Blank, FirstDataRow(bytes).GetCell(2).CellType);
    }

    [TestMethod]
    public void DataTableDbNullIsBlank()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Amount", typeof(double));
        table.Rows.Add("a", 1.25);
        table.Rows.Add("b", System.DBNull.Value);

        var bytes = new ExcelExporter().ObjectToExcelBytes(table, ExcelType.Xlsx);
        Assert.IsNotNull(bytes);
        using var stream = new MemoryStream(bytes);
        var sheet = WorkbookFactory.Create(stream).GetSheetAt(0);
        Assert.AreEqual(CellType.Numeric, sheet.GetRow(1).GetCell(1).CellType);
        Assert.AreEqual(1.25d, sheet.GetRow(1).GetCell(1).NumericCellValue);
        Assert.AreEqual(CellType.Blank, sheet.GetRow(2).GetCell(1).CellType);
    }
}
