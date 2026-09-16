using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Options;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     导入前的两件事：模型上的列在表头里有没有，以及表头本身是什么。
/// </summary>
[TestClass]
public class ImportDiagnosticsTest
{
    public class Order
    {
        [ExcelTitle("订单号")] public string No { get; set; } = "";
        [ExcelTitle("金额")] public decimal Amount { get; set; }
        [ExcelTitle("备注")] public string? Memo { get; set; }
    }

    /// <summary>表头由调用方写定，好把「对不上」这件事摆出来。</summary>
    private static byte[] Workbook(string[] titles, string[][] rows, string sheetTitle = "数据")
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet(sheetTitle);
        var header = sheet.CreateRow(0);
        for (var i = 0; i < titles.Length; i++) header.CreateCell(i).SetCellValue(titles[i]);
        for (var r = 0; r < rows.Length; r++)
        {
            var row = sheet.CreateRow(r + 1);
            for (var i = 0; i < rows[r].Length; i++) row.CreateCell(i).SetCellValue(rows[r][i]);
        }

        using var bytes = new MemoryStream();
        workbook.Write(bytes, true);
        return bytes.ToArray();
    }

    private static byte[] Sample()
    {
        // 「金额」在表里叫「金额（元）」——导入不会报错，只是那一列悄悄全为空
        return Workbook(new[] {"订单号", "金额（元）", "备注"},
            new[] {new[] {"00123", "100.5", "线上"}, new[] {"00124", "200", ""}});
    }

    [TestMethod]
    public void AColumnTheHeaderDoesNotHaveIsReported()
    {
        var missing = new List<ExcelColumnMissing>();
        var orders = new ExcelImporter()
            .ExcelToObject<Order>(Sample(), options => options.OnMissingColumn = missing.Add).ToList();

        // 沿用既有行为：那一列取默认值，导入照常
        Assert.AreEqual(2, orders.Count);
        Assert.AreEqual(0m, orders[0].Amount);
        Assert.AreEqual("00123", orders[0].No);

        Assert.AreEqual(1, missing.Count);
        Assert.AreEqual("金额", missing[0].Title);
        Assert.AreEqual(nameof(Order.Amount), missing[0].PropertyName);
        Assert.AreEqual("数据", missing[0].SheetTitle);
        CollectionAssert.AreEqual(new[] {"订单号", "金额（元）", "备注"}, missing[0].HeaderTitles.ToArray());
        // 相近的标题一并给出，否则「差在哪里」还得自己翻文件
        CollectionAssert.AreEqual(new[] {"金额（元）"}, missing[0].SimilarTitles.ToArray());
        StringAssert.Contains(missing[0].ToString(), "金额（元）");
    }

    [TestMethod]
    public void NothingIsReportedWhenEveryColumnIsFound()
    {
        var missing = new List<ExcelColumnMissing>();
        var bytes = Workbook(new[] {"订单号", "金额", "备注"}, new[] {new[] {"00123", "100.5", "线上"}});
        var orders = new ExcelImporter()
            .ExcelToObject<Order>(bytes, options => options.OnMissingColumn = missing.Add).ToList();

        Assert.AreEqual(100.5m, orders[0].Amount);
        Assert.AreEqual(0, missing.Count);
    }

    [TestMethod]
    public void WithoutTheCallbackTheImportIsUnchanged()
    {
        // 不设回调时一切照旧：这是既有版本的行为
        var orders = new ExcelImporter().ExcelToObject<Order>(Sample()).ToList();
        Assert.AreEqual(2, orders.Count);
        Assert.AreEqual(0m, orders[0].Amount);
    }

    [TestMethod]
    public void ThrowingFromTheCallbackRefusesTheFile()
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            new ExcelImporter().ExcelToObject<Order>(Sample(),
                    options => options.OnMissingColumn = m => throw new Excel2ObjectException(m.ToString()))
                .ToList());

        StringAssert.Contains(e.Message, "[金额]");
    }

    [TestMethod]
    public void TheStreamingImportReportsItToo()
    {
        var missing = new List<ExcelColumnMissing>();
        using var input = new MemoryStream(Sample());
        var orders = new ExcelImporter()
            .ExcelStreamToObject<Order>(input, options => options.OnMissingColumn = missing.Add).ToList();

        Assert.AreEqual(2, orders.Count);
        Assert.AreEqual(1, missing.Count);
        Assert.AreEqual("金额", missing[0].Title);
        CollectionAssert.AreEqual(new[] {"金额（元）"}, missing[0].SimilarTitles.ToArray());
    }

    [TestMethod]
    public void TwoColumnsMissingAreReportedOnceEach()
    {
        var missing = new List<ExcelColumnMissing>();
        var bytes = Workbook(new[] {"订单号"}, new[] {new[] {"00123"}, new[] {"00124"}, new[] {"00125"}});
        new ExcelImporter().ExcelToObject<Order>(bytes, options => options.OnMissingColumn = missing.Add).ToList();

        // 每个对不上的标题只上报一次，与行数无关
        CollectionAssert.AreEquivalent(new[] {"金额", "备注"}, missing.Select(m => m.Title).ToArray());
        Assert.AreEqual(0, missing[0].SimilarTitles.Count, "没有相近的就不要硬凑");
    }

    [TestMethod]
    public void AnEmptySheetReportsEveryColumnAsMissing()
    {
        var workbook = new XSSFWorkbook();
        workbook.CreateSheet("空表");
        using var bytes = new MemoryStream();
        workbook.Write(bytes, true);

        var missing = new List<ExcelColumnMissing>();
        var orders = new ExcelImporter()
            .ExcelToObject<Order>(bytes.ToArray(), options => options.OnMissingColumn = missing.Add).ToList();

        // 一行都没有，表头自然也没有：模型上的每个标题都对不上
        Assert.AreEqual(0, orders.Count);
        CollectionAssert.AreEquivalent(new[] {"订单号", "金额", "备注"}, missing.Select(m => m.Title).ToArray());
        Assert.AreEqual("空表", missing[0].SheetTitle);
        Assert.AreEqual(0, missing[0].HeaderTitles.Count);
    }

    [TestMethod]
    public void TheHeaderCanBeReadOnItsOwn()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var bytes = new ExcelExporter().ObjectToExcelBytes(
                              new[] {new Order {No = "00123", Amount = 100.5m, Memo = "线上"}}, options =>
                              {
                                  options.ExcelType = excelType;
                                  options.SheetTitle = "本月";
                              })
                          ?? throw new AssertFailedException($"{excelType} 导出应有内容");

            using var input = new MemoryStream(bytes);
            var header = ExcelHelper.ReadHeader(input);
            Assert.AreEqual("本月", header.SheetTitle, excelType.ToString());
            CollectionAssert.AreEqual(new[] {"订单号", "金额", "备注"}, header.Columns.ToArray(), excelType.ToString());
        }
    }

    [TestMethod]
    public void TheHeaderHonoursSheetTitleAndSkippedLines()
    {
        var bytes = Workbook(new[] {"订单号", "金额"}, new[] {new[] {"00123", "1"}}, "上月");
        var workbook = new XSSFWorkbook(new MemoryStream(bytes));
        var second = workbook.CreateSheet("本月");
        second.CreateRow(0).CreateCell(0).SetCellValue("导出说明");
        var header = second.CreateRow(1);
        header.CreateCell(0).SetCellValue("城市");
        header.CreateCell(1).SetCellValue("数量");
        using var both = new MemoryStream();
        workbook.Write(both, true);

        using var input = new MemoryStream(both.ToArray());
        var read = ExcelHelper.ReadHeader(input, options =>
        {
            options.SheetTitle = "本月";
            options.TitleSkipLine = 1;
        });

        Assert.AreEqual("本月", read.SheetTitle);
        CollectionAssert.AreEqual(new[] {"城市", "数量"}, read.Columns.ToArray());
    }

    [TestMethod]
    public void AnEmptySheetStillYieldsItsName()
    {
        var workbook = new XSSFWorkbook();
        workbook.CreateSheet("空表");
        using var bytes = new MemoryStream();
        workbook.Write(bytes, true);

        using var input = new MemoryStream(bytes.ToArray());
        var header = ExcelHelper.ReadHeader(input);
        Assert.AreEqual("空表", header.SheetTitle);
        Assert.AreEqual(0, header.Columns.Count);
    }
}
