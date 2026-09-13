using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Internal;
using Chsword.Excel2Object.Options;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     单元格读取失败时，导入不再向标准输出写日志，而是通过回调上报，并可由调用方决定是否中止。
/// </summary>
[TestClass]
public class ImportErrorTest : BaseExcelTest
{
    public class Model
    {
        [ExcelTitle("名称")] public string Name { get; set; } = "";
        [ExcelTitle("日期")] public DateTime? When { get; set; }
    }

    /// <summary>
    ///     序列号 -1 超出 Excel 日历，NPOI 读取该单元格时抛出异常。此处构造这样一份文件。
    /// </summary>
    private static byte[] WorkbookWithBrokenDate()
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("数据");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("名称");
        header.CreateCell(1).SetCellValue("日期");

        var style = workbook.CreateCellStyle();
        style.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd");

        var good = sheet.CreateRow(1);
        good.CreateCell(0).SetCellValue("正常");
        good.CreateCell(1).SetCellValue(new DateTime(2026, 9, 13));
        good.GetCell(1).CellStyle = style;

        var broken = sheet.CreateRow(2);
        broken.CreateCell(0).SetCellValue("越界");
        broken.CreateCell(1).SetCellValue(-1d);
        broken.GetCell(1).CellStyle = style;

        var stream = new MemoryStream();
        workbook.Write(stream, true);
        return stream.ToArray();
    }

    /// <summary>未设置回调时，该单元格取默认值，其余数据照常读取。</summary>
    [TestMethod]
    public void ABrokenCellDoesNotStopTheImport()
    {
        var list = ExcelHelper.ExcelToObject<Model>(WorkbookWithBrokenDate()).ToArray();

        Assert.AreEqual(2, list.Length);
        Assert.AreEqual(new DateTime(2026, 9, 13), list[0].When);
        Assert.IsNull(list[1].When, "越界的日期取 null");
        Assert.AreEqual("越界", list[1].Name, "同一行的其他列不受影响");
    }

    /// <summary>设置回调后，可得知是哪一个单元格、因何失败。</summary>
    [TestMethod]
    public void TheErrorIsReportedWithItsCell()
    {
        var errors = new List<ExcelImportError>();
        var list = new ExcelImporter()
            .ExcelToObject<Model>(WorkbookWithBrokenDate(), options => options.OnCellError = errors.Add)
            .ToArray();

        Assert.AreEqual(2, list.Length);
        Assert.AreEqual(1, errors.Count);
        Assert.AreEqual("数据", errors[0].SheetTitle);
        Assert.AreEqual("B3", errors[0].CellReference);
        Assert.AreEqual(2, errors[0].RowIndex);
        Assert.AreEqual(1, errors[0].ColumnIndex);
        Assert.IsNotNull(errors[0].Exception);
        StringAssert.Contains(errors[0].ToString(), "B3");
    }

    /// <summary>在回调中抛出异常即可使导入中止。</summary>
    [TestMethod]
    public void ThrowingInTheCallbackStopsTheImport()
    {
        var e = Assert.ThrowsException<InvalidOperationException>(() =>
            new ExcelImporter()
                .ExcelToObject<Model>(WorkbookWithBrokenDate(),
                    options => options.OnCellError = error =>
                        throw new InvalidOperationException(error.CellReference, error.Exception))
                .ToArray());

        Assert.AreEqual("B3", e.Message);
        Assert.IsNotNull(e.InnerException);
    }

    /// <summary>
    ///     回调抛出异常以中止导入时，回调只应被调用一次：读取路径有嵌套，异常向外传播时会途经外层的
    ///     catch，若不加区分便会被再次当作读取失败上报。
    /// </summary>
    [TestMethod]
    public void TheCallbackRunsOnceWhenItAborts()
    {
        var calls = 0;
        var e = Assert.ThrowsException<InvalidOperationException>(() =>
            new ExcelImporter()
                .ExcelToObject<Model>(WorkbookWithBrokenDate(), options => options.OnCellError = error =>
                {
                    calls++;
                    throw new InvalidOperationException(error.CellReference);
                })
                .ToArray());

        Assert.AreEqual("B3", e.Message);
        Assert.AreEqual(1, calls, "回调被调用的次数");
    }

    /// <summary>公式求值器按工作簿创建一次，不再逐单元格创建。</summary>
    [TestMethod]
    public void OneFormulaEvaluatorPerWorkbook()
    {
        var context = new ImportContext(new ExcelImporterOptions());
        var first = new XSSFWorkbook();
        var second = new XSSFWorkbook();

        Assert.AreSame(context.Evaluator(first), context.Evaluator(first));
        Assert.AreNotSame(context.Evaluator(first), context.Evaluator(second));
    }

    public class FormulaModel
    {
        [ExcelTitle("数量")] public int Count { get; set; }
        [ExcelTitle("合计")] public double Total { get; set; }
    }

    /// <summary>公式单元格仍按求值结果读取。</summary>
    [TestMethod]
    public void FormulaCellsAreStillEvaluated()
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Sheet1");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("数量");
        header.CreateCell(1).SetCellValue("合计");

        for (var i = 1; i <= 20; i++)
        {
            var row = sheet.CreateRow(i);
            row.CreateCell(0).SetCellValue(i);
            row.CreateCell(1).SetCellFormula($"A{i + 1}*2");
        }

        var stream = new MemoryStream();
        workbook.Write(stream, true);

        var list = ExcelHelper.ExcelToObject<FormulaModel>(stream.ToArray()).ToArray();
        Assert.AreEqual(20, list.Length);
        Assert.AreEqual(2d, list[0].Total);
        Assert.AreEqual(40d, list[19].Total);
    }
}
