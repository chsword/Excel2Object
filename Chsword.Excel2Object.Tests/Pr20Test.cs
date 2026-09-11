using System;
using System.Collections.Generic;
using System.IO;
using Chsword.Excel2Object.Styles;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using Alignment = Chsword.Excel2Object.Styles.HorizontalAlignment;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class Pr20Test : BaseExcelTest
{
    /// <summary>
    ///     Every style declared on [ExcelColumn] has to survive the export: the header styles on their
    ///     header cell, the cell styles on the data cells, and [ExcelTitle] on the class as the sheet name.
    /// </summary>
    [TestMethod]
    public void UseExcelColumnAttr()
    {
        var list = new List<Pr20Model>
        {
            new() {Fullname = "AAA", Mobile = "123456798123"},
            new() {Fullname = "BBB", Mobile = "234"}
        };
        var bytes = ExcelHelper.ObjectToExcelBytes(list, ExcelType.Xlsx);
        Assert.IsNotNull(bytes);

        using var stream = new MemoryStream(bytes);
        var workbook = WorkbookFactory.Create(stream);
        var sheet = workbook.GetSheet("SheetX");
        Assert.IsNotNull(sheet, "[ExcelTitle] on the class names the sheet");

        var mobileHeader = sheet.GetRow(0).GetCell(1);
        var headerFont = mobileHeader.CellStyle.GetFont(workbook);
        Assert.AreEqual("手机", mobileHeader.StringCellValue);
        Assert.AreEqual("宋体", headerFont.FontName);
        Assert.IsTrue(headerFont.IsBold);
        Assert.IsTrue(headerFont.IsItalic);
        Assert.AreEqual(30, headerFont.FontHeightInPoints);
        Assert.AreEqual((short) ExcelStyleColor.Blue, headerFont.Color);
        Assert.AreEqual(FontUnderlineType.Single, headerFont.Underline);
        Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.Right, mobileHeader.CellStyle.Alignment);

        var dataRow = sheet.GetRow(1);
        Assert.AreEqual((short) ExcelStyleColor.Red,
            dataRow.GetCell(0).CellStyle.GetFont(workbook).Color);
        Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.Justify, dataRow.GetCell(1).CellStyle.Alignment);
    }

    [ExcelTitle("SheetX")]
    public class Pr20Model
    {
        [ExcelColumn("姓名", CellFontColor = ExcelStyleColor.Red)]
        public string? Fullname { get; set; }

        [ExcelColumn("手机",
            HeaderFontFamily = "宋体",
            HeaderBold = true,
            HeaderFontHeight = 30,
            HeaderItalic = true,
            HeaderFontColor = ExcelStyleColor.Blue,
            HeaderUnderline = true,
            HeaderAlignment = Alignment.Right,
            //cell
            CellAlignment = Alignment.Justify
        )]
        public string Mobile { get; set; } = null!;
    }
}