using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Styles;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using Alignment = Chsword.Excel2Object.Styles.HorizontalAlignment;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     Cell styles declared on [ExcelColumn] have to reach the data cells, and must not leak into
///     columns that did not ask for them.
/// </summary>
[TestClass]
public class CellStyleTest : BaseExcelTest
{
    public class StyledModel
    {
        [ExcelColumn("Styled", CellBold = true, CellFontColor = ExcelStyleColor.Red,
            CellAlignment = Alignment.Right)]
        public string Styled { get; set; } = "s";

        [ExcelTitle("Plain")] public string Plain { get; set; } = "p";

        [ExcelColumn("HeaderOnly", HeaderBold = true)]
        public string HeaderOnly { get; set; } = "h";

        [ExcelColumn("SameAsStyled", CellBold = true, CellFontColor = ExcelStyleColor.Red,
            CellAlignment = Alignment.Right)]
        public string SameAsStyled { get; set; } = "x";

        [ExcelColumn("Qty", CellBold = true, CellFontColor = ExcelStyleColor.Red,
            CellAlignment = Alignment.Right, Format = "0.00")]
        public int Qty { get; set; } = 7;

        [ExcelColumn("Done", CellBold = true, CellFontColor = ExcelStyleColor.Red,
            CellAlignment = Alignment.Right)]
        public bool Done { get; set; } = true;

        [ExcelTitle("PlainQty")] public int PlainQty { get; set; } = 9;

        [ExcelColumn("Link", CellBold = true, CellFontColor = ExcelStyleColor.Red,
            CellAlignment = Alignment.Right)]
        public Uri Link { get; set; } = new("https://github.com/chsword/Excel2Object");

        [ExcelTitle("PlainLink")] public Uri PlainLink { get; set; } = new("https://example.com/");

        [ExcelColumn("BareQty")] public int BareQty { get; set; } = 11;

        [ExcelColumn("Money", Format = "#,##0.000")]
        public decimal Money { get; set; } = 1.5m;

        [ExcelColumn("BareLink")] public Uri BareLink { get; set; } = new("https://example.org/");
    }

    private static ISheet Export(ExcelType excelType, out IWorkbook workbook)
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(new List<StyledModel> {new()},
            options => options.ExcelType = excelType);
        Assert.IsNotNull(bytes);
        workbook = WorkbookFactory.Create(new MemoryStream(bytes));
        return workbook.GetSheetAt(0);
    }

    [TestMethod]
    public void CellStyleOnTheAttributeReachesTheCell()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var cell = Export(excelType, out var workbook).GetRow(1).GetCell(0);
            var font = cell.CellStyle.GetFont(workbook);

            Assert.IsTrue(font.IsBold, excelType.ToString());
            Assert.AreEqual((short) ExcelStyleColor.Red, font.Color, excelType.ToString());
            Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.Right, cell.CellStyle.Alignment,
                excelType.ToString());
        }
    }

    [TestMethod]
    public void StyleDoesNotLeakToOtherColumnsOrTheWorkbookDefault()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var row = Export(excelType, out var workbook).GetRow(1);

            Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.General,
                workbook.GetCellStyleAt(0).Alignment, excelType.ToString());
            foreach (var index in new[] {1, 2})
            {
                var cell = row.GetCell(index);
                Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.General, cell.CellStyle.Alignment,
                    $"{excelType} column {index}");
                Assert.IsFalse(cell.CellStyle.GetFont(workbook).IsBold, $"{excelType} column {index}");
            }
        }
    }

    /// <summary>
    ///     A column that carries [ExcelColumn] only for its title must still look like any other column;
    ///     the style machinery must not quietly impose a font size of its own.
    /// </summary>
    [TestMethod]
    public void AColumnWithoutFontSettingsKeepsTheDefaultFontHeight()
    {
        var sheet = Export(ExcelType.Xlsx, out var workbook);
        var styled = sheet.GetRow(1).GetCell(2); // HeaderOnly: [ExcelColumn] with header styles only
        var plain = sheet.GetRow(1).GetCell(1); // Plain: [ExcelTitle]

        Assert.AreEqual(plain.CellStyle.GetFont(workbook).FontHeightInPoints,
            styled.CellStyle.GetFont(workbook).FontHeightInPoints);
    }

    /// <summary>
    ///     Numeric and boolean cells take a different path through the exporter than text, so they need
    ///     their own check that the style survives - and that they stay numeric/boolean cells.
    /// </summary>
    [TestMethod]
    public void NumberAndBooleanCellsAreStyledToo()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var row = Export(excelType, out var workbook).GetRow(1);

            var qty = row.GetCell(4);
            Assert.AreEqual(CellType.Numeric, qty.CellType, excelType.ToString());
            Assert.AreEqual(7d, qty.NumericCellValue, excelType.ToString());
            Assert.IsTrue(qty.CellStyle.GetFont(workbook).IsBold, excelType.ToString());
            Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.Right, qty.CellStyle.Alignment);
            Assert.AreEqual(NPOI.HSSF.UserModel.HSSFDataFormat.GetBuiltinFormat("0.00"),
                qty.CellStyle.DataFormat, $"{excelType} number format");

            var done = row.GetCell(5);
            Assert.AreEqual(CellType.Boolean, done.CellType, excelType.ToString());
            Assert.IsTrue(done.BooleanCellValue, excelType.ToString());
            Assert.IsTrue(done.CellStyle.GetFont(workbook).IsBold, excelType.ToString());
            Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.Right, done.CellStyle.Alignment);
        }
    }

    /// <summary>
    ///     A number column that declared no style keeps the workbook default instead of gaining an empty
    ///     style of its own, so existing exports are unchanged.
    /// </summary>
    /// <summary>
    ///     [ExcelColumn] is handed to the column as its CellStyle even when it only carries a title, so
    ///     "did this column ask for a style" cannot be a null check - a bare [ExcelColumn] number column
    ///     must stay on the workbook default just like an [ExcelTitle] one.
    /// </summary>
    [TestMethod]
    public void NumberColumnWithoutAStyleKeepsTheWorkbookDefault()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var row = Export(excelType, out var workbook).GetRow(1);
            var byTitle = row.GetCell(6); // PlainQty: [ExcelTitle], never had a style
            var bareAttribute = row.GetCell(9); // BareQty: [ExcelColumn] carrying nothing but a title

            Assert.AreEqual(CellType.Numeric, bareAttribute.CellType, excelType.ToString());
            // the same style object, not merely one that happens to look the same: a bare [ExcelColumn]
            // must not allocate a style of its own (.xlsx and .xls number their default differently, so
            // compare the two cells instead of using a literal index)
            Assert.AreEqual(byTitle.CellStyle.Index, bareAttribute.CellStyle.Index, excelType.ToString());
            Assert.AreEqual(0, bareAttribute.CellStyle.DataFormat, $"{excelType} format");
            Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.General, bareAttribute.CellStyle.Alignment,
                excelType.ToString());
            Assert.IsFalse(bareAttribute.CellStyle.GetFont(workbook).IsBold, excelType.ToString());
        }
    }

    /// <summary>
    ///     A Uri column becomes a hyperlink cell on a branch of its own, which also has to carry the
    ///     column's style - while a link nobody styled keeps the workbook default.
    /// </summary>
    [TestMethod]
    public void HyperlinkCellsAreStyledToo()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var row = Export(excelType, out var workbook).GetRow(1);

            var link = row.GetCell(7);
            Assert.IsNotNull(link.Hyperlink, excelType.ToString());
            Assert.AreEqual("https://github.com/chsword/Excel2Object", link.Hyperlink.Address);
            Assert.IsTrue(link.CellStyle.GetFont(workbook).IsBold, excelType.ToString());
            Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.Right, link.CellStyle.Alignment);

            // the default style index differs between .xlsx and .xls, so compare against another
            // column that declared no style rather than against a literal
            // a link nobody styled keeps the workbook default - including one whose [ExcelColumn]
            // carries nothing but a title, which is not distinguishable by a null check
            var defaultStyle = row.GetCell(6).CellStyle.Index;
            foreach (var index in new[] {8, 11})
            {
                var plain = row.GetCell(index);
                Assert.IsNotNull(plain.Hyperlink, $"{excelType} column {index}");
                Assert.AreEqual(defaultStyle, plain.CellStyle.Index, $"{excelType} column {index}");
            }
        }
    }

    /// <summary>
    ///     A number format Excel has no builtin for has to be registered on the workbook, not dropped.
    /// </summary>
    [TestMethod]
    public void ACustomNumberFormatReachesTheCell()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var cell = Export(excelType, out _).GetRow(1).GetCell(10);
            Assert.AreEqual(CellType.Numeric, cell.CellType, excelType.ToString());
            Assert.AreEqual(1.5d, cell.NumericCellValue, excelType.ToString());
            Assert.AreEqual("#,##0.000", cell.CellStyle.GetDataFormatString(), excelType.ToString());
        }
    }

    [TestMethod]
    public void ColumnsAskingForTheSameLookShareOneCellStyle()
    {
        var row = Export(ExcelType.Xlsx, out _).GetRow(1);
        Assert.AreEqual(row.GetCell(0).CellStyle.Index, row.GetCell(3).CellStyle.Index);
    }

    [TestMethod]
    public void StringCellsKeepTheTextFormat()
    {
        var row = Export(ExcelType.Xlsx, out _).GetRow(1);
        var text = NPOI.HSSF.UserModel.HSSFDataFormat.GetBuiltinFormat("text");
        var stringCells = new[] {0, 1, 2, 3}.Select(row.GetCell).ToList();
        Assert.IsTrue(stringCells.All(c => c.CellStyle.DataFormat == text),
            string.Join(",", stringCells.Select(c => c.CellStyle.DataFormat)));
    }
}
