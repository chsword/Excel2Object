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
    [TestMethod]
    public void NumberColumnWithoutAStyleKeepsTheWorkbookDefault()
    {
        var cell = Export(ExcelType.Xlsx, out _).GetRow(1).GetCell(6);
        Assert.AreEqual(CellType.Numeric, cell.CellType);
        Assert.AreEqual(0, cell.CellStyle.Index);
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
