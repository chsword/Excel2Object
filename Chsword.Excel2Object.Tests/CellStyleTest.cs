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
        Assert.IsTrue(row.Cells.All(c => c.CellStyle.DataFormat == text),
            string.Join(",", row.Cells.Select(c => c.CellStyle.DataFormat)));
    }
}
