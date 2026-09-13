using System;
using System.Collections.Generic;
using System.IO;
using Chsword.Excel2Object.Options;
using Chsword.Excel2Object.Styles;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     Styles are written by what they apply to - the header, every cell, one column, every other row -
///     instead of repeated on every column, and they layer over each other.
/// </summary>
[TestClass]
public class StyleSheetTest : BaseExcelTest
{
    public class Model
    {
        [ExcelTitle("城市")] public string City { get; set; } = "北京";
        [ExcelTitle("金额")] public decimal Amount { get; set; } = 1000;

        [ExcelColumn("备注", CellBold = true)] public string Note { get; set; } = "n/a";
    }

    private static readonly List<Model> Rows = new() {new(), new(), new()};

    private static ISheet Export(ExcelType excelType, Action<ExcelExporterOptions> configure,
        out IWorkbook workbook, List<Model>? rows = null)
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(rows ?? Rows, options =>
        {
            options.ExcelType = excelType;
            configure(options);
        });
        Assert.IsNotNull(bytes);
        workbook = WorkbookFactory.Create(new MemoryStream(bytes));
        return workbook.GetSheetAt(0);
    }

    private static string Hex(IColor? color)
    {
        var rgb = color?.RGB;
        return rgb == null ? "" : $"#{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2}";
    }

    [TestMethod]
    public void TheHeaderTakesTheStyleWrittenForIt()
    {
        var sheet = Export(ExcelType.Xlsx,
            options => options.Styles.Header(s => s.Bold().Background("#4472C4").Color("#FFF").Center()),
            out var workbook);

        var cell = sheet.GetRow(0).GetCell(0);
        var font = cell.CellStyle.GetFont(workbook);
        Assert.IsTrue(font.IsBold);
        Assert.AreEqual("#FFFFFF", Hex(((XSSFFont) font).GetXSSFColor()));
        Assert.AreEqual("#4472C4", Hex(((XSSFCellStyle) cell.CellStyle).FillForegroundColorColor));
        Assert.AreEqual(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
        Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.Center, cell.CellStyle.Alignment);
    }

    /// <summary>A hex colour is stored as it is in .xlsx; .xls has to pick the nearest of its 56.</summary>
    [TestMethod]
    public void AHexColourSurvivesInXlsxAndIsMatchedInXls()
    {
        var xlsx = Export(ExcelType.Xlsx, options => options.Styles.Cells(s => s.Background("#4472C4")), out _);
        Assert.AreEqual("#4472C4",
            Hex(((XSSFCellStyle) xlsx.GetRow(1).GetCell(0).CellStyle).FillForegroundColorColor));

        var xls = Export(ExcelType.Xls, options => options.Styles.Cells(s => s.Background("#4472C4")),
            out var workbook);
        var index = xls.GetRow(1).GetCell(0).CellStyle.FillForegroundColor;
        Assert.AreNotEqual(0, index);
        Assert.AreEqual(FillPattern.SolidForeground, xls.GetRow(1).GetCell(0).CellStyle.FillPattern);
        // #4472C4 is a blue, and the palette colour chosen for it has to be one too
        var rgb = ((NPOI.HSSF.UserModel.HSSFWorkbook) workbook).GetCustomPalette().GetColor(index)!.RGB;
        Assert.IsTrue(rgb[2] > rgb[0], $"#{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2} is no blue");
    }

    /// <summary>
    ///     A light grey has no exact match in the .xls palette, and the nearest is measured by eye rather
    ///     than by raw RGB distance - which would land on a pale lavender.
    /// </summary>
    [TestMethod]
    public void TheNearestPaletteColourOfAGreyIsAGrey()
    {
        var sheet = Export(ExcelType.Xls, options => options.Styles.EvenRows(s => s.Background("#D9D9D9")),
            out var workbook);

        var index = sheet.GetRow(2).GetCell(0).CellStyle.FillForegroundColor;
        var rgb = ((NPOI.HSSF.UserModel.HSSFWorkbook) workbook).GetCustomPalette().GetColor(index)!.RGB;
        Assert.AreEqual(rgb[0], rgb[1], $"#{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2} is no grey");
        Assert.AreEqual(rgb[1], rgb[2], $"#{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2} is no grey");
    }

    /// <summary>Every other row takes its own colour, and the first data row counts as odd.</summary>
    [TestMethod]
    public void RowsStripe()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.Styles.EvenRows(s => s.Background("#F2F2F2")), out _);

        Assert.AreEqual("", Hex(((XSSFCellStyle) sheet.GetRow(1).GetCell(0).CellStyle).FillForegroundColorColor));
        Assert.AreEqual("#F2F2F2",
            Hex(((XSSFCellStyle) sheet.GetRow(2).GetCell(0).CellStyle).FillForegroundColorColor));
        Assert.AreEqual("", Hex(((XSSFCellStyle) sheet.GetRow(3).GetCell(0).CellStyle).FillForegroundColorColor));
    }

    /// <summary>Cells under stripes under the attribute under the column's own style.</summary>
    [TestMethod]
    public void StylesLayerFromTheWidestToTheNarrowest()
    {
        var sheet = Export(ExcelType.Xlsx, options =>
        {
            options.Styles.Cells(s => s.Italic().FontFamily("宋体"));
            options.Styles.EvenRows(s => s.Background("#F2F2F2"));
            options.Styles.Column("金额", s => s.Right().Format("#,##0.00"));
        }, out var workbook);

        var amount = sheet.GetRow(2).GetCell(1);
        var font = amount.CellStyle.GetFont(workbook);
        Assert.IsTrue(font.IsItalic, "from Cells");
        Assert.AreEqual("宋体", font.FontName, "from Cells");
        Assert.AreEqual("#F2F2F2", Hex(((XSSFCellStyle) amount.CellStyle).FillForegroundColorColor), "from EvenRows");
        Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.Right, amount.CellStyle.Alignment, "from Column");
        Assert.AreEqual("#,##0.00", amount.CellStyle.GetDataFormatString(), "from Column");

        // the attribute still reaches the cells of its own column
        Assert.IsTrue(sheet.GetRow(1).GetCell(2).CellStyle.GetFont(workbook).IsBold, "from [ExcelColumn]");
    }

    /// <summary>The style written for a column is the most deliberate, so it wins over the attribute.</summary>
    [TestMethod]
    public void AColumnStyleOverridesTheAttribute()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.Styles.Column("备注", s => s.Bold(false).Underline()),
            out var workbook);

        var font = sheet.GetRow(1).GetCell(2).CellStyle.GetFont(workbook);
        Assert.IsFalse(font.IsBold);
        Assert.AreEqual(FontUnderlineType.Single, font.Underline);
    }

    [TestMethod]
    public void BordersAreWrittenTheWayCssWritesThem()
    {
        var sheet = Export(ExcelType.Xlsx, options =>
        {
            options.Styles.Cells(s => s.Border("1px solid #D0D0D0"));
            options.Styles.Column("金额", s => s.BorderBottom("2px dashed #FF0000"));
        }, out _);

        var city = sheet.GetRow(1).GetCell(0).CellStyle;
        Assert.AreEqual(BorderStyle.Thin, city.BorderTop);
        Assert.AreEqual(BorderStyle.Thin, city.BorderLeft);
        Assert.AreEqual("#D0D0D0", Hex(((XSSFCellStyle) city).TopBorderXSSFColor));

        var amount = (XSSFCellStyle) sheet.GetRow(1).GetCell(1).CellStyle;
        Assert.AreEqual(BorderStyle.MediumDashed, amount.BorderBottom, "2px dashed");
        Assert.AreEqual("#FF0000", Hex(amount.BottomBorderXSSFColor));
        Assert.AreEqual(BorderStyle.Thin, amount.BorderTop, "the other sides keep what Cells gave them");
    }

    /// <summary>A format written for a column applies whatever the column holds, numbers included.</summary>
    [TestMethod]
    public void AFormatReachesANumberColumn()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.Styles.Column("金额", s => s.Format("#,##0.00")), out _);

        var cell = sheet.GetRow(1).GetCell(1);
        Assert.AreEqual(CellType.Numeric, cell.CellType);
        Assert.AreEqual("#,##0.00", cell.CellStyle.GetDataFormatString());
    }

    /// <summary>Text keeps its own format unless the style names one, so a leading zero survives.</summary>
    [TestMethod]
    public void ATextColumnKeepsTheTextFormat()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.Styles.Cells(s => s.Italic()), out _);
        Assert.AreEqual("@", sheet.GetRow(1).GetCell(0).CellStyle.GetDataFormatString());
    }

    [TestMethod]
    public void WrapAndVerticalAlignmentReachTheCell()
    {
        var sheet = Export(ExcelType.Xlsx,
            options => options.Styles.Cells(s => s.Wrap().VerticalAlign(ExcelVerticalAlignment.Middle)), out _);

        var style = sheet.GetRow(1).GetCell(0).CellStyle;
        Assert.IsTrue(style.WrapText);
        Assert.AreEqual(VerticalAlignment.Center, style.VerticalAlignment);
    }

    /// <summary>
    ///     Cells that look alike share one cell style: a workbook holds a limited number of them, and a
    ///     striped table would otherwise want one per row.
    /// </summary>
    [TestMethod]
    public void CellsThatLookAlikeShareOneStyle()
    {
        var many = new List<Model>();
        for (var i = 0; i < 400; i++) many.Add(new Model());

        var sheet = Export(ExcelType.Xlsx, options =>
        {
            options.Styles.Cells(s => s.Border("1px solid #CCC"));
            options.Styles.EvenRows(s => s.Background("#F2F2F2"));
        }, out var workbook, many);

        Assert.AreEqual(sheet.GetRow(1).GetCell(0).CellStyle.Index, sheet.GetRow(3).GetCell(0).CellStyle.Index);
        Assert.AreEqual(sheet.GetRow(2).GetCell(0).CellStyle.Index, sheet.GetRow(4).GetCell(0).CellStyle.Index);
        // header, text, number, striped variants of each - a handful, not one per cell
        Assert.IsTrue(workbook.NumCellStyles < 20, $"{workbook.NumCellStyles} cell styles");
    }

    /// <summary>A colour picked from the palette keeps meaning that palette entry, as it always has.</summary>
    [TestMethod]
    public void APaletteColourStaysIndexed()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.Styles.Cells(s => s.Color(ExcelStyleColor.Red)),
            out var workbook);

        Assert.AreEqual((short) ExcelStyleColor.Red, sheet.GetRow(1).GetCell(0).CellStyle.GetFont(workbook).Color);
    }

    /// <summary>A blank cell is still part of the table, stripes and borders included.</summary>
    [TestMethod]
    public void ABlankCellKeepsTheLookOfItsRow()
    {
        var rows = new List<Model> {new(), new() {Amount = 0}};
        var sheet = Export(ExcelType.Xlsx, options =>
        {
            options.Styles.EvenRows(s => s.Background("#F2F2F2"));
            options.Styles.Cells(s => s.Border("1px solid #CCC"));
        }, out _, rows);

        // a null in a nullable column would blank the cell; this model blanks it through an empty value
        var blank = sheet.GetRow(2).GetCell(1);
        Assert.AreEqual("#F2F2F2", Hex(((XSSFCellStyle) blank.CellStyle).FillForegroundColorColor));
        Assert.AreEqual(BorderStyle.Thin, blank.CellStyle.BorderTop);
    }

    /// <summary>The width of an auto-sized column follows what the format makes of the number.</summary>
    [TestMethod]
    public void AutoColumnWidthFollowsTheNumberFormat()
    {
        var rows = new List<Model> {new() {Amount = 1000}};
        var narrow = Export(ExcelType.Xlsx, options =>
        {
            options.AutoColumnWidth = true;
            options.MinColumnWidth = 1;
        }, out _, rows);
        var wide = Export(ExcelType.Xlsx, options =>
        {
            options.AutoColumnWidth = true;
            options.MinColumnWidth = 1;
            options.Styles.Column("金额", s => s.Format("#,##0.00"));
        }, out _, rows);

        // "1000" against "1,000.00"
        Assert.IsTrue(wide.GetColumnWidth(1) > narrow.GetColumnWidth(1),
            $"{wide.GetColumnWidth(1)} vs {narrow.GetColumnWidth(1)}");
    }

    /// <summary>A header size written in the stylesheet is not undone by the legacy default.</summary>
    [TestMethod]
    public void TheStylesheetSetsTheHeaderSize()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.Styles.Header(s => s.FontSize(20)),
            out var workbook);

        // 备注 declares [ExcelColumn] and names no header size of its own
        Assert.AreEqual(20, sheet.GetRow(0).GetCell(2).CellStyle.GetFont(workbook).FontHeightInPoints);
    }

    /// <summary>A format written for every cell reaches a date column too.</summary>
    [TestMethod]
    public void ADateTakesAFormatFromAnyScope()
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(new List<Dated> {new()}, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.Styles.Cells(s => s.Format("yyyy-MM-dd"));
        })!;
        var sheet = WorkbookFactory.Create(new MemoryStream(bytes)).GetSheetAt(0);

        Assert.AreEqual("yyyy-mm-dd", sheet.GetRow(1).GetCell(0).CellStyle.GetDataFormatString());
    }

    public class Dated
    {
        [ExcelTitle("日期")] public DateTime When { get; set; } = new(2026, 9, 13, 14, 30, 45);
    }

    public enum Grade
    {
        A,
        B
    }

    public class MixedModel
    {
        [ExcelTitle("等级")] public Grade Grade { get; set; } = Grade.A;
        [ExcelTitle("编号")] public Guid Id { get; set; } = Guid.NewGuid();
        [ExcelTitle("用时")] public TimeSpan Took { get; set; } = TimeSpan.FromMinutes(3);
    }

    /// <summary>
    ///     Columns of a type the exporter writes as text - an enum, a Guid, a TimeSpan - are part of the
    ///     table too, so they take its stripes and borders.
    /// </summary>
    [TestMethod]
    public void EveryColumnTypeTakesTheStyle()
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(new List<MixedModel> {new()}, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.Styles.Cells(s => s.Background("#F2F2F2"));
        })!;
        var row = WorkbookFactory.Create(new MemoryStream(bytes)).GetSheetAt(0).GetRow(1);

        for (var i = 0; i < 3; i++)
            Assert.AreEqual("#F2F2F2", Hex(((XSSFCellStyle) row.GetCell(i).CellStyle).FillForegroundColorColor),
                row.GetCell(i).ToString());
    }

    /// <summary>
    ///     A format written for every cell is about the numbers; a date column keeps showing a date rather
    ///     than the serial number underneath it.
    /// </summary>
    [TestMethod]
    public void ANumberFormatDoesNotReachADateColumn()
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(new List<Dated> {new()}, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.Styles.Cells(s => s.Format("#,##0.00"));
        })!;
        var cell = WorkbookFactory.Create(new MemoryStream(bytes)).GetSheetAt(0).GetRow(1).GetCell(0);

        Assert.AreEqual("yyyy-mm-dd hh:mm:ss", cell.CellStyle.GetDataFormatString());
        Assert.IsTrue(DateUtil.IsCellDateFormatted(cell));
    }

    /// <summary>One header row, one size: the legacy default steps aside for a header style.</summary>
    [TestMethod]
    public void AStyledHeaderRowHasOneSize()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.Styles.Header(s => s.Bold()), out var workbook);

        var plain = sheet.GetRow(0).GetCell(0).CellStyle.GetFont(workbook).FontHeightInPoints;
        var declared = sheet.GetRow(0).GetCell(2).CellStyle.GetFont(workbook).FontHeightInPoints;
        Assert.AreEqual(plain, declared, "[ExcelTitle] and [ExcelColumn] headers");
    }

    /// <summary>Without a header style, an [ExcelColumn] header keeps the size it always had.</summary>
    [TestMethod]
    public void WithoutAHeaderStyleTheLegacySizeStands()
    {
        var sheet = Export(ExcelType.Xlsx, _ => { }, out var workbook);
        Assert.AreEqual(10, sheet.GetRow(0).GetCell(2).CellStyle.GetFont(workbook).FontHeightInPoints);
    }

    public class Timed
    {
        [ExcelColumn("时间", Format = "yyyy-MM-dd HH:mm:ss")]
        public DateTime When { get; set; } = new(2026, 9, 13, 14, 30, 45);

        [ExcelTitle("金额")] public decimal Amount { get; set; } = 12.5m;
        [ExcelTitle("编号")] public string No { get; set; } = "00123";
    }

    private static ISheet ExportTimed(Action<ExcelExporterOptions> configure)
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(new List<Timed> {new()}, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            configure(options);
        })!;
        return WorkbookFactory.Create(new MemoryStream(bytes)).GetSheetAt(0);
    }

    /// <summary>
    ///     A format written for every cell is a default. The column's own - here the attribute's - is the
    ///     more specific statement and keeps its time of day.
    /// </summary>
    [TestMethod]
    public void AColumnsOwnFormatOutranksTheSheetWideOne()
    {
        var sheet = ExportTimed(options => options.Styles.Cells(s => s.Format("yyyy-MM-dd")));
        Assert.AreEqual("yyyy-mm-dd hh:mm:ss", sheet.GetRow(1).GetCell(0).CellStyle.GetDataFormatString());
    }

    /// <summary>And the stylesheet's own Column(...) outranks the attribute in turn.</summary>
    [TestMethod]
    public void AColumnStyleOutranksTheAttributesFormat()
    {
        var sheet = ExportTimed(options => options.Styles.Column("时间", s => s.Format("yyyy-MM-dd")));
        Assert.AreEqual("yyyy-mm-dd", sheet.GetRow(1).GetCell(0).CellStyle.GetDataFormatString());
    }

    /// <summary>A date format written for every cell has nothing to say to a number column.</summary>
    [TestMethod]
    public void ADateFormatDoesNotReachANumberColumn()
    {
        var sheet = ExportTimed(options => options.Styles.Cells(s => s.Format("yyyy-MM-dd")));

        var amount = sheet.GetRow(1).GetCell(1);
        Assert.AreEqual(CellType.Numeric, amount.CellType);
        Assert.AreEqual(0, amount.CellStyle.DataFormat, "12.5 would show as 1900-01-12");
    }

    /// <summary>Nor does a number format take the text format away from a string column.</summary>
    [TestMethod]
    public void ASheetWideNumberFormatLeavesTextAlone()
    {
        var sheet = ExportTimed(options => options.Styles.Cells(s => s.Format("#,##0.00")));
        Assert.AreEqual("@", sheet.GetRow(1).GetCell(2).CellStyle.GetDataFormatString());
    }

    /// <summary>A format written for that column reaches it whatever it holds.</summary>
    [TestMethod]
    public void AColumnStyleFormatReachesATextColumn()
    {
        var sheet = ExportTimed(options => options.Styles.Column("编号", s => s.Format("000000")));
        Assert.AreEqual("000000", sheet.GetRow(1).GetCell(2).CellStyle.GetDataFormatString());
    }

    /// <summary>
    ///     A date column given a number format shows the serial number Excel stores, and an auto-sized
    ///     column is measured as what it shows rather than as the date underneath.
    /// </summary>
    [TestMethod]
    public void AutoColumnWidthFollowsADateShownAsANumber()
    {
        byte[] Export(Action<ExcelExporterOptions> configure)
        {
            return new ExcelExporter().ObjectToExcelBytes(new List<Timed> {new()}, options =>
            {
                options.ExcelType = ExcelType.Xlsx;
                options.AutoColumnWidth = true;
                options.MinColumnWidth = 1;
                configure(options);
            })!;
        }

        var asDate = WorkbookFactory.Create(new MemoryStream(Export(_ => { }))).GetSheetAt(0);
        var asNumber = WorkbookFactory
            .Create(new MemoryStream(Export(o => o.Styles.Column("时间", s => s.Format("#,##0.00")))))
            .GetSheetAt(0);

        // "2026-09-13 14:30:45" (19 characters) against the serial "46,278.60" (9)
        Assert.IsTrue(asNumber.GetColumnWidth(0) < asDate.GetColumnWidth(0),
            $"{asNumber.GetColumnWidth(0)} vs {asDate.GetColumnWidth(0)}");
    }

    /// <summary>CSS measures a border in pixels or names its width.</summary>
    [DataTestMethod]
    [DataRow("thin solid #CCC", BorderStyle.Thin)]
    [DataRow("medium solid #CCC", BorderStyle.Medium)]
    [DataRow("thick solid #CCC", BorderStyle.Thick)]
    public void ABorderWidthCanBeNamed(string border, BorderStyle expected)
    {
        var sheet = Export(ExcelType.Xlsx, options => options.Styles.Cells(s => s.Border(border)), out _);
        Assert.AreEqual(expected, sheet.GetRow(1).GetCell(0).CellStyle.BorderTop);
    }

    [TestMethod]
    public void ABorderThatIsNoneSaysWhatIsWrong()
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options => options.Styles.Cells(s => s.Border("1px solid chartreuse")), out _));

        StringAssert.Contains(e.Message, "chartreuse");
        StringAssert.Contains(e.Message, "1px solid #D0D0D0");
    }

    /// <summary>
    ///     样式的缓存键必须把每个属性都算进去，哪怕第一个属性（字体色）没有设置——.NET Framework 的
    ///     string.Join(string, params object[]) 在首个元素为 null 时返回空字符串，曾使所有未设字体色的
    ///     样式共用同一个单元格样式，隔行底色、边框、加粗因此全部失效。
    /// </summary>
    [TestMethod]
    public void StylesWithoutAFontColourStillGetDistinctKeys()
    {
        var plain = new ExcelStyle();
        var filled = new ExcelStyle().Background("#F2F2F2");
        var bold = new ExcelStyle().Bold();

        Assert.AreNotEqual(plain.Key(), filled.Key());
        Assert.AreNotEqual(plain.Key(), bold.Key());
        Assert.AreNotEqual(filled.Key(), bold.Key());
        StringAssert.Contains(filled.Key(), "F2F2F2");
    }

    /// <summary>CSS names a handful of colours, and a stylesheet may as well take them.</summary>
    [TestMethod]
    public void AColourCanBeNamed()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.Styles.Cells(s => s.Border("2px dashed red")), out _);

        var style = (XSSFCellStyle) sheet.GetRow(1).GetCell(0).CellStyle;
        Assert.AreEqual(BorderStyle.MediumDashed, style.BorderTop);
        Assert.AreEqual("#FF0000", Hex(style.TopBorderXSSFColor));
    }

    [DataTestMethod]
    [DataRow("")]
    [DataRow("nope")]
    [DataRow("#12345")]
    public void AColourThatIsNoneSaysSo(string color)
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options => options.Styles.Cells(s => s.Background(color)), out _));

        StringAssert.Contains(e.Message, "#4472C4");
    }

    /// <summary>Nothing declared means nothing written, so an export without styles is what it always was.</summary>
    [TestMethod]
    public void WithoutStylesTheCellsKeepTheWorkbookDefault()
    {
        var sheet = Export(ExcelType.Xlsx, _ => { }, out var workbook);

        Assert.AreEqual(0, sheet.GetRow(0).GetCell(0).CellStyle.Index, "the header keeps the default style");
        Assert.AreEqual(0, sheet.GetRow(1).GetCell(1).CellStyle.Index, "so does a plain number cell");
        Assert.IsTrue(workbook.NumCellStyles < 5, $"{workbook.NumCellStyles} cell styles");
    }
}
