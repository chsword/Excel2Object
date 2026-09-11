using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Options;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     DateTime columns are exported as real date cells - a serial number with a date format - rather
///     than as text, so Excel can sort, filter and calculate with them.
/// </summary>
[TestClass]
public class DateCellExportTest : BaseExcelTest
{
    private static readonly DateTime When = new(2026, 9, 11, 14, 30, 45);

    public class Model
    {
        [ExcelTitle("Name")] public string Name { get; set; } = "a";

        [ExcelTitle("Plain")] public DateTime Plain { get; set; } = When;

        [ExcelColumn("Chinese", Format = "yyyy年MM月dd日")]
        public DateTime Chinese { get; set; } = When;

        [ExcelColumn("DayOnly", Format = "yyyy-MM-dd")]
        public DateTime? DayOnly { get; set; } = When.Date;

        [ExcelColumn("Unset", Format = "yyyy-MM-dd")]
        public DateTime Unset { get; set; }

        [ExcelTitle("UnsetPlain")] public DateTime UnsetPlain { get; set; }

        [ExcelColumn("ExcelSpelled", Format = "yyyy-mm-dd")]
        public DateTime ExcelSpelled { get; set; } = When;
    }

    private static byte[] ExportBytes(ExcelType excelType, Action<ExcelExporterOptions>? configure = null)
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(new List<Model> {new()}, options =>
        {
            options.ExcelType = excelType;
            configure?.Invoke(options);
        });
        Assert.IsNotNull(bytes);
        return bytes;
    }

    private static IRow Export(ExcelType excelType, Action<ExcelExporterOptions>? configure = null)
    {
        return FirstDataRow(ExportBytes(excelType, configure));
    }

    private static IRow FirstDataRow(byte[] bytes)
    {
        return WorkbookFactory.Create(new MemoryStream(bytes)).GetSheetAt(0).GetRow(1);
    }

    [TestMethod]
    public void DatesAreWrittenAsDateCellsInBothFileFormats()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var cell = Export(excelType).GetCell(1);
            Assert.AreEqual(CellType.Numeric, cell.CellType, excelType.ToString());
            Assert.IsTrue(DateUtil.IsCellDateFormatted(cell), excelType.ToString());
            Assert.AreEqual(When, cell.DateCellValue, excelType.ToString());
            Assert.AreEqual("yyyy-mm-dd hh:mm:ss", cell.CellStyle.GetDataFormatString(), excelType.ToString());
        }
    }

    [TestMethod]
    public void TheFormatIsTranslatedIntoExcelsSpelling()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var row = Export(excelType);
            Assert.AreEqual("yyyy\"年\"mm\"月\"dd\"日\"", row.GetCell(2).CellStyle.GetDataFormatString(),
                excelType.ToString());
            Assert.AreEqual("yyyy-mm-dd", row.GetCell(3).CellStyle.GetDataFormatString(), excelType.ToString());
            Assert.AreEqual(When.Date, row.GetCell(3).DateCellValue, excelType.ToString());
        }
    }

    /// <summary>
    ///     Excel's calendar starts in 1900, so default(DateTime) cannot be a date cell. It is kept as text
    ///     rendered with the column's Format - not silently turned into the #### Excel shows for a
    ///     negative serial number.
    /// </summary>
    [TestMethod]
    public void DatesBefore1900FallBackToText()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var row = Export(excelType);

            var formatted = row.GetCell(4);
            Assert.AreEqual(CellType.String, formatted.CellType, excelType.ToString());
            Assert.AreEqual("0001-01-01", formatted.StringCellValue, excelType.ToString());

            var plain = row.GetCell(5);
            Assert.AreEqual(CellType.String, plain.CellType, excelType.ToString());
            Assert.AreEqual(default(DateTime).ToString(), plain.StringCellValue, excelType.ToString());
        }
    }

    /// <summary>
    ///     A workbook on the 1904 date system starts four years later; a date Excel could otherwise hold
    ///     must become text there too, not the serial -1 NPOI would silently write.
    /// </summary>
    [TestMethod]
    public void The1904DateSystemMovesTheTextFallback()
    {
        var source = new XSSFWorkbook();
        var workbookPr = source.GetCTWorkbook().workbookPr ??= new NPOI.OpenXmlFormats.Spreadsheet.CT_WorkbookPr();
        workbookPr.date1904 = true;
        source.CreateSheet("Other");
        var sourceStream = new MemoryStream();
        source.Write(sourceStream, true);

        var model = new Model {Plain = new DateTime(1901, 6, 1), ExcelSpelled = new DateTime(1905, 1, 1)};
        var bytes = ExcelHelper.AppendObjectToExcelBytes(sourceStream.ToArray(), new List<Model> {model}, "Dates")!;
        var row = WorkbookFactory.Create(new MemoryStream(bytes)).GetSheet("Dates").GetRow(1);

        Assert.AreEqual(CellType.String, row.GetCell(1).CellType);
        Assert.AreEqual(model.Plain.ToString(), row.GetCell(1).StringCellValue);
        Assert.AreEqual(CellType.Numeric, row.GetCell(6).CellType);
        Assert.AreEqual(model.ExcelSpelled, row.GetCell(6).DateCellValue);
    }

    /// <summary>DateTimeAsText restores the export of v2.3 and earlier.</summary>
    [TestMethod]
    public void DateTimeAsTextWritesTheRenderedTextInstead()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var row = Export(excelType, options => options.DateTimeAsText = true);

            var chinese = row.GetCell(2);
            Assert.AreEqual(CellType.String, chinese.CellType, excelType.ToString());
            Assert.AreEqual("2026年09月11日", chinese.StringCellValue, excelType.ToString());
            Assert.AreEqual(0, chinese.CellStyle.DataFormat, excelType.ToString());

            var plain = row.GetCell(1);
            Assert.AreEqual(CellType.String, plain.CellType, excelType.ToString());
            Assert.AreEqual(When.ToString(), plain.StringCellValue, excelType.ToString());
        }
    }

    /// <summary>
    ///     .NET reads an Excel-spelled Format as minutes where Excel means months; rendered as text, such a
    ///     column gets the ISO form rather than "2026-30-11" - at the length the format shows.
    /// </summary>
    [TestMethod]
    public void AnExcelSpelledFormatIsNotRenderedByDotNet()
    {
        var row = Export(ExcelType.Xlsx, options => options.DateTimeAsText = true);
        Assert.AreEqual("2026-09-11", row.GetCell(6).StringCellValue);
    }

    [TestMethod]
    public void DateCellsRoundTripThroughTheImporter()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var back = ExcelHelper.ExcelToObject<Model>(ExportBytes(excelType)).Single();

            Assert.AreEqual(When, back.Plain, excelType.ToString());
            Assert.AreEqual(When, back.Chinese, excelType.ToString());
            Assert.AreEqual(When.Date, back.DayOnly, excelType.ToString());
            Assert.AreEqual(default, back.Unset, excelType.ToString());
        }
    }

    public class TextModel
    {
        [ExcelTitle("Plain")] public string Plain { get; set; } = "";
        [ExcelTitle("DayOnly")] public string DayOnly { get; set; } = "";
    }

    /// <summary>
    ///     Read as text - into a string property or a dictionary - a date cell must come out as a date,
    ///     not as the serial number Excel stores (46276.6...).
    /// </summary>
    [TestMethod]
    public void ADateCellReadAsTextIsNotASerialNumber()
    {
        var bytes = ExportBytes(ExcelType.Xlsx);

        var text = ExcelHelper.ExcelToObject<TextModel>(bytes).Single();
        Assert.AreEqual("2026-09-11 14:30:45", text.Plain);
        Assert.AreEqual("2026-09-11", text.DayOnly);

        var dictionary = ExcelHelper.ExcelToObject<Dictionary<string, object>>(bytes).Single();
        Assert.AreEqual("2026-09-11 14:30:45", dictionary["Plain"]);
        Assert.AreEqual("2026-09-11", dictionary["DayOnly"]);
    }

    public class NumberModel
    {
        [ExcelTitle("Serial")] public double Serial { get; set; }
        [ExcelTitle("Hours")] public decimal Hours { get; set; }
        [ExcelTitle("Clock")] public string Clock { get; set; } = "";
    }

    /// <summary>
    ///     A numeric property wants the number a date cell stores, whatever format the cell wears; a
    ///     string reads a time-only cell as a time, and an elapsed-time cell as the number it is.
    /// </summary>
    [TestMethod]
    public void ANumericPropertyStillGetsTheSerialNumber()
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Sheet1");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Serial");
        header.CreateCell(1).SetCellValue("Hours");
        header.CreateCell(2).SetCellValue("Clock");
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue(When);
        row.GetCell(0).CellStyle = DateStyle(workbook, "yyyy-mm-dd");
        row.CreateCell(1).SetCellValue(1.5);
        row.GetCell(1).CellStyle = DateStyle(workbook, "[h]:mm");
        row.CreateCell(2).SetCellValue(When);
        row.GetCell(2).CellStyle = DateStyle(workbook, "h:mm");
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        var bytes = stream.ToArray();

        var numbers = ExcelHelper.ExcelToObject<NumberModel>(bytes).Single();
        Assert.AreEqual(DateUtil.GetExcelDate(When), numbers.Serial);
        Assert.AreEqual(1.5m, numbers.Hours);
        Assert.AreEqual("14:30:45", numbers.Clock);

        var dictionary = ExcelHelper.ExcelToObject<Dictionary<string, object>>(bytes).Single();
        Assert.AreEqual("2026-09-11", dictionary["Serial"]);
        Assert.AreEqual("1.5", dictionary["Hours"]);
        Assert.AreEqual("14:30:45", dictionary["Clock"]);
    }

    private static ICellStyle DateStyle(IWorkbook workbook, string format)
    {
        var style = workbook.CreateCellStyle();
        style.DataFormat = workbook.CreateDataFormat().GetFormat(format);
        return style;
    }

    [TestMethod]
    public void DataTableDateColumnsAreDateCellsToo()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("When", typeof(DateTime));
        table.Rows.Add("a", When);
        table.Rows.Add("b", DBNull.Value);

        var bytes = ExcelHelper.ObjectToExcelBytes(table, ExcelType.Xlsx);
        Assert.IsNotNull(bytes);
        var sheet = WorkbookFactory.Create(new MemoryStream(bytes)).GetSheetAt(0);

        var cell = sheet.GetRow(1).GetCell(1);
        Assert.AreEqual(CellType.Numeric, cell.CellType);
        Assert.AreEqual(When, cell.DateCellValue);
        Assert.AreEqual("yyyy-mm-dd hh:mm:ss", cell.CellStyle.GetDataFormatString());
        Assert.AreEqual(CellType.Blank, sheet.GetRow(2).GetCell(1).CellType);
    }

    /// <summary>A DataTable export reaches the exporter options too.</summary>
    [TestMethod]
    public void DataTableExportTakesOptions()
    {
        var table = new DataTable();
        table.Columns.Add("When", typeof(DateTime));
        table.Rows.Add(When);

        var bytes = ExcelHelper.ObjectToExcelBytes(table, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.DateTimeAsText = true;
        });
        Assert.IsNotNull(bytes);

        var cell = FirstDataRow(bytes).GetCell(0);
        Assert.AreEqual(CellType.String, cell.CellType);
        Assert.AreEqual(When.ToString(), cell.StringCellValue);
    }

    /// <summary>A dictionary export types every column string, so the value has to decide.</summary>
    [TestMethod]
    public void DictionaryDateValuesAreDateCellsToo()
    {
        var rows = new List<Dictionary<string, object>>
        {
            new() {["Name"] = "a", ["When"] = When}
        };
        var bytes = ExcelHelper.ObjectToExcelBytes(rows, ExcelType.Xlsx);
        Assert.IsNotNull(bytes);

        var cell = FirstDataRow(bytes).GetCell(1);
        Assert.AreEqual(CellType.Numeric, cell.CellType);
        Assert.AreEqual(When, cell.DateCellValue);
        Assert.AreEqual("yyyy-mm-dd hh:mm:ss", cell.CellStyle.GetDataFormatString());
    }

    /// <summary>
    ///     The width of an auto-sized date column follows what the cell displays, i.e. its Format.
    /// </summary>
    [TestMethod]
    public void AutoColumnWidthFitsTheDisplayedDate()
    {
        var bytes = ExportBytes(ExcelType.Xlsx, options =>
        {
            options.AutoColumnWidth = true;
            options.MinColumnWidth = 1;
        });
        var sheet = WorkbookFactory.Create(new MemoryStream(bytes)).GetSheetAt(0);

        // "2026-09-11 14:30:45" (19 characters) against "2026-09-11" (10)
        Assert.IsTrue(sheet.GetColumnWidth(1) > sheet.GetColumnWidth(3),
            $"{sheet.GetColumnWidth(1)} vs {sheet.GetColumnWidth(3)}");
    }
}
