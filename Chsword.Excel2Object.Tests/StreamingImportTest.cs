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
///     流式导入：逐行读出 .xlsx 的一张工作表，不把整个工作簿建进内存。转换那一套与整份读入共用，
///     故这里验的是「读出来的东西一样」，以及流式独有的几处：公式取缓存值、表名定位、惰性取用。
/// </summary>
[TestClass]
public class StreamingImportTest
{
    public enum Status
    {
        待付款 = 1,
        已发货 = 2
    }

    public class Model
    {
        [ExcelColumn("文本")] public string Text { get; set; } = "";
        [ExcelColumn("金额", Format = "#,##0.00")] public decimal Amount { get; set; }
        [ExcelColumn("数量")] public int Count { get; set; }
        [ExcelColumn("下单时间", Format = "yyyy-MM-dd HH:mm")] public DateTime PlacedAt { get; set; }
        [ExcelColumn("完成时间")] public DateTime? FinishedAt { get; set; }
        [ExcelColumn("已付")] public bool Paid { get; set; }
        [ExcelColumn("状态")] public Status Status { get; set; }
        [ExcelColumn("链接")] public Uri? Link { get; set; }
    }

    private static List<Model> Rows(int count)
    {
        return Enumerable.Range(0, count).Select(i => new Model
        {
            Text = i % 3 == 0 ? "值" + i : "重复的值",
            Amount = 100.5m + i,
            Count = i,
            PlacedAt = new DateTime(2026, 1, 1, 9, 30, 0).AddMinutes(i),
            FinishedAt = i % 2 == 0 ? new DateTime(2026, 2, 1).AddDays(i) : null,
            Paid = i % 2 == 0,
            Status = i % 2 == 0 ? Status.已发货 : Status.待付款,
            Link = new Uri("https://example.com/" + i)
        }).ToList();
    }

    private static byte[] Export(IEnumerable<Model> rows, ExcelType excelType = ExcelType.Xlsx,
        string? sheetTitle = null)
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(rows, options =>
        {
            options.ExcelType = excelType;
            options.SheetTitle = sheetTitle;
        });
        Assert.IsNotNull(bytes);
        return bytes!;
    }

    private static List<Model> Streamed(byte[] bytes, Action<ExcelImporterOptions>? configure = null)
    {
        using var input = new MemoryStream(bytes);
        return new ExcelImporter().ExcelStreamToObject<Model>(input, configure).ToList();
    }

    private static void AssertSame(IReadOnlyList<Model> expected, IReadOnlyList<Model> actual)
    {
        Assert.AreEqual(expected.Count, actual.Count, "行数");
        for (var i = 0; i < expected.Count; i++)
        {
            Assert.AreEqual(expected[i].Text, actual[i].Text, $"第 {i} 行的文本");
            Assert.AreEqual(expected[i].Amount, actual[i].Amount, $"第 {i} 行的金额");
            Assert.AreEqual(expected[i].Count, actual[i].Count, $"第 {i} 行的数量");
            Assert.AreEqual(expected[i].PlacedAt, actual[i].PlacedAt, $"第 {i} 行的下单时间");
            Assert.AreEqual(expected[i].FinishedAt, actual[i].FinishedAt, $"第 {i} 行的完成时间");
            Assert.AreEqual(expected[i].Paid, actual[i].Paid, $"第 {i} 行的已付");
            Assert.AreEqual(expected[i].Status, actual[i].Status, $"第 {i} 行的状态");
            Assert.AreEqual(expected[i].Link, actual[i].Link, $"第 {i} 行的链接");
        }
    }

    [TestMethod]
    public void StreamedImportMatchesTheInMemoryOne()
    {
        // 共享字符串表版本：内存导出写出的文件把文本收进 sharedStrings.xml
        var bytes = Export(Rows(50));
        var inMemory = new ExcelImporter().ExcelToObject<Model>(bytes).ToList();

        AssertSame(inMemory, Streamed(bytes));
        AssertSame(Rows(50), Streamed(bytes));
    }

    [TestMethod]
    public void InlineStringsAreReadToo()
    {
        // 流式导出写出的文件把文本直接写在单元格里，不进共享字符串表
        using var output = new MemoryStream();
        new ExcelExporter().ObjectToExcelStream(Rows(30), output, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.StreamingRowWindow = 5;
        });

        var bytes = output.ToArray();
        Assert.IsFalse(Parts(bytes).Contains("xl/sharedStrings.xml"), "前提：该文件确实没有共享字符串表");
        AssertSame(Rows(30), Streamed(bytes));
    }

    [TestMethod]
    public void BlanksAndGapsKeepTheirColumns()
    {
        var rows = Rows(4);
        rows[1].Text = "";
        rows[2].Link = null;
        rows[3].FinishedAt = null;

        var streamed = Streamed(Export(rows));
        Assert.AreEqual("", streamed[1].Text);
        Assert.IsNull(streamed[2].Link);
        Assert.IsNull(streamed[3].FinishedAt);
        // 空白格不影响同一行上其他列的对位
        Assert.AreEqual(rows[2].Amount, streamed[2].Amount);
        Assert.AreEqual(rows[3].Status, streamed[3].Status);
    }

    [TestMethod]
    public void TheSheetIsFoundByTitle()
    {
        var first = Export(Rows(3), sheetTitle: "上月");
        var both = new ExcelExporter().AppendObjectToExcelBytes(first, Rows(6), options =>
        {
            options.SheetTitle = "本月";
            options.ExcelType = ExcelType.Xlsx;
        });
        Assert.IsNotNull(both);

        Assert.AreEqual(3, Streamed(both!, o => o.SheetTitle = "上月").Count);
        Assert.AreEqual(6, Streamed(both!, o => o.SheetTitle = "本月").Count);
        // 不指定表名时取第一张
        Assert.AreEqual(3, Streamed(both!).Count);
    }

    [TestMethod]
    public void AnUnknownSheetIsRefusedBeforeAnyRowIsRead()
    {
        using var input = new MemoryStream(Export(Rows(3)));

        // 表名写错当场即报，不必等到取第一行
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            new ExcelImporter().ExcelStreamToObject<Model>(input, o => o.SheetTitle = "没有这张表"));
        StringAssert.Contains(e.Message, "没有这张表");
    }

    [TestMethod]
    public void TitleSkipLineIsHonoured()
    {
        // 表头之上多出两行说明文字
        var bytes = Export(Rows(5));
        var workbook = new XSSFWorkbook(new MemoryStream(bytes));
        var sheet = workbook.GetSheetAt(0);
        sheet.ShiftRows(0, sheet.LastRowNum, 2);
        sheet.CreateRow(0).CreateCell(0).SetCellValue("导出说明");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("第二行说明");
        using var shifted = new MemoryStream();
        workbook.Write(shifted, true);

        var rows = Streamed(shifted.ToArray(), o => o.TitleSkipLine = 2);
        Assert.AreEqual(5, rows.Count);
        Assert.AreEqual("值0", rows[0].Text);
    }

    [TestMethod]
    public void AFormulaCellReadsItsStoredResult()
    {
        // Excel 存盘时会写下公式的计算结果，流式读取用的正是这个值
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("数据");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("文本");
        header.CreateCell(1).SetCellValue("数量");
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue("甲");
        var formula = row.CreateCell(1);
        formula.SetCellFormula("2*21");
        formula.SetCellValue(42);
        using var bytes = new MemoryStream();
        workbook.Write(bytes, true);

        var streamed = Streamed(bytes.ToArray());
        Assert.AreEqual(42, streamed[0].Count);
    }

    /// <summary>公式列对应到一个可空属性，好看清「读作空白」究竟读成了什么。</summary>
    public class FormulaModel
    {
        [ExcelColumn("文本")] public string Text { get; set; } = "";
        [ExcelColumn("数量")] public int? Count { get; set; }
    }

    [TestMethod]
    public void AFormulaWithoutAStoredResultReadsAsBlank()
    {
        // 本库导出的文件里公式没有算出的结果，要等 Excel 打开时才算
        var bytes = new ExcelExporter().ObjectToExcelBytes(
            new[] {new FormulaModel {Text = "甲"}, new FormulaModel {Text = "乙"}}, options =>
            {
                options.ExcelType = ExcelType.Xlsx;
                options.FormulaColumns.Add(new FormulaColumn {Title = "数量", Formula = c => c["文本"]});
            });
        Assert.IsNotNull(bytes);

        var errors = new List<ExcelImportError>();
        using var input = new MemoryStream(bytes!);
        var streamed = new ExcelImporter()
            .ExcelStreamToObject<FormulaModel>(input, options => options.OnCellError = errors.Add).ToList();

        Assert.AreEqual(2, streamed.Count);
        Assert.IsNull(streamed[0].Count, "没有存下结果的公式读作空白");
        // 空白不是读取失败：这样的格常写成一个空的 <v/>，不应当作读不出来上报
        Assert.AreEqual(0, errors.Count, string.Join("；", errors.Select(e => e.Exception.Message)));
        // 其余各列照常
        Assert.AreEqual("甲", streamed[0].Text);
    }

    [TestMethod]
    public void ABlankCellBehavesTheSameWayAsWhenReadWhole()
    {
        // 空白格落在非可空的数值列上，两条路都以转换失败中止——此处只验二者一致
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("数据");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("文本");
        header.CreateCell(1).SetCellValue("数量");
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue("甲");
        row.CreateCell(1).SetBlank();
        using var bytes = new MemoryStream();
        workbook.Write(bytes, true);

        Assert.ThrowsException<FormatException>(() =>
            new ExcelImporter().ExcelToObject<Model>(bytes.ToArray()).ToList());
        Assert.ThrowsException<FormatException>(() => Streamed(bytes.ToArray()));

        // 同一格落在可空列上则读作 null
        using var input = new MemoryStream(bytes.ToArray());
        var nullable = new ExcelImporter().ExcelStreamToObject<FormulaModel>(input).ToList();
        Assert.IsNull(nullable[0].Count);
    }

    [TestMethod]
    public void PrefixedElementsMissingReferencesAndEmptyRowsAreRead()
    {
        // 带前缀的写法（<x:row>）同样合法，行与格的 r 属性也可以不写，空行可以自闭合。
        // 这类文件由别的工具写出，此处手工造一份。
        var bytes = HandWritten();

        var rows = Streamed(bytes, options => options.TitleSkipLine = 1);
        Assert.AreEqual(2, rows.Count);
        Assert.AreEqual("甲", rows[0].Text);
        Assert.AreEqual(12, rows[0].Count);
        Assert.AreEqual("乙", rows[1].Text);
        Assert.AreEqual(34, rows[1].Count);
    }

    /// <summary>一份手写的 .xlsx：元素带前缀、行与格不写 r、开头有一个自闭合的空行。</summary>
    private static byte[] HandWritten()
    {
        const string main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string rels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                "</Types>",
            ["_rels/.rels"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rIdWb\" Type=\"" + rels + "/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                "</Relationships>",
            ["xl/_rels/workbook.xml.rels"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"" + rels + "/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                "</Relationships>",
            ["xl/workbook.xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<x:workbook xmlns:x=\"" + main + "\" xmlns:r=\"" + rels + "\">" +
                "<x:sheets><x:sheet name=\"数据\" sheetId=\"1\" r:id=\"rId1\"/></x:sheets></x:workbook>",
            ["xl/worksheets/sheet1.xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<x:worksheet xmlns:x=\"" + main + "\"><x:sheetData>" +
                "<x:row/>" +
                "<x:row><x:c t=\"inlineStr\"><x:is><x:t>文本</x:t></x:is></x:c>" +
                "<x:c t=\"inlineStr\"><x:is><x:t>数量</x:t></x:is></x:c></x:row>" +
                "<x:row><x:c t=\"inlineStr\"><x:is><x:t>甲</x:t></x:is></x:c><x:c><x:v>12</x:v></x:c></x:row>" +
                "<x:row><x:c t=\"inlineStr\"><x:is><x:t>乙</x:t></x:is></x:c><x:c><x:v>34</x:v></x:c></x:row>" +
                "</x:sheetData></x:worksheet>"
        };

        using var buffer = new MemoryStream();
        using (var zip = new System.IO.Compression.ZipArchive(buffer, System.IO.Compression.ZipArchiveMode.Create, true))
            foreach (var part in parts)
            {
                using var entry = new StreamWriter(zip.CreateEntry(part.Key).Open());
                entry.Write(part.Value);
            }

        return buffer.ToArray();
    }

    [TestMethod]
    public void ReadFailuresAreReportedWithTheirPosition()
    {
        var bytes = Export(Rows(3));
        var workbook = new XSSFWorkbook(new MemoryStream(bytes));
        // 第 3 行的「下单时间」改写成一段不是日期的文本
        var cell = workbook.GetSheetAt(0).GetRow(3).GetCell(3);
        cell.SetCellType(CellType.String);
        cell.SetCellValue("待定");
        using var broken = new MemoryStream();
        workbook.Write(broken, true);

        var errors = new List<ExcelImportError>();
        using var input = new MemoryStream(broken.ToArray());
        var rows = new ExcelImporter()
            .ExcelStreamToObject<Model>(input, options => options.OnCellError = errors.Add).ToList();

        Assert.AreEqual(3, rows.Count);
        Assert.AreEqual(1, errors.Count);
        Assert.AreEqual("D4", errors[0].CellReference);
        StringAssert.Contains(errors[0].Exception.Message, "待定");
        Assert.AreEqual(default, rows[2].PlacedAt);
    }

    [TestMethod]
    public void RowsAreTakenOneAtATime()
    {
        // 惰性取用：取够了就可以不再往下读
        using var input = new MemoryStream(Export(Rows(500)));
        var first = new ExcelImporter().ExcelStreamToObject<Model>(input).Take(3).ToList();

        Assert.AreEqual(3, first.Count);
        Assert.AreEqual("值0", first[0].Text);
    }

    [TestMethod]
    public void XlsGoesThroughTheSameEntry()
    {
        // .xls 无从逐行读出，整份读入后结果一致
        AssertSame(Rows(10), Streamed(Export(Rows(10), ExcelType.Xls)));
    }

    [TestMethod]
    public void TheDictionaryFormIsReadToo()
    {
        using var input = new MemoryStream(Export(Rows(3)));
        var rows = new ExcelImporter().ExcelStreamToObject<Dictionary<string, object>>(input).ToList();

        Assert.AreEqual(3, rows.Count);
        Assert.AreEqual("值0", rows[0]["文本"]);
        Assert.AreEqual("100.5", rows[0]["金额"]);
        // 日期列读作它显示的样子，而非 Excel 存的序列号
        Assert.AreEqual("2026-01-01 09:30:00", rows[0]["下单时间"]);
    }

    private static string[] Parts(byte[] bytes)
    {
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        return archive.Entries.Select(e => e.FullName).ToArray();
    }
}
