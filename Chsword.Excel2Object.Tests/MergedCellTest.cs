using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Options;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     合并单元格：某列中连续相同的值并成一格，以及合并调用方指定的区域。
/// </summary>
[TestClass]
public class MergedCellTest : BaseExcelTest
{
    public class Model
    {
        [ExcelTitle("省份")] public string Province { get; set; } = "";
        [ExcelTitle("城市")] public string City { get; set; } = "";
        [ExcelTitle("金额")] public decimal Amount { get; set; }
    }

    private static readonly List<Model> Rows = new()
    {
        new() {Province = "广东", City = "广州", Amount = 100},
        new() {Province = "广东", City = "广州", Amount = 200},
        new() {Province = "广东", City = "深圳", Amount = 300},
        new() {Province = "江苏", City = "南京", Amount = 400}
    };

    private static ISheet Export(ExcelType excelType, Action<ExcelExporterOptions> configure,
        List<Model>? rows = null)
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(rows ?? Rows, options =>
        {
            options.ExcelType = excelType;
            configure(options);
        });
        Assert.IsNotNull(bytes);
        return WorkbookFactory.Create(new MemoryStream(bytes)).GetSheetAt(0);
    }

    private static string[] Regions(ISheet sheet)
    {
        return sheet.MergedRegions.Select(r => r.FormatAsString()).OrderBy(r => r, StringComparer.Ordinal).ToArray();
    }

    [TestMethod]
    public void RepeatedValuesMergeInBothFileFormats()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var sheet = Export(excelType, options => options.MergeRepeatedColumns.Add("省份"));

            // 广东占第 2 至第 4 行，江苏只有一行故不合并
            CollectionAssert.AreEqual(new[] {"A2:A4"}, Regions(sheet), excelType.ToString());
        }
    }

    /// <summary>各列彼此独立判断。</summary>
    [TestMethod]
    public void EachColumnIsJudgedOnItsOwn()
    {
        var sheet = Export(ExcelType.Xlsx, options =>
        {
            options.MergeRepeatedColumns.Add("省份");
            options.MergeRepeatedColumns.Add("城市");
        });

        CollectionAssert.AreEqual(new[] {"A2:A4", "B2:B3"}, Regions(sheet));
    }

    /// <summary>被并入的单元格仍保留各自的值，导回对象时每一行的数据依然完整。</summary>
    [TestMethod]
    public void TheRowsStillImportInFull()
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(Rows, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.MergeRepeatedColumns.Add("省份");
        })!;

        var back = ExcelHelper.ExcelToObject<Model>(bytes).ToArray();
        Assert.AreEqual(4, back.Length);
        CollectionAssert.AreEqual(new[] {"广东", "广东", "广东", "江苏"}, back.Select(m => m.Province).ToArray());
        Assert.AreEqual(300m, back[2].Amount);
    }

    /// <summary>相同的数值同样合并，且不受文本表示的影响。</summary>
    [TestMethod]
    public void NumbersMergeToo()
    {
        var rows = new List<Model>
        {
            new() {Province = "甲", City = "A", Amount = 100},
            new() {Province = "乙", City = "B", Amount = 100},
            new() {Province = "丙", City = "C", Amount = 250}
        };
        var sheet = Export(ExcelType.Xlsx, options => options.MergeRepeatedColumns.Add("金额"), rows);

        CollectionAssert.AreEqual(new[] {"C2:C3"}, Regions(sheet));
    }

    /// <summary>空单元格不参与合并：几行都没有值并不意味着它们是同一组。</summary>
    [TestMethod]
    public void BlankCellsDoNotMerge()
    {
        var rows = new List<Model>
        {
            new() {Province = "", City = "A", Amount = 1},
            new() {Province = "", City = "B", Amount = 2},
            new() {Province = "丙", City = "C", Amount = 3}
        };
        var sheet = Export(ExcelType.Xlsx, options => options.MergeRepeatedColumns.Add("省份"), rows);

        Assert.AreEqual(0, sheet.NumMergedRegions);
    }

    /// <summary>只有一行数据时无从合并。</summary>
    [TestMethod]
    public void ASingleRowMergesNothing()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.MergeRepeatedColumns.Add("省份"),
            new List<Model> {new() {Province = "广东", City = "广州", Amount = 1}});

        Assert.AreEqual(0, sheet.NumMergedRegions);
    }

    /// <summary>
    ///     列以标题指定：列在表中的位置由 Order、特性顺序与公式列的插入位置决定，编写导出代码时
    ///     无从得知，故不必数到 A、B、C。
    /// </summary>
    [TestMethod]
    public void ARegionIsWrittenWithColumnTitles()
    {
        var sheet = Export(ExcelType.Xlsx, options =>
            options.MergedRegions.Add(new MergedRegion("省份", "金额")));

        // 省略行即指表头行
        CollectionAssert.AreEqual(new[] {"A1:C1"}, Regions(sheet));
    }

    /// <summary>行以数据行的序号指定，自 1 计起。</summary>
    [TestMethod]
    public void RowsAreWrittenAsDataRowOrdinals()
    {
        var sheet = Export(ExcelType.Xlsx, options =>
            options.MergedRegions.Add(new MergedRegion("金额") {FirstRow = 1, LastRow = 3}));

        CollectionAssert.AreEqual(new[] {"C2:C4"}, Regions(sheet));
    }

    /// <summary>单独一行也可以，省略结束行即可。</summary>
    [TestMethod]
    public void ASingleDataRowCanSpanColumns()
    {
        var sheet = Export(ExcelType.Xlsx, options =>
            options.MergedRegions.Add(new MergedRegion("省份", "城市") {FirstRow = 2}));

        CollectionAssert.AreEqual(new[] {"A3:B3"}, Regions(sheet));
    }

    /// <summary>确已知道布局时，仍可直接写地址。</summary>
    [TestMethod]
    public void AnAddressStillWorks()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.MergedRegions.Add("A1:C1"));
        CollectionAssert.AreEqual(new[] {"A1:C1"}, Regions(sheet));
    }

    [TestMethod]
    public void ARowBeyondTheDataSaysSo()
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options =>
                options.MergedRegions.Add(new MergedRegion("金额") {FirstRow = 1, LastRow = 99})));

        StringAssert.Contains(e.Message, "4 行数据");
    }

    [TestMethod]
    public void ARowOrdinalBelowOneSaysSo()
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options => options.MergedRegions.Add(new MergedRegion("金额") {FirstRow = 0})));

        StringAssert.Contains(e.Message, "自 1 计起");
    }

    [TestMethod]
    public void AnUnknownColumnSaysSo()
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options => options.MergeRepeatedColumns.Add("不存在")));

        StringAssert.Contains(e.Message, "不存在");
    }

    [TestMethod]
    public void ARangeThatIsNoneSaysSo()
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options => options.MergedRegions.Add("A1-C1")));

        StringAssert.Contains(e.Message, "A1:C1");
    }

    public class FormulaModel
    {
        [ExcelTitle("数量")] public int Count { get; set; }
        [ExcelTitle("合计")] public double Total { get; set; }
    }

    /// <summary>
    ///     公式列写进单元格的是公式，其结果由 Excel 打开时才算出，导出时无从比较，故明确拒绝而非
    ///     按公式文本比较——后者会因每行的引用不同而永不合并。
    /// </summary>
    [TestMethod]
    public void AFormulaColumnCannotMergeByValue()
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            new ExcelExporter().ObjectToExcelBytes(
                new List<FormulaModel> {new() {Count = 1}, new() {Count = 1}}, options =>
                {
                    options.ExcelType = ExcelType.Xlsx;
                    options.FormulaColumns.Add("合计", c => c["数量"]);
                    options.MergeRepeatedColumns.Add("合计");
                }));

        StringAssert.Contains(e.Message, "合计");
        StringAssert.Contains(e.Message, "公式");
    }

    /// <summary>同一列写两次不应产生两段相同的区域（那将彼此重叠）。</summary>
    [TestMethod]
    public void TheSameColumnTwiceIsHarmless()
    {
        var sheet = Export(ExcelType.Xlsx, options =>
        {
            options.MergeRepeatedColumns.Add("省份");
            options.MergeRepeatedColumns.Add("省份");
        });

        CollectionAssert.AreEqual(new[] {"A2:A4"}, Regions(sheet));
    }

    /// <summary>
    ///     分组表动辄数万行，合并区域的数量与行数同阶。此处确认其耗时不随行数平方增长——
    ///     NPOI 自带的 AddMergedRegion 会对每个新区域扫描已有全部区域，8000 个区域即需数秒。
    /// </summary>
    [TestMethod]
    public void ManyRunsStayFast()
    {
        var rows = new List<Model>();
        for (var i = 0; i < 20000; i++)
            rows.Add(new Model {Province = $"省{i / 2}", City = "甲", Amount = i});

        var watch = System.Diagnostics.Stopwatch.StartNew();
        var sheet = Export(ExcelType.Xlsx, options => options.MergeRepeatedColumns.Add("省份"), rows);
        watch.Stop();

        Assert.AreEqual(10000, sheet.NumMergedRegions);
        Assert.IsTrue(watch.ElapsedMilliseconds < 20000, $"耗时 {watch.ElapsedMilliseconds} ms");
    }

    public class DoubleModel
    {
        [ExcelTitle("数值")] public double Value { get; set; }
    }

    /// <summary>
    ///     数值直接比较其值，不经由文本：.NET Framework 上 "R" 与 G17 两种格式都可能把两个不同的数
    ///     渲染成同一串字符。
    ///     <para>
    ///     两个取值选在 13 位有效数字上，而非相邻的两个 double：decimal 转 double、以及 NPOI 把
    ///     double 写进 XML，在 .NET Framework 上都只保留约 15 位有效数字，需要 17 位才能分辨的两个数
    ///     在文件里就已经相同，那样测到的便不是本意了。
    ///     </para>
    /// </summary>
    [TestMethod]
    public void NumbersThatOnlyLookAlikeDoNotMerge()
    {
        var rows = new List<DoubleModel>
        {
            new() {Value = 1.000000000001d},
            new() {Value = 1.000000000002d},
            new() {Value = 2d}
        };
        var bytes = new ExcelExporter().ObjectToExcelBytes(rows, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.MergeRepeatedColumns.Add("数值");
        })!;
        var sheet = WorkbookFactory.Create(new MemoryStream(bytes)).GetSheetAt(0);

        Assert.AreNotEqual(sheet.GetRow(1).GetCell(0).NumericCellValue,
            sheet.GetRow(2).GetCell(0).NumericCellValue, "前提：两个单元格装的确实是不同的数");
        Assert.AreEqual(0, sheet.NumMergedRegions, "两个数并不相同，不应合并");
    }

    /// <summary>只给结束行会把表头一并并进去，故要求一并给出起始行。</summary>
    [TestMethod]
    public void OnlyALastRowSaysSo()
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options => options.MergedRegions.Add(new MergedRegion("金额") {LastRow = 3})));

        StringAssert.Contains(e.Message, "FirstRow");
    }

    /// <summary>只占一个单元格的区域无从合并。</summary>
    [DataTestMethod]
    [DataRow("A1:A1")]
    [DataRow(null)]
    public void ASingleCellRegionSaysSo(string? address)
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options => options.MergedRegions.Add(
                address == null ? new MergedRegion("省份") {FirstRow = 1} : address)));

        StringAssert.Contains(e.Message, "一个单元格");
    }

    /// <summary>重叠的区域会被 Excel 视为损坏，故提前拒绝并指出与哪一个重叠。</summary>
    [TestMethod]
    public void OverlappingRegionsAreRejected()
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options =>
            {
                options.MergeRepeatedColumns.Add("省份");
                options.MergedRegions.Add(new MergedRegion("省份") {FirstRow = 2, LastRow = 3});
            }));

        StringAssert.Contains(e.Message, "A2:A4");
    }
}
