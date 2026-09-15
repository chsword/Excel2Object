using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Options;
using Chsword.Excel2Object.Styles;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     流式导出：边取数据边写，内存中只留若干行。已刷出内存的行无从读回，故凡是要回看数据的特性
///     （自动列宽、按值合并）都改在写入过程中算出，这里逐一验证其结果与内存导出一致。
/// </summary>
[TestClass]
public class StreamingExportTest
{
    /// <summary>刻意远小于行数，好让写入过程中确有大批行被刷出内存。</summary>
    private const int Window = 5;

    public class Model
    {
        [ExcelColumn("省份")] public string Province { get; set; } = "";
        [ExcelColumn("城市")] public string City { get; set; } = "";
        [ExcelColumn("金额", Format = "#,##0.00")] public decimal Amount { get; set; }
        [ExcelColumn("下单时间", Format = "yyyy-MM-dd")] public DateTime PlacedAt { get; set; }
        [ExcelColumn("状态")] public string Status { get; set; } = "";
    }

    private static List<Model> Rows(int count)
    {
        return Enumerable.Range(0, count).Select(i => new Model
        {
            // 每 10 行换一个省份、每 5 行换一个城市，于是合并区域必然跨越被刷出内存的行
            Province = "省" + (i / 10),
            City = "市" + (i / 5),
            Amount = 100 + i,
            PlacedAt = new DateTime(2026, 1, 1).AddDays(i),
            Status = i % 2 == 0 ? "已发货" : "待付款"
        }).ToList();
    }

    private static ISheet Streamed(IEnumerable<Model> rows, Action<ExcelExporterOptions>? configure = null,
        int window = Window)
    {
        using var output = new MemoryStream();
        new ExcelExporter().ObjectToExcelStream(rows, output, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.StreamingRowWindow = window;
            configure?.Invoke(options);
        });

        return WorkbookFactory.Create(new MemoryStream(output.ToArray())).GetSheetAt(0);
    }

    private static ISheet InMemory(IEnumerable<Model> rows, Action<ExcelExporterOptions>? configure = null)
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(rows, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            configure?.Invoke(options);
        });
        Assert.IsNotNull(bytes);
        return WorkbookFactory.Create(new MemoryStream(bytes!)).GetSheetAt(0);
    }

    /// <summary>合并区域，按列再按行排序，与写入顺序一致。</summary>
    private static string[] Regions(ISheet sheet)
    {
        return sheet.MergedRegions.OrderBy(r => r.FirstColumn).ThenBy(r => r.FirstRow)
            .Select(r => r.FormatAsString()).ToArray();
    }

    /// <summary>逐格比对两张表：值、数字格式、字体与底色都取自文件，而非只看值。</summary>
    private static void AssertSameSheet(ISheet expected, ISheet actual)
    {
        Assert.AreEqual(expected.LastRowNum, actual.LastRowNum, "行数");
        for (var r = 0; r <= expected.LastRowNum; r++)
        {
            var expectedRow = expected.GetRow(r);
            var actualRow = actual.GetRow(r);
            Assert.AreEqual(expectedRow.LastCellNum, actualRow.LastCellNum, $"第 {r} 行的列数");
            for (var c = 0; c < expectedRow.LastCellNum; c++)
            {
                var a = expectedRow.GetCell(c);
                var b = actualRow.GetCell(c);
                Assert.AreEqual(a.CellType, b.CellType, $"{a.Address} 的类型");
                Assert.AreEqual(a.ToString(), b.ToString(), $"{a.Address} 的值");
                Assert.AreEqual(a.CellStyle.GetDataFormatString(), b.CellStyle.GetDataFormatString(),
                    $"{a.Address} 的数字格式");
                Assert.AreEqual(a.CellStyle.FillForegroundColor, b.CellStyle.FillForegroundColor,
                    $"{a.Address} 的底色");
                Assert.AreEqual(a.CellStyle.GetFont(expected.Workbook).IsBold,
                    b.CellStyle.GetFont(actual.Workbook).IsBold, $"{a.Address} 的字体");
            }
        }
    }

    [TestMethod]
    public void StreamedFileMatchesTheInMemoryOne()
    {
        var rows = Rows(120);
        Action<ExcelExporterOptions> configure = options => options.Styles
            .Header(s => s.Bold().Background("#4472C4").Color("#FFFFFF"))
            .EvenRows(s => s.Background("#F2F2F2"))
            .Column("金额", s => s.Right());

        AssertSameSheet(InMemory(rows, configure), Streamed(rows, configure));
    }

    [TestMethod]
    public void MergesSpanRowsAlreadyFlushedFromMemory()
    {
        var rows = Rows(30);
        Action<ExcelExporterOptions> configure = options =>
        {
            options.MergeRepeatedColumns.Add("省份");
            options.MergeRepeatedColumns.Add("城市");
        };

        // 每段 10 行、窗口 5 行：任何一段都跨越已刷出内存的行
        var streamed = Regions(Streamed(rows, configure));
        var provinces = streamed.Where(r => r.StartsWith("A")).ToArray();
        CollectionAssert.AreEqual(new[] {"A2:A11", "A12:A21", "A22:A31"}, provinces, string.Join(" ", streamed));
        CollectionAssert.AreEqual(Regions(InMemory(rows, configure)), streamed, string.Join(" ", streamed));
    }

    [TestMethod]
    public void DeclaredRegionsAreResolvedAgainstTheRowsWritten()
    {
        var sheet = Streamed(Rows(30), options =>
            options.MergedRegions.Add(new MergedRegion("状态") {FirstRow = 1, LastRow = 30}));

        CollectionAssert.AreEqual(new[] {"E2:E31"}, Regions(sheet));
    }

    [TestMethod]
    public void HeaderViewDropdownAndConditionalFormatSurvive()
    {
        var sheet = Streamed(Rows(40), options =>
        {
            options.FreezeHeader = true;
            options.AutoFilter = true;
            options.Dropdowns["状态"] = new[] {"已发货", "待付款", "已取消"};
            options.ConditionalFormats.Add(new ConditionalFormat("金额")
            {
                Operator = ConditionalOperator.GreaterThan,
                Value = 120,
                FontColor = ExcelStyleColor.Red
            });
        });

        Assert.AreEqual(1, sheet.PaneInformation.HorizontalSplitPosition, "冻结位置");
        Assert.AreEqual(1, sheet.GetDataValidations().Count, "数据验证");
        Assert.AreEqual(1, sheet.SheetConditionalFormatting.NumConditionalFormattings, "条件格式");
        Assert.IsNotNull(sheet.Workbook.GetName("_xlnm._FilterDatabase"), "筛选区域");
    }

    [TestMethod]
    public void TwoLongDropdownsBothLandOnTheHiddenSheet()
    {
        // 超过 255 字符的列表放不进验证本身，要写到隐藏表上。两张列表即两列，流式导出的工作表只能
        // 自上而下写一遍，故须一次写完，否则后写的那一列会丢。
        var first = Enumerable.Range(0, 60).Select(i => "甲选项" + i).ToArray();
        var second = Enumerable.Range(0, 80).Select(i => "乙选项" + i).ToArray();

        using var output = new MemoryStream();
        new ExcelExporter().ObjectToExcelStream(Rows(20), output, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.StreamingRowWindow = Window;
            options.Dropdowns["省份"] = first;
            options.Dropdowns["城市"] = second;
        });

        var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray()));
        var lists = workbook.GetSheet("_excel2object_lists");
        Assert.IsNotNull(lists, "隐藏的列表工作表");
        for (var i = 0; i < second.Length; i++)
        {
            var row = lists.GetRow(i);
            Assert.AreEqual(i < first.Length ? first[i] : null, row.GetCell(0)?.StringCellValue, $"第 {i} 行第一张列表");
            Assert.AreEqual(second[i], row.GetCell(1)?.StringCellValue, $"第 {i} 行第二张列表");
        }
    }

    [TestMethod]
    public void ALongDropdownGoesToItsOwnListSheetWhenAppending()
    {
        // 源工作簿里已有一张写着长列表的隐藏表。流式写入无从往那张表里再添一列——它的行不在流式
        // 这一层里，既读不回也无从从首行重写——故应另起一张，原有的列表原样保留。
        var existing = Enumerable.Range(0, 60).Select(i => "旧选项" + i).ToArray();
        var added = Enumerable.Range(0, 70).Select(i => "新选项" + i).ToArray();

        var first = new ExcelExporter().ObjectToExcelBytes(Rows(5), options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.SheetTitle = "上月";
            options.Dropdowns["省份"] = existing;
        });

        using var output = new MemoryStream();
        new ExcelExporter().ObjectToExcelStream(Rows(20), output, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.StreamingRowWindow = Window;
            options.SheetTitle = "本月";
            options.SourceExcelBytes = first;
            options.Dropdowns["城市"] = added;
        });

        var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray()));
        var old = workbook.GetSheet("_excel2object_lists");
        var fresh = workbook.GetSheet("_excel2object_lists2");
        Assert.IsNotNull(fresh, "应另起一张列表工作表");
        Assert.AreEqual(existing[0], old.GetRow(0).GetCell(0).StringCellValue, "原有的列表应原样保留");
        Assert.AreEqual(existing[59], old.GetRow(59).GetCell(0).StringCellValue, "原有的列表应原样保留");
        Assert.AreEqual(added[0], fresh.GetRow(0).GetCell(0).StringCellValue);
        Assert.AreEqual(added[69], fresh.GetRow(69).GetCell(0).StringCellValue);
        Assert.AreEqual(SheetVisibility.Hidden, workbook.GetSheetVisibility(workbook.GetSheetIndex(fresh)));

        // 两张列表各有自己的定义名称，先前那个仍指向原处
        Assert.AreEqual("_excel2object_lists!$A$1:$A$60", workbook.GetName("_excel2object_list1").RefersToFormula);
        Assert.AreEqual("_excel2object_lists2!$A$1:$A$70", workbook.GetName("_excel2object_list2").RefersToFormula);
    }

    [TestMethod]
    public void ManyMergedRegionsAreWrittenOnAStreamedSheet()
    {
        // 分组表的合并区域数量与行数同阶：2000 行、每 5 行一组即 400 个区域。流式工作表同样绕开
        // NPOI 自带的两两比对，否则这一步会退化为平方级。
        var sheet = Streamed(Rows(2000), options => options.MergeRepeatedColumns.Add("城市"), 100);

        var regions = Regions(sheet);
        Assert.AreEqual(400, regions.Length);
        Assert.AreEqual("B2:B6", regions[0]);
        Assert.AreEqual("B1997:B2001", regions[399]);
    }

    [TestMethod]
    public void AutoColumnWidthMatchesTheInMemoryExport()
    {
        var rows = Rows(60);
        Action<ExcelExporterOptions> configure = options => options.AutoColumnWidth = true;

        var streamed = Streamed(rows, configure);
        var inMemory = InMemory(rows, configure);
        for (var i = 0; i < 5; i++)
            Assert.AreEqual(inMemory.GetColumnWidth(i), streamed.GetColumnWidth(i), $"第 {i} 列的宽度");
    }

    [TestMethod]
    public void TheDataIsEnumeratedOnlyOnce()
    {
        var source = new CountingSequence(Rows(40));
        using var output = new MemoryStream();
        new ExcelExporter().ObjectToExcelStream(source, output, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.StreamingRowWindow = Window;
            // 自动列宽此前要先把数据过一遍量宽度，如今与写入同一遍完成
            options.AutoColumnWidth = true;
        });

        Assert.AreEqual(1, source.Enumerations);
    }

    [TestMethod]
    public void TheOutputStreamIsLeftOpen()
    {
        using var output = new MemoryStream();
        new ExcelExporter().ObjectToExcelStream(Rows(20), output, options => options.StreamingRowWindow = Window);

        // 关闭了的 MemoryStream 连长度都读不到
        Assert.IsTrue(output.Length > 0);
        output.WriteByte(0);
    }

    [TestMethod]
    public void NothingIsWrittenWhenAMergeColumnIsUnknown()
    {
        using var output = new MemoryStream();
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            new ExcelExporter().ObjectToExcelStream(Rows(40), output, options =>
            {
                options.StreamingRowWindow = Window;
                options.MergeRepeatedColumns.Add("不存在的列");
            }));

        StringAssert.Contains(e.Message, "不存在的列");
        // 列名在写入开始前即解析，故调用方的流上不会留下半个文件
        Assert.AreEqual(0, output.Length);
    }

    [TestMethod]
    public void TheRowWindowMustBePositive()
    {
        // .xls 那条路上虽用不着这个值，取值不合法同样说明调用方想要的与得到的并不一致
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            using var output = new MemoryStream();
            var e = Assert.ThrowsException<Excel2ObjectException>(() =>
                new ExcelExporter().ObjectToExcelStream(Rows(10), output, options =>
                {
                    options.ExcelType = excelType;
                    options.StreamingRowWindow = 0;
                }), excelType.ToString());

            StringAssert.Contains(e.Message, "StreamingRowWindow");
            Assert.AreEqual(0, output.Length, excelType.ToString());
        }
    }

    [TestMethod]
    public void TheSourceIsReleasedWhenTheExportNeverStarts()
    {
        // 字典入口为取列名要先读一行，枚举器因而在写入开始之前即已打开。写入若根本没能开始，
        // 它也须被释放——数据库游标之类的东西正是这样悬着的。
        var badWindow = new TrackedDictionaries();
        Assert.ThrowsException<Excel2ObjectException>(() =>
            new ExcelExporter().ObjectToExcelStream(badWindow, new MemoryStream(),
                options => options.StreamingRowWindow = 0));
        Assert.IsTrue(badWindow.Disposed, "行窗口不合法时");

        var badSource = new TrackedDictionaries();
        Assert.ThrowsException<Excel2ObjectException>(() =>
            new ExcelExporter().ObjectToExcelStream(badSource, new MemoryStream(), options =>
            {
                options.StreamingRowWindow = Window;
                options.SourceExcelBytes = new byte[] {1, 2, 3};
            }));
        Assert.IsTrue(badSource.Disposed, "源工作簿读不出来时");
    }

    [TestMethod]
    public void XlsIsWrittenToTheStreamToo()
    {
        // .xls 无从流式写入，仍应写出一份可读的文件
        using var output = new MemoryStream();
        new ExcelExporter().ObjectToExcelStream(Rows(20), output, options => options.ExcelType = ExcelType.Xls);

        var sheet = WorkbookFactory.Create(new MemoryStream(output.ToArray())).GetSheetAt(0);
        Assert.AreEqual(20, sheet.LastRowNum);
        Assert.AreEqual("省0", sheet.GetRow(1).GetCell(0).StringCellValue);
    }

    [TestMethod]
    public void TheStreamedFileImportsBack()
    {
        // 流式写入的文本不进共享字符串表，而是直接写在单元格里（SXSSF 的做法，否则各不相同的字符串
        // 会全数留在内存中）。读的一侧不受影响，此处以导回对象确认。
        var rows = Rows(30);
        using var output = new MemoryStream();
        new ExcelExporter().ObjectToExcelStream(rows, output, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.StreamingRowWindow = Window;
        });

        var imported = new ExcelImporter().ExcelToObject<Model>(output.ToArray()).ToList();
        Assert.AreEqual(rows.Count, imported.Count);
        for (var i = 0; i < rows.Count; i++)
        {
            Assert.AreEqual(rows[i].Province, imported[i].Province, $"第 {i} 行的省份");
            Assert.AreEqual(rows[i].Amount, imported[i].Amount, $"第 {i} 行的金额");
            Assert.AreEqual(rows[i].PlacedAt, imported[i].PlacedAt, $"第 {i} 行的下单时间");
        }
    }

    [TestMethod]
    public void TempFilesAreWrittenWhileExportingAndRemovedAtTheEnd()
    {
        // 临时目录已由 TestTempDirectory 改指到本进程自己的目录，故这里数到的只会是自己写下的文件
        var duringExport = -1;
        var rows = Rows(100).Select((row, i) =>
        {
            if (i == 90) duringExport = PoiFiles();
            return row;
        });

        using var output = new MemoryStream();
        new ExcelExporter().ObjectToExcelStream(rows, output, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.StreamingRowWindow = Window;
        });

        Assert.AreEqual(1, duringExport, "写入过程中应有一个临时文件装着已刷出内存的行");
        Assert.AreEqual(0, PoiFiles(), "导出结束后临时文件应已删除");
    }

    /// <summary>NPOI 把刷出内存的行写在临时目录下的这个位置。</summary>
    private static int PoiFiles()
    {
        var dir = Path.Combine(Path.GetTempPath(), "poifiles");
        return Directory.Exists(dir) ? Directory.GetFiles(dir, "poi-sxssf-sheet*").Length : 0;
    }

    [TestMethod]
    public void FormulaColumnsAreWrittenPerRow()
    {
        var sheet = Streamed(Rows(30), options => options.FormulaColumns.Add(new FormulaColumn
        {
            Title = "两倍金额",
            Formula = c => c["金额"] * 2
        }));

        // 公式按各自所在的行写出，行号不因先前的行离开内存而错位
        Assert.AreEqual("C2*2", sheet.GetRow(1).GetCell(5).CellFormula);
        Assert.AreEqual("C21*2", sheet.GetRow(20).GetCell(5).CellFormula);
        Assert.AreEqual("C31*2", sheet.GetRow(30).GetCell(5).CellFormula);
    }

    [TestMethod]
    public void FormulasCanReferToASheetAlreadyInTheWorkbook()
    {
        // 先写出一张工作簿，再以流式导出向其追加一张表，其中的公式引用前一张表的列
        var first = new ExcelExporter().ObjectToExcelBytes(Rows(10), options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.SheetTitle = "上月";
        });

        using var output = new MemoryStream();
        new ExcelExporter().ObjectToExcelStream(Rows(30), output, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.StreamingRowWindow = Window;
            options.SheetTitle = "本月";
            options.SourceExcelBytes = first;
            options.FormulaColumns.Add(new FormulaColumn
            {
                Title = "环比",
                Formula = c => c["金额"] - c.Sheet("上月")["金额"]
            });
        });

        var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray()));
        var sheet = workbook.GetSheet("本月");
        // 表名含非 ASCII 字符，读回时 NPOI 会给它加上引号
        Assert.AreEqual("C2-'上月'!C2", sheet.GetRow(1).GetCell(5).CellFormula);
        Assert.AreEqual(10, workbook.GetSheet("上月").LastRowNum, "原有的表应原样保留");
    }

    /// <summary>记下自己有没有被释放：字典入口会预先取走一行，那个枚举器的归属正是要验的东西。</summary>
    private sealed class TrackedDictionaries : IEnumerable<Dictionary<string, object>>
    {
        public bool Disposed { get; private set; }

        public IEnumerator<Dictionary<string, object>> GetEnumerator()
        {
            return new Enumerator(this);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private sealed class Enumerator : IEnumerator<Dictionary<string, object>>
        {
            private readonly TrackedDictionaries _owner;
            private int _index;

            public Enumerator(TrackedDictionaries owner)
            {
                _owner = owner;
            }

            public Dictionary<string, object> Current { get; private set; } = new();

            object IEnumerator.Current => Current;

            public bool MoveNext()
            {
                if (_index >= 3) return false;

                Current = new Dictionary<string, object> {{"省份", "省" + _index}, {"金额", 100 + _index}};
                _index++;
                return true;
            }

            public void Reset()
            {
                _index = 0;
            }

            public void Dispose()
            {
                _owner.Disposed = true;
            }
        }
    }

    /// <summary>记下被遍历的次数，遍历第二遍即失败——流式导出只应过一遍数据。</summary>
    private sealed class CountingSequence : IEnumerable<Model>
    {
        private readonly List<Model> _rows;

        public CountingSequence(List<Model> rows)
        {
            _rows = rows;
        }

        public int Enumerations { get; private set; }

        public IEnumerator<Model> GetEnumerator()
        {
            Enumerations++;
            return _rows.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
