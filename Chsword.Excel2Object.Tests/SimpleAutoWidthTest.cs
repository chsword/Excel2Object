using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>
/// 简单的功能验证程序，用于测试自动列宽功能
/// </summary>
[TestClass]
public class SimpleAutoWidthTest
{
    [TestMethod]
    public void BasicAutoWidthTest()
    {
        // 准备测试数据
        var testData = new List<Dictionary<string, object>>
        {
            new() { ["Name"] = "张三", ["Age"] = 25, ["Description"] = "Short" },
            new() { ["Name"] = "李四有一个很长的名字", ["Age"] = 30, ["Description"] = "This is a much longer description text" },
            new() { ["Name"] = "王五", ["Age"] = 35, ["Description"] = "Medium length text" }
        };

        // 测试自动列宽
        var bytesAuto = ExcelHelper.ObjectToExcelBytes(testData, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.AutoColumnWidth = true;
            options.MinColumnWidth = 8;
            options.MaxColumnWidth = 30;
        });

        // 测试固定列宽
        var bytesFixed = ExcelHelper.ObjectToExcelBytes(testData, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.AutoColumnWidth = false;
            options.DefaultColumnWidth = 15;
        });

        // 验证结果
        Assert.IsNotNull(bytesAuto, "自动列宽导出失败");
        Assert.IsNotNull(bytesFixed, "固定列宽导出失败");
        Assert.IsTrue(bytesAuto.Length > 0, "自动列宽文件为空");
        Assert.IsTrue(bytesFixed.Length > 0, "固定列宽文件为空");

        // 验证导入功能
        var importer = new ExcelImporter();
        var resultAuto = importer.ExcelToObject<Dictionary<string, object>>(bytesAuto).ToList();
        var resultFixed = importer.ExcelToObject<Dictionary<string, object>>(bytesFixed).ToList();

        Assert.AreEqual(3, resultAuto.Count, "自动列宽导入数据数量不正确");
        Assert.AreEqual(3, resultFixed.Count, "固定列宽导入数据数量不正确");

        // 验证数据内容
        Assert.AreEqual("张三", resultAuto[0]["Name"].ToString(), "自动列宽数据内容不正确");
        Assert.AreEqual("张三", resultFixed[0]["Name"].ToString(), "固定列宽数据内容不正确");

        Console.WriteLine("✅ 自动列宽功能测试通过");
        Console.WriteLine($"自动列宽文件大小: {bytesAuto.Length} bytes");
        Console.WriteLine($"固定列宽文件大小: {bytesFixed.Length} bytes");
    }

    /// <summary>
    ///     Auto width is the header width or the widest cell, clamped to Min/MaxColumnWidth. A character
    ///     counts as 1, an upper case or wide letter as 1.2 and a full width character as 2, plus 2 for
    ///     padding; an empty string counts as 1 with no padding. NPOI stores the width in 1/256 of a
    ///     character. The header here is a single narrow character so the content decides the width.
    /// </summary>
    [TestMethod]
    public void TestColumnWidthCalculation()
    {
        var cases = new[]
        {
            new {Text = "A", Expected = 4}, // A counts 1.2, rounded up to 2, plus 2 padding
            new {Text = "Hello", Expected = 8}, // H(1.2) + ello(4) = 5.2 -> 6, plus 2
            new {Text = "中文", Expected = 6}, // two full width characters = 4, plus 2
            new {Text = "Mixed中文", Expected = 12}, // Mixed(5.2) + 中文(4) = 9.2 -> 10, plus 2
            new {Text = "", Expected = 3} // empty counts 1, so the header "x" (1 + 2) wins
        };

        foreach (var testCase in cases)
            Assert.AreEqual(testCase.Expected * 256d, WidthOf(testCase.Text, 1, 100),
                $"width of [{testCase.Text}]");

        // and the calculated width is clamped on both ends
        Assert.AreEqual(20 * 256d, WidthOf("Hello", 20, 100), "MinColumnWidth");
        Assert.AreEqual(5 * 256d, WidthOf("Mixed中文", 1, 5), "MaxColumnWidth");
    }

    private static double WidthOf(string text, int min, int max)
    {
        var bytes = ExcelHelper.ObjectToExcelBytes(
            new List<Dictionary<string, object>> {new() {["x"] = text}},
            options =>
            {
                options.ExcelType = ExcelType.Xlsx;
                options.AutoColumnWidth = true;
                options.MinColumnWidth = min;
                options.MaxColumnWidth = max;
            });
        Assert.IsNotNull(bytes);
        using var stream = new MemoryStream(bytes);
        return WorkbookFactory.Create(stream).GetSheetAt(0).GetColumnWidth(0);
    }
}
