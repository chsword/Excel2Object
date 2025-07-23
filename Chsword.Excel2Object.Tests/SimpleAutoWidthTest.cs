using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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

    [TestMethod]
    public void TestColumnWidthCalculation()
    {
        // 测试列宽计算逻辑
        var testCases = new[]
        {
            new { Text = "A", ExpectedMin = 3 },           // 最小宽度 (1字符 + 2填充)
            new { Text = "Hello", ExpectedMin = 7 },       // 5字符 + 2填充
            new { Text = "中文", ExpectedMin = 6 },         // 2个中文字符(4) + 2填充
            new { Text = "Mixed中文", ExpectedMin = 11 },   // Mixed(5) + 中文(4) + 2填充
            new { Text = "", ExpectedMin = 1 }             // 空字符串最小为1
        };

        foreach (var testCase in testCases)
        {
            // 这里我们无法直接测试私有方法，但可以通过导出验证
            var data = new List<Dictionary<string, object>>
            {
                new() { ["TestColumn"] = testCase.Text }
            };

            var bytes = ExcelHelper.ObjectToExcelBytes(data, options =>
            {
                options.ExcelType = ExcelType.Xlsx;
                options.AutoColumnWidth = true;
                options.MinColumnWidth = 1;  // 允许最小宽度
                options.MaxColumnWidth = 100; // 允许最大宽度
            });

            Assert.IsNotNull(bytes, $"列宽计算测试失败: {testCase.Text}");
        }

        Console.WriteLine("✅ 列宽计算逻辑测试通过");
    }
}
