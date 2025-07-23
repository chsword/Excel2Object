using System;
using System.Collections.Generic;
using System.Linq;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class AutoColumnWidthTest : BaseExcelTest
{
    [TestMethod]
    public void TestAutoColumnWidthEnabled()
    {
        var models = GetTestModels();
        
        // Test with auto column width enabled
        var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.AutoColumnWidth = true;
            options.MinColumnWidth = 10;
            options.MaxColumnWidth = 30;
        });
        
        Assert.IsNotNull(bytes);
        Assert.IsTrue(bytes.Length > 0);
        
        // Verify the export can be imported back
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();
        Assert.AreEqual(3, result.Count);
        Assert.AreEqual("张三", result[0].Name);
    }

    [TestMethod]
    public void TestAutoColumnWidthDisabled()
    {
        var models = GetTestModels();
        
        // Test with auto column width disabled (default behavior)
        var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.AutoColumnWidth = false;
            options.DefaultColumnWidth = 20;
        });
        
        Assert.IsNotNull(bytes);
        Assert.IsTrue(bytes.Length > 0);
        
        // Verify the export can be imported back
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();
        Assert.AreEqual(3, result.Count);
    }

    [TestMethod]
    public void TestAutoColumnWidthWithLongContent()
    {
        var models = GetLongContentModels();
        
        // Test with very long content
        var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.AutoColumnWidth = true;
            options.MinColumnWidth = 5;
            options.MaxColumnWidth = 50; // Should be capped at this value
        });
        
        Assert.IsNotNull(bytes);
        Assert.IsTrue(bytes.Length > 0);
        
        // Verify the export works
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();
        Assert.AreEqual(2, result.Count);
    }

    [TestMethod]
    public void TestAutoColumnWidthWithMixedContent()
    {
        var models = new List<Dictionary<string, object>>
        {
            new()
            {
                ["短名"] = "A",
                ["Medium Length Name"] = "This is a medium length content",
                ["很长的中文列名用于测试宽字符"] = "中文内容测试，包含宽字符"
            },
            new()
            {
                ["短名"] = "B",
                ["Medium Length Name"] = "Short",
                ["很长的中文列名用于测试宽字符"] = "English mixed 中文 content"
            }
        };
        
        var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.AutoColumnWidth = true;
            options.MinColumnWidth = 8;
            options.MaxColumnWidth = 40;
        });
        
        Assert.IsNotNull(bytes);
        Assert.IsTrue(bytes.Length > 0);
        
        // Verify the export can be imported back as dictionary
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<Dictionary<string, object>>(bytes).ToList();
        Assert.AreEqual(2, result.Count);
        Console.WriteLine(JsonConvert.SerializeObject(result, Formatting.Indented));
    }

    private IEnumerable<TestModelPerson> GetTestModels()
    {
        return new List<TestModelPerson>
        {
            new()
            {
                Name = "张三",
                Age = 25,
                CreateTime = DateTime.Now
            },
            new()
            {
                Name = "李四",
                Age = 30,
                CreateTime = DateTime.Now.AddDays(-1)
            },
            new()
            {
                Name = "王五",
                Age = 35,
                CreateTime = DateTime.Now.AddDays(-2)
            }
        };
    }

    private IEnumerable<TestModelPerson> GetLongContentModels()
    {
        return new List<TestModelPerson>
        {
            new()
            {
                Name = "This is a very long name that should test the maximum column width constraint functionality",
                Age = 25,
                CreateTime = DateTime.Now
            },
            new()
            {
                Name = "Another extremely long name with lots of characters to verify that the auto width calculation works properly with maximum limits",
                Age = 30,
                CreateTime = DateTime.Now.AddDays(-1)
            }
        };
    }
}
