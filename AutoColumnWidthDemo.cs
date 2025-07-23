using System;
using System.Collections.Generic;
using System.IO;
using Chsword.Excel2Object;
using Chsword.Excel2Object.Tests.Models;

namespace Chsword.Excel2Object.Demo;

/// <summary>
/// Demo program to showcase the new auto column width feature
/// </summary>
public class AutoColumnWidthDemo
{
    public static void RunDemo()
    {
        Console.WriteLine("=== Excel2Object 自动列宽功能演示 ===\n");
        
        // 准备测试数据
        var testData = new List<TestModelPerson>
        {
            new()
            {
                Name = "张三",
                Age = 25,
                CreateTime = DateTime.Now
            },
            new()
            {
                Name = "李四有一个很长的名字用来测试自动列宽功能",
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
        
        // 测试1：启用自动列宽
        Console.WriteLine("1. 测试启用自动列宽功能:");
        var bytesWithAutoWidth = ExcelHelper.ObjectToExcelBytes(testData, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.AutoColumnWidth = true;
            options.MinColumnWidth = 8;   // 最小宽度
            options.MaxColumnWidth = 40;  // 最大宽度
        });
        
        if (bytesWithAutoWidth != null)
        {
            File.WriteAllBytes("demo_auto_width.xlsx", bytesWithAutoWidth);
            Console.WriteLine("✅ 自动列宽Excel文件已生成: demo_auto_width.xlsx");
            Console.WriteLine($"   文件大小: {bytesWithAutoWidth.Length} bytes");
        }
        
        // 测试2：禁用自动列宽（使用固定宽度）
        Console.WriteLine("\n2. 测试固定列宽功能:");
        var bytesWithFixedWidth = ExcelHelper.ObjectToExcelBytes(testData, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.AutoColumnWidth = false;
            options.DefaultColumnWidth = 15; // 固定宽度
        });
        
        if (bytesWithFixedWidth != null)
        {
            File.WriteAllBytes("demo_fixed_width.xlsx", bytesWithFixedWidth);
            Console.WriteLine("✅ 固定列宽Excel文件已生成: demo_fixed_width.xlsx");
            Console.WriteLine($"   文件大小: {bytesWithFixedWidth.Length} bytes");
        }
        
        // 测试3：测试中英文混合内容
        Console.WriteLine("\n3. 测试中英文混合内容的自动列宽:");
        var mixedData = new List<Dictionary<string, object>>
        {
            new()
            {
                ["Short"] = "A",
                ["Medium Length Text"] = "This is medium length content",
                ["很长的中文列名测试"] = "中文内容测试，包含宽字符"
            },
            new()
            {
                ["Short"] = "B",
                ["Medium Length Text"] = "Short",
                ["很长的中文列名测试"] = "Mixed 中英文 content test"
            }
        };
        
        var bytesWithMixed = ExcelHelper.ObjectToExcelBytes(mixedData, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.AutoColumnWidth = true;
            options.MinColumnWidth = 5;
            options.MaxColumnWidth = 35;
        });
        
        if (bytesWithMixed != null)
        {
            File.WriteAllBytes("demo_mixed_content.xlsx", bytesWithMixed);
            Console.WriteLine("✅ 中英文混合内容Excel文件已生成: demo_mixed_content.xlsx");
            Console.WriteLine($"   文件大小: {bytesWithMixed.Length} bytes");
        }
        
        // 验证导入功能
        Console.WriteLine("\n4. 验证导入功能:");
        try
        {
            var importer = new ExcelImporter();
            var importedData = importer.ExcelToObject<TestModelPerson>(bytesWithAutoWidth!).ToList();
            Console.WriteLine($"✅ 成功导入 {importedData.Count} 条记录");
            
            foreach (var item in importedData)
            {
                Console.WriteLine($"   - {item.Name}, 年龄: {item.Age}, 创建时间: {item.CreateTime:yyyy-MM-dd}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ 导入失败: {ex.Message}");
        }
        
        Console.WriteLine("\n=== 演示完成 ===");
        Console.WriteLine("请检查生成的Excel文件，对比自动列宽和固定列宽的效果差异。");
    }
}

// 程序入口（如果需要独立运行）
public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            AutoColumnWidthDemo.RunDemo();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"程序执行出错: {ex.Message}");
            Console.WriteLine($"详细信息: {ex}");
        }
        
        Console.WriteLine("\n按任意键退出...");
        Console.ReadKey();
    }
}
