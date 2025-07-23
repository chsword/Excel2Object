using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Chsword.Excel2Object;
using Chsword.Excel2Object.Tests.Models;

namespace Chsword.Excel2Object.Demo;

/// <summary>
/// 演示 P2.1 流式处理大文件和 P2.3 异步流支持的功能
/// </summary>
public class StreamingDemo
{
    /// <summary>
    /// 运行完整的流式处理演示
    /// </summary>
    public static async Task RunDemo()
    {
        Console.WriteLine("=== Excel2Object 流式处理功能演示 ===\n");

        // 创建大量测试数据
        var largeTestData = CreateLargeTestData(5000); // 5000 条记录
        
        // 生成Excel文件
        Console.WriteLine("1. 生成大型Excel文件...");
        var bytes = ExcelHelper.ObjectToExcelBytes(largeTestData, ExcelType.Xlsx);
        if (bytes == null)
        {
            Console.WriteLine("❌ 生成Excel文件失败");
            return;
        }
        
        Console.WriteLine($"✅ 生成了包含 {largeTestData.Count} 条记录的Excel文件，大小: {bytes.Length / 1024.0:F2} KB\n");

        // 测试1：传统同步处理
        await TestTraditionalProcessing(bytes);
        
        // 测试2：异步批量处理
        await TestAsyncBatchProcessing(bytes);
        
#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
        // 测试3：流式异步处理
        await TestStreamProcessing(bytes);
        
        // 测试4：可取消的流式处理
        await TestCancellableStreamProcessing(bytes);
#endif
        
        Console.WriteLine("=== 演示完成 ===");
    }

    /// <summary>
    /// 测试传统同步处理方式
    /// </summary>
    private static async Task TestTraditionalProcessing(byte[] bytes)
    {
        Console.WriteLine("2. 传统同步处理测试:");
        var stopwatch = Stopwatch.StartNew();
        
        try
        {
            var results = ExcelHelper.ExcelToObject<TestModelPerson>(bytes);
            var count = results.Count();
            
            stopwatch.Stop();
            Console.WriteLine($"✅ 同步处理完成，处理了 {count} 条记录");
            Console.WriteLine($"   耗时: {stopwatch.ElapsedMilliseconds} ms");
            Console.WriteLine($"   内存使用: 一次性加载所有数据到内存\n");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ 同步处理失败: {ex.Message}\n");
        }
    }

    /// <summary>
    /// 测试异步批量处理
    /// </summary>
    private static async Task TestAsyncBatchProcessing(byte[] bytes)
    {
        Console.WriteLine("3. 异步批量处理测试:");
        var stopwatch = Stopwatch.StartNew();
        
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(1));
            var results = await ExcelHelper.ExcelToObjectAsync<TestModelPerson>(bytes, cancellationToken: cts.Token);
            var count = results.Count();
            
            stopwatch.Stop();
            Console.WriteLine($"✅ 异步批量处理完成，处理了 {count} 条记录");
            Console.WriteLine($"   耗时: {stopwatch.ElapsedMilliseconds} ms");
            Console.WriteLine($"   特点: 异步操作，但仍然一次性加载所有数据\n");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ 异步批量处理失败: {ex.Message}\n");
        }
    }

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    /// <summary>
    /// 测试流式异步处理
    /// </summary>
    private static async Task TestStreamProcessing(byte[] bytes)
    {
        Console.WriteLine("4. 流式异步处理测试:");
        var stopwatch = Stopwatch.StartNew();
        
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(1));
            var count = 0;
            var processedCount = 0;
            
            await foreach (var person in ExcelHelper.ExcelToObjectStreamAsync<TestModelPerson>(bytes, cancellationToken: cts.Token))
            {
                count++;
                
                // 模拟处理每条记录
                if (person.Age > 0) // 简单的数据验证
                {
                    processedCount++;
                }
                
                // 每处理1000条记录显示进度
                if (count % 1000 == 0)
                {
                    Console.WriteLine($"   已处理 {count} 条记录...");
                }
            }
            
            stopwatch.Stop();
            Console.WriteLine($"✅ 流式处理完成，处理了 {count} 条记录，有效记录 {processedCount} 条");
            Console.WriteLine($"   耗时: {stopwatch.ElapsedMilliseconds} ms");
            Console.WriteLine($"   特点: 逐行流式处理，内存占用低，支持大文件\n");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ 流式处理失败: {ex.Message}\n");
        }
    }

    /// <summary>
    /// 测试可取消的流式处理
    /// </summary>
    private static async Task TestCancellableStreamProcessing(byte[] bytes)
    {
        Console.WriteLine("5. 可取消流式处理测试:");
        var stopwatch = Stopwatch.StartNew();
        
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(2)); // 2秒后取消
            var count = 0;
            
            await foreach (var person in ExcelHelper.ExcelToObjectStreamAsync<TestModelPerson>(bytes, cancellationToken: cts.Token))
            {
                count++;
                
                // 每处理500条记录显示进度
                if (count % 500 == 0)
                {
                    Console.WriteLine($"   已处理 {count} 条记录...");
                }
                
                // 模拟处理延迟
                await Task.Delay(1, cts.Token);
            }
            
            stopwatch.Stop();
            Console.WriteLine($"✅ 处理完成，总共处理了 {count} 条记录");
        }
        catch (OperationCanceledException)
        {
            stopwatch.Stop();
            Console.WriteLine($"⚠️  流式处理被取消");
            Console.WriteLine($"   已处理时间: {stopwatch.ElapsedMilliseconds} ms");
            Console.WriteLine($"   特点: 支持优雅取消，可以中途停止大文件处理\n");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ 可取消流式处理失败: {ex.Message}\n");
        }
    }
#endif

    /// <summary>
    /// 创建大量测试数据
    /// </summary>
    private static List<TestModelPerson> CreateLargeTestData(int count)
    {
        var random = new Random();
        var firstNames = new[] { "张", "李", "王", "刘", "陈", "杨", "赵", "黄", "周", "吴" };
        var lastNames = new[] { "伟", "芳", "娜", "秀英", "敏", "静", "丽", "强", "磊", "军" };
        
        var data = new List<TestModelPerson>();
        
        for (int i = 0; i < count; i++)
        {
            data.Add(new TestModelPerson
            {
                Name = $"{firstNames[random.Next(firstNames.Length)]}{lastNames[random.Next(lastNames.Length)]}{i:D4}",
                Age = random.Next(18, 80),
                CreateTime = DateTime.Now.AddDays(-random.Next(1000))
            });
        }
        
        return data;
    }

    /// <summary>
    /// 演示传统方法 vs 流式方法的性能对比
    /// </summary>
    public static async Task RunPerformanceComparison()
    {
        Console.WriteLine("=== Excel2Object 流式处理性能演示 ===\n");

        // 生成测试数据
        var testData = GenerateTestData(10000); // 生成10,000条记录
        
        // 导出到Excel文件
        Console.WriteLine("1. 生成测试Excel文件...");
        var excelBytes = ExcelHelper.ObjectToExcelBytes(testData, ExcelType.Xlsx);
        if (excelBytes == null)
        {
            Console.WriteLine("❌ 生成Excel文件失败");
            return;
        }
        
        var filePath = "large_test_file.xlsx";
        File.WriteAllBytes(filePath, excelBytes);
        Console.WriteLine($"✅ 生成测试文件: {filePath} ({excelBytes.Length / 1024.0 / 1024.0:F2} MB)");

        await CompareReadingMethods(filePath);
        
        // 清理
        if (File.Exists(filePath))
            File.Delete(filePath);
    }

    /// <summary>
    /// 对比不同读取方法的性能
    /// </summary>
    private static async Task CompareReadingMethods(string filePath)
    {
        using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(5));
        var cancellationToken = cts.Token;

        Console.WriteLine("\n2. 性能对比测试:");
        
        // 方法1: 传统同步方法
        Console.WriteLine("\n--- 方法1: 传统同步读取 ---");
        await TestTraditionalMethod(filePath);

        // 方法2: 异步方法
        Console.WriteLine("\n--- 方法2: 异步读取 ---");
        await TestAsyncMethod(filePath, cancellationToken);

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
        // 方法3: 异步流方法
        Console.WriteLine("\n--- 方法3: 异步流式读取 ---");
        await TestStreamingMethod(filePath, cancellationToken);
#else
        Console.WriteLine("\n--- 方法3: 异步流式读取 (当前框架不支持) ---");
        Console.WriteLine("需要 .NET Standard 2.1 或更高版本才能使用 IAsyncEnumerable");
#endif
    }

    /// <summary>
    /// 测试传统同步方法
    /// </summary>
    private static async Task TestTraditionalMethod(string filePath)
    {
        var sw = Stopwatch.StartNew();
        var initialMemory = GC.GetTotalMemory(false);
        
        try
        {
            var importer = new ExcelImporter();
            var result = importer.ExcelToObject<TestModelPerson>(filePath)?.ToList();
            
            sw.Stop();
            var finalMemory = GC.GetTotalMemory(false);
            var memoryUsed = (finalMemory - initialMemory) / 1024.0 / 1024.0;
            
            Console.WriteLine($"✅ 处理完成: {result?.Count ?? 0} 条记录");
            Console.WriteLine($"⏱️ 耗时: {sw.ElapsedMilliseconds} ms");
            Console.WriteLine($"🧠 内存使用: {memoryUsed:F2} MB");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ 错误: {ex.Message}");
        }
        
        await Task.Delay(1000); // 等待GC
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    /// <summary>
    /// 测试异步方法
    /// </summary>
    private static async Task TestAsyncMethod(string filePath, CancellationToken cancellationToken)
    {
        var sw = Stopwatch.StartNew();
        var initialMemory = GC.GetTotalMemory(false);
        
        try
        {
            var result = await ExcelHelper.ExcelToObjectAsync<TestModelPerson>(filePath, cancellationToken: cancellationToken);
            var list = result?.ToList();
            
            sw.Stop();
            var finalMemory = GC.GetTotalMemory(false);
            var memoryUsed = (finalMemory - initialMemory) / 1024.0 / 1024.0;
            
            Console.WriteLine($"✅ 处理完成: {list?.Count ?? 0} 条记录");
            Console.WriteLine($"⏱️ 耗时: {sw.ElapsedMilliseconds} ms");
            Console.WriteLine($"🧠 内存使用: {memoryUsed:F2} MB");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ 错误: {ex.Message}");
        }
        
        await Task.Delay(1000); // 等待GC
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    /// <summary>
    /// 测试异步流方法
    /// </summary>
    private static async Task TestStreamingMethod(string filePath, CancellationToken cancellationToken)
    {
        var sw = Stopwatch.StartNew();
        var initialMemory = GC.GetTotalMemory(false);
        var count = 0;
        var peakMemory = initialMemory;
        
        try
        {
            await foreach (var person in ExcelHelper.ExcelToObjectStreamAsync<TestModelPerson>(filePath, cancellationToken: cancellationToken))
            {
                count++;
                
                // 监控内存使用
                if (count % 1000 == 0)
                {
                    var currentMemory = GC.GetTotalMemory(false);
                    peakMemory = Math.Max(peakMemory, currentMemory);
                    Console.WriteLine($"📊 已处理 {count} 条记录，当前内存: {(currentMemory - initialMemory) / 1024.0 / 1024.0:F2} MB");
                }
                
                // 模拟处理
                if (!string.IsNullOrEmpty(person.Name))
                {
                    // 简单的数据验证
                    _ = person.Name.Length;
                }
            }
            
            sw.Stop();
            var memoryUsed = (peakMemory - initialMemory) / 1024.0 / 1024.0;
            
            Console.WriteLine($"✅ 流式处理完成: {count} 条记录");
            Console.WriteLine($"⏱️ 耗时: {sw.ElapsedMilliseconds} ms");
            Console.WriteLine($"🧠 峰值内存使用: {memoryUsed:F2} MB");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ 错误: {ex.Message}");
        }
        
        await Task.Delay(1000); // 等待GC
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    /// <summary>
    /// 演示取消操作的功能
    /// </summary>
    public static async Task DemonstrateCancellation()
    {
        Console.WriteLine("\n=== 演示取消操作功能 ===");
        
        var testData = GenerateTestData(50000); // 生成更多数据
        var excelBytes = ExcelHelper.ObjectToExcelBytes(testData, ExcelType.Xlsx);
        if (excelBytes == null) return;
        
        var filePath = "cancellation_test.xlsx";
        File.WriteAllBytes(filePath, excelBytes);
        
        using var cts = new CancellationTokenSource();
        
        // 5秒后取消
        cts.CancelAfter(TimeSpan.FromSeconds(5));
        
        try
        {
            var count = 0;
            Console.WriteLine("开始处理，5秒后将自动取消...");
            
            await foreach (var person in ExcelHelper.ExcelToObjectStreamAsync<TestModelPerson>(filePath, cancellationToken: cts.Token))
            {
                count++;
                if (count % 1000 == 0)
                {
                    Console.WriteLine($"已处理 {count} 条记录");
                }
            }
            
            Console.WriteLine($"✅ 处理完成: {count} 条记录");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("⚠️ 操作被取消");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ 错误: {ex.Message}");
        }
        
        // 清理
        if (File.Exists(filePath))
            File.Delete(filePath);
    }
#endif

    /// <summary>
    /// 生成测试数据
    /// </summary>
    private static List<TestModelPerson> GenerateTestData(int count)
    {
        var random = new Random();
        var names = new[] { "张三", "李四", "王五", "赵六", "钱七", "孙八", "周九", "吴十" };
        
        var result = new List<TestModelPerson>();
        for (int i = 0; i < count; i++)
        {
            result.Add(new TestModelPerson
            {
                Name = $"{names[random.Next(names.Length)]}{i}",
                Age = random.Next(18, 65),
                CreateTime = DateTime.Now.AddDays(-random.Next(365 * 5))
            });
        }
        
        return result;
    }
}

// 程序入口（如果需要独立运行）
public class StreamingProgram
{
    public static async Task Main(string[] args)
    {
        try
        {
            Console.WriteLine("选择测试模式:");
            Console.WriteLine("1. 流式处理功能演示");
            Console.WriteLine("2. 性能对比测试");
            Console.WriteLine("请输入选择 (1-2):");
            
            var choice = Console.ReadLine();
            
            switch (choice)
            {
                case "1":
                    await StreamingDemo.RunDemo();
                    break;
                case "2":
                    await StreamingDemo.RunPerformanceComparison();
#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
                    await StreamingDemo.DemonstrateCancellation();
#endif
                    break;
                default:
                    Console.WriteLine("默认运行流式处理功能演示");
                    await StreamingDemo.RunDemo();
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"程序错误: {ex.Message}");
            Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
        }
        
        Console.WriteLine("\n按任意键退出...");
        Console.ReadKey();
    }
}
