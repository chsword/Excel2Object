using System.Diagnostics;
using Chsword.Excel2Object;
using Chsword.Excel2Object.Tests.Models;

Console.WriteLine("=== Excel2Object 流式处理功能测试 ===\n");

try
{
    // 创建测试数据
    var testData = CreateTestData(5000);
    Console.WriteLine($"✅ 创建了 {testData.Count} 条测试数据");

    // 生成Excel文件
    var excelBytes = ExcelHelper.ObjectToExcelBytes(testData, ExcelType.Xlsx);
    if (excelBytes == null)
    {
        Console.WriteLine("❌ 生成Excel文件失败");
        return;
    }
    
    Console.WriteLine($"✅ 生成Excel文件成功，大小: {excelBytes.Length / 1024.0:F2} KB");

    // 测试1：传统同步处理
    Console.WriteLine("\n--- 测试1: 传统同步处理 ---");
    var sw1 = Stopwatch.StartNew();
    var results1 = ExcelHelper.ExcelToObject<TestModelPerson>(excelBytes).ToList();
    sw1.Stop();
    Console.WriteLine($"✅ 同步处理完成: {results1.Count} 条记录，耗时: {sw1.ElapsedMilliseconds} ms");

    // 测试2：异步批量处理
    Console.WriteLine("\n--- 测试2: 异步批量处理 ---");
    var sw2 = Stopwatch.StartNew();
    var results2 = (await ExcelHelper.ExcelToObjectAsync<TestModelPerson>(excelBytes)).ToList();
    sw2.Stop();
    Console.WriteLine($"✅ 异步处理完成: {results2.Count} 条记录，耗时: {sw2.ElapsedMilliseconds} ms");

    // 测试3：流式异步处理 (仅支持.NET Standard 2.1+)
    Console.WriteLine("\n--- 测试3: 流式异步处理 ---");
    var sw3 = Stopwatch.StartNew();
    var count = 0;
    
    await foreach (var person in ExcelHelper.ExcelToObjectStreamAsync<TestModelPerson>(excelBytes))
    {
        count++;
        if (count % 1000 == 0)
        {
            Console.WriteLine($"   已处理 {count} 条记录...");
        }
    }
    
    sw3.Stop();
    Console.WriteLine($"✅ 流式处理完成: {count} 条记录，耗时: {sw3.ElapsedMilliseconds} ms");

    // 测试4：可取消的流式处理
    Console.WriteLine("\n--- 测试4: 可取消流式处理 (2秒后自动取消) ---");
    using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(2));
    var cancelCount = 0;
    
    try
    {
        await foreach (var person in ExcelHelper.ExcelToObjectStreamAsync<TestModelPerson>(excelBytes, cancellationToken: cts.Token))
        {
            cancelCount++;
            if (cancelCount % 500 == 0)
            {
                Console.WriteLine($"   已处理 {cancelCount} 条记录...");
            }
            await Task.Delay(1, cts.Token); // 模拟处理延迟
        }
        Console.WriteLine($"✅ 处理完成: {cancelCount} 条记录");
    }
    catch (OperationCanceledException)
    {
        Console.WriteLine($"⚠️ 流式处理被取消，已处理 {cancelCount} 条记录");
    }

    Console.WriteLine("\n=== 所有测试完成 ===");
}
catch (Exception ex)
{
    Console.WriteLine($"❌ 测试失败: {ex.Message}");
    Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
}

Console.WriteLine("\n按任意键退出...");
Console.ReadKey();

// 创建测试数据的辅助方法
List<TestModelPerson> CreateTestData(int count)
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
