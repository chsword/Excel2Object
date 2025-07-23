using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class StreamingTests : BaseExcelTest
{
    [TestMethod]
    public async Task ExcelToObjectAsync_ShouldWork()
    {
        // Arrange
        var testData = GetTestData();
        var bytes = ExcelHelper.ObjectToExcelBytes(testData, ExcelType.Xlsx);
        Assert.IsNotNull(bytes);

        // Act
        var result = await ExcelHelper.ExcelToObjectAsync<TestModelPerson>(bytes);
        var list = result?.ToList();

        // Assert
        Assert.IsNotNull(list);
        Assert.AreEqual(testData.Count, list.Count);
        
        for (int i = 0; i < testData.Count; i++)
        {
            Assert.AreEqual(testData[i].Name, list[i].Name);
            Assert.AreEqual(testData[i].Age, list[i].Age);
        }
    }

    [TestMethod]
    public async Task ExcelToObjectAsync_WithCancellation_ShouldRespectCancellation()
    {
        // Arrange
        var testData = GetTestData();
        var bytes = ExcelHelper.ObjectToExcelBytes(testData, ExcelType.Xlsx);
        Assert.IsNotNull(bytes);

        using var cts = new CancellationTokenSource();
        cts.Cancel(); // Cancel immediately

        // Act & Assert
        await Assert.ThrowsExceptionAsync<TaskCanceledException>(async () =>
        {
            await ExcelHelper.ExcelToObjectAsync<TestModelPerson>(bytes, cancellationToken: cts.Token);
        });
    }

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    [TestMethod]
    public async Task ExcelToObjectStreamAsync_ShouldWork()
    {
        // Arrange
        var testData = GetTestData();
        var bytes = ExcelHelper.ObjectToExcelBytes(testData, ExcelType.Xlsx);
        Assert.IsNotNull(bytes);

        // Act
        var result = new List<TestModelPerson>();
        await foreach (var item in ExcelHelper.ExcelToObjectStreamAsync<TestModelPerson>(bytes))
        {
            result.Add(item);
        }

        // Assert
        Assert.AreEqual(testData.Count, result.Count);
        
        for (int i = 0; i < testData.Count; i++)
        {
            Assert.AreEqual(testData[i].Name, result[i].Name);
            Assert.AreEqual(testData[i].Age, result[i].Age);
        }
    }

    [TestMethod]
    public async Task ExcelToObjectStreamAsync_WithCancellation_ShouldRespectCancellation()
    {
        // Arrange
        var testData = GetLargeTestData(1000); // Generate more data for testing
        var bytes = ExcelHelper.ObjectToExcelBytes(testData, ExcelType.Xlsx);
        Assert.IsNotNull(bytes);

        using var cts = new CancellationTokenSource();
        cts.CancelAfter(TimeSpan.FromMilliseconds(100)); // Cancel after 100ms

        // Act
        var processedCount = 0;
        var cancellationReceived = false;
        
        try
        {
            await foreach (var item in ExcelHelper.ExcelToObjectStreamAsync<TestModelPerson>(bytes, cancellationToken: cts.Token))
            {
                processedCount++;
                // Add some processing time to ensure cancellation can occur
                await Task.Delay(5, cts.Token);
            }
        }
        catch (OperationCanceledException)
        {
            cancellationReceived = true;
        }

        // Assert
        Assert.IsTrue(cancellationReceived, "Cancellation should have been triggered");
        Assert.IsTrue(processedCount < testData.Count, "Should not have processed all items due to cancellation");
    }

    [TestMethod]
    public async Task ExcelToObjectStreamAsync_MemoryEfficiency_Test()
    {
        // Arrange
        var testData = GetLargeTestData(5000); // Large dataset
        var bytes = ExcelHelper.ObjectToExcelBytes(testData, ExcelType.Xlsx);
        Assert.IsNotNull(bytes);

        // Act
        var initialMemory = GC.GetTotalMemory(false);
        var processedCount = 0;
        var maxMemoryIncrease = 0L;

        await foreach (var item in ExcelHelper.ExcelToObjectStreamAsync<TestModelPerson>(bytes))
        {
            processedCount++;
            
            if (processedCount % 500 == 0)
            {
                var currentMemory = GC.GetTotalMemory(false);
                var memoryIncrease = currentMemory - initialMemory;
                maxMemoryIncrease = Math.Max(maxMemoryIncrease, memoryIncrease);
            }
        }

        // Assert
        Assert.AreEqual(testData.Count, processedCount);
        
        // Memory should not increase dramatically (streaming should be memory efficient)
        var maxAllowedMemoryIncrease = 50 * 1024 * 1024; // 50 MB
        Assert.IsTrue(maxMemoryIncrease < maxAllowedMemoryIncrease, 
            $"Memory usage increased by {maxMemoryIncrease / 1024.0 / 1024.0:F2} MB, should be less than {maxAllowedMemoryIncrease / 1024.0 / 1024.0:F2} MB");
    }

    [TestMethod]
    public async Task ExcelToObjectStreamAsync_WithDictionary_ShouldWork()
    {
        // Arrange
        var testData = new List<Dictionary<string, object>>
        {
            new() { ["Name"] = "张三", ["Age"] = 25, ["City"] = "北京" },
            new() { ["Name"] = "李四", ["Age"] = 30, ["City"] = "上海" },
            new() { ["Name"] = "王五", ["Age"] = 35, ["City"] = "广州" }
        };
        
        var bytes = ExcelHelper.ObjectToExcelBytes(testData, ExcelType.Xlsx);
        Assert.IsNotNull(bytes);

        // Act
        var result = new List<Dictionary<string, object>>();
        await foreach (var item in ExcelHelper.ExcelToObjectStreamAsync<Dictionary<string, object>>(bytes))
        {
            result.Add(item);
        }

        // Assert
        Assert.AreEqual(testData.Count, result.Count);
        
        for (int i = 0; i < testData.Count; i++)
        {
            Assert.AreEqual(testData[i]["Name"], result[i]["Name"]);
            Assert.AreEqual(testData[i]["Age"].ToString(), result[i]["Age"]);
            Assert.AreEqual(testData[i]["City"], result[i]["City"]);
        }
    }
#endif

    private List<TestModelPerson> GetTestData()
    {
        return new List<TestModelPerson>
        {
            new() { Name = "张三", Age = 25, CreateTime = DateTime.Now },
            new() { Name = "李四", Age = 30, CreateTime = DateTime.Now.AddDays(-1) },
            new() { Name = "王五", Age = 35, CreateTime = DateTime.Now.AddDays(-2) }
        };
    }

    private List<TestModelPerson> GetLargeTestData(int count)
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
                CreateTime = DateTime.Now.AddDays(-random.Next(365))
            });
        }
        
        return result;
    }
}
