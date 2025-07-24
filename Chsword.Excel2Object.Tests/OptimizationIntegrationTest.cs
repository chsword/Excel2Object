using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class OptimizationIntegrationTest
    {
        [TestMethod]
        public void Test_AllOptimizationComponents_Available()
        {
            // 验证所有优化组件都可用
            var assembly = Assembly.LoadFrom(@"Chsword.Excel2Object.dll");
            
            // 检查 ObjectPoolManager
            var objectPoolType = assembly.GetType("Chsword.Excel2Object.Internal.ObjectPoolManager");
            Assert.IsNotNull(objectPoolType, "ObjectPoolManager 应该存在");
            
            // 检查 ExpressionCache
            var expressionCacheType = assembly.GetType("Chsword.Excel2Object.Internal.ExpressionCache");
            Assert.IsNotNull(expressionCacheType, "ExpressionCache 应该存在");
            
            // 检查 PerformanceMonitor
            var performanceMonitorType = assembly.GetType("Chsword.Excel2Object.Internal.PerformanceMonitor");
            Assert.IsNotNull(performanceMonitorType, "PerformanceMonitor 应该存在");
            
            // 检查 ParallelProcessor
            var parallelProcessorType = assembly.GetType("Chsword.Excel2Object.Internal.ParallelProcessor");
            Assert.IsNotNull(parallelProcessorType, "ParallelProcessor 应该存在");

            Console.WriteLine("✅ 所有优化组件验证成功");
        }

        [TestMethod] 
        public void Test_ObjectPoolManager_Methods()
        {
            // 通过反射测试 ObjectPoolManager 的方法
            var assembly = Assembly.LoadFrom(@"Chsword.Excel2Object.dll");
            var objectPoolType = assembly.GetType("Chsword.Excel2Object.Internal.ObjectPoolManager");
            
            // 检查关键方法存在
            var getStringBuilderMethod = objectPoolType.GetMethod("GetStringBuilder", BindingFlags.Public | BindingFlags.Static);
            Assert.IsNotNull(getStringBuilderMethod, "GetStringBuilder 方法应该存在");
            
            var returnStringBuilderMethod = objectPoolType.GetMethod("ReturnStringBuilder", BindingFlags.Public | BindingFlags.Static);
            Assert.IsNotNull(returnStringBuilderMethod, "ReturnStringBuilder 方法应该存在");

            Console.WriteLine("✅ ObjectPoolManager 方法验证成功");
        }

        [TestMethod]
        public void Test_ExpressionCache_Methods()
        {
            // 通过反射测试 ExpressionCache 的方法
            var assembly = Assembly.LoadFrom(@"Chsword.Excel2Object.dll");
            var expressionCacheType = assembly.GetType("Chsword.Excel2Object.Internal.ExpressionCache");
            
            // 检查关键方法存在
            var clearMethod = expressionCacheType.GetMethod("Clear", BindingFlags.Public | BindingFlags.Static);
            Assert.IsNotNull(clearMethod, "Clear 方法应该存在");
            
            var getStatisticsMethod = expressionCacheType.GetMethod("GetStatistics", BindingFlags.Public | BindingFlags.Static);
            Assert.IsNotNull(getStatisticsMethod, "GetStatistics 方法应该存在");

            Console.WriteLine("✅ ExpressionCache 方法验证成功");
        }

        [TestMethod]
        public void Test_PerformanceMonitor_Methods()
        {
            // 通过反射测试 PerformanceMonitor 的方法
            var assembly = Assembly.LoadFrom(@"Chsword.Excel2Object.dll");
            var performanceMonitorType = assembly.GetType("Chsword.Excel2Object.Internal.PerformanceMonitor");
            
            // 检查关键方法存在
            var monitorMethods = performanceMonitorType.GetMethods(BindingFlags.Public | BindingFlags.Static)
                .Where(m => m.Name == "Monitor").ToArray();
            Assert.IsTrue(monitorMethods.Length > 0, "Monitor 方法应该存在");

            Console.WriteLine("✅ PerformanceMonitor 方法验证成功");
        }

        [TestMethod]
        public void Test_OptimizationIntegration_WithRealData()
        {
            // 创建测试数据
            var testData = new List<OptimizationTestData>();
            for (int i = 0; i < 500; i++)
            {
                testData.Add(new OptimizationTestData
                {
                    Id = i,
                    Name = $"测试数据 {i}",
                    Value = i * 1.5m,
                    CreateDate = DateTime.Now.AddDays(-i),
                    IsActive = i % 2 == 0
                });
            }

            // 测试导出
            var exportBytes = ExcelHelper.ObjectToExcelBytes(testData, ExcelType.Xlsx);
            Assert.IsNotNull(exportBytes, "Excel导出应该成功");
            Assert.IsTrue(exportBytes.Length > 0, "导出的字节数组不应为空");

            // 测试导入
            var importedData = ExcelHelper.ExcelToObject<OptimizationTestData>(exportBytes);
            Assert.IsNotNull(importedData, "Excel导入应该成功");
            
            var importedList = importedData.ToList();
            Assert.AreEqual(testData.Count, importedList.Count, "导入的数据数量应该与导出的一致");

            // 验证数据完整性
            for (int i = 0; i < Math.Min(10, testData.Count); i++)
            {
                Assert.AreEqual(testData[i].Id, importedList[i].Id, $"第{i}条记录的ID应该一致");
                Assert.AreEqual(testData[i].Name, importedList[i].Name, $"第{i}条记录的Name应该一致");
                Assert.AreEqual(testData[i].Value, importedList[i].Value, $"第{i}条记录的Value应该一致");
                Assert.AreEqual(testData[i].IsActive, importedList[i].IsActive, $"第{i}条记录的IsActive应该一致");
            }

            Console.WriteLine($"✅ 优化集成测试成功：处理了 {testData.Count} 条记录");
            Console.WriteLine($"  - 导出字节大小: {exportBytes.Length:N0} bytes");
            Console.WriteLine($"  - 导入数据数量: {importedList.Count} 条");
        }

        [TestMethod]
        public void Test_LargeDataset_PerformanceCheck()
        {
            // 创建大量测试数据
            var largeTestData = new List<OptimizationTestData>();
            for (int i = 0; i < 2000; i++)
            {
                largeTestData.Add(new OptimizationTestData
                {
                    Id = i,
                    Name = $"大数据集测试 {i}",
                    Value = (decimal)(i * Math.PI),
                    CreateDate = DateTime.Now.AddMinutes(-i),
                    IsActive = i % 3 == 0
                });
            }

            var startTime = DateTime.Now;

            // 执行导出导入操作
            var exportBytes = ExcelHelper.ObjectToExcelBytes(largeTestData, ExcelType.Xlsx);
            var importedData = ExcelHelper.ExcelToObject<OptimizationTestData>(exportBytes);
            var importedList = importedData.ToList();

            var elapsedTime = DateTime.Now - startTime;

            // 验证结果
            Assert.AreEqual(largeTestData.Count, importedList.Count, "大数据集处理后数量应该一致");
            Assert.IsTrue(elapsedTime.TotalSeconds < 30, "大数据集处理应该在30秒内完成");

            Console.WriteLine($"✅ 大数据集性能测试成功：");
            Console.WriteLine($"  - 处理数据量: {largeTestData.Count:N0} 条");
            Console.WriteLine($"  - 耗时: {elapsedTime.TotalMilliseconds:F0} ms");
            Console.WriteLine($"  - 平均每条记录: {elapsedTime.TotalMilliseconds / largeTestData.Count:F2} ms");
        }
    }

    public class OptimizationTestData
    {
        [ExcelColumn("编号")]
        public int Id { get; set; }

        [ExcelColumn("名称")]
        public string Name { get; set; }

        [ExcelColumn("值")]
        public decimal Value { get; set; }

        [ExcelColumn("创建日期")]
        public DateTime CreateDate { get; set; }

        [ExcelColumn("是否活跃")]
        public bool IsActive { get; set; }
    }
}
