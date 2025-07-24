using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using Chsword.Excel2Object;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class BasicOptimizationIntegrationTest
    {
        [TestMethod]
        public void Test_ExcelImporter_WithOptimizations()
        {
            // 创建一个简单的Excel测试数据
            var testData = new List<TestPerson>
            {
                new TestPerson { Name = "张三", Age = 25, Score = 85.5 },
                new TestPerson { Name = "李四", Age = 30, Score = 92.0 },
                new TestPerson { Name = "王五", Age = 28, Score = 78.5 }
            };

            var filePath = System.IO.Path.GetTempFileName() + ".xlsx";
            
            try
            {
                // 导出数据到Excel
                ExcelHelper.ObjectToExcel(testData, filePath);

                // 导入数据
                var importer = new ExcelImporter();
                var importedData = importer.ExcelToObject<TestPerson>(filePath).ToList();

                // 验证数据
                Assert.AreEqual(testData.Count, importedData.Count);
                
                for (int i = 0; i < testData.Count; i++)
                {
                    Assert.AreEqual(testData[i].Name, importedData[i].Name);
                    Assert.AreEqual(testData[i].Age, importedData[i].Age);
                    Assert.AreEqual(testData[i].Score, importedData[i].Score, 0.01);
                }

                Console.WriteLine($"成功测试了 {testData.Count} 条记录的导入导出");
            }
            finally
            {
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }
            }
        }

        [TestMethod]
        public void Test_LargeDataSet_Performance()
        {
            // 测试大数据集的性能
            var largeTestData = new List<TestPerson>();
            for (int i = 0; i < 1000; i++)
            {
                largeTestData.Add(new TestPerson 
                { 
                    Name = $"测试用户{i}", 
                    Age = 20 + (i % 50), 
                    Score = 60 + (i % 40) 
                });
            }

            var filePath = System.IO.Path.GetTempFileName() + ".xlsx";
            
            try
            {
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();

                // 导出大数据集
                ExcelHelper.ObjectToExcel(largeTestData, filePath);

                // 导入大数据集
                var importer = new ExcelImporter();
                var importedData = importer.ExcelToObject<TestPerson>(filePath).ToList();

                stopwatch.Stop();

                // 验证数据完整性
                Assert.AreEqual(largeTestData.Count, importedData.Count);
                
                Console.WriteLine($"处理 {largeTestData.Count} 条记录耗时: {stopwatch.ElapsedMilliseconds}ms");
                
                // 性能期望：1000条记录应该在合理时间内完成
                Assert.IsTrue(stopwatch.ElapsedMilliseconds < 10000, "大数据集处理应该在10秒内完成");
            }
            finally
            {
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }
            }
        }
    }

    public class TestPerson
    {
        [ExcelColumn("姓名")]
        public string Name { get; set; }

        [ExcelColumn("年龄")]
        public int Age { get; set; }

        [ExcelColumn("分数")]
        public double Score { get; set; }
    }
}
