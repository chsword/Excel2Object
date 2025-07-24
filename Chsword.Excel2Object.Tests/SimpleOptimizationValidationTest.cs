using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class SimpleOptimizationValidationTest
    {
        [TestMethod]
        public void Test_ObjectPoolManager_Available()
        {
            // Test that we can access the ObjectPoolManager via reflection
            var objectPoolType = Type.GetType("Chsword.Excel2Object.Internal.ObjectPoolManager, Chsword.Excel2Object");
            Assert.IsNotNull(objectPoolType, "ObjectPoolManager应该可以通过反射访问");
            
            Console.WriteLine("✅ ObjectPoolManager 类型可用");
        }

        [TestMethod]
        public void Test_ExpressionCache_Available()
        {
            // Test that we can access the ExpressionCache via reflection
            var expressionCacheType = Type.GetType("Chsword.Excel2Object.Internal.ExpressionCache, Chsword.Excel2Object");
            Assert.IsNotNull(expressionCacheType, "ExpressionCache应该可以通过反射访问");
            
            Console.WriteLine("✅ ExpressionCache 类型可用");
        }

        [TestMethod]
        public void Test_PerformanceMonitor_Available()
        {
            // Test that we can access the PerformanceMonitor via reflection
            var performanceMonitorType = Type.GetType("Chsword.Excel2Object.Internal.PerformanceMonitor, Chsword.Excel2Object");
            Assert.IsNotNull(performanceMonitorType, "PerformanceMonitor应该可以通过反射访问");
            
            Console.WriteLine("✅ PerformanceMonitor 类型可用");
        }

        [TestMethod]
        public void Test_ParallelProcessor_Available()
        {
            // Test that we can access the ParallelProcessor via reflection
            var parallelProcessorType = Type.GetType("Chsword.Excel2Object.Internal.ParallelProcessor, Chsword.Excel2Object");
            Assert.IsNotNull(parallelProcessorType, "ParallelProcessor应该可以通过反射访问");
            
            Console.WriteLine("✅ ParallelProcessor 类型可用");
        }

        [TestMethod] 
        public void Test_BasicExcelFunctionality_Works()
        {
            // Test basic Excel functionality still works
            var testData = new System.Collections.Generic.List<TestItem>
            {
                new TestItem { Name = "测试1", Value = 100 },
                new TestItem { Name = "测试2", Value = 200 }
            };

            var bytes = ExcelHelper.ObjectToExcelBytes(testData, ExcelType.Xlsx);
            Assert.IsNotNull(bytes, "Excel导出应该成功");
            Assert.IsTrue(bytes.Length > 0, "导出的Excel字节不应为空");

            var imported = ExcelHelper.ExcelToObject<TestItem>(bytes);
            Assert.IsNotNull(imported, "Excel导入应该成功");
            var importedList = imported.ToList();
            Assert.AreEqual(2, importedList.Count, "应该导入2条记录");

            Console.WriteLine($"✅ 基础Excel功能正常，处理了 {importedList.Count} 条记录");
        }
    }

    public class TestItem
    {
        [ExcelColumn("名称")]
        public string Name { get; set; }

        [ExcelColumn("值")]
        public int Value { get; set; }
    }
}
