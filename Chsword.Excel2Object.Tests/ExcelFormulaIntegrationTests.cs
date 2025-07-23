using System;
using System.IO;
using Chsword.Excel2Object;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExcelFormulaIntegrationTests
    {
        private string _testFilePath;

        [TestInitialize]
        public void Setup()
        {
            _testFilePath = Path.Combine(Path.GetTempPath(), $"test_formula_{Guid.NewGuid()}.xlsx");
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (File.Exists(_testFilePath))
            {
                File.Delete(_testFilePath);
            }
        }

        [TestMethod]
        public void TestMathFormulas_WriteAndCalculate()
        {
            // 创建工作簿和工作表
            using var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("MathTest");

            // 设置测试数据
            var row0 = sheet.CreateRow(0);
            row0.CreateCell(0).SetCellValue(10.5);  // A1
            row0.CreateCell(1).SetCellValue(5.2);   // B1
            row0.CreateCell(2).SetCellValue(-3.7);  // C1

            var row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellValue(20);    // A2
            row1.CreateCell(1).SetCellValue(8);     // B2

            // 设置公式
            var row2 = sheet.CreateRow(2);
            row2.CreateCell(0).SetCellFormula("SUM(A1:B1)");      // A3: SUM(10.5, 5.2) = 15.7
            row2.CreateCell(1).SetCellFormula("ABS(C1)");         // B3: ABS(-3.7) = 3.7
            row2.CreateCell(2).SetCellFormula("ROUND(A1,0)");     // C3: ROUND(10.5, 0) = 11
            row2.CreateCell(3).SetCellFormula("MAX(A1:C1)");      // D3: MAX(10.5, 5.2, -3.7) = 10.5
            row2.CreateCell(4).SetCellFormula("AVERAGE(A1:B2)");  // E3: AVERAGE(10.5, 5.2, 20, 8) = 10.925

            // 保存文件
            using (var fileStream = new FileStream(_testFilePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }

            // 重新打开文件并验证计算结果
            using var readWorkbook = new XSSFWorkbook(new FileStream(_testFilePath, FileMode.Open, FileAccess.Read));
            var readSheet = readWorkbook.GetSheetAt(0);
            
            // 强制重新计算公式
            readWorkbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

            // 验证计算结果
            var resultRow = readSheet.GetRow(2);
            
            // SUM(A1:B1) = 15.7
            Assert.AreEqual(15.7, resultRow.GetCell(0).NumericCellValue, 0.001, "SUM formula calculation failed");
            
            // ABS(C1) = 3.7
            Assert.AreEqual(3.7, resultRow.GetCell(1).NumericCellValue, 0.001, "ABS formula calculation failed");
            
            // ROUND(A1,0) = 11
            Assert.AreEqual(11, resultRow.GetCell(2).NumericCellValue, 0.001, "ROUND formula calculation failed");
            
            // MAX(A1:C1) = 10.5
            Assert.AreEqual(10.5, resultRow.GetCell(3).NumericCellValue, 0.001, "MAX formula calculation failed");
            
            // AVERAGE(A1:B2) = 10.925
            Assert.AreEqual(10.925, resultRow.GetCell(4).NumericCellValue, 0.001, "AVERAGE formula calculation failed");
        }

        [TestMethod]
        public void TestTextFormulas_WriteAndCalculate()
        {
            using var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("TextTest");

            // 设置测试数据
            var row0 = sheet.CreateRow(0);
            row0.CreateCell(0).SetCellValue("Hello");    // A1
            row0.CreateCell(1).SetCellValue("World");    // B1
            row0.CreateCell(2).SetCellValue("  Test  "); // C1

            // 设置文本公式
            var row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellFormula("CONCATENATE(A1,\" \",B1)");  // A2: "Hello World"
            row1.CreateCell(1).SetCellFormula("LEN(A1)");                   // B2: 5
            row1.CreateCell(2).SetCellFormula("UPPER(A1)");                 // C2: "HELLO"
            row1.CreateCell(3).SetCellFormula("TRIM(C1)");                  // D2: "Test"
            row1.CreateCell(4).SetCellFormula("LEFT(A1,3)");                // E2: "Hel"

            // 保存文件
            using (var fileStream = new FileStream(_testFilePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }

            // 重新打开文件并验证计算结果
            using var readWorkbook = new XSSFWorkbook(new FileStream(_testFilePath, FileMode.Open, FileAccess.Read));
            var readSheet = readWorkbook.GetSheetAt(0);
            readWorkbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

            var resultRow = readSheet.GetRow(1);
            
            // CONCATENATE(A1," ",B1) = "Hello World"
            Assert.AreEqual("Hello World", resultRow.GetCell(0).StringCellValue, "CONCATENATE formula calculation failed");
            
            // LEN(A1) = 5
            Assert.AreEqual(5, resultRow.GetCell(1).NumericCellValue, "LEN formula calculation failed");
            
            // UPPER(A1) = "HELLO"
            Assert.AreEqual("HELLO", resultRow.GetCell(2).StringCellValue, "UPPER formula calculation failed");
            
            // TRIM(C1) = "Test"
            Assert.AreEqual("Test", resultRow.GetCell(3).StringCellValue, "TRIM formula calculation failed");
            
            // LEFT(A1,3) = "Hel"
            Assert.AreEqual("Hel", resultRow.GetCell(4).StringCellValue, "LEFT formula calculation failed");
        }

        [TestMethod]
        public void TestDateTimeFormulas_WriteAndCalculate()
        {
            using var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("DateTest");

            // 设置日期数据
            var row0 = sheet.CreateRow(0);
            row0.CreateCell(0).SetCellValue(new DateTime(2023, 12, 25)); // A1: 2023-12-25
            row0.CreateCell(1).SetCellValue(new DateTime(2023, 1, 1));   // B1: 2023-01-01

            // 设置日期公式
            var row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellFormula("YEAR(A1)");              // A2: 2023
            row1.CreateCell(1).SetCellFormula("MONTH(A1)");             // B2: 12
            row1.CreateCell(2).SetCellFormula("DAY(A1)");               // C2: 25
            row1.CreateCell(3).SetCellFormula("A1-B1");                 // D2: 日期差（简单减法）
            row1.CreateCell(4).SetCellFormula("DATE(2024,1,1)");        // E2: 2024-01-01

            // 保存文件
            using (var fileStream = new FileStream(_testFilePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }

            // 重新打开文件并验证计算结果
            using var readWorkbook = new XSSFWorkbook(new FileStream(_testFilePath, FileMode.Open, FileAccess.Read));
            var readSheet = readWorkbook.GetSheetAt(0);
            readWorkbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

            var resultRow = readSheet.GetRow(1);
            
            // YEAR(A1) = 2023
            Assert.AreEqual(2023, resultRow.GetCell(0).NumericCellValue, "YEAR formula calculation failed");
            
            // MONTH(A1) = 12
            Assert.AreEqual(12, resultRow.GetCell(1).NumericCellValue, "MONTH formula calculation failed");
            
            // DAY(A1) = 25
            Assert.AreEqual(25, resultRow.GetCell(2).NumericCellValue, "DAY formula calculation failed");
            
            // DAYS(A1,B1) = 358 (2023年有365天，从1月1日到12月25日)
            Assert.AreEqual(358, resultRow.GetCell(3).NumericCellValue, "Date difference calculation failed");
            
            // DATE(2024,1,1) = 2024-01-01
            var dateCell = resultRow.GetCell(4);
            var expectedDate = new DateTime(2024, 1, 1);
            Assert.AreEqual(expectedDate, dateCell.DateCellValue, "DATE formula calculation failed");
        }

        [TestMethod]
        public void TestConditionalFormulas_WriteAndCalculate()
        {
            using var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("ConditionalTest");

            // 设置测试数据
            var row0 = sheet.CreateRow(0);
            row0.CreateCell(0).SetCellValue(85);     // A1: 85
            row0.CreateCell(1).SetCellValue(92);     // B1: 92
            row0.CreateCell(2).SetCellValue(78);     // C1: 78

            // 设置条件公式
            var row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellFormula("IF(A1>=90,\"A\",IF(A1>=80,\"B\",\"C\"))");  // A2: "B"
            row1.CreateCell(1).SetCellFormula("IF(B1>=90,\"A\",IF(B1>=80,\"B\",\"C\"))");  // B2: "A"
            row1.CreateCell(2).SetCellFormula("IF(C1>=90,\"A\",IF(C1>=80,\"B\",\"C\"))");  // C2: "C"
            row1.CreateCell(3).SetCellFormula("COUNTIF(A1:C1,\">80\")");                   // D2: 2
            row1.CreateCell(4).SetCellFormula("SUMIF(A1:C1,\">80\")");                     // E2: 177

            // 保存文件
            using (var fileStream = new FileStream(_testFilePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }

            // 重新打开文件并验证计算结果
            using var readWorkbook = new XSSFWorkbook(new FileStream(_testFilePath, FileMode.Open, FileAccess.Read));
            var readSheet = readWorkbook.GetSheetAt(0);
            readWorkbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

            var resultRow = readSheet.GetRow(1);
            
            // IF formula for A1 (85) = "B"
            Assert.AreEqual("B", resultRow.GetCell(0).StringCellValue, "IF formula for A1 failed");
            
            // IF formula for B1 (92) = "A"
            Assert.AreEqual("A", resultRow.GetCell(1).StringCellValue, "IF formula for B1 failed");
            
            // IF formula for C1 (78) = "C"
            Assert.AreEqual("C", resultRow.GetCell(2).StringCellValue, "IF formula for C1 failed");
            
            // COUNTIF(A1:C1,">80") = 2 (85, 92)
            Assert.AreEqual(2, resultRow.GetCell(3).NumericCellValue, "COUNTIF formula calculation failed");
            
            // SUMIF(A1:C1,">80") = 177 (85 + 92)
            Assert.AreEqual(177, resultRow.GetCell(4).NumericCellValue, "SUMIF formula calculation failed");
        }

        [TestMethod]
        public void TestComplexFormulas_WriteAndCalculate()
        {
            using var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("ComplexTest");

            // 创建一个小的数据表
            var headers = new[] { "Name", "Score1", "Score2", "Score3" };
            var row0 = sheet.CreateRow(0);
            for (int i = 0; i < headers.Length; i++)
            {
                row0.CreateCell(i).SetCellValue(headers[i]);
            }

            // 添加数据
            var data = new object[,]
            {
                { "Alice", 85, 92, 88 },
                { "Bob", 78, 85, 90 },
                { "Charlie", 92, 88, 85 }
            };

            for (int row = 0; row < data.GetLength(0); row++)
            {
                var dataRow = sheet.CreateRow(row + 1);
                for (int col = 0; col < data.GetLength(1); col++)
                {
                    var cell = dataRow.CreateCell(col);
                    if (data[row, col] is string str)
                        cell.SetCellValue(str);
                    else if (data[row, col] is int num)
                        cell.SetCellValue(num);
                }
            }

            // 添加复杂公式
            var row5 = sheet.CreateRow(5);
            row5.CreateCell(0).SetCellValue("Average:");
            row5.CreateCell(1).SetCellFormula("AVERAGE(B2:B4)");     // Score1平均值
            row5.CreateCell(2).SetCellFormula("AVERAGE(C2:C4)");     // Score2平均值
            row5.CreateCell(3).SetCellFormula("AVERAGE(D2:D4)");     // Score3平均值

            var row6 = sheet.CreateRow(6);
            row6.CreateCell(0).SetCellValue("Max:");
            row6.CreateCell(1).SetCellFormula("MAX(B2:D4)");         // 所有分数中的最大值

            var row7 = sheet.CreateRow(7);
            row7.CreateCell(0).SetCellValue("Count >85:");
            row7.CreateCell(1).SetCellFormula("COUNTIF(B2:D4,\">85\")"); // 大于85分的数量

            // 保存文件
            using (var fileStream = new FileStream(_testFilePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }

            // 重新打开文件并验证计算结果
            using var readWorkbook = new XSSFWorkbook(new FileStream(_testFilePath, FileMode.Open, FileAccess.Read));
            var readSheet = readWorkbook.GetSheetAt(0);
            readWorkbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

            // 验证平均值计算
            var avgRow = readSheet.GetRow(5);
            Assert.AreEqual(85.0, avgRow.GetCell(1).NumericCellValue, 0.001, "Score1 average calculation failed"); // (85+78+92)/3 = 85
            Assert.AreEqual(88.33, avgRow.GetCell(2).NumericCellValue, 0.01, "Score2 average calculation failed");  // (92+85+88)/3 = 88.33
            Assert.AreEqual(87.67, avgRow.GetCell(3).NumericCellValue, 0.01, "Score3 average calculation failed");  // (88+90+85)/3 = 87.67

            // 验证最大值
            var maxRow = readSheet.GetRow(6);
            Assert.AreEqual(92, maxRow.GetCell(1).NumericCellValue, "Max value calculation failed");

            // 验证条件计数
            var countRow = readSheet.GetRow(7);
            Assert.AreEqual(5, countRow.GetCell(1).NumericCellValue, "Count >85 calculation failed"); // 92,88,92,88,90 = 5个
        }
    }
}
