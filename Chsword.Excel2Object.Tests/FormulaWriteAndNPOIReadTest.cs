using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Chsword.Excel2Object;

namespace Chsword.Excel2Object.Tests
{
    /// <summary>
    /// 测试Excel公式写入和NPOI读取功能
    /// </summary>
    [TestClass]
    public class FormulaWriteAndNPOIReadTest : BaseExcelTest
    {
        /// <summary>
        /// 测试数学公式：直接使用NPOI写入公式，验证计算结果
        /// </summary>
        [TestMethod]
        public void Test_DirectNPOI_MathFormulas()
        {
            var testPath = GetFilePath("math_formulas.xlsx");

            // 使用NPOI直接创建包含公式的Excel文件
            using (var workbook = new XSSFWorkbook())
            {
                var sheet = workbook.CreateSheet("数学公式");
                
                // 创建标题行
                var headerRow = sheet.CreateRow(0);
                headerRow.CreateCell(0).SetCellValue("数值1");
                headerRow.CreateCell(1).SetCellValue("数值2");
                headerRow.CreateCell(2).SetCellValue("求和");
                headerRow.CreateCell(3).SetCellValue("乘积");
                headerRow.CreateCell(4).SetCellValue("平均值");

                // 创建数据行
                var dataRow = sheet.CreateRow(1);
                dataRow.CreateCell(0).SetCellValue(10);
                dataRow.CreateCell(1).SetCellValue(20);
                
                // 写入公式
                dataRow.CreateCell(2).SetCellFormula("A2+B2");       // 求和：10+20=30
                dataRow.CreateCell(3).SetCellFormula("A2*B2");       // 乘积：10*20=200
                dataRow.CreateCell(4).SetCellFormula("(A2+B2)/2");   // 平均值：(10+20)/2=15

                // 强制重新计算公式
                sheet.ForceFormulaRecalculation = true;

                // 保存文件
                using (var fs = new FileStream(testPath, FileMode.Create))
                {
                    workbook.Write(fs);
                }
            }

            // 使用NPOI读取并验证计算结果
            using (var fs = new FileStream(testPath, FileMode.Open))
            using (var workbook = new XSSFWorkbook(fs))
            {
                var sheet = workbook.GetSheetAt(0);
                var dataRow = sheet.GetRow(1);

                // 创建公式计算器
                var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();

                // 验证求和公式结果
                var sumCell = dataRow.GetCell(2);
                Assert.AreEqual(CellType.Formula, sumCell.CellType);
                Assert.AreEqual("A2+B2", sumCell.CellFormula);
                
                var sumResult = evaluator.Evaluate(sumCell);
                Assert.AreEqual(30.0, sumResult.NumberValue, 0.001);

                // 验证乘积公式结果
                var productCell = dataRow.GetCell(3);
                Assert.AreEqual(CellType.Formula, productCell.CellType);
                Assert.AreEqual("A2*B2", productCell.CellFormula);
                
                var productResult = evaluator.Evaluate(productCell);
                Assert.AreEqual(200.0, productResult.NumberValue, 0.001);

                // 验证平均值公式结果
                var avgCell = dataRow.GetCell(4);
                Assert.AreEqual(CellType.Formula, avgCell.CellType);
                Assert.AreEqual("(A2+B2)/2", avgCell.CellFormula);
                
                var avgResult = evaluator.Evaluate(avgCell);
                Assert.AreEqual(15.0, avgResult.NumberValue, 0.001);
            }
        }

        /// <summary>
        /// 测试统计公式
        /// </summary>
        [TestMethod]
        public void Test_DirectNPOI_StatisticsFormulas()
        {
            var testPath = GetFilePath("statistics_formulas.xlsx");

            using (var workbook = new XSSFWorkbook())
            {
                var sheet = workbook.CreateSheet("统计公式");
                
                // 创建标题行
                var headerRow = sheet.CreateRow(0);
                headerRow.CreateCell(0).SetCellValue("科目");
                headerRow.CreateCell(1).SetCellValue("成绩1");
                headerRow.CreateCell(2).SetCellValue("成绩2");
                headerRow.CreateCell(3).SetCellValue("成绩3");
                headerRow.CreateCell(4).SetCellValue("最高分");
                headerRow.CreateCell(5).SetCellValue("最低分");
                headerRow.CreateCell(6).SetCellValue("计数");

                // 创建数据行
                var dataRow = sheet.CreateRow(1);
                dataRow.CreateCell(0).SetCellValue("数学");
                dataRow.CreateCell(1).SetCellValue(85);
                dataRow.CreateCell(2).SetCellValue(90);
                dataRow.CreateCell(3).SetCellValue(78);
                
                // 写入统计公式
                dataRow.CreateCell(4).SetCellFormula("MAX(B2:D2)");     // 最高分：MAX(85,90,78)=90
                dataRow.CreateCell(5).SetCellFormula("MIN(B2:D2)");     // 最低分：MIN(85,90,78)=78
                dataRow.CreateCell(6).SetCellFormula("COUNT(B2:D2)");   // 计数：COUNT(85,90,78)=3

                sheet.ForceFormulaRecalculation = true;

                using (var fs = new FileStream(testPath, FileMode.Create))
                {
                    workbook.Write(fs);
                }
            }

            // 验证结果
            using (var fs = new FileStream(testPath, FileMode.Open))
            using (var workbook = new XSSFWorkbook(fs))
            {
                var sheet = workbook.GetSheetAt(0);
                var dataRow = sheet.GetRow(1);
                var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();

                // 验证最高分
                var maxCell = dataRow.GetCell(4);
                Assert.AreEqual("MAX(B2:D2)", maxCell.CellFormula);
                var maxResult = evaluator.Evaluate(maxCell);
                Assert.AreEqual(90.0, maxResult.NumberValue, 0.001);

                // 验证最低分
                var minCell = dataRow.GetCell(5);
                Assert.AreEqual("MIN(B2:D2)", minCell.CellFormula);
                var minResult = evaluator.Evaluate(minCell);
                Assert.AreEqual(78.0, minResult.NumberValue, 0.001);

                // 验证计数
                var countCell = dataRow.GetCell(6);
                Assert.AreEqual("COUNT(B2:D2)", countCell.CellFormula);
                var countResult = evaluator.Evaluate(countCell);
                Assert.AreEqual(3.0, countResult.NumberValue, 0.001);
            }
        }

        /// <summary>
        /// 测试文本公式
        /// </summary>
        [TestMethod]
        public void Test_DirectNPOI_TextFormulas()
        {
            var testPath = GetFilePath("text_formulas.xlsx");

            using (var workbook = new XSSFWorkbook())
            {
                var sheet = workbook.CreateSheet("文本公式");
                
                var headerRow = sheet.CreateRow(0);
                headerRow.CreateCell(0).SetCellValue("姓");
                headerRow.CreateCell(1).SetCellValue("名");
                headerRow.CreateCell(2).SetCellValue("全名");
                headerRow.CreateCell(3).SetCellValue("长度");
                headerRow.CreateCell(4).SetCellValue("大写");

                var dataRow = sheet.CreateRow(1);
                dataRow.CreateCell(0).SetCellValue("张");
                dataRow.CreateCell(1).SetCellValue("三");
                
                // 文本公式
                dataRow.CreateCell(2).SetCellFormula("A2&B2");         // 连接：张三
                dataRow.CreateCell(3).SetCellFormula("LEN(A2&B2)");    // 长度：LEN("张三")=2
                dataRow.CreateCell(4).SetCellFormula("UPPER(A2&B2)");  // 大写：UPPER("张三")="张三"

                sheet.ForceFormulaRecalculation = true;

                using (var fs = new FileStream(testPath, FileMode.Create))
                {
                    workbook.Write(fs);
                }
            }

            // 验证结果
            using (var fs = new FileStream(testPath, FileMode.Open))
            using (var workbook = new XSSFWorkbook(fs))
            {
                var sheet = workbook.GetSheetAt(0);
                var dataRow = sheet.GetRow(1);
                var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();

                // 验证全名连接
                var fullNameCell = dataRow.GetCell(2);
                Assert.AreEqual("A2&B2", fullNameCell.CellFormula);
                var fullNameResult = evaluator.Evaluate(fullNameCell);
                Assert.AreEqual("张三", fullNameResult.StringValue);

                // 验证长度
                var lengthCell = dataRow.GetCell(3);
                Assert.AreEqual("LEN(A2&B2)", lengthCell.CellFormula);
                var lengthResult = evaluator.Evaluate(lengthCell);
                Assert.AreEqual(2.0, lengthResult.NumberValue, 0.001);

                // 验证大写（中文字符大写转换可能不变）
                var upperCell = dataRow.GetCell(4);
                Assert.AreEqual("UPPER(A2&B2)", upperCell.CellFormula);
                var upperResult = evaluator.Evaluate(upperCell);
                Assert.AreEqual("张三", upperResult.StringValue);
            }
        }

        /// <summary>
        /// 测试日期时间公式
        /// </summary>
        [TestMethod]
        public void Test_DirectNPOI_DateTimeFormulas()
        {
            var testPath = GetFilePath("datetime_formulas.xlsx");

            using (var workbook = new XSSFWorkbook())
            {
                var sheet = workbook.CreateSheet("日期公式");
                
                var headerRow = sheet.CreateRow(0);
                headerRow.CreateCell(0).SetCellValue("开始日期");
                headerRow.CreateCell(1).SetCellValue("结束日期");
                headerRow.CreateCell(2).SetCellValue("天数差");
                headerRow.CreateCell(3).SetCellValue("年份");
                headerRow.CreateCell(4).SetCellValue("月份");

                var dataRow = sheet.CreateRow(1);
                // 使用Excel日期序列号：2024-01-01 和 2024-01-31
                dataRow.CreateCell(0).SetCellValue(new DateTime(2024, 1, 1));
                dataRow.CreateCell(1).SetCellValue(new DateTime(2024, 1, 31));
                
                // 设置日期格式
                var dateStyle = workbook.CreateCellStyle();
                var format = workbook.CreateDataFormat();
                dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
                dataRow.GetCell(0).CellStyle = dateStyle;
                dataRow.GetCell(1).CellStyle = dateStyle;
                
                // 日期公式
                dataRow.CreateCell(2).SetCellFormula("B2-A2");         // 天数差：31-1=30
                dataRow.CreateCell(3).SetCellFormula("YEAR(A2)");      // 年份：YEAR(2024-01-01)=2024
                dataRow.CreateCell(4).SetCellFormula("MONTH(A2)");     // 月份：MONTH(2024-01-01)=1

                sheet.ForceFormulaRecalculation = true;

                using (var fs = new FileStream(testPath, FileMode.Create))
                {
                    workbook.Write(fs);
                }
            }

            // 验证结果
            using (var fs = new FileStream(testPath, FileMode.Open))
            using (var workbook = new XSSFWorkbook(fs))
            {
                var sheet = workbook.GetSheetAt(0);
                var dataRow = sheet.GetRow(1);
                var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();

                // 验证天数差
                var daysCell = dataRow.GetCell(2);
                Assert.AreEqual("B2-A2", daysCell.CellFormula);
                var daysResult = evaluator.Evaluate(daysCell);
                Assert.AreEqual(30.0, daysResult.NumberValue, 0.001);

                // 验证年份
                var yearCell = dataRow.GetCell(3);
                Assert.AreEqual("YEAR(A2)", yearCell.CellFormula);
                var yearResult = evaluator.Evaluate(yearCell);
                Assert.AreEqual(2024.0, yearResult.NumberValue, 0.001);

                // 验证月份
                var monthCell = dataRow.GetCell(4);
                Assert.AreEqual("MONTH(A2)", monthCell.CellFormula);
                var monthResult = evaluator.Evaluate(monthCell);
                Assert.AreEqual(1.0, monthResult.NumberValue, 0.001);
            }
        }
    }
}