using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Chsword.Excel2Object;

namespace AsyncExample
{
    /// <summary>
    /// 演示 Excel2Object 异步功能的示例
    /// </summary>
    public class AsyncUsageExample
    {
        public class Employee
        {
            [ExcelTitle("姓名")]
            public string Name { get; set; } = string.Empty;

            [ExcelTitle("年龄")]
            public int Age { get; set; }

            [ExcelTitle("部门")]
            public string Department { get; set; } = string.Empty;

            [ExcelTitle("薪资")]
            public decimal Salary { get; set; }
        }

        /// <summary>
        /// 异步读取Excel文件示例
        /// </summary>
        public static async Task<IEnumerable<Employee>?> ReadExcelFileAsync(string filePath)
        {
            try
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(5));
                
                // 使用异步方法读取Excel文件
                var employees = await ExcelHelper.ExcelToObjectAsync<Employee>(
                    filePath, 
                    sheetTitle: null, 
                    cancellationToken: cts.Token);
                
                Console.WriteLine($"成功异步读取 {employees?.Count() ?? 0} 条员工记录");
                return employees;
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("操作被取消");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"读取Excel文件时发生错误: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 异步写入Excel文件示例
        /// </summary>
        public static async Task WriteExcelFileAsync(IEnumerable<Employee> employees, string outputPath)
        {
            try
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(3));
                
                // 使用异步方法写入Excel文件
                await ExcelHelper.ObjectToExcelAsync(
                    employees, 
                    outputPath, 
                    cancellationToken: cts.Token);
                
                Console.WriteLine($"成功异步写入Excel文件: {outputPath}");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("写入操作被取消");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"写入Excel文件时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 异步处理大型Excel文件的示例
        /// </summary>
        public static async Task ProcessLargeExcelFileAsync(string inputPath, string outputPath)
        {
            try
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(10));
                
                Console.WriteLine("开始异步处理大型Excel文件...");
                
                // 异步读取
                var employees = await ExcelHelper.ExcelToObjectAsync<Employee>(
                    inputPath, 
                    cancellationToken: cts.Token);
                
                if (employees == null)
                {
                    Console.WriteLine("没有读取到数据");
                    return;
                }

                // 处理数据（例如：给所有员工涨薪10%）
                var processedEmployees = employees.Select(emp => new Employee
                {
                    Name = emp.Name,
                    Age = emp.Age,
                    Department = emp.Department,
                    Salary = emp.Salary * 1.1m // 涨薪10%
                }).ToList();

                Console.WriteLine($"处理了 {processedEmployees.Count} 条记录");

                // 异步写入处理后的数据
                await ExcelHelper.ObjectToExcelAsync(
                    processedEmployees, 
                    outputPath, 
                    ExcelType.Xlsx,
                    cancellationToken: cts.Token);
                
                Console.WriteLine("大型Excel文件处理完成！");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("大型文件处理操作被取消");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理大型Excel文件时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 演示并行处理多个Excel文件
        /// </summary>
        public static async Task ProcessMultipleFilesAsync(string[] filePaths, string outputDirectory)
        {
            var tasks = filePaths.Select(async (filePath, index) =>
            {
                try
                {
                    using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(5));
                    
                    var employees = await ExcelHelper.ExcelToObjectAsync<Employee>(
                        filePath, 
                        cancellationToken: cts.Token);
                    
                    if (employees != null)
                    {
                        var outputPath = Path.Combine(outputDirectory, $"processed_{index}.xlsx");
                        await ExcelHelper.ObjectToExcelAsync(
                            employees, 
                            outputPath, 
                            ExcelType.Xlsx,
                            cancellationToken: cts.Token);
                        
                        Console.WriteLine($"文件 {filePath} 处理完成，输出到 {outputPath}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"处理文件 {filePath} 时发生错误: {ex.Message}");
                }
            });

            await Task.WhenAll(tasks);
            Console.WriteLine("所有文件处理完成！");
        }
    }
}
