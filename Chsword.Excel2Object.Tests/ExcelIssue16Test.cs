using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;

namespace Chsword.Excel2Object.Tests
{
    /// <summary>
    /// for config and formula
    /// </summary>
    [TestClass]
    public class ExcelIssue16Test: BaseExcelTest
    {
        [TestMethod]
        public void FormulaColumnExport()
        {

            var list = new List<Dictionary<string, object>>
            {
                new Dictionary<string, object> {["姓名"] = "吴老狗", ["Age"] = "19"
                    //, ["BirthYear"] = null

                },
                new Dictionary<string, object> {["姓名"] = "老林", ["Age"] = "50"}
            };
            var bytes = new ExcelExporter().ObjectToExcelBytes(list, options =>
            {
                options.ExcelType = ExcelType.Xlsx;
                options.FormulaColumns.Add(new FormulaColumn
                {
                    Title = "BirthYear",
                    Formula = c => (int) c["Age"] + DateTime.Now.Year,
                    AfterColumnTitle = "姓名"
                });
            });
            var path = GetFilePath("test.xlsx");
            File.WriteAllBytes(path, bytes);
            var result = ExcelHelper.ExcelToObject<Dictionary<string, object>>(bytes).ToList();
            Console.WriteLine(JsonConvert.SerializeObject(result));
            Assert.AreEqual(
                "2039",
                result[0]["BirthYear"]
            );
        }
    }
}