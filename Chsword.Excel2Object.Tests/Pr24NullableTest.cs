using System;
using System.Collections.Generic;
using System.Linq;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class Pr24NullableTest : BaseExcelTest
    {

        [TestMethod]
        public void ImportExcelNullableType()
        {
            try
            {
                var path = GetLocalFilePath("test.person.xlsx");
                var importer = new ExcelImporter();
                var result = importer.ExcelToObject<TestModelPerson>(path).ToList();
                Assert.AreEqual(2, result.Count);
                Console.WriteLine(JsonConvert.SerializeObject(result));
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        private List<TestModelPerson> GetPersonList()
        {
            return new List<TestModelPerson>
            {
                new TestModelPerson {Name = "张三", Age = 18, Birthday = null},
                new TestModelPerson {Name = "李四", Age = null, Birthday = new DateTime(2021, 10, 10)}
            };
        }

        [TestMethod]
        public void ExportExcelNullableType()
        {
            try
            {
                var personList = GetPersonList();
                var bytes = ExcelHelper.ObjectToExcelBytes(personList);
                Assert.IsTrue(bytes.Length > 0);
                var importer = new ExcelImporter();
                var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();
                Console.WriteLine(JsonConvert.SerializeObject(result));
            }
            catch
            {
                Assert.Fail();
            }
            
        }
    }
}