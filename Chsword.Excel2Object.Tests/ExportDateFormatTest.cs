using System;
using System.Collections.Generic;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.UserModel;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExportDateFormatTest : BaseExcelTest
{
    [TestMethod]
    public void ExportDateTest()
    {
        var list = new List<TestModelDatePerson>
        {
            new()
            {
                Age = 18,
                Birthday = DateTime.Now,
                Birthday2 = DateTime.Now,
                Name = "test"
            },
            new()
            {
                Age = 18,
                Birthday = DateTime.Now,

                Name = "test2"
            }
        };
        ExcelHelper.ObjectToExcel(list, GetFilePath(DateTime.Now.Ticks + "test.xls"));
    }

    [TestMethod]
    public void MyTestMethod()
    {
        var list = HSSFDataFormat.GetBuiltinFormats();
        foreach (var item in list) Console.WriteLine(item);
    }
}