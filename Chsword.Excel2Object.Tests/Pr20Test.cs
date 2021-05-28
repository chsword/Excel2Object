using System;
using System.Collections.Generic;
using System.IO;
using Chsword.Excel2Object.Styles;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.XWPF.UserModel;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class Pr20Test : BaseExcelTest
    {
        [TestMethod]
        public void UseExcelColumnAttr()
        {
            var list = new List<Pr20Model>
            {
                new Pr20Model
                {
                    Fullname = "AAA", Mobile = "123456798123"
                },
                new Pr20Model
                {
                    Fullname = "BBB", Mobile = "234"
                }
            };
            var bytes = ExcelHelper.ObjectToExcelBytes(list, ExcelType.Xlsx);
            var path = GetFilePath("test.xlsx");
            File.WriteAllBytes(path, bytes);
            Console.WriteLine(path);
        }

        [ExcelTitle("SheetX")]
        public class Pr20Model
        {
            [ExcelColumn("姓名", CellFontColor = ExcelStyleColor.Red)]
            public string Fullname { get; set; }

            [ExcelColumn("手机",
                HeaderFontFamily = "宋体",
                HeaderBold = true,
                HeaderFontHeight = 30,
                HeaderItalic = true, 
                HeaderFontColor = ExcelStyleColor.Blue, 
                HeaderUnderline = true,
                HeaderAlignment = HorizontalAlignment.Right,
                //cell
                CellAlignment = HorizontalAlignment.Justify
                )]
            public string Mobile { get; set; }

        }
    }
}