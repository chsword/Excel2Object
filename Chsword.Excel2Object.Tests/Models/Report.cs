using System;

namespace Chsword.Excel2Object.Tests.Models
{
    [ExcelTitle("Test Sheetname")]
    public class ReportModel
    {
        [ExcelTitle("Open")] public bool? Enabled { get; set; }

        [ExcelTitle("User Name")] public string Name { get; set; }

        [ExcelTitle("Document Title")] public string Title { get; set; }

        [ExcelTitle("Type")] public MyEnum Type { get; set; }

        [ExcelTitle("Address")] public Uri Uri { get; set; }
    }

    public enum MyEnum
    {
        Unkonw = 0,
        一 = 1,
        二 = 2,
        三 = 3
    }
}