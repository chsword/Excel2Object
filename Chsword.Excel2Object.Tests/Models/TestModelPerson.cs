using System;

namespace Chsword.Excel2Object.Tests.Models
{
    [ExcelTitle("Test Person")]
    public class TestModelPerson
    {
        [ExcelTitle("姓名")] public string Name { get; set; }
        [ExcelTitle("年龄")] public int? Age { get; set; }
        [ExcelTitle("出生日期")] public DateTime? Birthday { get; set; }
    }

    [ExcelTitle("Test Strict Person")]
    public class TestModelStrictPerson
    {
        [ExcelTitle("姓名")] public string Name { get; set; }
        [ExcelTitle("年龄")] public int Age { get; set; }
        [ExcelTitle("出生日期")] public DateTime Birthday { get; set; }
    }
}
