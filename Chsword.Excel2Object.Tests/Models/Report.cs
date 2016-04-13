namespace Chsword.Excel2Object.Tests.Models
{
    public class ReportModel
    {
        [ExcelTitle("标题")]
        public string Title { get; set; }

        [ExcelTitle("用户")]
        public string Name { get; set; }

        [ExcelTitle("启用")]
        public bool? Enabled { get; set; }
    }
}