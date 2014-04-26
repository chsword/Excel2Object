namespace Chsword.Excel2Object.Tests.Models
{
    public class ReportModel
    {
        [Excel("标题")]
        public string Title { get; set; }
        [Excel("用户")]
        public string Name { get; set; }
    }
}