using System;

namespace Chsword.Excel2Object.Tests.Models
{
	[ExcelTitle("Test Sheetname")]
	public class ReportModel
    {
        [ExcelTitle("Document Title")]
        public string Title { get; set; }

        [ExcelTitle("User Name")]
        public string Name { get; set; }

        [ExcelTitle("Open")]
        public bool? Enabled { get; set; }
        [ExcelTitle("Address")]
        public Uri Uri { get; set; }
    }
}