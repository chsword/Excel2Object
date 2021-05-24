namespace Chsword.Excel2Object.Options
{
    public class ExcelExporterOptions
    {
        /// <summary>
        /// Sheet Title default:null
        /// </summary>
        public string SheetTitle { get; set; }
        /// <summary>
        /// Excel file type default:xlsx
        /// </summary>
        public ExcelType ExcelType { get; set; } = ExcelType.Xlsx;

        public FormulaColumnsCollection FormulaColumns { get; set; }=new FormulaColumnsCollection();
    }
}