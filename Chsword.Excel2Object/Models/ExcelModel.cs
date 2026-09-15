namespace Chsword.Excel2Object;

internal class ExcelModel
{
    public List<SheetModel>? Sheets { get; set; }

    /// <summary>
    ///     取数据的枚举器，仅字典入口用得到：列由首行的键决定，故建模型时就得先取一行，枚举器因而
    ///     在写入开始之前即已打开。写入若未能开始（例如选项不合法），它也须被释放，故由模型带着，
    ///     导出结束时一并释放。
    /// </summary>
    public IDisposable? RowSource { get; set; }
}