using System.Globalization;
using Chsword.Excel2Object.Options;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     合并单元格：把某列中连续相同的值并成一格，以及合并调用方指定的区域。
/// </summary>
internal static class MergedRegions
{
    public static void Apply(ISheet sheet, ExcelColumn[] columns, int lastDataRowIndex,
        ExcelExporterOptions options)
    {
        foreach (var title in options.MergeRepeatedColumns)
            MergeRepeated(sheet, IndexOf(columns, title), lastDataRowIndex);

        foreach (var region in options.MergedRegions)
            Add(sheet, Resolve(region, columns, lastDataRowIndex), region.ToString());
    }

    /// <summary>把按标题与数据行序号写下的区域，换算成工作表中的坐标。</summary>
    private static CellRangeAddress Resolve(MergedRegion region, ExcelColumn[] columns, int lastDataRowIndex)
    {
        if (region.Address != null) return Parse(region.Address);

        if (string.IsNullOrEmpty(region.Column))
            throw new Excel2ObjectException($"合并区域 [{region}] 没有指定列，请给出列标题或 A1:C1 这样的地址。");

        var first = IndexOf(columns, region.Column!);
        var last = region.LastColumn == null ? first : IndexOf(columns, region.LastColumn);
        if (last < first) (first, last) = (last, first);

        // 行以数据行的序号给出，自 1 计起；省略则指表头行
        var firstRow = Row(region, region.FirstRow, lastDataRowIndex);
        var lastRow = region.LastRow == null ? firstRow : Row(region, region.LastRow, lastDataRowIndex);
        if (lastRow < firstRow) (firstRow, lastRow) = (lastRow, firstRow);

        return new CellRangeAddress(firstRow, lastRow, first, last);
    }

    private static int Row(MergedRegion region, int? dataRow, int lastDataRowIndex)
    {
        if (dataRow == null) return ExcelConstants.DefaultHeaderRowIndex;
        if (dataRow < 1)
            throw new Excel2ObjectException($"合并区域 [{region}] 的行序号自 1 计起，1 即第一行数据。");

        var rowIndex = ExcelConstants.DefaultDataStartRowIndex + dataRow.Value - 1;
        if (rowIndex > lastDataRowIndex)
            throw new Excel2ObjectException(
                $"合并区域 [{region}] 超出了数据的范围，本次导出共 " +
                $"{lastDataRowIndex - ExcelConstants.DefaultDataStartRowIndex + 1} 行数据。");

        return rowIndex;
    }

    private static int IndexOf(ExcelColumn[] columns, string title)
    {
        for (var i = 0; i < columns.Length; i++)
            if (columns[i].Title == title)
                return i;

        throw new Excel2ObjectException($"要合并的列 [{title}] 不在该工作表中。");
    }

    /// <summary>把该列中连续相同的值并成一格，仅相邻且相等的行参与。</summary>
    private static void MergeRepeated(ISheet sheet, int column, int lastDataRowIndex)
    {
        var start = ExcelConstants.DefaultDataStartRowIndex;
        if (lastDataRowIndex <= start) return;

        var previous = Text(sheet, start, column);
        for (var row = start + 1; row <= lastDataRowIndex + 1; row++)
        {
            // 多走一行，好让最后一段也能收尾
            var current = row > lastDataRowIndex ? null : Text(sheet, row, column);
            if (current != null && current == previous) continue;

            if (row - start > 1) Add(sheet, new CellRangeAddress(start, row - 1, column, column), null);
            previous = current;
            start = row;
        }
    }

    /// <summary>单元格写进去的内容，用于判断相邻两行是否相同。空单元格不参与合并。</summary>
    private static string? Text(ISheet sheet, int rowIndex, int column)
    {
        var cell = sheet.GetRow(rowIndex)?.GetCell(column);
        if (cell == null) return null;

        switch (cell.CellType)
        {
            case CellType.Numeric:
                return cell.NumericCellValue.ToString("R", CultureInfo.InvariantCulture);
            case CellType.String:
                var text = cell.StringCellValue;
                return string.IsNullOrEmpty(text) ? null : text;
            case CellType.Boolean:
                return cell.BooleanCellValue ? "TRUE" : "FALSE";
            case CellType.Formula:
                return "=" + cell.CellFormula;
            default:
                return null;
        }
    }

    private static CellRangeAddress Parse(string range)
    {
        try
        {
            return CellRangeAddress.ValueOf(range);
        }
        catch (Exception e)
        {
            throw new Excel2ObjectException($"[{range}] 不是合法的区域，应写作 A1:C1 这样的形式。", e);
        }
    }

    private static void Add(ISheet sheet, CellRangeAddress range, string? declared)
    {
        foreach (var existing in sheet.MergedRegions)
            if (Intersects(existing, range))
                throw new Excel2ObjectException(
                    $"区域 [{declared ?? range.FormatAsString()}] 与已合并的 [{existing.FormatAsString()}] 重叠。");

        sheet.AddMergedRegion(range);
    }

    private static bool Intersects(CellRangeAddress a, CellRangeAddress b)
    {
        return a.FirstRow <= b.LastRow && b.FirstRow <= a.LastRow &&
               a.FirstColumn <= b.LastColumn && b.FirstColumn <= a.LastColumn;
    }
}
