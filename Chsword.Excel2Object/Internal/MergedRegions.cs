using System.Globalization;
using Chsword.Excel2Object.Options;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     合并单元格：把某列中连续相同的值并成一格，以及合并调用方另行指定的区域。
/// </summary>
internal static class MergedRegions
{
    public static void Apply(ISheet sheet, ExcelColumn[] columns, int lastDataRowIndex,
        ExcelExporterOptions options)
    {
        // 同一列内的若干段天然互不相交，故按列分别算出后整体写入，无需两两比对
        var byColumn = new Dictionary<int, List<CellRangeAddress>>();
        foreach (var title in options.MergeRepeatedColumns.Distinct(StringComparer.Ordinal))
        {
            var column = IndexOf(columns, title);
            RejectFormula(columns[column]);
            byColumn[column] = Runs(sheet, column, lastDataRowIndex);
        }

        var declared = new List<CellRangeAddress>();
        foreach (var region in options.MergedRegions)
        {
            var range = Resolve(region, columns, lastDataRowIndex);
            var clash = Overlapping(range, byColumn, declared);
            if (clash != null)
                throw new Excel2ObjectException(
                    $"区域 [{region}] 与已合并的 [{clash.FormatAsString()}] 重叠。");

            declared.Add(range);
        }

        foreach (var range in byColumn.Values.SelectMany(list => list).Concat(declared))
            // 重叠已在此处校验过，故绕开 NPOI 自带的两两比对：它对每个新区域都要扫描已有全部区域，
            // 十万行的分组表会因此退化为平方级
            AddUnsafe(sheet, range);
    }

    private static void AddUnsafe(ISheet sheet, CellRangeAddress range)
    {
        switch (sheet)
        {
            case XSSFSheet xssf:
                xssf.AddMergedRegionUnsafe(range);
                break;
            case HSSFSheet hssf:
                hssf.AddMergedRegionUnsafe(range);
                break;
            default:
                sheet.AddMergedRegion(range);
                break;
        }
    }

    /// <summary>该区域与已算出的区域是否相交；返回相交的那一个。</summary>
    private static CellRangeAddress? Overlapping(CellRangeAddress range,
        Dictionary<int, List<CellRangeAddress>> byColumn, List<CellRangeAddress> declared)
    {
        for (var column = range.FirstColumn; column <= range.LastColumn; column++)
        {
            if (!byColumn.TryGetValue(column, out var runs)) continue;

            // 各段按行递增，越过该区域即可停下
            foreach (var run in runs)
            {
                if (run.FirstRow > range.LastRow) break;
                if (run.LastRow >= range.FirstRow) return run;
            }
        }

        return declared.FirstOrDefault(existing => Intersects(existing, range));
    }

    private static void RejectFormula(ExcelColumn column)
    {
        if (column.Type != typeof(System.Linq.Expressions.Expression)) return;

        throw new Excel2ObjectException(
            $"公式列 [{column.Title}] 不能按相同值合并：单元格里写的是公式，其结果由 Excel 打开时才算出，" +
            "导出时无从比较。");
    }

    private static int IndexOf(ExcelColumn[] columns, string title)
    {
        for (var i = 0; i < columns.Length; i++)
            if (columns[i].Title == title)
                return i;

        throw new Excel2ObjectException($"要合并的列 [{title}] 不在该工作表中。");
    }

    /// <summary>该列中连续相同的值各占一段，仅相邻且相等的行参与。</summary>
    private static List<CellRangeAddress> Runs(ISheet sheet, int column, int lastDataRowIndex)
    {
        var runs = new List<CellRangeAddress>();
        var start = ExcelConstants.DefaultDataStartRowIndex;
        if (lastDataRowIndex <= start) return runs;

        var previous = Text(sheet, start, column);
        for (var row = start + 1; row <= lastDataRowIndex + 1; row++)
        {
            // 多走一行，好让最后一段也能收尾
            var current = row > lastDataRowIndex ? null : Text(sheet, row, column);
            if (current != null && current == previous) continue;

            if (row - start > 1) runs.Add(new CellRangeAddress(start, row - 1, column, column));
            previous = current;
            start = row;
        }

        return runs;
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
            default:
                return null;
        }
    }

    /// <summary>把按标题与数据行序号写下的区域，换算成工作表中的坐标。</summary>
    private static CellRangeAddress Resolve(MergedRegion region, ExcelColumn[] columns, int lastDataRowIndex)
    {
        if (region.Address != null) return Check(Parse(region.Address), region.Address);

        if (string.IsNullOrEmpty(region.Column))
            throw new Excel2ObjectException($"合并区域 [{region}] 没有指定列，请给出列标题或 A1:C1 这样的地址。");

        var first = IndexOf(columns, region.Column!);
        var last = region.LastColumn == null ? first : IndexOf(columns, region.LastColumn);
        if (last < first) (first, last) = (last, first);

        // 行以数据行的序号给出，自 1 计起；省略则指表头行
        var firstRow = Row(region, region.FirstRow, lastDataRowIndex);
        var lastRow = region.LastRow == null ? firstRow : Row(region, region.LastRow, lastDataRowIndex);
        if (lastRow < firstRow) (firstRow, lastRow) = (lastRow, firstRow);

        return Check(new CellRangeAddress(firstRow, lastRow, first, last), region.ToString());
    }

    /// <summary>
    ///     只占一个单元格的区域无从合并。NPOI 自带的 AddMergedRegion 会拒绝它，而此处为避开其平方级
    ///     的两两比对走的是 Unsafe 版本，故须自行把关，否则会写出一个 Excel 认为异常的区域。
    /// </summary>
    private static CellRangeAddress Check(CellRangeAddress range, string declared)
    {
        if (range.FirstRow == range.LastRow && range.FirstColumn == range.LastColumn)
            throw new Excel2ObjectException(
                $"区域 [{declared}] 只有一个单元格，无从合并：请给出结束列或结束行。");

        return range;
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

    private static bool Intersects(CellRangeAddress a, CellRangeAddress b)
    {
        return a.FirstRow <= b.LastRow && b.FirstRow <= a.LastRow &&
               a.FirstColumn <= b.LastColumn && b.FirstColumn <= a.LastColumn;
    }
}
