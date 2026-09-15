using Chsword.Excel2Object.Options;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     合并单元格：把某列中连续相同的值并成一格，以及合并调用方另行指定的区域。
/// </summary>
/// <remarks>
///     各段是在写入过程中逐格记下的，而非事后回看工作表：流式导出会把先前的行刷出内存，届时已无从读回。
/// </remarks>
internal sealed class MergedRegions
{
    private readonly ExcelColumn[] _columns;
    private readonly ExcelExporterOptions _options;

    /// <summary>各列当前那一段的起始行与取值，按列序号存放；未跟踪的列为空。</summary>
    private readonly Dictionary<int, Run> _open = new();

    private readonly Dictionary<int, List<CellRangeAddress>> _runs = new();

    public MergedRegions(ExcelColumn[] columns, ExcelExporterOptions options)
    {
        _columns = columns;
        _options = options;

        // 列名与公式列在此处即校验：写入尚未开始，报错也就不会留下写了一半的工作表
        foreach (var title in options.MergeRepeatedColumns.Distinct(StringComparer.Ordinal))
        {
            var column = IndexOf(columns, title);
            RejectFormula(columns[column]);
            _open[column] = new Run(ExcelConstants.DefaultDataStartRowIndex, null);
            _runs[column] = new List<CellRangeAddress>();
        }
    }

    /// <summary>是否有按值合并的列，没有则连每格的记录都不必做。</summary>
    public bool TracksAnyColumn => _open.Count > 0;

    public bool Tracks(int column)
    {
        return _open.ContainsKey(column);
    }

    /// <summary>记下刚写好的一格。行号须递增，即按写入顺序调用。</summary>
    public void Observe(int column, int rowIndex, ICell cell)
    {
        if (!_open.TryGetValue(column, out var run)) return;

        var value = CellValue.Of(cell);
        if (value != null && run.Value != null && value.Equals(run.Value)) return;

        CloseRun(column, run, rowIndex - 1);
        _open[column] = new Run(rowIndex, value);
    }

    public void Apply(ISheet sheet, int lastDataRowIndex)
    {
        foreach (var pair in _open) CloseRun(pair.Key, pair.Value, lastDataRowIndex);

        var declared = new List<CellRangeAddress>();
        foreach (var region in _options.MergedRegions)
        {
            var range = Resolve(region, _columns, lastDataRowIndex);
            var clash = Overlapping(range, _runs, declared);
            if (clash != null)
                throw new Excel2ObjectException(
                    $"区域 [{region}] 与已合并的 [{clash.FormatAsString()}] 重叠。");

            declared.Add(range);
        }

        foreach (var range in _runs.Values.SelectMany(list => list).Concat(declared))
            // 重叠已在此处校验过，故绕开 NPOI 自带的两两比对：它对每个新区域都要扫描已有全部区域，
            // 十万行的分组表会因此退化为平方级
            AddUnsafe(sheet, range);
    }

    /// <summary>一段就此收尾：只有跨越多行时才算一处合并区域。</summary>
    private void CloseRun(int column, Run run, int lastRow)
    {
        if (run.Value != null && lastRow > run.FirstRow)
            _runs[column].Add(new CellRangeAddress(run.FirstRow, lastRow, column, column));
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

    /// <summary>把按标题与数据行序号写下的区域，换算成工作表中的坐标。</summary>
    private static CellRangeAddress Resolve(MergedRegion region, ExcelColumn[] columns, int lastDataRowIndex)
    {
        if (region.Address != null) return Check(Parse(region.Address), region.Address);

        if (string.IsNullOrEmpty(region.Column))
            throw new Excel2ObjectException($"合并区域 [{region}] 没有指定列，请给出列标题或 A1:C1 这样的地址。");

        var first = IndexOf(columns, region.Column!);
        var last = region.LastColumn == null ? first : IndexOf(columns, region.LastColumn);
        if (last < first) (first, last) = (last, first);

        if (region.FirstRow == null && region.LastRow != null)
            throw new Excel2ObjectException(
                $"合并区域 [{region}] 只给了结束行：省略起始行即指表头行，如此会把表头一并并进去。" +
                "请一并给出 FirstRow。");

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

    /// <summary>某一列中正在延续的一段。</summary>
    private readonly struct Run
    {
        public Run(int firstRow, CellValue? value)
        {
            FirstRow = firstRow;
            Value = value;
        }

        public int FirstRow { get; }

        /// <summary>该段的取值；为空表示这一行不参与合并（空单元格、公式等）。</summary>
        public CellValue? Value { get; }
    }

    /// <summary>
    ///     一格的值，于写入时即记下。数值记其值而不记文本：.NET Framework 上 "R" 与 G17 两种格式都可能
    ///     把不同的数渲染成同一串字符，比较文本会把它们并成一格。
    /// </summary>
    private sealed class CellValue
    {
        private readonly bool _flag;
        private readonly double _number;
        private readonly string? _text;
        private readonly CellType _type;

        private CellValue(CellType type, double number, string? text, bool flag)
        {
            _type = type;
            _number = number;
            _text = text;
            _flag = flag;
        }

        /// <summary>该格参与合并时的取值；空单元格、公式等一律返回 null，即不参与。</summary>
        public static CellValue? Of(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    return new CellValue(CellType.Numeric, cell.NumericCellValue, null, false);
                case CellType.String:
                    var text = cell.StringCellValue;
                    return string.IsNullOrEmpty(text) ? null : new CellValue(CellType.String, 0, text, false);
                case CellType.Boolean:
                    return new CellValue(CellType.Boolean, 0, null, cell.BooleanCellValue);
                default:
                    return null;
            }
        }

        public bool Equals(CellValue other)
        {
            if (_type != other._type) return false;

            return _type switch
            {
                CellType.Numeric => _number.Equals(other._number),
                CellType.String => _text == other._text,
                CellType.Boolean => _flag == other._flag,
                _ => false
            };
        }
    }
}
