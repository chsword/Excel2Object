using System.Globalization;
using System.Xml;
using Chsword.Excel2Object.Options;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.UserModel;
using NPOI.XSSF.EventUserModel;
using NPOI.XSSF.Model;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     逐行读出 <c>.xlsx</c> 的一张工作表，不把整个工作簿建进内存。
/// </summary>
/// <remarks>
///     <para>
///         读的是工作表的那份 XML：一行解出来即交给调用方，随后便可回收。留在内存里的只有共享字符串表
///         与样式表——前者是文件中所有不重复的字符串，后者通常只有几十条，二者都远小于整份工作簿。
///     </para>
///     <para>
///         公式格读的是文件中存着的上一次计算结果。本库导出的文件里公式没有这个结果（要等 Excel 打开时
///         才算出），此时该格读作空白。
///     </para>
/// </remarks>
internal static class XlsxRowReader
{
    private const string RelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static IEnumerable<IImportRow> Rows(Stream input, ExcelImporterOptions options, ImportContext context)
    {
        OPCPackage package;
        try
        {
            package = OPCPackage.Open(input);
        }
        catch (Exception e)
        {
            throw new Excel2ObjectException("这不是一份能够打开的 .xlsx 文件。", e);
        }

        // 打开与定位在此处即完成：表名写错应当当场报错，而不是等到第一次取行
        try
        {
            var reader = new XSSFReader(package);
            var workbook = ReadWorkbook(reader);
            var sheet = Locate(workbook, options.SheetTitle);
            return Iterate(package, reader, sheet, workbook.Date1904, context);
        }
        catch
        {
            package.Close();
            throw;
        }
    }

    /// <summary>
    ///     自此按行取用。工作簿由这条序列持有：遍历完毕或中途放弃（<c>foreach</c> 均会释放枚举器）
    ///     时关闭；与 <see cref="File.ReadLines(string)" /> 一样，取到的序列须被遍历。
    /// </summary>
    private static IEnumerable<IImportRow> Iterate(OPCPackage package, XSSFReader reader, SheetInfo sheet,
        bool date1904, ImportContext context)
    {
        try
        {
            var styles = new NumberFormats(reader.StylesTable);
            var strings = new ReadOnlySharedStringsTable(package);
            using var sheetData = reader.GetSheet(sheet.RelationshipId);

            foreach (var row in ReadRows(sheetData, sheet.Title, date1904, styles, strings, context))
                yield return row;
        }
        finally
        {
            package.Close();
        }
    }

    /// <summary>工作簿本身那份 XML：各表的名字与其部件编号，以及所用的日期系统。</summary>
    private static WorkbookInfo ReadWorkbook(XSSFReader reader)
    {
        var sheets = new List<SheetInfo>();
        var date1904 = false;

        using var data = reader.WorkbookData;
        using var xml = XmlReader.Create(data, new XmlReaderSettings {IgnoreWhitespace = true, IgnoreComments = true});
        while (xml.Read())
        {
            if (xml.NodeType != XmlNodeType.Element) continue;

            switch (xml.Name)
            {
                case "workbookPr":
                    date1904 = IsTrue(xml.GetAttribute("date1904")) || IsTrue(xml.GetAttribute("date1904Compat"));
                    break;
                case "sheet":
                    var id = xml.GetAttribute("id", RelationshipNamespace) ?? xml.GetAttribute("r:id");
                    if (id != null) sheets.Add(new SheetInfo(xml.GetAttribute("name") ?? string.Empty, id));
                    break;
            }
        }

        return new WorkbookInfo(sheets, date1904);
    }

    private static SheetInfo Locate(WorkbookInfo workbook, string? title)
    {
        if (workbook.Sheets.Count == 0) throw new Excel2ObjectException("该工作簿中没有工作表。");

        if (string.IsNullOrEmpty(title)) return workbook.Sheets[0];

        foreach (var sheet in workbook.Sheets)
            if (sheet.Title == title)
                return sheet;

        throw new Excel2ObjectException($"The specified sheet:[{title}] does not exist");
    }

    private static IEnumerable<IImportRow> ReadRows(Stream sheetData, string sheetTitle, bool date1904,
        NumberFormats styles, ReadOnlySharedStringsTable strings, ImportContext context)
    {
        using var xml = XmlReader.Create(sheetData,
            new XmlReaderSettings {IgnoreWhitespace = true, IgnoreComments = true});

        var cells = new Dictionary<int, CellData>();
        var rowIndex = -1;
        var column = -1;

        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.Element && xml.Name == "row")
            {
                // 行号自 1 计起；缺失时按上一行顺延
                rowIndex = Number(xml.GetAttribute("r")) - 1;
                if (rowIndex < 0) rowIndex = 0;
                cells = new Dictionary<int, CellData>();
                column = -1;
                if (!xml.IsEmptyElement) continue;
            }

            if (xml.NodeType == XmlNodeType.Element && xml.Name == "c")
            {
                var reference = xml.GetAttribute("r");
                column = reference == null ? column + 1 : ColumnOf(reference);
                var type = xml.GetAttribute("t");
                var style = xml.GetAttribute("s");
                var cell = ReadCell(xml, type, style, date1904, styles, strings, sheetTitle, rowIndex, column,
                    context);
                if (cell.Exists || cell.Failed) cells[column] = cell;
                continue;
            }

            if (xml.NodeType == XmlNodeType.EndElement && xml.Name == "row")
                yield return new StreamedRow(sheetTitle, rowIndex, cells);
        }
    }

    /// <summary>
    ///     一格的取值。<c>t</c> 说明它是什么：<c>s</c> 指向共享字符串表，<c>inlineStr</c> 与 <c>str</c>
    ///     自带文本，<c>b</c> 是布尔，<c>e</c> 是错误值，其余为数值。
    /// </summary>
    private static CellData ReadCell(XmlReader xml, string? type, string? style, bool date1904,
        NumberFormats styles, ReadOnlySharedStringsTable strings, string sheetTitle, int rowIndex, int column,
        ImportContext context)
    {
        string? value = null;
        var inline = (string?) null;

        if (!xml.IsEmptyElement)
        {
            var depth = xml.Depth;
            // ReadElementContentAsString 会一并越过该元素，故读到内容之后不再 Read，否则将跳过下一个节点
            while (!(xml.NodeType == XmlNodeType.EndElement && xml.Depth == depth))
            {
                if (xml.NodeType == XmlNodeType.Element)
                {
                    if (xml.Name == "v")
                    {
                        value = xml.ReadElementContentAsString();
                        continue;
                    }

                    if (xml.Name == "is")
                    {
                        inline = InlineText(xml);
                        continue;
                    }
                }

                if (!xml.Read()) break;
            }
        }

        var text = inline ?? value;
        try
        {
            switch (type)
            {
                case "s":
                    // 共享字符串表：值是其中的序号
                    return text == null ? CellData.Blank() : CellData.OfText(strings.GetEntryAt(int.Parse(text, CultureInfo.InvariantCulture)));
                case "inlineStr":
                case "str":
                    return text == null ? CellData.Blank() : CellData.OfText(text);
                case "b":
                    return text == null ? CellData.Blank() : CellData.OfBoolean(text == "1" || text == "TRUE");
                case "e":
                    return CellData.OfError(text);
                case "d":
                    // ISO 8601 写法的日期，少见
                    return text == null
                        ? CellData.Blank()
                        : CellData.OfText(text);
                default:
                    if (text == null) return CellData.Blank();

                    var format = styles.Of(style);
                    return CellData.OfNumber(double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture),
                        format.Index, format.Text, date1904);
            }
        }
        catch (Exception e)
        {
            context.Report(sheetTitle, rowIndex, column, e);
            return CellData.Failure;
        }
    }

    /// <summary>行内字符串：<c>&lt;is&gt;</c> 之下可有若干段 <c>&lt;t&gt;</c>，依次相接。</summary>
    private static string InlineText(XmlReader xml)
    {
        var depth = xml.Depth;
        var text = string.Empty;
        while (!(xml.NodeType == XmlNodeType.EndElement && xml.Depth == depth))
        {
            if (xml.NodeType == XmlNodeType.Element && xml.Name == "t")
            {
                text += xml.ReadElementContentAsString();
                continue;
            }

            if (!xml.Read()) break;
        }

        return text;
    }

    private static bool IsTrue(string? value)
    {
        return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }

    private static int Number(string? value)
    {
        return value != null && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)
            ? n
            : 0;
    }

    /// <summary>由 <c>B12</c> 这样的地址取出列序号，自 0 计起。</summary>
    private static int ColumnOf(string reference)
    {
        var column = 0;
        foreach (var c in reference)
        {
            if (c is < 'A' or > 'Z')
            {
                if (c is >= 'a' and <= 'z')
                {
                    column = column * 26 + (c - 'a' + 1);
                    continue;
                }

                break;
            }

            column = column * 26 + (c - 'A' + 1);
        }

        return column - 1;
    }

    /// <summary>流式读出的一行：各格已解好，按列序号存放。</summary>
    private sealed class StreamedRow : IImportRow
    {
        private readonly Dictionary<int, CellData> _cells;

        public StreamedRow(string? sheetTitle, int rowIndex, Dictionary<int, CellData> cells)
        {
            SheetTitle = sheetTitle;
            RowIndex = rowIndex;
            _cells = cells;
        }

        public string? SheetTitle { get; }

        public int RowIndex { get; }

        public IEnumerable<KeyValuePair<int, CellData>> Cells => _cells.OrderBy(c => c.Key);

        public int CellCount => _cells.Count;

        public CellData Cell(int columnIndex)
        {
            return _cells.TryGetValue(columnIndex, out var cell) ? cell : CellData.None;
        }
    }

    /// <summary>
    ///     各样式所用的数字格式。判断一格是不是日期要看它的格式，而样式通常只有几十条，逐格去问
    ///     样式表并无必要。
    /// </summary>
    private sealed class NumberFormats
    {
        private readonly Dictionary<int, Format> _cache = new();
        private readonly StylesTable _styles;

        public NumberFormats(StylesTable styles)
        {
            _styles = styles;
        }

        public Format Of(string? styleIndex)
        {
            if (styleIndex == null) return default;

            var index = Number(styleIndex);
            if (_cache.TryGetValue(index, out var format)) return format;

            try
            {
                var style = _styles.GetStyleAt(index);
                format = style == null ? default : new Format(style.DataFormat, style.GetDataFormatString());
            }
            catch (Exception)
            {
                // 样式表里没有这一条：按没有格式处理，读到的仍是那个数
                format = default;
            }

            _cache[index] = format;
            return format;
        }
    }

    private readonly struct Format
    {
        public Format(short index, string? text)
        {
            Index = index;
            Text = text;
        }

        public short Index { get; }

        public string? Text { get; }
    }

    private readonly struct SheetInfo
    {
        public SheetInfo(string title, string relationshipId)
        {
            Title = title;
            RelationshipId = relationshipId;
        }

        public string Title { get; }

        public string RelationshipId { get; }
    }

    private readonly struct WorkbookInfo
    {
        public WorkbookInfo(List<SheetInfo> sheets, bool date1904)
        {
            Sheets = sheets;
            Date1904 = date1904;
        }

        public List<SheetInfo> Sheets { get; }

        public bool Date1904 { get; }
    }
}
