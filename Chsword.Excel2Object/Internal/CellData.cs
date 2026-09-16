using System.Globalization;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     一格读到的东西，与它从哪里来无关：可以是 NPOI 的单元格，也可以是流式读取解出的一段 XML。
/// </summary>
/// <remarks>
///     导入的类型转换原本直接写在 <see cref="ICell" /> 之上，而流式读取拿到的是原始的 XML，并没有
///     单元格对象。把两者都归到这个结构上，转换那一套便只有一份。
/// </remarks>
internal readonly struct CellData
{
    /// <summary>该格不存在——空单元格并不占位。</summary>
    public static readonly CellData None = default;

    /// <summary>
    ///     这一格读取失败，且已上报。与「不存在」不同：不存在读作空白，读取失败则应取该类型的默认值，
    ///     不再按空字符串继续转换，否则真正的原因会被转换时抛出的异常掩盖。
    /// </summary>
    public static readonly CellData Failure = new(true);

    /// <summary>
    ///     文本以可空字段存放而非自动属性：结构的默认值（不存在的那一格）本就没有文本，让它读作空串，
    ///     各处便无须为此判空。
    /// </summary>
    private readonly string? _text;

    private CellData(bool failed)
    {
        Failed = failed;
        Exists = false;
        Kind = CellType.Blank;
        _text = null;
        Number = 0;
        Flag = false;
        FormatIndex = 0;
        Format = null;
        Date1904 = false;
    }

    private CellData(CellType kind, string? text, double number, bool flag, short formatIndex, string? format,
        bool date1904)
    {
        Failed = false;
        Exists = true;
        Kind = kind;
        _text = text;
        Number = number;
        Flag = flag;
        FormatIndex = formatIndex;
        Format = format;
        Date1904 = date1904;
    }

    public bool Exists { get; }

    public bool Failed { get; }

    public CellType Kind { get; }

    /// <summary>文本格即其文本；数值、布尔与错误格则是其渲染出的样子。</summary>
    public string Text => _text ?? string.Empty;

    public double Number { get; }

    public bool Flag { get; }

    /// <summary>数字格式的序号与格式串，用于判断这一格是不是日期。</summary>
    public short FormatIndex { get; }

    public string? Format { get; }

    /// <summary>该工作簿是否采用 1904 日期系统；序列号换算成日期时须据此而定。</summary>
    public bool Date1904 { get; }

    public static CellData Blank(short formatIndex = 0, string? format = null)
    {
        return new CellData(CellType.Blank, null, 0, false, formatIndex, format, false);
    }

    public static CellData OfNumber(double value, short formatIndex, string? format, bool date1904)
    {
        return new CellData(CellType.Numeric, null, value, false, formatIndex, format, date1904);
    }

    public static CellData OfText(string? text)
    {
        return new CellData(CellType.String, text, 0, false, 0, null, false);
    }

    public static CellData OfBoolean(bool value)
    {
        // 与 NPOI 的单元格渲染一致
        return new CellData(CellType.Boolean, value ? "TRUE" : "FALSE", 0, value, 0, null, false);
    }

    public static CellData OfError(string? text)
    {
        return new CellData(CellType.Error, text, 0, false, 0, null, false);
    }

    /// <summary>数值格的文本表示；必须能原样解析回同一个 double，数值列正是经由这段文本转换的。</summary>
    public string NumberText()
    {
        var text = Number.ToString("R", CultureInfo.InvariantCulture);
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var back) &&
            back.Equals(Number))
            return text;

        // .NET Framework 的 "R" 对少数值无法往返（已知缺陷），此时退回 17 位有效数字
        return Number.ToString("G17", CultureInfo.InvariantCulture);
    }

    /// <summary>该格作为日期的取值；序列号超出 Excel 日历时为 null。</summary>
    public DateTime? Date()
    {
        if (Kind != CellType.Numeric || !DateUtil.IsValidExcelDate(Number)) return null;

        return DateUtil.GetJavaDate(Number, Date1904);
    }
}
