using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     流式写入开始前的一道检查。
/// </summary>
/// <remarks>
///     NPOI 的 SXSSF 每把一行刷出内存，都要先量一次默认字符宽度（为自动列宽备着的），而这一步用的是
///     SkiaSharp。NPOI 把 SkiaSharp 列为不随包传递的依赖，应用若未自行引用，缺失要到第一次刷行时才
///     暴露，届时已写下若干行，抛出的又是 <c>FileNotFoundException</c> 之类与 Excel 无关的异常。
///     故在此处先量一次：代价是一次字体度量，换来的是在任何内容写出之前给出一句说得清的错。
/// </remarks>
internal static class StreamingSupport
{
    private const string SkiaSharp = "SkiaSharp";

    public static void Ensure(IWorkbook workbook)
    {
        try
        {
            SheetUtil.GetDefaultCharWidth(workbook);
        }
        catch (Exception e) when (NeedsSkiaSharp(e))
        {
            throw new Excel2ObjectException(
                "流式导出需要 SkiaSharp：NPOI 把行刷出内存时要测量字符宽度，而它把 SkiaSharp 列为不随包" +
                "传递的依赖。请在应用中引用 SkiaSharp（Linux 上还需 SkiaSharp.NativeAssets.Linux." +
                "NoDependencies），或改用 ObjectToExcelBytes 在内存中完成导出。", e);
        }
    }

    private static bool NeedsSkiaSharp(Exception e)
    {
        for (var current = e; current != null; current = current.InnerException)
            if (current is TypeLoadException or FileNotFoundException or BadImageFormatException &&
                current.Message.IndexOf(SkiaSharp, StringComparison.OrdinalIgnoreCase) >= 0)
                return true;

        return false;
    }
}
