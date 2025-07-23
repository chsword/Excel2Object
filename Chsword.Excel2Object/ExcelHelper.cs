using System.Data;
using Chsword.Excel2Object.Options;

namespace Chsword.Excel2Object;

public static class ExcelHelper
{
    public static byte[]? AppendObjectToExcelBytes<TModel>(byte[] sourceExcelBytes, IEnumerable<TModel> data,
        string sheetTitle)
    {
        var excelExporter = new ExcelExporter();
        return excelExporter.AppendObjectToExcelBytes(sourceExcelBytes, data, sheetTitle);
    }

    /// <summary>
    ///     convert a excel file(bytes) to IEnumerable of TModel
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <param name="bytes">the excel file bytes</param>
    /// <param name="sheetTitle">specify sheet  name which wants to import</param>
    /// <returns></returns>
    public static IEnumerable<TModel> ExcelToObject<TModel>(byte[] bytes, string? sheetTitle = null)
        where TModel : class, new()
    {
        var importer = new ExcelImporter();
        return importer.ExcelToObject<TModel>(bytes, sheetTitle);
    }

    /// <summary>
    ///     import file excel file to a IEnumerable of TModel
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <param name="path">excel full path</param>
    /// <param name="sheetTitle">specify sheet  name which wants to import</param>
    /// <returns></returns>
    public static IEnumerable<TModel>? ExcelToObject<TModel>(string path, string? sheetTitle = null)
        where TModel : class, new()
    {
        var importer = new ExcelImporter();
        return importer.ExcelToObject<TModel>(path, sheetTitle);
    }

    /// <summary>
    ///     Asynchronously convert excel file bytes to IEnumerable of TModel
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <param name="bytes">the excel file bytes</param>
    /// <param name="sheetTitle">specify sheet name which wants to import</param>
    /// <param name="cancellationToken">cancellation token for async operation</param>
    /// <returns></returns>
    public static async Task<IEnumerable<TModel>> ExcelToObjectAsync<TModel>(byte[] bytes, 
        string? sheetTitle = null, CancellationToken cancellationToken = default)
        where TModel : class, new()
    {
        var importer = new ExcelImporter();
        return await importer.ExcelToObjectAsync<TModel>(bytes, sheetTitle, cancellationToken);
    }

    /// <summary>
    ///     Asynchronously import excel file to a IEnumerable of TModel
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <param name="path">excel full path</param>
    /// <param name="sheetTitle">specify sheet name which wants to import</param>
    /// <param name="cancellationToken">cancellation token for async operation</param>
    /// <returns></returns>
    public static async Task<IEnumerable<TModel>?> ExcelToObjectAsync<TModel>(string path, 
        string? sheetTitle = null, CancellationToken cancellationToken = default)
        where TModel : class, new()
    {
        var importer = new ExcelImporter();
        return await importer.ExcelToObjectAsync<TModel>(path, sheetTitle, cancellationToken);
    }

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    /// <summary>
    ///     Asynchronously stream excel file bytes to IAsyncEnumerable of TModel
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <param name="bytes">the excel file bytes</param>
    /// <param name="sheetTitle">specify sheet name which wants to import</param>
    /// <param name="cancellationToken">cancellation token for async operation</param>
    /// <returns></returns>
    public static IAsyncEnumerable<TModel> ExcelToObjectStreamAsync<TModel>(byte[] bytes, 
        string? sheetTitle = null, CancellationToken cancellationToken = default)
        where TModel : class, new()
    {
        var importer = new ExcelImporter();
        return importer.ExcelToObjectStreamAsync<TModel>(bytes, sheetTitle, cancellationToken);
    }

    /// <summary>
    ///     Asynchronously stream excel file to IAsyncEnumerable of TModel
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <param name="path">excel full path</param>
    /// <param name="sheetTitle">specify sheet name which wants to import</param>
    /// <param name="cancellationToken">cancellation token for async operation</param>
    /// <returns></returns>
    public static IAsyncEnumerable<TModel> ExcelToObjectStreamAsync<TModel>(string path, 
        string? sheetTitle = null, CancellationToken cancellationToken = default)
        where TModel : class, new()
    {
        var importer = new ExcelImporter();
        return importer.ExcelToObjectStreamAsync<TModel>(path, sheetTitle, cancellationToken);
    }
#endif

    /// <summary>
    ///     Export object to excel file
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <param name="data">a IEnumerable of TModel</param>
    /// <param name="path">excel full path</param>
    public static void ObjectToExcel<TModel>(IEnumerable<TModel> data, string path) where TModel : class, new()
    {
        var excelExporter = new ExcelExporter();
        var bytes = excelExporter.ObjectToExcelBytes(data);
        WriteExcelBytesToFile(bytes, path);
    }

    /// <summary>
    ///     Asynchronously export object to excel file
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <param name="data">a IEnumerable of TModel</param>
    /// <param name="path">excel full path</param>
    /// <param name="cancellationToken">cancellation token for async operation</param>
    public static async Task ObjectToExcelAsync<TModel>(IEnumerable<TModel> data, string path, 
        CancellationToken cancellationToken = default) where TModel : class, new()
    {
        var excelExporter = new ExcelExporter();
        var bytes = excelExporter.ObjectToExcelBytes(data);
        await WriteExcelBytesToFileAsync(bytes, path, cancellationToken);
    }

    /// <summary>
    ///     Export object to excel file
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <param name="data">a IEnumerable of TModel</param>
    /// <param name="path">excel full path</param>
    /// <param name="excelType"></param>
    public static void ObjectToExcel<TModel>(IEnumerable<TModel> data, string path, ExcelType excelType)
        where TModel : class, new()
    {
        var excelExporter = new ExcelExporter();
        var bytes = excelExporter.ObjectToExcelBytes(data, excelType);
        WriteExcelBytesToFile(bytes, path);
    }

    /// <summary>
    ///     Asynchronously export object to excel file
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <param name="data">a IEnumerable of TModel</param>
    /// <param name="path">excel full path</param>
    /// <param name="excelType">excel type (xls or xlsx)</param>
    /// <param name="cancellationToken">cancellation token for async operation</param>
    public static async Task ObjectToExcelAsync<TModel>(IEnumerable<TModel> data, string path, 
        ExcelType excelType, CancellationToken cancellationToken = default)
        where TModel : class, new()
    {
        var excelExporter = new ExcelExporter();
        var bytes = excelExporter.ObjectToExcelBytes(data, excelType);
        await WriteExcelBytesToFileAsync(bytes, path, cancellationToken);
    }

    /// <summary>
    ///     Export object to excel bytes
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <param name="data"></param>
    /// <param name="excelType"></param>
    /// <param name="sheetTitle"></param>
    public static byte[]? ObjectToExcelBytes<TModel>(IEnumerable<TModel> data,
        ExcelType excelType = ExcelType.Xls,
        string? sheetTitle = null)
        where TModel : class, new()
    {
        var excelExporter = new ExcelExporter();
        return excelExporter.ObjectToExcelBytes(data, excelType, sheetTitle);
    }

    public static byte[]? ObjectToExcelBytes(DataTable dt, ExcelType excelType,
        string? sheetTitle = null)
    {
        var excelExporter = new ExcelExporter();
        return excelExporter.ObjectToExcelBytes(dt, excelType, sheetTitle);
    }

    public static byte[]? ObjectToExcelBytes<TModel>(IEnumerable<TModel> data,
        Action<ExcelExporterOptions> optionsAction)
    {
        var excelExporter = new ExcelExporter();
        return excelExporter.ObjectToExcelBytes(data, optionsAction);
    }

    private static void WriteExcelBytesToFile(byte[]? bytes, string path)
    {
        if (bytes != null) 
            File.WriteAllBytes(path, bytes);
    }

    private static async Task WriteExcelBytesToFileAsync(byte[]? bytes, string path, 
        CancellationToken cancellationToken = default)
    {
        if (bytes != null)
        {
#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
            await File.WriteAllBytesAsync(path, bytes, cancellationToken);
#else
            await Task.Run(() => File.WriteAllBytes(path, bytes), cancellationToken);
#endif
        }
    }
}