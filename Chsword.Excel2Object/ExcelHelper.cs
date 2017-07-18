using System.Collections.Generic;
using System.IO;

namespace Chsword.Excel2Object
{
    public class ExcelHelper
    {
        /// <summary>
        ///     import file excel file to a IEnumerable of TModel
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="path">excel full path</param>
        /// <returns></returns>
        public static IEnumerable<TModel> ExcelToObject<TModel>(string path) where TModel : class, new()
        {
            var importer = new ExcelImporter();
            return importer.ExcelToObject<TModel>(path);
        }

        /// <summary>
        ///     Export object to excel file
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="data">a IEnumerable of TModel</param>
        /// <param name="path">excel full path</param>
        public static void ObjectToExcel<TModel>(IEnumerable<TModel> data, string path) where TModel : class, new()
        {
            var importer = new ExcelExporter();
            var bytes = importer.ObjectToExcelBytes(data);
            File.WriteAllBytes(path, bytes);
        }

        /// <summary>
        ///     Export object to excel bytes
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="data"></param>
        public static byte[] ObjectToExcelBytes<TModel>(IEnumerable<TModel> data) where TModel : class, new()
        {
            var importer = new ExcelExporter();
            return importer.ObjectToExcelBytes(data);
        }

        /// <summary>
        ///     Export object to excel bytes
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="data"></param>
        /// <param name="excelType"></param>
        public static byte[] ObjectToExcelBytes<TModel>(IEnumerable<TModel> data, ExcelType excelType)
            where TModel : class, new()
        {
            var excelExporter = new ExcelExporter();
            return excelExporter.ObjectToExcelBytes(data, excelType);
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
            File.WriteAllBytes(path, bytes);
        }
    }
}