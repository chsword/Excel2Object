using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Chsword.Excel2Object
{
    public class ExcelHelper
    {
        public static byte[] AppendObjectToExcelBytes<TModel>(byte[] sourceExcelBytes, IEnumerable<TModel> data,
            string sheetTitle)
        {
            var excelExporter = new ExcelExporter();
            
            return excelExporter.AppendObjectToExcelBytes(sourceExcelBytes, data, sheetTitle);
        }

        /// <summary>
        /// convert a excel file(bytes) to IEnumerable of TModel
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="bytes">the excel file bytes</param>
        /// <param name="sheetTitle">specify sheet  name which wants to import</param>
        /// <returns></returns>
        public static IEnumerable<TModel> ExcelToObject<TModel>(byte[] bytes, string sheetTitle = null)
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
        public static IEnumerable<TModel> ExcelToObject<TModel>(string path, string sheetTitle = null)
            where TModel : class, new()
        {
            var importer = new ExcelImporter();
            return importer.ExcelToObject<TModel>(path, sheetTitle);
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

        /// <summary>
        ///     Export object to excel bytes
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="data"></param>
        /// <param name="excelType"></param>
        /// <param name="sheetTitle"></param>
        public static byte[] ObjectToExcelBytes<TModel>(IEnumerable<TModel> data, ExcelType excelType = ExcelType.Xls,
            string sheetTitle = null)
            where TModel : class, new()
        {
            var excelExporter = new ExcelExporter();
            return excelExporter.ObjectToExcelBytes(data, excelType, sheetTitle);
        }

        public static byte[] ObjectToExcelBytes(DataTable dt, ExcelType excelType,
            string sheetTitle = null)

        {
            var excelExporter = new ExcelExporter();
            return excelExporter.ObjectToExcelBytes(dt, excelType, sheetTitle);
        }
    }
}