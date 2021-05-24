using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using Chsword.Excel2Object.Internal;
using Chsword.Excel2Object.Options;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Chsword.Excel2Object
{
    public class ExcelExporter
	{
        #region Public

        /// <summary>
        /// Export a excel file from a List of T generic list
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="data"></param>
        /// <param name="excelType"></param>
        /// <param name="sheetTitle"></param>
        /// <returns></returns>
        public byte[] ObjectToExcelBytes<TModel>(IEnumerable<TModel> data, ExcelType excelType,
            string sheetTitle = null)
        {
            return ObjectToExcelBytes(data, options =>
            {
                options.ExcelType = excelType;
                options.SheetTitle = sheetTitle;
            });
        }
        public byte[] ObjectToExcelBytes<TModel>(IEnumerable<TModel> data, Action<ExcelExporterOptions> optionsAction)
        {
            var options = new ExcelExporterOptions();
            optionsAction(options);
            ExcelModel excel;
            if (data is IEnumerable<Dictionary<string, object>> models)
            {
                excel = TypeConvert.ConvertDictionaryToExcelModel(models, options);
            }
            else
            {
                excel = TypeConvert.ConvertObjectToExcelModel(data, options);
            }

            return ObjectToExcelBytes(excel, options);
        }

        public byte[] ObjectToExcelBytes(DataTable dt, ExcelType excelType, string sheetTitle = null)
        {
            var options = new ExcelExporterOptions
            {
                ExcelType = excelType,
                SheetTitle = sheetTitle
            };
            var excel = TypeConvert.ConvertDataSetToExcelModel(dt, options);
            return ObjectToExcelBytes(excel, options);
        }

        #endregion

        #region Core

        internal byte[] ObjectToExcelBytes(ExcelModel excel, ExcelExporterOptions options)
        {
            ExcelType excelType = options.ExcelType;
            var workbook = Workbook(excelType);
            CheckExcelModel(excel);
            foreach (var excelSheet in excel.Sheets)
            {
                var sheet = string.IsNullOrWhiteSpace(excelSheet.Title)
                    ? workbook.CreateSheet()
                    : workbook.CreateSheet(excelSheet.Title);
                sheet.ForceFormulaRecalculation = true;
                var columns = excelSheet.Columns.OrderBy(c => c.Order).ToArray();
                for (var i = 0; i < columns.Length; i++)
                { 
                    sheet.SetColumnWidth(i, 16 * 256); 
                    // todo 此处可统计字节数Min(50,Max(16,标题与内容最大长))
                }
                var headerRow = sheet.CreateRow(0);
                for (var i = 0; i < columns.Length; i++)
                {
                    var cell = headerRow.CreateCell(i);
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(columns[i].Title);
                }
                var columnTitles = columns.Select(c=>c.Title).ToArray();
                var rowNumber = 1;
                var data = excelSheet.Rows;
                foreach (var item in data)
                {
                    var row = sheet.CreateRow(rowNumber++);
                    for (var i = 0; i < columns.Length; i++)
                    {
                        var column = columns[i];
                        var cell = row.CreateCell(i);
                        var val = item.ContainsKey(column.Title)?(item?[column.Title] ?? "").ToString():"";
                        SetCellValue(excelType, column, cell, val, columnTitles);
                     
                    }
                }
            }

            return ToBytes(workbook);
        }

        private void CheckExcelModel(ExcelModel excel)
        {
            //todo validate

        }

        #endregion

        #region Factory

        private void SetCellValue(ExcelType excelType, ExcelColumn column, ICell cell, string val,
            string[] columnTitles)
        {
            if (column.Type == typeof(Uri))
            {
                cell.Hyperlink = Switch<IHyperlink>(
                    excelType,
                    () => new HSSFHyperlink(HyperlinkType.Url)
                    {
                        Address = val
                    },
                    () => new XSSFHyperlink(HyperlinkType.Url)
                    {
                        Address = val
                    }
                );
            }
            else if (column.Type == typeof(Expression))
            {
                var convert = new ExpressionConvert(columnTitles, cell.RowIndex);
                cell.SetCellFormula(convert.Convert(column.Formula));
                if (column.ResultType != null)
                {
                    if (column.ResultType == typeof(DateTime))
                    {
                        cell.CellStyle = CreateStyle("datetime", cell);
                    }
                }
                return;
            }
            else if (column.Type == typeof(string))
            {
                cell.SetCellType(CellType.String);
                cell.CellStyle = CreateStyle("text", cell);
            }

            //cell.Hyperlink=new HSSFHyperlink

            cell.SetCellValue(val);
        }



        private ConcurrentDictionary<string, ICellStyle> CellStyleDict = new ConcurrentDictionary<string, ICellStyle>();
        ICellStyle CreateStyle(string key,ICell cell)
        {
            if (CellStyleDict.TryGetValue(key, out var val))
            {
                return val;
            }

            if (key == "text")
            {
                var style = cell.Sheet.Workbook.CreateCellStyle();
                style.DataFormat = HSSFDataFormat.GetBuiltinFormat("text");
                CellStyleDict.AddOrUpdate(key, style, (k, s) => style);
                return style;
            }
            if (key == "datatime")
            {
                var style = cell.Sheet.Workbook.CreateCellStyle();
                style.DataFormat = HSSFDataFormat.GetBuiltinFormat("m/d/yy");
                CellStyleDict.AddOrUpdate(key, style, (k, s) => style);
                return style;
            }
            return null;
        }

        private static T Switch<T>(ExcelType excelType, Func<T> funcXlsHssf, Func<T> funcXlsxXssf)
		{
			T obj;
			switch (excelType)
			{
				case ExcelType.Xls:
					obj = funcXlsHssf();
					break;
				case ExcelType.Xlsx:
					obj = funcXlsxXssf();
					break;
				default:
					throw new ArgumentOutOfRangeException(nameof(excelType));
			}
			return obj;
		}

		private static IWorkbook Workbook(ExcelType excelType)
		{
			IWorkbook workbook;
			switch (excelType)
			{
				case ExcelType.Xls:
					workbook = new HSSFWorkbook();
					break;
				case ExcelType.Xlsx:
					workbook = new XSSFWorkbook();
					break;
				default:
					throw new ArgumentOutOfRangeException(nameof(excelType));
			}
			return workbook;
		}

        #endregion

        #region Utils

        public byte[] ObjectToExcelBytes<TModel>(IEnumerable<TModel> data)
		{
			return ObjectToExcelBytes(data, ExcelType.Xls);
		}

        private static byte[] ToBytes(IWorkbook workbook)
        {
            using (var output = new MemoryStream())
            {
                workbook.Write(output);
                var bytes = output.ToArray();
                return bytes;
            }
        }

        #endregion
    }
}