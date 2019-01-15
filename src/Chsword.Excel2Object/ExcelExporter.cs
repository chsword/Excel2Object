using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Internal;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Chsword.Excel2Object
{
	public class ExcelExporter
	{
		/// <summary>
		/// Export a excel file from a List of T generic list
		/// </summary>
		/// <typeparam name="TModel"></typeparam>
		/// <param name="data"></param>
		/// <param name="excelType"></param>
		/// <param name="sheetTitle"></param>
		/// <returns></returns>
		public byte[] ObjectToExcelBytes<TModel>(IEnumerable<TModel> data, ExcelType excelType, string sheetTitle = null)
		{
			var workbook = Workbook(excelType);
			ISheet sheet;

			if (string.IsNullOrWhiteSpace(sheetTitle))
			{
				var classAttr = ExcelUtil.GetClassExportAttribute<TModel>();
				sheet = classAttr == null ? workbook.CreateSheet() : workbook.CreateSheet(classAttr.Title);
			}
			else
			{
				sheet = workbook.CreateSheet(sheetTitle);
			}

			var attrDict = ExcelUtil.GetPropertiesAttributesDict<TModel>();
			var attrArray = attrDict.OrderBy(c => c.Value.Order).ToArray();
			for (var i = 0; i < attrArray.Length; i++)
				sheet.SetColumnWidth(i, 16 * 256);// todo 此处可统计字节数Min(50,Max(16,标题与内容最大长))
			var headerRow = sheet.CreateRow(0);
			for (var i = 0; i < attrArray.Length; i++)
			{
				var cell = headerRow.CreateCell(i);
				cell.SetCellType(CellType.String);
				cell.SetCellValue(attrArray[i].Value.Title);
			}
			var rowNumber = 1;
			foreach (var item in data.Where(c => c != null))
			{
				var row = sheet.CreateRow(rowNumber++);
				for (var i = 0; i < attrArray.Length; i++)
				{
					var cell = row.CreateCell(i);
					var prop = attrArray[i].Key;
					var val = (prop.GetValue(item, null) ?? "").ToString();
					
					SetCellValue(excelType, prop.PropertyType, cell, val);
				}
			}
			return ToBytes(workbook);
		}

		public byte[] ObjectToExcelBytes(DataTable dt, ExcelType excelType, string sheetTitle = null)
		{
			var workbook = Workbook(excelType);

			var sheet = string.IsNullOrWhiteSpace(sheetTitle) ? workbook.CreateSheet() : workbook.CreateSheet(sheetTitle);

			var attrArray = dt.Columns.Cast<DataColumn>().ToArray();
			for (var i = 0; i < attrArray.Length; i++)
				sheet.SetColumnWidth(i, 16 * 256);// todo 此处可统计字节数Min(50,Max(16,标题与内容最大长))
			var headerRow = sheet.CreateRow(0);
			for (var i = 0; i < attrArray.Length; i++)
			{
				var cell = headerRow.CreateCell(i);
				cell.SetCellType(CellType.String);
				cell.SetCellValue(attrArray[i].ColumnName);
			}
			var rowNumber = 1;
			var data = dt.Rows.Cast<DataRow>().ToArray();
			foreach (var item in data.Where(c => c != null))
			{
				var row = sheet.CreateRow(rowNumber++);
				for (var i = 0; i < attrArray.Length; i++)
				{
					var cell = row.CreateCell(i);
					var type = attrArray[i].DataType;
					var val = (item[attrArray[i].ColumnName] ?? "").ToString();
					SetCellValue(excelType, type, cell, val);
				}
			}
			return ToBytes(workbook);
		}

		private static void SetCellValue(ExcelType excelType, Type type, ICell cell, string val)
		{
			if (type == typeof(Uri))
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

			//cell.Hyperlink=new HSSFHyperlink
			cell.SetCellValue(val);
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

		public byte[] ObjectToExcelBytes<TModel>(IEnumerable<TModel> data)
		{
			return ObjectToExcelBytes(data, ExcelType.Xls);
		}
	}
}