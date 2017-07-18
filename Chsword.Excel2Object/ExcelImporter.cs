using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using Chsword.Excel2Object.Internal;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object
{
    public class ExcelImporter
    {
        public IEnumerable<TModel> ExcelToObject<TModel>(string path) where TModel : class, new()
        {
            var result = GetDataRows(path);
            return ExcelToObject<TModel>(result);
        }

        public IEnumerable<TModel> ExcelToObject<TModel>(byte[] bytes) where TModel : class, new()
        {
            var result = GetDataRows(bytes);
            return ExcelToObject<TModel>(result);
        }

        private static readonly Dictionary<Type, Func<IRow, int, object>> SpecialConvertDict =
            new Dictionary<Type, Func<IRow, int, object>>
            {
                [typeof(DateTime)] = GetCellDateTime,
                [typeof(bool)] = GetCellBoolean,
                [typeof(Uri)] = GetCellUri
            };

        private static object GetCellUri(IRow row, int key)
        {
            var cellValue = GetCellValue(row, key);
            if (string.IsNullOrEmpty(cellValue)) return null;
            return new Uri(cellValue);
        }

        private static IEnumerable<TModel> ExcelToObject<TModel>(IEnumerator result) where TModel : class, new()
        {
            var dict = ExcelUtil.GetExportAttrDict<TModel>();
            var dictColumns = new Dictionary<int, KeyValuePair<PropertyInfo, ExcelTitleAttribute>>();

            var rows = result;

            var titleRow = (IRow) rows.Current;
            foreach (var cell in titleRow.Cells)
            {
                var prop = dict.FirstOrDefault(c => cell.StringCellValue == c.Value.Title);
                if (prop.Key != null && !dictColumns.ContainsKey(cell.ColumnIndex))
                    dictColumns.Add(cell.ColumnIndex, prop);
            }
            while (rows.MoveNext())
            {
                var row = (IRow) rows.Current;
                var firstCell = row.GetCell(0);
                if (firstCell == null || firstCell.CellType == CellType.Blank ||
                    string.IsNullOrWhiteSpace(firstCell.ToString()))
                    continue;

                var model = new TModel();

                foreach (var pair in dictColumns)
                {
                    var propType = pair.Value.Key.PropertyType;
                    var type = TypeUtil.GetUnNullableType(propType);
                    if (SpecialConvertDict.ContainsKey(type))
                    {
                        var specialValue = SpecialConvertDict[type](row, pair.Key);
                        pair.Value.Key.SetValue(model, specialValue, null);
                    }
                    else
                    {
                        var val = Convert.ChangeType(GetCellValue(row, pair.Key), propType);
                        pair.Value.Key.SetValue(model, val, null);
                    }
                }
                yield return model;
            }
        }

        private static object GetCellBoolean(IRow row, int key)
        {
            var cellValue = GetCellValue(row, key);
            if (string.IsNullOrEmpty(cellValue)) return null;
            bool value;
            if (bool.TryParse(cellValue, out value)) return value;
            switch (cellValue.ToLower())
            {
                case "1":
                case "是":
                case "yes":
                case "true":
                    return true;
                case "0":
                case "否":
                case "no":
                case "false":
                    return false;
                default:
                    return Convert.ToBoolean(cellValue);
            }
        }

        private static string GetCellValue(IRow row, int index)
        {
            var result = string.Empty;
            try
            {
                switch (row.GetCell(index).CellType)
                {
                    case CellType.Numeric:
                        result = row.GetCell(index).NumericCellValue.ToString(CultureInfo.InvariantCulture);
                        break;
                    case CellType.String:
                        result = row.GetCell(index).StringCellValue;
                        break;
                    case CellType.Blank:
                        result = string.Empty;
                        break;
                    //case CellType.Formula:
                    //    result = row.GetCell(index).CellFormula;
                    //    break;
                    //case CellType.Boolean:
                    //    result = row.GetCell(index).NumericCellValue.ToString();
                    //    break;
                    //case CellType.Error:
                    //    result = row.GetCell(index).NumericCellValue.ToString();
                    //    break;
                    //case CellType.Unknown:
                    //    result = row.GetCell(index).NumericCellValue.ToString();
                    //    break;
                    default:
                        result = row.GetCell(index).ToString();
                        break;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return (result ?? "").Trim();
        }

        private static IEnumerator GetDataRows(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return null;
            IWorkbook workbook;

            try
            {
                using (var file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(file);
                }
            }
            catch
            {
                return null;
            }
            var sheet = workbook.GetSheetAt(0);
            var rows = sheet.GetRowEnumerator();
            rows.MoveNext();
            return rows;
        }

        private static IEnumerator GetDataRows(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                return null;
            IWorkbook workbook;
            try
            {
                using (var memoryStream = new MemoryStream(bytes))
                {
                    workbook = WorkbookFactory.Create(memoryStream);
                }
            }
            catch
            {
                return null;
            }
            var sheet = workbook.GetSheetAt(0);
            var rows = sheet.GetRowEnumerator();
            rows.MoveNext();
            return rows;
        }

        private static object GetCellDateTime(IRow row, int index)
        {
            DateTime? result = null;
            try
            {
                switch (row.GetCell(index).CellType)
                {
                    case CellType.Numeric:
                        try
                        {
                            result = row.GetCell(index).DateCellValue;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }
                        break;
                    case CellType.String:
                        var str = row.GetCell(index).StringCellValue;
                        result = GetDateTimeFromString(str);
                        break;
                    case CellType.Blank:
                        break;
                    case CellType.Unknown:
                        break;
                    case CellType.Formula:
                        break;
                    case CellType.Boolean:
                        break;
                    case CellType.Error:
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return result;
        }

        private static DateTime? GetDateTimeFromString(string str)
        {
            DateTime dt;
            if (str.EndsWith("年"))
            {
                if (DateTime.TryParse((str + "-01-01").Replace("年", ""), out dt))
                    return dt;
            }
            else if (str.EndsWith("月"))
            {
                if (DateTime.TryParse((str + "-01").Replace("年", "").Replace("月", ""), out dt))
                    return dt;
            }
            else if (!str.Contains("年") && !str.Contains("月") && !str.Contains("日"))
            {
                if (DateTime.TryParse(str, out dt))
                    return dt;
                if (DateTime.TryParse((str + "-01-01").Replace("年", "").Replace("月", ""), out dt))
                    return dt;
            }
            else
            {
                if (DateTime.TryParse(str.Replace("年", "").Replace("月", ""), out dt))
                    return dt;
            }
            return null;
        }
    }
}