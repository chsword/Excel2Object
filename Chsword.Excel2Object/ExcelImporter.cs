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
                [typeof(bool)] = GetCellBoolean
            };

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
                    var unNullableType = TypeUtil.GetUnNullableType(propType);
                    if (SpecialConvertDict.ContainsKey(unNullableType))
                    {
                        var specialValue = SpecialConvertDict[unNullableType](row, pair.Key);
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

                    #region

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

                    #endregion

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
                        if (str.EndsWith("年"))
                        {
                            DateTime dt;
                            if (DateTime.TryParse((str + "-01-01").Replace("年", ""), out dt))
                                result = dt;
                        }
                        else if (str.EndsWith("月"))
                        {
                            DateTime dt;
                            if (DateTime.TryParse((str + "-01").Replace("年", "").Replace("月", ""), out dt))
                                result = dt;
                        }
                        else if (!str.Contains("年") && !str.Contains("月") && !str.Contains("日"))
                        {
                            DateTime dt;
                            if (DateTime.TryParse(str, out dt))
                                result = dt;
                            else if (DateTime.TryParse((str + "-01-01").Replace("年", "").Replace("月", ""), out dt))
                                result = dt;
                        }
                        else
                        {
                            DateTime dt;
                            if (DateTime.TryParse(str.Replace("年", "").Replace("月", ""), out dt))
                                result = dt;
                        }
                        break;
                    case CellType.Blank:
                        break;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return result;
        }
    }
}