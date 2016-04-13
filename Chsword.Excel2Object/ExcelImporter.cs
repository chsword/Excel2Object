using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Chsword.Excel2Object.Internal;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

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
        IEnumerable<TModel> ExcelToObject<TModel>(IEnumerator result) where TModel : class, new()
        {
            var dict = ExcelUtil.GetExportAttrDict<TModel>();
            var dictColumns = new Dictionary<int, KeyValuePair<PropertyInfo, ExcelTitleAttribute>>();

            IEnumerator rows = result;

            var titleRow = (IRow)rows.Current;
            foreach (var cell in titleRow.Cells)
            {
                var prop = dict.FirstOrDefault(c => cell.StringCellValue == c.Value.Title);
                if (prop.Key != null && !dictColumns.ContainsKey(cell.ColumnIndex))
                {
                    dictColumns.Add(cell.ColumnIndex, prop);
                }
            }
            while (rows.MoveNext())
            {
                var row = (IRow)rows.Current;
                ICell firstCell = row.GetCell(0);
                if (firstCell == null || firstCell.CellType == CellType.Blank ||
                    string.IsNullOrWhiteSpace(firstCell.ToString()))
                    continue;

                var model = new TModel();

                foreach (var pair in dictColumns)
                {
                    var propType = pair.Value.Key.PropertyType;
                    if (propType == typeof(DateTime?) ||
                        propType == typeof(DateTime))
                    {
                        pair.Value.Key.SetValue(model, GetCellDateTime(row, pair.Key), null);
                    }
                    else if (propType == typeof(Boolean) || propType == typeof(bool?))
                    {
                        var cellValue = GetCellValue(row, pair.Key);
                        if (!String.IsNullOrEmpty(cellValue))
                        {
                            var value = false;
                            if (!bool.TryParse(cellValue, out value))
                            {
                                switch (cellValue.ToLower())
                                {
                                    case "1":
                                    case "是":
                                    case "yes":
                                    case "true":
                                        value = true;
                                        break;
                                    case "0":
                                    case "否":
                                    case "no":
                                    case "false":
                                        value = false;
                                        break;
                                    default:
                                        value = Convert.ToBoolean(cellValue);
                                        break;
                                }
                            }
                            pair.Value.Key.SetValue(model, value, null);
                        }
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

        string GetCellValue(IRow row, int index)
        {
            var result = string.Empty;
            try
            {
                switch (row.GetCell(index).CellType)
                {
                    case CellType.Numeric:
                        result = row.GetCell(index).NumericCellValue.ToString();
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
        IEnumerator GetDataRows(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return null;
            IWorkbook workbook = null;
           
            try
            {
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    workbook = WorkbookFactory.Create(file);
                }
            }
            catch
            {
                return null;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            IEnumerator rows = sheet.GetRowEnumerator();
            rows.MoveNext();
            return rows;
        }
        IEnumerator GetDataRows(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                return null;
            IWorkbook workbook;
            try
            {
                using (MemoryStream memoryStream = new MemoryStream(bytes))
                {
                    workbook = WorkbookFactory.Create(memoryStream);
                }
            }
            catch
            {
                return null;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            IEnumerator rows = sheet.GetRowEnumerator();
            rows.MoveNext();
            return rows;
        }
        DateTime? GetCellDateTime(IRow row, int index)
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
                            {
                                result = dt;
                            }
                        }
                        else if (str.EndsWith("月"))
                        {
                            DateTime dt;
                            if (DateTime.TryParse((str + "-01").Replace("年", "").Replace("月", ""), out dt))
                            {
                                result = dt;
                            }
                        }
                        else if (!str.Contains("年") && !str.Contains("月") && !str.Contains("日"))
                        {

                            DateTime dt;
                            if (DateTime.TryParse(str, out dt))
                            {
                                result = dt;
                            }
                            else if (DateTime.TryParse((str + "-01-01").Replace("年", "").Replace("月", ""), out dt))
                            {
                                result = dt;
                            }
                            else
                            {
                                result = null;
                            }

                        }
                        else
                        {
                            DateTime dt;
                            if (DateTime.TryParse(str.Replace("年", "").Replace("月", ""), out dt))
                            {
                                result = dt;
                            }
                        }
                        break;
                    case CellType.Blank:
                        break;
                    default:
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