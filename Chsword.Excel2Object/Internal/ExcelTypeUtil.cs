using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Internal
{
    internal static class ExcelTypeUtil
    {
        public static object GetEnum(Dictionary<string, ExcelCell> row, string title, Type enumType)
        {
            var val = row[title].Value;
            if (val == null) return null;
            if (Enum.GetNames(enumType).Contains(val))
            {
                return Enum.Parse(enumType, val.ToString());
            }
            return Enum.Parse(enumType, "0");
        }

        public static object GetEnum(IRow row, int key, Type enumType)
        {
            var cellValue = GetCellValue(row, key);
            if (String.IsNullOrEmpty(cellValue)) return null;
            if (Enum.GetNames(enumType).Contains(cellValue))
            {
                return Enum.Parse(enumType, cellValue);
            }

            return Enum.Parse(enumType, "0");
        }

        public static readonly Dictionary<Type, Func<IRow, int, object>> SpecialConvertDict =
            new Dictionary<Type, Func<IRow, int, object>>
            {
                [typeof(DateTime)] = GetCellDateTime,
                [typeof(bool)] = GetCellBoolean,
                [typeof(Uri)] = GetCellUri,

            };

        private static object GetCellUri(ICell cell)
        {
            var cellValue = GetCellValue(cell);
            if (String.IsNullOrEmpty(cellValue)) return null;
            return new Uri(cellValue);
        }

        private static object GetCellBoolean(ICell cell)
        {
            var cellValue = GetCellValue(cell);
            if (String.IsNullOrEmpty(cellValue)) return null;
            if (Boolean.TryParse(cellValue, out var value)) return value;
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
                    return Convert.ToBoolean((string) cellValue);
            }
        }

        public static string GetCellValue(ICell cell)
        {
            var result = String.Empty;
            try
            {
                switch (cell.CellType)
                {
                    case CellType.Numeric:
                        result = cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                        break;
                    case CellType.String:
                        result = cell.StringCellValue;
                        break;
                    case CellType.Blank:
                        result = String.Empty;
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
                        result = cell.ToString();
                        break;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            return (result ?? "").Trim();
        }

        public static IEnumerable<ISheet> GetSheets(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                throw new InvalidDataException("bad excel file");
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
                throw new InvalidDataException("bad excel file");
            }

            for (var i = 0; i < workbook.NumberOfSheets; i++)
            {
                yield return workbook.GetSheetAt(i);
            }
            //var sheet = workbook.GetSheetAt(0);
            //var rows = sheet.GetRowEnumerator();
            //rows.MoveNext();
            //return rows;
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