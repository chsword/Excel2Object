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
            if (string.IsNullOrWhiteSpace(path))
                return null;
            var bytes = File.ReadAllBytes(path);
            return ExcelToObject<TModel>(bytes);
        }

        public IEnumerable<TModel> ExcelToObject<TModel>(byte[] bytes) where TModel : class, new()
        {
            var sheets = ExcelTypeUtil.GetSheets(bytes).ToList();
            if(sheets.Count==0)
                throw new InvalidDataException("Currently there is no Sheet in Excel");

            var list = ExcelToExcelModel(sheets[0], Sheme.GetSheetSheme<TModel>());
            return TypeConvert.SheetModelToObject<TModel>(list);
        }

        //private static IEnumerable<TModel> ExcelToObject<TModel>(ISheet sheet) where TModel : class, new()
        //{ 
        //    var rows = sheet.GetRowEnumerator();
        //    rows.MoveNext();

        //    var dict = ExcelUtil.GetPropertiesAttributesDict<TModel>();
        //    var dictColumns = new Dictionary<int, KeyValuePair<PropertyInfo, ExcelTitleAttribute>>();
    
        //    var titleRow = (IRow) rows.Current;
        //    if (titleRow != null)
        //    {
        //        foreach (var cell in titleRow.Cells)
        //        {
        //            var prop = dict.FirstOrDefault(c => cell.StringCellValue == c.Value.Title);
        //            if (prop.Key != null && !dictColumns.ContainsKey(cell.ColumnIndex))
        //                dictColumns.Add(cell.ColumnIndex, prop);
        //        }
        //    }

        //    while (rows.MoveNext())
        //    {
        //        var row = (IRow) rows.Current;
        //        if (row?.Cells?.Count == 0)
        //            continue;

        //        var model = new TModel();

        //        foreach (var pair in dictColumns)
        //        {
        //            var propType = pair.Value.Key.PropertyType;
        //            var type = TypeUtil.GetUnNullableType(propType);
        //            if (type.IsEnum)
        //            {
        //                var specialValue = GetEnum(row, pair.Key, type);
        //                pair.Value.Key.SetValue(model, specialValue, null);
        //            }
        //            else
        //            {
        //                if (SpecialConvertDict.ContainsKey(type))
        //                {
        //                    var specialValue = SpecialConvertDict[type](row, pair.Key);
        //                    pair.Value.Key.SetValue(model, specialValue, null);
        //                }
        //                else
        //                {
        //                    var val = Convert.ChangeType(GetCellValue(row, pair.Key), propType);
        //                    pair.Value.Key.SetValue(model, val, null);
        //                }
        //            }
        //        }

        //        yield return model;
        //    }
        //}

        internal static SheetModel ExcelToExcelModel(ISheet sheet, SheetModel sheetModel = null)
        {
            var rows = sheet.GetRowEnumerator();
            rows.MoveNext();
            if (sheetModel == null)
            {
                sheetModel = SheetModel.Create(sheet.SheetName);
                var titleRow = (IRow) rows.Current;
                if (titleRow != null)
                {
                    for (var i = 0; i < titleRow.Cells.Count; i++)
                    {
                        var cell = titleRow.Cells[i];
                        sheetModel.Columns.Add(new ExcelColumn
                        {
                            Order = cell.ColumnIndex,
                            Title = cell.StringCellValue,
                            Type = null, //cell.CellType todo
                        });
                    }
                }
            }

            while (rows.MoveNext())
            {
                var row = (IRow) rows.Current;
                if (row == null || row?.Cells?.Count == 0)
                    continue;
                var line = new Dictionary<string, ExcelCell>();
                foreach (var column in sheetModel.Columns)
                {
                    var cell = row.GetCell(column.Order);
                    var propType = column.Type;
                    if (propType == null)
                    {
                        var specialValue = ExcelTypeUtil.GetCellValue(row, column.Order);
                        line[column.Title] = new ExcelCell(specialValue, cell.CellType);
                        continue;
                    }

                    var type = TypeUtil.GetUnNullableType(propType);
                    if (type.IsEnum)
                    {
                        var specialValue = ExcelTypeUtil.GetEnum(row, column.Order, type);
                        line[column.Title] = new ExcelCell(specialValue, cell.CellType);
                        continue;
                    }

                    if (ExcelTypeUtil.SpecialConvertDict.ContainsKey(type))
                    {
                        var specialValue = ExcelTypeUtil.SpecialConvertDict[type](row, column.Order);
                        line[column.Title] = new ExcelCell(specialValue, cell.CellType);
                        continue;
                    }

                    var val = Convert.ChangeType(ExcelTypeUtil.GetCellValue(row, column.Order), propType);
                    line[column.Title] = new ExcelCell(val,cell.CellType);

                }

                sheetModel.Rows.Add(line);

            }

            return sheetModel;
        }
    }
}