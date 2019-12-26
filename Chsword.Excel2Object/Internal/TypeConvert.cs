using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Internal
{
    internal class TypeConvert
    {
        internal static ExcelModel ConvertDataSetToExcelModel(DataTable dt, string sheetTitle)
        {
            var excel = new ExcelModel {Sheets = new List<SheetModel>()};
            var sheet = SheetModel.Create(sheetTitle);
            excel.Sheets.Add(sheet);
            var dataSetColumnArray = dt.Columns.Cast<DataColumn>().ToArray();

            var columns = dataSetColumnArray.Select((item, i) =>
                new ExcelColumn
                {
                    Order = i,
                    Title = item.ColumnName,
                    Type = item.DataType
                }).ToList();
            sheet.Columns = columns;

            var data = dt.Rows.Cast<DataRow>().ToArray();
            foreach (var item in data.Where(c => c != null))
            {
                var row = new Dictionary<string, ExcelCell>();
                foreach (var column in dataSetColumnArray)
                {
                    
                    row[column.ColumnName] = new ExcelCell(item[column.ColumnName]);
                }

                sheet.Rows.Add(row);
            }

            return excel;
        }

        public static ExcelModel ConvertObjectToExcelModel<TModel>(IEnumerable<TModel> data, string sheetTitle)
        {
            var excel = new ExcelModel {Sheets = new List<SheetModel>()};


            if (string.IsNullOrWhiteSpace(sheetTitle))
            {
                var classAttr = ExcelUtil.GetClassExportAttribute<TModel>();
                sheetTitle = classAttr == null ? sheetTitle : classAttr.Title;
            }

            var sheet = SheetModel.Create(sheetTitle);

            excel.Sheets.Add(sheet);
            var attrDict = ExcelUtil.GetPropertiesAttributesDict<TModel>();
            var objKeysArray = attrDict.OrderBy(c => c.Value.Order).ToArray();


            for (var i = 0; i < objKeysArray.Length; i++)
            {
                var column = new ExcelColumn
                {
                    Title = objKeysArray[i].Value.Title, Type = objKeysArray[i].Key.PropertyType, Order = i
                };
                sheet.Columns.Add(column);
            }

            foreach (var item in data.Where(c => c != null))
            {
                var row = new Dictionary<string, ExcelCell>();
                foreach (var column in objKeysArray)
                {
                    var prop = column.Key;
                    row[column.Value.Title] = new ExcelCell(prop.GetValue(item, null)) ;
                }

                sheet.Rows.Add(row);
            }

            return excel;
        }

        public static List<TModel> SheetModelToObject<TModel>(SheetModel sheet) where TModel : new()
        {

            var dict = ExcelUtil.GetPropertiesAttributesDict<TModel>();
            var dictColumns = new Dictionary<int, KeyValuePair<PropertyInfo, ExcelTitleAttribute>>();

            var columns = sheet.Columns.OrderBy(c => c.Order).ToList();
            for (var i = 0; i < columns.Count; i++)
            {
                var cell = sheet.Columns[i];
                var prop = dict.FirstOrDefault(c => cell.Title == c.Value.Title);
                if (prop.Key != null && !dictColumns.ContainsKey(i))
                    dictColumns.Add(i, prop);
            }

            var list = new List<TModel>();

            foreach (var row in sheet.Rows)
            {
                if (row.Count == 0)
                    continue;
                var model = new TModel();

                foreach (var pair in dictColumns)
                {
                    var propType = pair.Value.Key.PropertyType;
                    var title = pair.Value.Value.Title;
                    var type = TypeUtil.GetUnNullableType(propType);
                    if (type.IsEnum)
                    {
                        var specialValue = ExcelTypeUtil.GetEnum(row, title, type);
                        pair.Value.Key.SetValue(model, specialValue, null);
                        continue;

                    }

                    if (ExcelTypeUtil.SpecialConvertDict.ContainsKey(type))
                    {
                        var specialValue = ExcelTypeUtil.SpecialConvertDict[type](row, pair.Key);
                        pair.Value.Key.SetValue(model, specialValue, null);
                        continue;
                    }

                    var val = Convert.ChangeType(ExcelTypeUtil.GetCellValue(row, pair.Key), propType);
                    pair.Value.Key.SetValue(model, val, null);
                    continue;
                }

                list.Add(model);
            }
            return list;
        }
    }
}